"""
app.py — Customer Lead Dashboard
==================================
UNI-T North America | Signal-driven B2B prospecting dashboard.

Tab build order (deploy and verify each before starting the next):
  1. Upload        ← this session
  2. New Contacts
  3. Watch List
  4. Signal Feed   (Phase 3)
  5. Outreach Generator (Phase 3)
  6. HAM           (Phase 3)
  7. VV Review Queue (Phase 3)

Do not put credentials in this file. All secrets live in Streamlit Cloud
Settings → Secrets (TOML format). See CLD_HANDOFF_V2.md for schema.
"""

import io
import math
import time
import datetime
import streamlit as st

try:
    import gspread
    GSPREAD_OK = True
except ImportError:
    GSPREAD_OK = False

# Optional heavy imports — graceful fallback if not yet installed
try:
    import pandas as pd
    PD_OK = True
except ImportError:
    PD_OK = False

try:
    import openpyxl  # required for pd.read_excel engine="openpyxl"
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

from utils import (
    get_secret, gsheet_open, gsheet_ensure_tabs,
    sheet_get_all, sheet_append_rows, sheet_update_row, sheet_get_domains,
    get_anthropic_client,
    SHEET_MASTER, SHEET_NEW_CONTACTS, SHEET_ORDERS,
    SHEET_SIGNALS, SHEET_VV, SHEET_OUTREACH, SHEET_ERRORS,
    MASTER_HEADERS, NEW_CONTACTS_HEADERS, ORDER_HEADERS,
    DEFAULT_MONITOR, OUTREACH_ENABLED,
    CLASS_LABEL, CLASS_COLOR, MANUAL_REVIEW_DOMAINS,
    INDUSTRY_LABELS, ICP_LABELS,
    now_str, today_str, extract_domain, format_currency, tier_badge,
)

# SUPPRESS_OUTREACH derived from OUTREACH_ENABLED (not a separate export in utils)
SUPPRESS_OUTREACH = {cls: not enabled for cls, enabled in OUTREACH_ENABLED.items()}

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="UNI-T Customer Lead Dashboard",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────────────────────────────────
# PERSONAL EMAIL DOMAINS — strip entirely during upload
# ─────────────────────────────────────────────────────────────────────────────

PERSONAL_EMAIL_DOMAINS = {
    "gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "aol.com",
    "icloud.com", "me.com", "mac.com", "live.com", "msn.com",
    "comcast.net", "att.net", "cox.net", "verizon.net", "earthlink.net",
    "sbcglobal.net", "bellsouth.net",
}

OWN_DOMAINS = {
    "uni-trendus.com",
    "instruments.uni-trend.com",
    "uni-trend.com",
}

# ─────────────────────────────────────────────────────────────────────────────
# UPLOAD CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

BATCH_SIZE  = 200   # rows per Sheets API append call
BATCH_DELAY = 1.0   # seconds between append batches


# ─────────────────────────────────────────────────────────────────────────────
# NaN SANITIZER
# ─────────────────────────────────────────────────────────────────────────────

def _sanitize_row(row: list) -> list:
    """
    Replace float nan / inf with empty string.
    The Sheets API rejects JSON-non-compliant floats; pandas emits them for
    empty numeric cells.
    """
    result = []
    for v in row:
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            result.append("")
        else:
            result.append(v)
    return result


# ─────────────────────────────────────────────────────────────────────────────
# STARTUP — SHEETS CONNECTION
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_resource(ttl=300)
def get_sheet():
    """Open the CLD Google Sheet. Cached for 5 minutes."""
    return gsheet_open()


def startup():
    """Run once per session: open Sheet and ensure all tabs exist."""
    sh = get_sheet()
    if sh:
        gsheet_ensure_tabs(sh)
    return sh


# ─────────────────────────────────────────────────────────────────────────────
# FILE-TYPE DETECTION
# ─────────────────────────────────────────────────────────────────────────────

def detect_file_type(df: "pd.DataFrame") -> str | None:
    """
    Identify upload type from column headers.

    Returns one of:
      "brevo"            — Brevo CSV export
      "shopify_contacts" — Shopify contacts XLSX
      "shopify_orders"   — Shopify orders CSV
      None               — unrecognised

    Detection logic:
      - Brevo:            requires EMAIL, FIRSTNAME, LASTNAME, COMPANY (all caps)
      - Shopify contacts: requires Email, First Name, Last Name (mixed case)
                          AND must NOT have all-caps EMAIL (guards against
                          Brevo variants that include mixed-case aliases)
      - Shopify orders:   requires Name, Financial Status, Lineitem name,
                          Lineitem price
    """
    cols = set(df.columns)

    brevo_required           = {"EMAIL", "FIRSTNAME", "LASTNAME", "COMPANY"}
    shopify_contact_required = {"Email", "First Name", "Last Name"}
    shopify_orders_required  = {"Name", "Financial Status", "Lineitem name", "Lineitem price"}

    if brevo_required.issubset(cols):
        return "brevo"
    if shopify_contact_required.issubset(cols) and "EMAIL" not in cols:
        return "shopify_contacts"
    if shopify_orders_required.issubset(cols):
        return "shopify_orders"

    return None


# ─────────────────────────────────────────────────────────────────────────────
# FILE READERS
# ─────────────────────────────────────────────────────────────────────────────

def read_uploaded_file(f) -> "pd.DataFrame | None":
    """
    Read an uploaded file into a DataFrame.
    Supports: .csv, .xlsx
    Returns None on failure with a user-visible warning.
    """
    if not PD_OK:
        st.error("pandas is not installed. Check requirements.txt.")
        return None

    name = f.name.lower()
    try:
        if name.endswith(".xlsx"):
            if not OPENPYXL_OK:
                st.error("openpyxl is not installed. Check requirements.txt.")
                return None
            return pd.read_excel(io.BytesIO(f.read()), engine="openpyxl")
        else:
            # Default: CSV — try UTF-8 first, fall back to latin-1
            raw = f.read()
            try:
                return pd.read_csv(io.BytesIO(raw), dtype=str)
            except UnicodeDecodeError:
                return pd.read_csv(io.BytesIO(raw), dtype=str, encoding="latin-1")
    except Exception as ex:
        st.warning(f"Could not read {f.name}: {ex}")
        return None


# ─────────────────────────────────────────────────────────────────────────────
# DOMAIN CLASSIFICATION HELPERS
# ─────────────────────────────────────────────────────────────────────────────

EDUCATION_TLDS = {".edu", ".ac.uk", ".edu.au", ".edu.ca"}
GOVERNMENT_TLDS = {".gov", ".mil", ".gov.uk", ".gc.ca"}

DEFENSE_DOMAINS = {
    "lmco.com", "l3harris.com", "harris.com", "boeing.com",
    "northropgrumman.com", "ngc.com", "raytheon.com", "rtx.com",
    "baesystems.com", "textron.com",
}

HAM_SIGNALS = {"arrl", "qrz", "hamradio", "ham"}  # substring hints — enrichment corrects


def classify_domain(domain: str) -> str:
    """
    Derive domain_class from the domain string alone.
    Returns: commercial / education / government / defense / ham.
    Never returns 'distributor' — that requires master list context.
    Distributor domains uploaded via new exports land as 'commercial' and
    should be corrected during New Contacts review.
    """
    d = domain.lower().strip()

    if d in DEFENSE_DOMAINS:
        return "defense"

    for tld in EDUCATION_TLDS:
        if d.endswith(tld):
            return "education"

    for tld in GOVERNMENT_TLDS:
        if d.endswith(tld):
            return "government"

    for hint in HAM_SIGNALS:
        if hint in d:
            return "ham"

    return "commercial"


# ─────────────────────────────────────────────────────────────────────────────
# BREVO PROCESSOR
# ─────────────────────────────────────────────────────────────────────────────

def process_brevo(df: "pd.DataFrame", existing_domains: set) -> dict:
    """
    Parse a Brevo contacts CSV export.

    Returns dict with keys:
      new_master_rows, new_contact_rows, update_map, filtered, filter_reasons
    """
    from collections import Counter

    new_master   = []
    new_contacts = []
    update_map   = {}
    filtered     = 0
    reasons      = Counter()
    seen         = set()

    for _, row in df.iterrows():
        email   = str(row.get("EMAIL", "")).strip().lower()
        company = str(row.get("COMPANY", "")).strip()

        if not email or "@" not in email:
            filtered += 1
            reasons["no email"] += 1
            continue

        domain = extract_domain(email)

        if not domain:
            filtered += 1
            reasons["no domain"] += 1
            continue

        if domain in PERSONAL_EMAIL_DOMAINS:
            filtered += 1
            reasons["personal email"] += 1
            continue

        if domain in OWN_DOMAINS:
            filtered += 1
            reasons["own domain"] += 1
            continue

        if domain in seen:
            continue  # keep first occurrence per domain within this file
        seen.add(domain)

        first = str(row.get("FIRSTNAME", "")).strip()
        last  = str(row.get("LASTNAME", "")).strip()
        name  = f"{first} {last}".strip()
        tags  = str(row.get("TAGS", "") or row.get("Tags", "")).strip()

        if domain in existing_domains:
            update_map[domain] = {
                "brevo_contacts": 1,
                "last_activity":  now_str(),
                "in_brevo":       "True",
            }
            if tags:
                update_map[domain]["tags"] = tags
        else:
            cls             = classify_domain(domain)
            monitor_default = DEFAULT_MONITOR.get(cls, True)
            suppress        = SUPPRESS_OUTREACH.get(cls, False)

            master_row = _blank_master_row()
            master_row.update({
                "domain":             domain,
                "company_name":       company,
                "domain_class":       cls,
                "customer_status":    "prospect",
                "watch_tier":         2,
                "score":              1,
                "score_domain":       1,
                "score_contact":      1,
                "score_engagement":   1,
                "monitor":            str(monitor_default),
                "enrich":             "False",
                "suppress_outreach":  str(suppress),
                "brevo_contacts":     1,
                "shopify_contacts":   0,
                "total_spent":        0,
                "total_orders":       0,
                "best_contact_name":  name,
                "best_contact_email": email,
                "tags":               tags,
                "has_event_tag":      "True" if tags else "False",
                "in_brevo":           "True",
                "in_shopify":         "False",
                "has_purchases":      "False",
                "first_seen":         today_str(),
                "last_activity":      now_str(),
                "enriched":           "False",
                "outreach_status":    "none",
            })

            if domain in MANUAL_REVIEW_DOMAINS:
                master_row["notes"] = MANUAL_REVIEW_DOMAINS[domain]

            new_master.append(master_row)
            new_contacts.append(_new_contact_row(master_row))
            existing_domains.add(domain)

    return {
        "new_master_rows":  new_master,
        "new_contact_rows": new_contacts,
        "update_map":       update_map,
        "filtered":         filtered,
        "filter_reasons":   reasons,
    }


# ─────────────────────────────────────────────────────────────────────────────
# SHOPIFY CONTACTS PROCESSOR
# ─────────────────────────────────────────────────────────────────────────────

def process_shopify_contacts(df: "pd.DataFrame", existing_domains: set) -> dict:
    """
    Parse a Shopify contacts XLSX export.

    Column names: Email, First Name, Last Name, Tags, ...
    Company is in 'Company' when present or 'Address Company' as fallback.
    Same filtering rules as Brevo, different column names.
    """
    from collections import Counter

    new_master   = []
    new_contacts = []
    update_map   = {}
    filtered     = 0
    reasons      = Counter()
    seen         = set()

    for _, row in df.iterrows():
        email = str(row.get("Email", "")).strip().lower()

        # Company: top-level preferred, address fallback
        company = str(row.get("Company", "") or row.get("Address Company", "")).strip()

        if not email or "@" not in email:
            filtered += 1
            reasons["no email"] += 1
            continue

        domain = extract_domain(email)

        if not domain:
            filtered += 1
            reasons["no domain"] += 1
            continue

        if domain in PERSONAL_EMAIL_DOMAINS:
            filtered += 1
            reasons["personal email"] += 1
            continue

        if domain in OWN_DOMAINS:
            filtered += 1
            reasons["own domain"] += 1
            continue

        if domain in seen:
            continue
        seen.add(domain)

        first = str(row.get("First Name", "")).strip()
        last  = str(row.get("Last Name", "")).strip()
        name  = f"{first} {last}".strip()
        tags  = str(row.get("Tags", "")).strip()

        if domain in existing_domains:
            update_map[domain] = {
                "shopify_contacts": 1,
                "last_activity":    now_str(),
                "in_shopify":       "True",
            }
            if tags:
                update_map[domain]["tags"] = tags
        else:
            cls             = classify_domain(domain)
            monitor_default = DEFAULT_MONITOR.get(cls, True)
            suppress        = SUPPRESS_OUTREACH.get(cls, False)

            master_row = _blank_master_row()
            master_row.update({
                "domain":             domain,
                "company_name":       company,
                "domain_class":       cls,
                "customer_status":    "prospect",
                "watch_tier":         2,
                "score":              1,
                "score_domain":       1,
                "score_contact":      1,
                "score_engagement":   1,
                "monitor":            str(monitor_default),
                "enrich":             "False",
                "suppress_outreach":  str(suppress),
                "brevo_contacts":     0,
                "shopify_contacts":   1,
                "total_spent":        0,
                "total_orders":       0,
                "best_contact_name":  name,
                "best_contact_email": email,
                "tags":               tags,
                "has_event_tag":      "True" if tags else "False",
                "in_brevo":           "False",
                "in_shopify":         "True",
                "has_purchases":      "False",
                "first_seen":         today_str(),
                "last_activity":      now_str(),
                "enriched":           "False",
                "outreach_status":    "none",
            })

            if domain in MANUAL_REVIEW_DOMAINS:
                master_row["notes"] = MANUAL_REVIEW_DOMAINS[domain]

            new_master.append(master_row)
            new_contacts.append(_new_contact_row(master_row))
            existing_domains.add(domain)

    return {
        "new_master_rows":  new_master,
        "new_contact_rows": new_contacts,
        "update_map":       update_map,
        "filtered":         filtered,
        "filter_reasons":   reasons,
    }


# ─────────────────────────────────────────────────────────────────────────────
# SHOPIFY ORDERS PROCESSOR
# ─────────────────────────────────────────────────────────────────────────────

def process_shopify_orders(df: "pd.DataFrame", existing_domains: set) -> dict:
    """
    Parse a Shopify orders CSV export.

    Shopify orders use a multi-row format: Email / Total / Financial Status /
    Tags appear only on the FIRST row of each order. Continuation rows carry
    additional line items for the same order and have a blank Email field.
    State is carried forward via current_* variables.

    Skips:
      - Orders where Tags contains "amazon" (case-insensitive)
      - Orders where Financial Status == "refunded"
      - Personal / own email domains
    """
    from collections import Counter

    order_rows    = []
    domain_totals = {}
    filtered      = 0
    reasons       = Counter()

    current_email   = ""
    current_total   = 0.0
    current_status  = ""
    current_tags    = ""
    current_order   = ""
    current_date    = ""
    current_company = ""

    for _, row in df.iterrows():

        email_raw = str(row.get("Email", "")).strip()

        if email_raw and "@" in email_raw:
            # ── first row of a new order ──
            current_email   = email_raw.lower()
            current_order   = str(row.get("Name", "")).strip()
            current_date    = str(
                row.get("Created at", "") or row.get("Paid at", "")
            ).strip()
            current_company = str(row.get("Billing Company", "")).strip()

            try:
                current_total = float(
                    str(row.get("Total", 0) or 0).replace(",", "")
                )
            except ValueError:
                current_total = 0.0

            current_status = str(row.get("Financial Status", "")).strip().lower()
            current_tags   = str(row.get("Tags", "")).strip()

            # order-level filters
            if current_status == "refunded":
                filtered += 1
                reasons["refunded order"] += 1
                current_email = ""
                continue

            if "amazon" in current_tags.lower():
                filtered += 1
                reasons["amazon order"] += 1
                current_email = ""
                continue

            domain = extract_domain(current_email)
            if not domain:
                current_email = ""
                continue

            if domain in PERSONAL_EMAIL_DOMAINS:
                filtered += 1
                reasons["personal email"] += 1
                current_email = ""
                continue

            if domain in OWN_DOMAINS:
                filtered += 1
                reasons["own domain"] += 1
                current_email = ""
                continue

        # ── line item row (first or continuation) ──
        if not current_email:
            continue

        domain = extract_domain(current_email)
        if not domain:
            continue

        product   = str(row.get("Lineitem name", "")).strip()
        sku       = str(row.get("Lineitem sku", "")).strip()
        qty_raw   = str(row.get("Lineitem quantity", 1) or 1)
        price_raw = str(row.get("Lineitem price", 0) or 0)

        try:
            qty = int(float(qty_raw))
        except ValueError:
            qty = 1
        try:
            price = float(price_raw.replace(",", ""))
        except ValueError:
            price = 0.0

        fulfil = str(row.get("Fulfillment Status", "")).strip()

        order_rows.append([
            domain, current_order, current_date,
            product, sku, price, qty, current_total,
            fulfil, current_status, current_company, current_email,
        ])

        if domain not in domain_totals:
            domain_totals[domain] = {
                "total_spent":    0.0,
                "total_orders":   0,
                "last_date":      current_date,
                "email":          current_email,
                "company":        current_company,
                "counted_orders": set(),
            }
        stats = domain_totals[domain]
        if current_order not in stats["counted_orders"]:
            stats["total_orders"] += 1
            stats["total_spent"]  += current_total
            stats["counted_orders"].add(current_order)
            if current_date > stats["last_date"]:
                stats["last_date"] = current_date

    return {
        "order_rows":     order_rows,
        "domain_totals":  domain_totals,
        "filtered":       filtered,
        "filter_reasons": reasons,
    }


# ─────────────────────────────────────────────────────────────────────────────
# ROW BUILDERS
# ─────────────────────────────────────────────────────────────────────────────

def _blank_master_row() -> dict:
    """Return a dict with all MASTER_HEADERS keys initialised to empty string."""
    return {h: "" for h in MASTER_HEADERS}


def _master_row_to_list(row: dict) -> list:
    """Convert a master_row dict to an ordered list matching MASTER_HEADERS."""
    return [row.get(h, "") for h in MASTER_HEADERS]


def _new_contact_row(master_row: dict) -> list:
    """Build a new_contacts list row from a master_row dict."""
    return [
        master_row.get("domain", ""),
        master_row.get("company_name", ""),
        master_row.get("domain_class", ""),
        master_row.get("customer_status", ""),
        master_row.get("total_spent", 0),
        master_row.get("total_orders", 0),
        master_row.get("brevo_contacts", 0),
        master_row.get("tags", ""),
        master_row.get("has_event_tag", "False"),
        master_row.get("best_contact_name", ""),
        master_row.get("best_contact_email", ""),
        master_row.get("best_contact_title", ""),
        master_row.get("first_seen", today_str()),
        master_row.get("monitor", "True"),
        master_row.get("enrich", "False"),
        today_str(),  # added_date
        "False",      # reviewed
    ]


# ─────────────────────────────────────────────────────────────────────────────
# BATCH WRITER
# ─────────────────────────────────────────────────────────────────────────────

def batch_write(sh, tab_name: str, rows: list[list],
                progress_label: str = "") -> bool:
    """
    Append rows to a Sheet tab in chunks of BATCH_SIZE, with a short delay
    between chunks to avoid Sheets API write-quota errors.
    Sanitizes each row to remove nan / inf before sending.
    Returns True if all batches succeeded.
    """
    if not rows:
        return True

    total   = len(rows)
    success = True

    for start in range(0, total, BATCH_SIZE):
        batch = [_sanitize_row(r) for r in rows[start : start + BATCH_SIZE]]
        ok = sheet_append_rows(sh, tab_name, batch)
        if not ok:
            success = False
        if start + BATCH_SIZE < total:
            time.sleep(BATCH_DELAY)

    return success


# ─────────────────────────────────────────────────────────────────────────────
# UPLOAD SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

def _render_upload_summary(results: list[dict]) -> None:
    """Render metric cards and filter breakdown after all files are processed."""
    total_new      = sum(r.get("new_added", 0)        for r in results)
    total_updated  = sum(r.get("existing_updated", 0) for r in results)
    total_orders   = sum(r.get("orders_written", 0)   for r in results)
    total_filtered = sum(r.get("filtered", 0)         for r in results)

    from collections import Counter
    merged: Counter = Counter()
    for r in results:
        merged += r.get("filter_reasons", Counter())

    st.success("✅ Upload complete")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("New companies",      total_new)
    c2.metric("Existing updated",   total_updated)
    c3.metric("Orders processed",   total_orders)
    c4.metric("Filtered / skipped", total_filtered)

    if merged:
        with st.expander("Filter breakdown"):
            for reason, count in sorted(merged.items(), key=lambda x: -x[1]):
                st.write(f"• **{reason}** — {count:,}")


# ─────────────────────────────────────────────────────────────────────────────
# TAB 1 — UPLOAD
# ─────────────────────────────────────────────────────────────────────────────

def tab_upload(sh) -> None:
    """Upload tab: accept one or more files, auto-detect type, process all."""

    st.header("Upload Contact & Order Data")
    st.caption(
        "Accepts Brevo CSV exports, Shopify contact XLSX exports, and "
        "Shopify order CSV exports. File type is detected automatically from "
        "column headers — filenames don't matter."
    )

    if not PD_OK:
        st.error(
            "**pandas** is not installed. "
            "Add `pandas` and `openpyxl` to requirements.txt and redeploy."
        )
        return

    uploaded_files = st.file_uploader(
        "Drop files here (CSV or XLSX)",
        type=["csv", "xlsx"],
        accept_multiple_files=True,
        key="upload_files",
    )

    if not uploaded_files:
        st.info(
            "No files selected yet. Drag one or more files above to begin. "
            "You can upload all three file types in a single session."
        )
        return

    st.write(f"**{len(uploaded_files)} file(s) ready to process:**")
    for f in uploaded_files:
        st.write(f"  • `{f.name}` ({f.size:,} bytes)")

    if not st.button("🚀 Process uploads", type="primary"):
        return

    existing_domains = sheet_get_domains(sh)
    session_results  = []
    progress         = st.progress(0.0, text="Starting…")

    for file_idx, f in enumerate(uploaded_files):
        pct_base = file_idx / len(uploaded_files)
        progress.progress(pct_base, text=f"Reading {f.name}…")

        df = read_uploaded_file(f)
        if df is None:
            session_results.append({
                "file": f.name, "error": "Could not read file",
                "new_added": 0, "existing_updated": 0,
                "orders_written": 0, "filtered": 0, "filter_reasons": {},
            })
            continue

        file_type = detect_file_type(df)

        if file_type is None:
            st.warning(
                f"⚠️ `{f.name}` — unrecognised column layout. "
                "Expected Brevo CSV, Shopify contacts XLSX, or Shopify orders CSV. "
                "Skipping."
            )
            session_results.append({
                "file": f.name, "error": "unrecognised file type",
                "new_added": 0, "existing_updated": 0,
                "orders_written": 0, "filtered": 0, "filter_reasons": {},
            })
            continue

        type_labels = {
            "brevo":            "Brevo contacts",
            "shopify_contacts": "Shopify contacts",
            "shopify_orders":   "Shopify orders",
        }
        progress.progress(
            pct_base + 0.1 / len(uploaded_files),
            text=f"Processing {type_labels[file_type]}: {f.name}…",
        )

        if file_type == "brevo":
            result      = process_brevo(df, existing_domains)
            file_result = _write_contact_results(
                sh, result, f.name, progress, pct_base, len(uploaded_files)
            )
        elif file_type == "shopify_contacts":
            result      = process_shopify_contacts(df, existing_domains)
            file_result = _write_contact_results(
                sh, result, f.name, progress, pct_base, len(uploaded_files)
            )
        else:  # shopify_orders
            result      = process_shopify_orders(df, existing_domains)
            file_result = _write_order_results(
                sh, result, f.name, progress, pct_base, len(uploaded_files)
            )

        session_results.append(file_result)

    progress.progress(1.0, text="Done.")
    _render_upload_summary(session_results)


def _write_contact_results(sh, result: dict, filename: str,
                            progress, pct_base: float, n_files: int) -> dict:
    """
    Write new master + new_contacts rows, and silently update existing domains,
    for a contact file (Brevo or Shopify contacts).
    """
    new_master   = result["new_master_rows"]
    new_contacts = result["new_contact_rows"]
    update_map   = result["update_map"]

    master_lists  = [_master_row_to_list(r) for r in new_master]
    contact_lists = new_contacts  # already lists

    new_added = 0
    if master_lists:
        progress.progress(
            pct_base + 0.4 / n_files,
            text=f"Writing {len(master_lists):,} new companies…",
        )
        if batch_write(sh, SHEET_MASTER, master_lists):
            new_added = len(master_lists)

    if contact_lists:
        progress.progress(
            pct_base + 0.6 / n_files,
            text=f"Writing {len(contact_lists):,} new_contacts rows…",
        )
        batch_write(sh, SHEET_NEW_CONTACTS, contact_lists)

    existing_updated = 0
    if update_map:
        progress.progress(
            pct_base + 0.8 / n_files,
            text=f"Updating {len(update_map):,} existing companies…",
        )
        for domain, updates in update_map.items():
            sheet_update_row(sh, SHEET_MASTER, "domain", domain, updates)
            existing_updated += 1
            time.sleep(0.12)  # ~8 writes/sec — well within 60/min quota

    return {
        "file":             filename,
        "new_added":        new_added,
        "existing_updated": existing_updated,
        "orders_written":   0,
        "filtered":         result["filtered"],
        "filter_reasons":   result["filter_reasons"],
    }


def _write_order_results(sh, result: dict, filename: str,
                          progress, pct_base: float, n_files: int) -> dict:
    """
    Write order_history rows and update master_companies spend data.

    Spend updates use a single values_batch_update call (one API request for
    all domains) instead of one call per domain, avoiding the 429 quota error.
    """
    order_rows    = result["order_rows"]
    domain_totals = result["domain_totals"]

    orders_written = 0
    if order_rows:
        progress.progress(
            pct_base + 0.4 / n_files,
            text=f"Writing {len(order_rows):,} order line items…",
        )
        if batch_write(sh, SHEET_ORDERS, order_rows):
            orders_written = len(order_rows)

    if domain_totals and GSPREAD_OK:
        progress.progress(
            pct_base + 0.6 / n_files,
            text=f"Reading master list for spend update…",
        )
        try:
            ws       = sh.worksheet(SHEET_MASTER)
            all_vals = ws.get_all_values()  # single read — entire sheet

            if all_vals:
                headers = all_vals[0]
                col     = {h: i for i, h in enumerate(headers)}
                dom_col = col.get("domain", 0)
                dom_idx = {
                    all_vals[i][dom_col]: i
                    for i in range(1, len(all_vals))
                }

                updates_needed = []  # (row_1based, col_1based, value)

                for domain, stats in domain_totals.items():
                    if domain not in dom_idx:
                        continue
                    r = dom_idx[domain]

                    for field, value in [
                        ("total_spent",   round(stats["total_spent"], 2)),
                        ("total_orders",  stats["total_orders"]),
                        ("has_purchases", "True"),
                        ("in_shopify",    "True"),
                        ("last_activity", now_str()),
                    ]:
                        if field in col:
                            updates_needed.append((r + 1, col[field] + 1, value))

                    if stats["email"] and "best_contact_email" in col:
                        updates_needed.append(
                            (r + 1, col["best_contact_email"] + 1, stats["email"])
                        )
                    if stats["company"] and "company_name" in col:
                        updates_needed.append(
                            (r + 1, col["company_name"] + 1, stats["company"])
                        )

                if updates_needed:
                    progress.progress(
                        pct_base + 0.85 / n_files,
                        text=f"Writing spend updates for "
                             f"{len(domain_totals):,} domains…",
                    )
                    body = {
                        "data": [
                            {
                                "range":  gspread.utils.rowcol_to_a1(r, c),
                                "values": [[v]],
                            }
                            for r, c, v in updates_needed
                        ],
                        "valueInputOption": "USER_ENTERED",
                    }
                    ws.spreadsheet.values_batch_update(body)

        except Exception as ex:
            st.warning(f"Spend update error: {ex}")

    return {
        "file":             filename,
        "new_added":        0,
        "existing_updated": len(domain_totals),
        "orders_written":   orders_written,
        "filtered":         result["filtered"],
        "filter_reasons":   result["filter_reasons"],
    }


# ─────────────────────────────────────────────────────────────────────────────
# PLACEHOLDER TABS (built next session)
# ─────────────────────────────────────────────────────────────────────────────

def tab_new_contacts(sh) -> None:
    st.header("New Contacts")
    st.info("🔜 Coming next — review and approve newly uploaded companies.")


def tab_watch_list(sh) -> None:
    st.header("Watch List")
    st.info("🔜 Coming soon — filterable master company table with detail panel.")


def tab_ham(sh) -> None:
    st.header("HAM Radio")
    st.info("🔜 Coming soon — HAM segment signal feed and outreach queue.")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    # ── header ──
    st.markdown(
        """
        <div style="
            background: #0066CC;
            padding: 14px 24px 10px 24px;
            border-radius: 6px;
            margin-bottom: 20px;
        ">
            <span style="color:#FFFFFF; font-size:1.5rem; font-weight:700;
                         letter-spacing:0.02em;">
                📡 UNI-T Customer Lead Dashboard
            </span>
            <span style="color:#B3D1F5; font-size:0.9rem; margin-left:16px;">
                Signal-driven B2B prospecting · North America
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── sheets connection ──
    sh = startup()

    if sh is None:
        st.error(
            "**Google Sheets connection failed.** "
            "Check that GOOGLE_SERVICE_ACCOUNT and GSHEET_ID are set in "
            "Streamlit Cloud → Settings → Secrets."
        )
        st.stop()

    # ── tabs ──
    tabs = st.tabs([
        "📤 Upload",
        "🆕 New Contacts",
        "👁 Watch List",
        "📻 HAM",
    ])

    with tabs[0]:
        tab_upload(sh)

    with tabs[1]:
        tab_new_contacts(sh)

    with tabs[2]:
        tab_watch_list(sh)

    with tabs[3]:
        tab_ham(sh)


if __name__ == "__main__":
    main()
