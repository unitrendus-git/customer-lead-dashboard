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
import time
import datetime
import streamlit as st

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

BATCH_SIZE   = 200   # rows per Sheets API write
BATCH_DELAY  = 1.0   # seconds between batches


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
      "brevo"           — Brevo CSV export
      "shopify_contacts"— Shopify contacts XLSX
      "shopify_orders"  — Shopify orders CSV
      None              — unrecognised
    """
    cols = set(df.columns)

    brevo_required = {"EMAIL", "FIRSTNAME", "LASTNAME", "COMPANY"}
    shopify_contact_required = {"Email", "First Name", "Last Name", "Company"}
    shopify_orders_required  = {"Name", "Financial Status", "Lineitem name", "Lineitem price"}

    if brevo_required.issubset(cols):
        return "brevo"
    if shopify_contact_required.issubset(cols):
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

HAM_SIGNALS = {"arrl", "qrz", "hamradio", "ham"}  # substring hints only


def classify_domain(domain: str) -> str:
    """
    Derive domain_class from the domain string.
    Returns one of: commercial / education / government / defense / ham / distributor.

    NOTE: 'distributor' classification cannot be inferred from a domain alone —
    it must come from the master list.  This function will never return 'distributor'.
    Distributor rows uploaded via new exports will land as 'commercial' and
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

    # Rough HAM heuristic — enrichment will correct if wrong
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

    Returns:
      {
        "new_master_rows": [...],   # rows for master_companies
        "new_contact_rows": [...],  # rows for new_contacts
        "update_map": {...},        # domain → {updates} for existing rows
        "filtered": int,            # count of skipped rows
        "filter_reasons": Counter,  # why rows were skipped
      }
    """
    from collections import Counter

    new_master   = []
    new_contacts = []
    update_map   = {}
    filtered     = 0
    reasons      = Counter()
    seen         = set()  # deduplicate within this file

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
            continue  # keep first occurrence per domain
        seen.add(domain)

        first = str(row.get("FIRSTNAME", "")).strip()
        last  = str(row.get("LASTNAME", "")).strip()
        name  = f"{first} {last}".strip()
        tags  = str(row.get("TAGS", "") or row.get("Tags", "")).strip()

        if domain in existing_domains:
            # Existing — silent update
            update_map[domain] = {
                "brevo_contacts": 1,   # will be incremented in sheet
                "last_activity": now_str(),
                "in_brevo": "True",
            }
            if tags:
                update_map[domain]["tags"] = tags
        else:
            # New domain
            cls = classify_domain(domain)
            monitor_default = DEFAULT_MONITOR.get(cls, True)
            suppress        = SUPPRESS_OUTREACH.get(cls, False)

            master_row = _blank_master_row()
            master_row["domain"]             = domain
            master_row["company_name"]       = company
            master_row["domain_class"]       = cls
            master_row["customer_status"]    = "prospect"
            master_row["watch_tier"]         = 2
            master_row["score"]              = 1
            master_row["score_domain"]       = 1
            master_row["score_contact"]      = 1
            master_row["score_engagement"]   = 1
            master_row["monitor"]            = str(monitor_default)
            master_row["enrich"]             = "False"
            master_row["suppress_outreach"]  = str(suppress)
            master_row["brevo_contacts"]     = 1
            master_row["shopify_contacts"]   = 0
            master_row["total_spent"]        = 0
            master_row["total_orders"]       = 0
            master_row["best_contact_name"]  = name
            master_row["best_contact_email"] = email
            master_row["tags"]               = tags
            master_row["has_event_tag"]      = "True" if tags else "False"
            master_row["in_brevo"]           = "True"
            master_row["in_shopify"]         = "False"
            master_row["has_purchases"]      = "False"
            master_row["first_seen"]         = today_str()
            master_row["last_activity"]      = now_str()
            master_row["enriched"]           = "False"
            master_row["outreach_status"]    = "none"

            # Handle tektronix manually
            if domain in MANUAL_REVIEW_DOMAINS:
                master_row["notes"] = MANUAL_REVIEW_DOMAINS[domain]

            new_master.append(master_row)

            nc_row = _new_contact_row(master_row)
            new_contacts.append(nc_row)

            existing_domains.add(domain)

    return {
        "new_master_rows":   new_master,
        "new_contact_rows":  new_contacts,
        "update_map":        update_map,
        "filtered":          filtered,
        "filter_reasons":    reasons,
    }


# ─────────────────────────────────────────────────────────────────────────────
# SHOPIFY CONTACTS PROCESSOR
# ─────────────────────────────────────────────────────────────────────────────

def process_shopify_contacts(df: "pd.DataFrame", existing_domains: set) -> dict:
    """
    Parse a Shopify contacts XLSX export.

    Column names: Email, First Name, Last Name, Company, Tags, ...
    Same rules as Brevo, different column names.
    """
    from collections import Counter

    new_master   = []
    new_contacts = []
    update_map   = {}
    filtered     = 0
    reasons      = Counter()
    seen         = set()

    for _, row in df.iterrows():
        email   = str(row.get("Email", "")).strip().lower()
        company = str(row.get("Company", "")).strip()

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
                "last_activity": now_str(),
                "in_shopify": "True",
            }
            if tags:
                update_map[domain]["tags"] = tags
        else:
            cls = classify_domain(domain)
            monitor_default = DEFAULT_MONITOR.get(cls, True)
            suppress        = SUPPRESS_OUTREACH.get(cls, False)

            master_row = _blank_master_row()
            master_row["domain"]             = domain
            master_row["company_name"]       = company
            master_row["domain_class"]       = cls
            master_row["customer_status"]    = "prospect"
            master_row["watch_tier"]         = 2
            master_row["score"]              = 1
            master_row["score_domain"]       = 1
            master_row["score_contact"]      = 1
            master_row["score_engagement"]   = 1
            master_row["monitor"]            = str(monitor_default)
            master_row["enrich"]             = "False"
            master_row["suppress_outreach"]  = str(suppress)
            master_row["brevo_contacts"]     = 0
            master_row["shopify_contacts"]   = 1
            master_row["total_spent"]        = 0
            master_row["total_orders"]       = 0
            master_row["best_contact_name"]  = name
            master_row["best_contact_email"] = email
            master_row["tags"]               = tags
            master_row["has_event_tag"]      = "True" if tags else "False"
            master_row["in_brevo"]           = "False"
            master_row["in_shopify"]         = "True"
            master_row["has_purchases"]      = "False"
            master_row["first_seen"]         = today_str()
            master_row["last_activity"]      = now_str()
            master_row["enriched"]           = "False"
            master_row["outreach_status"]    = "none"

            if domain in MANUAL_REVIEW_DOMAINS:
                master_row["notes"] = MANUAL_REVIEW_DOMAINS[domain]

            new_master.append(master_row)
            new_contacts.append(_new_contact_row(master_row))
            existing_domains.add(domain)

    return {
        "new_master_rows":   new_master,
        "new_contact_rows":  new_contacts,
        "update_map":        update_map,
        "filtered":          filtered,
        "filter_reasons":    reasons,
    }


# ─────────────────────────────────────────────────────────────────────────────
# SHOPIFY ORDERS PROCESSOR
# ─────────────────────────────────────────────────────────────────────────────

def process_shopify_orders(df: "pd.DataFrame", existing_domains: set) -> dict:
    """
    Parse a Shopify orders CSV export.

    Shopify orders are multi-row: Email/Total/Status/Tags appear only on the
    FIRST row of each order. Continuation rows have blank Email and carry
    additional line items for the same order.

    Skips:
      - Rows where Tags contains "amazon" (case-insensitive)
      - Rows where Financial Status == "refunded"
      - Personal / own email domains
    """
    from collections import Counter

    order_rows     = []   # rows for order_history tab
    domain_totals  = {}   # domain → {total_spent, total_orders, last_date, email, company}
    filtered       = 0
    reasons        = Counter()

    # Carry-forward state for multi-row orders
    current_email  = ""
    current_total  = 0.0
    current_status = ""
    current_tags   = ""
    current_order  = ""
    current_date   = ""
    current_company= ""

    for _, row in df.iterrows():

        # ── detect first row of a new order (Email is populated) ──
        email_raw = str(row.get("Email", "")).strip()

        if email_raw and "@" in email_raw:
            # New order — reset carry-forward state
            current_email   = email_raw.lower()
            current_order   = str(row.get("Name", "")).strip()
            current_date    = str(row.get("Created at", "") or row.get("Paid at", "")).strip()
            current_company = str(row.get("Billing Company", "")).strip()

            try:
                current_total = float(str(row.get("Total", 0) or 0).replace(",", ""))
            except ValueError:
                current_total = 0.0

            current_status = str(row.get("Financial Status", "")).strip().lower()
            current_tags   = str(row.get("Tags", "")).strip()

            # ── order-level filters ──
            if current_status == "refunded":
                filtered += 1
                reasons["refunded order"] += 1
                current_email = ""  # suppress line items
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

        # ── process line item (first row or continuation) ──
        if not current_email:
            continue  # order was filtered — skip all its line items

        domain = extract_domain(current_email)
        if not domain:
            continue

        product  = str(row.get("Lineitem name", "")).strip()
        sku      = str(row.get("Lineitem sku", "")).strip()
        qty_raw  = str(row.get("Lineitem quantity", 1) or 1)
        price_raw= str(row.get("Lineitem price", 0) or 0)

        try:
            qty   = int(float(qty_raw))
        except ValueError:
            qty   = 1
        try:
            price = float(price_raw.replace(",", ""))
        except ValueError:
            price = 0.0

        fulfil  = str(row.get("Fulfillment Status", "")).strip()

        order_row = [
            domain,
            current_order,
            current_date,
            product,
            sku,
            price,
            qty,
            current_total,
            fulfil,
            current_status,
            current_company,
            current_email,
        ]
        order_rows.append(order_row)

        # Accumulate domain-level stats
        if domain not in domain_totals:
            domain_totals[domain] = {
                "total_spent":  0.0,
                "total_orders": 0,
                "last_date":    current_date,
                "email":        current_email,
                "company":      current_company,
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
        "order_rows":    order_rows,
        "domain_totals": domain_totals,
        "filtered":      filtered,
        "filter_reasons":reasons,
    }


# ─────────────────────────────────────────────────────────────────────────────
# ROW BUILDERS
# ─────────────────────────────────────────────────────────────────────────────

def _blank_master_row() -> dict:
    """Return a dict with all MASTER_HEADERS keys set to empty string."""
    return {h: "" for h in MASTER_HEADERS}


def _master_row_to_list(row: dict) -> list:
    """Convert a master_row dict to an ordered list matching MASTER_HEADERS."""
    return [row.get(h, "") for h in MASTER_HEADERS]


def _new_contact_row(master_row: dict) -> list:
    """Build a new_contacts row (list) from a master_row dict."""
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
        today_str(),   # added_date
        "False",       # reviewed
    ]


# ─────────────────────────────────────────────────────────────────────────────
# BATCH WRITER
# ─────────────────────────────────────────────────────────────────────────────

def batch_write(sh, tab_name: str, rows: list[list], progress_label: str = "") -> bool:
    """
    Write rows to a sheet tab in batches of BATCH_SIZE with delay between
    batches to stay within Sheets API quota.
    Returns True if all batches succeeded.
    """
    if not rows:
        return True

    total   = len(rows)
    success = True

    for start in range(0, total, BATCH_SIZE):
        batch = rows[start : start + BATCH_SIZE]
        ok    = sheet_append_rows(sh, tab_name, batch)
        if not ok:
            success = False
        if start + BATCH_SIZE < total:
            time.sleep(BATCH_DELAY)

    return success


# ─────────────────────────────────────────────────────────────────────────────
# UPLOAD SUMMARY BUILDER
# ─────────────────────────────────────────────────────────────────────────────

def _render_upload_summary(results: list[dict]) -> None:
    """
    Render a clean summary card after processing one or more files.
    results = list of per-file result dicts (merged before display).
    """
    total_new     = sum(r.get("new_added", 0)        for r in results)
    total_updated = sum(r.get("existing_updated", 0) for r in results)
    total_orders  = sum(r.get("orders_written", 0)   for r in results)
    total_filtered= sum(r.get("filtered", 0)         for r in results)

    # Merge filter reasons
    from collections import Counter
    merged_reasons: Counter = Counter()
    for r in results:
        merged_reasons += r.get("filter_reasons", Counter())

    st.success("✅ Upload complete")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("New companies",      total_new)
    col2.metric("Existing updated",   total_updated)
    col3.metric("Orders processed",   total_orders)
    col4.metric("Filtered / skipped", total_filtered)

    if merged_reasons:
        with st.expander("Filter breakdown"):
            for reason, count in sorted(merged_reasons.items(),
                                        key=lambda x: -x[1]):
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

    # Show what's queued
    st.write(f"**{len(uploaded_files)} file(s) ready to process:**")
    for f in uploaded_files:
        st.write(f"  • `{f.name}` ({f.size:,} bytes)")

    if not st.button("🚀 Process uploads", type="primary"):
        return

    # ── process each file ──────────────────────────────────────────────────

    existing_domains = sheet_get_domains(sh)
    session_results  = []

    progress = st.progress(0.0, text="Starting…")

    for file_idx, f in enumerate(uploaded_files):
        pct_base = file_idx / len(uploaded_files)
        progress.progress(pct_base, text=f"Reading {f.name}…")

        df = read_uploaded_file(f)
        if df is None:
            session_results.append({
                "file": f.name,
                "error": "Could not read file",
                "new_added": 0, "existing_updated": 0,
                "orders_written": 0, "filtered": 0,
                "filter_reasons": {},
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
                "file": f.name,
                "error": "unrecognised file type",
                "new_added": 0, "existing_updated": 0,
                "orders_written": 0, "filtered": 0,
                "filter_reasons": {},
            })
            continue

        type_labels = {
            "brevo":             "Brevo contacts",
            "shopify_contacts":  "Shopify contacts",
            "shopify_orders":    "Shopify orders",
        }
        progress.progress(pct_base + 0.1 / len(uploaded_files),
                          text=f"Processing {type_labels[file_type]}: {f.name}…")

        # ── route to correct processor ────────────────────────────────────

        if file_type == "brevo":
            result = process_brevo(df, existing_domains)
            file_result = _write_contact_results(sh, result, f.name, progress,
                                                  pct_base, len(uploaded_files))

        elif file_type == "shopify_contacts":
            result = process_shopify_contacts(df, existing_domains)
            file_result = _write_contact_results(sh, result, f.name, progress,
                                                  pct_base, len(uploaded_files))

        else:  # shopify_orders
            result = process_shopify_orders(df, existing_domains)
            file_result = _write_order_results(sh, result, f.name, progress,
                                                pct_base, len(uploaded_files))

        session_results.append(file_result)

    progress.progress(1.0, text="Done.")
    _render_upload_summary(session_results)


def _write_contact_results(sh, result: dict, filename: str,
                            progress, pct_base: float, n_files: int) -> dict:
    """
    Write new master rows, new_contacts rows, and handle existing-domain
    updates for a contact file (Brevo or Shopify contacts).
    Returns a file-level summary dict.
    """
    new_master   = result["new_master_rows"]
    new_contacts = result["new_contact_rows"]
    update_map   = result["update_map"]

    # Convert dicts to ordered lists
    master_lists  = [_master_row_to_list(r) for r in new_master]
    contact_lists = new_contacts  # already lists from _new_contact_row()

    new_added = 0
    if master_lists:
        progress.progress(pct_base + 0.4 / n_files,
                          text=f"Writing {len(master_lists):,} new companies…")
        ok = batch_write(sh, SHEET_MASTER, master_lists, "master_companies")
        if ok:
            new_added = len(master_lists)

    if contact_lists:
        progress.progress(pct_base + 0.6 / n_files,
                          text=f"Writing {len(contact_lists):,} new_contacts rows…")
        batch_write(sh, SHEET_NEW_CONTACTS, contact_lists, "new_contacts")

    # Existing-domain updates — one API call per domain (quiet)
    existing_updated = 0
    if update_map:
        progress.progress(pct_base + 0.8 / n_files,
                          text=f"Updating {len(update_map):,} existing companies…")
        for domain, updates in update_map.items():
            sheet_update_row(sh, SHEET_MASTER, "domain", domain, updates)
            existing_updated += 1

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
    Write order_history rows and update master_companies spend/order counts
    for a Shopify orders file.
    Returns a file-level summary dict.
    """
    order_rows    = result["order_rows"]
    domain_totals = result["domain_totals"]

    orders_written = 0
    if order_rows:
        progress.progress(pct_base + 0.5 / n_files,
                          text=f"Writing {len(order_rows):,} order line items…")
        ok = batch_write(sh, SHEET_ORDERS, order_rows, "order_history")
        if ok:
            orders_written = len(order_rows)

    # Update master_companies spend data
    if domain_totals:
        progress.progress(pct_base + 0.8 / n_files,
                          text=f"Updating spend data for {len(domain_totals):,} domains…")
        for domain, stats in domain_totals.items():
            updates = {
                "total_spent":   round(stats["total_spent"], 2),
                "total_orders":  stats["total_orders"],
                "has_purchases": "True",
                "in_shopify":    "True",
                "last_activity": now_str(),
            }
            if stats["email"]:
                updates["best_contact_email"] = stats["email"]
            if stats["company"]:
                updates["company_name"] = stats["company"]

            sheet_update_row(sh, SHEET_MASTER, "domain", domain, updates)

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
    # ── TEMPORARY DIAGNOSTIC ──────────────────────────────────────────────
    import traceback, json as _json
    try:
        import gspread as _gs
        from google.oauth2.service_account import Credentials as _Creds
        sa_raw    = get_secret("GOOGLE_SERVICE_ACCOUNT")
        gsheet_id = get_secret("GSHEET_ID")
        with st.expander("🔧 Sheets diagnostic", expanded=True):
            st.write(f"GSHEET_ID: `{gsheet_id[:12] if gsheet_id else 'MISSING'}`")
            st.write(f"SA JSON length: `{len(sa_raw) if sa_raw else 0}`")
            sa_info = _json.loads(sa_raw)
            st.write(f"client_email: `{sa_info.get('client_email')}`")
            pk = sa_info.get("private_key","")
            if r"\n" in pk:
                sa_info["private_key"] = pk.replace(r"\n", chr(10))
            creds = _Creds.from_service_account_info(sa_info, scopes=[
                "https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/drive",
            ])
            gc = _gs.authorize(creds)
            sh = gc.open_by_key(gsheet_id)
            st.success(f"✅ Connected: **{sh.title}**")
    except Exception:
        st.error("Diagnostic failed:")
        st.code(traceback.format_exc())
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
