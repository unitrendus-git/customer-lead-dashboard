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

try:
    import pandas as pd
    PD_OK = True
except ImportError:
    PD_OK = False

try:
    import openpyxl
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
# CONSTANTS
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

EDUCATION_TLDS  = {".edu", ".ac.uk", ".edu.au", ".edu.ca"}
GOVERNMENT_TLDS = {".gov", ".mil", ".gov.uk", ".gc.ca"}

DEFENSE_DOMAINS = {
    "lmco.com", "l3harris.com", "harris.com", "boeing.com",
    "northropgrumman.com", "ngc.com", "raytheon.com", "rtx.com",
    "baesystems.com", "textron.com",
}

HAM_SIGNALS = {"arrl", "qrz", "hamradio", "ham"}

BATCH_SIZE  = 200   # rows per append call
BATCH_DELAY = 1.0   # seconds between append batches


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _sanitize_row(row: list) -> list:
    """Replace float nan/inf with '' — Sheets API rejects JSON-non-compliant floats."""
    result = []
    for v in row:
        if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
            result.append("")
        else:
            result.append(v)
    return result


def _batch_update(ws, updates_needed: list) -> None:
    """
    Fire a single values_batch_update for a list of (row_1based, col_1based, value)
    tuples. One API call regardless of how many cells — avoids 429 quota errors.
    """
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


def _sheet_index(sh) -> tuple:
    """
    Read master_companies once and return (ws, col_map, domain_row_index).
      col_map          — {header_name: 0-based column index}
      domain_row_index — {domain_string: 0-based row index in all_vals}
    """
    ws       = sh.worksheet(SHEET_MASTER)
    all_vals = ws.get_all_values()
    if not all_vals:
        return ws, {}, {}
    headers  = all_vals[0]
    col      = {h: i for i, h in enumerate(headers)}
    dom_col  = col.get("domain", 0)
    dom_idx  = {all_vals[i][dom_col]: i for i in range(1, len(all_vals))}
    return ws, col, dom_idx


# ─────────────────────────────────────────────────────────────────────────────
# STARTUP
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_resource(ttl=300)
def get_sheet():
    return gsheet_open()


def startup():
    sh = get_sheet()
    if sh:
        gsheet_ensure_tabs(sh)
    return sh


# ─────────────────────────────────────────────────────────────────────────────
# FILE-TYPE DETECTION
# ─────────────────────────────────────────────────────────────────────────────

def detect_file_type(df: "pd.DataFrame") -> "str | None":
    """
    Identify upload type from column headers.
      "brevo"            — Brevo CSV  (all-caps EMAIL, FIRSTNAME, LASTNAME, COMPANY)
      "shopify_contacts" — Shopify contacts XLSX  (mixed-case Email, First Name, Last Name)
      "shopify_orders"   — Shopify orders CSV  (Name, Financial Status, Lineitem name/price)
      None               — unrecognised
    """
    cols = set(df.columns)
    brevo_req  = {"EMAIL", "FIRSTNAME", "LASTNAME", "COMPANY"}
    sc_req     = {"Email", "First Name", "Last Name"}
    orders_req = {"Name", "Financial Status", "Lineitem name", "Lineitem price"}

    if brevo_req.issubset(cols):
        return "brevo"
    # Guard: Shopify contacts must NOT have all-caps EMAIL (avoids Brevo overlap)
    if sc_req.issubset(cols) and "EMAIL" not in cols:
        return "shopify_contacts"
    if orders_req.issubset(cols):
        return "shopify_orders"
    return None


# ─────────────────────────────────────────────────────────────────────────────
# FILE READER
# ─────────────────────────────────────────────────────────────────────────────

def read_uploaded_file(f) -> "pd.DataFrame | None":
    if not PD_OK:
        st.error("pandas not installed.")
        return None
    name = f.name.lower()
    try:
        if name.endswith(".xlsx"):
            if not OPENPYXL_OK:
                st.error("openpyxl not installed.")
                return None
            return pd.read_excel(io.BytesIO(f.read()), engine="openpyxl")
        raw = f.read()
        try:
            return pd.read_csv(io.BytesIO(raw), dtype=str)
        except UnicodeDecodeError:
            return pd.read_csv(io.BytesIO(raw), dtype=str, encoding="latin-1")
    except Exception as ex:
        st.warning(f"Could not read {f.name}: {ex}")
        return None


# ─────────────────────────────────────────────────────────────────────────────
# DOMAIN CLASSIFIER
# ─────────────────────────────────────────────────────────────────────────────

def classify_domain(domain: str) -> str:
    """
    Infer domain_class from domain string alone.
    Never returns 'distributor' — that requires master list context.
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
# ROW BUILDERS
# ─────────────────────────────────────────────────────────────────────────────

def _blank_master_row() -> dict:
    return {h: "" for h in MASTER_HEADERS}


def _master_row_to_list(row: dict) -> list:
    return [row.get(h, "") for h in MASTER_HEADERS]


def _new_contact_row(master_row: dict) -> list:
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
        today_str(),
        "False",
    ]


def _build_master_row(domain, company, email, name, tags, cls, source) -> dict:
    """Build a complete new master_companies row dict from parsed contact data."""
    monitor_default = DEFAULT_MONITOR.get(cls, True)
    suppress        = SUPPRESS_OUTREACH.get(cls, False)
    is_brevo        = source == "brevo"
    is_shopify      = source == "shopify"

    row = _blank_master_row()
    row.update({
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
        "brevo_contacts":     1 if is_brevo else 0,
        "shopify_contacts":   1 if is_shopify else 0,
        "total_spent":        0,
        "total_orders":       0,
        "best_contact_name":  name,
        "best_contact_email": email,
        "tags":               tags,
        "has_event_tag":      "True" if tags else "False",
        "in_brevo":           str(is_brevo),
        "in_shopify":         str(is_shopify),
        "has_purchases":      "False",
        "first_seen":         today_str(),
        "last_activity":      now_str(),
        "enriched":           "False",
        "outreach_status":    "none",
    })
    if domain in MANUAL_REVIEW_DOMAINS:
        row["notes"] = MANUAL_REVIEW_DOMAINS[domain]
    return row


# ─────────────────────────────────────────────────────────────────────────────
# CONTACT FILE PARSER (shared by Brevo + Shopify contacts)
# ─────────────────────────────────────────────────────────────────────────────

def _parse_contact_file(df: "pd.DataFrame", existing_domains: set,
                        source: str) -> dict:
    """
    Core contact-parsing logic shared by Brevo and Shopify contacts.
    source: "brevo" | "shopify"
    """
    from collections import Counter

    if source == "brevo":
        col_email   = "EMAIL"
        col_company = "COMPANY"
        col_first   = "FIRSTNAME"
        col_last    = "LASTNAME"
        col_tags    = "TAGS"
    else:
        col_email   = "Email"
        col_company = "Company"
        col_first   = "First Name"
        col_last    = "Last Name"
        col_tags    = "Tags"

    new_master   = []
    new_contacts = []
    update_map   = {}
    filtered     = 0
    reasons      = Counter()
    seen         = set()

    for _, row in df.iterrows():
        email = str(row.get(col_email, "")).strip().lower()

        # Shopify: company fallback to Address Company
        company = str(
            row.get(col_company, "") or row.get("Address Company", "")
        ).strip()

        if not email or "@" not in email:
            filtered += 1; reasons["no email"] += 1; continue

        domain = extract_domain(email)
        if not domain:
            filtered += 1; reasons["no domain"] += 1; continue
        if domain in PERSONAL_EMAIL_DOMAINS:
            filtered += 1; reasons["personal email"] += 1; continue
        if domain in OWN_DOMAINS:
            filtered += 1; reasons["own domain"] += 1; continue
        if domain in seen:
            continue
        seen.add(domain)

        first = str(row.get(col_first, "")).strip()
        last  = str(row.get(col_last, "")).strip()
        name  = f"{first} {last}".strip()
        tags  = str(row.get(col_tags, "") or "").strip()

        if domain in existing_domains:
            upd = {"last_activity": now_str()}
            if source == "brevo":
                upd["in_brevo"] = "True"
            else:
                upd["in_shopify"] = "True"
            if tags:
                upd["tags"] = tags
            update_map[domain] = upd
        else:
            cls        = classify_domain(domain)
            master_row = _build_master_row(domain, company, email, name,
                                           tags, cls, source)
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


def process_brevo(df, existing_domains):
    return _parse_contact_file(df, existing_domains, "brevo")


def process_shopify_contacts(df, existing_domains):
    return _parse_contact_file(df, existing_domains, "shopify")


# ─────────────────────────────────────────────────────────────────────────────
# SHOPIFY ORDERS PROCESSOR
# ─────────────────────────────────────────────────────────────────────────────

def process_shopify_orders(df: "pd.DataFrame", existing_domains: set) -> dict:
    """
    Parse a Shopify orders CSV.

    Multi-row format: Email/Total/Status/Tags on the FIRST row of each order;
    continuation rows have blank Email. Carry-forward state via current_* vars.

    Skips amazon-tagged orders, refunded orders, personal/own domains.
    """
    from collections import Counter

    order_rows    = []
    domain_totals = {}
    filtered      = 0
    reasons       = Counter()

    current_email = current_order = current_date = current_company = ""
    current_total  = 0.0
    current_status = current_tags = ""

    for _, row in df.iterrows():
        email_raw = str(row.get("Email", "")).strip()

        if email_raw and "@" in email_raw:
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

            if current_status == "refunded":
                filtered += 1; reasons["refunded order"] += 1
                current_email = ""; continue
            if "amazon" in current_tags.lower():
                filtered += 1; reasons["amazon order"] += 1
                current_email = ""; continue

            domain = extract_domain(current_email)
            if not domain:
                current_email = ""; continue
            if domain in PERSONAL_EMAIL_DOMAINS:
                filtered += 1; reasons["personal email"] += 1
                current_email = ""; continue
            if domain in OWN_DOMAINS:
                filtered += 1; reasons["own domain"] += 1
                current_email = ""; continue

        if not current_email:
            continue

        domain = extract_domain(current_email)
        if not domain:
            continue

        try:
            qty = int(float(str(row.get("Lineitem quantity", 1) or 1)))
        except ValueError:
            qty = 1
        try:
            price = float(str(row.get("Lineitem price", 0) or 0).replace(",", ""))
        except ValueError:
            price = 0.0

        order_rows.append([
            domain,
            current_order,
            current_date,
            str(row.get("Lineitem name", "")).strip(),
            str(row.get("Lineitem sku", "")).strip(),
            price,
            qty,
            current_total,
            str(row.get("Fulfillment Status", "")).strip(),
            current_status,
            current_company,
            current_email,
        ])

        if domain not in domain_totals:
            domain_totals[domain] = {
                "total_spent": 0.0, "total_orders": 0,
                "last_date": current_date, "email": current_email,
                "company": current_company, "counted_orders": set(),
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
# BATCH APPEND WRITER
# ─────────────────────────────────────────────────────────────────────────────

def batch_write(sh, tab_name: str, rows: list, progress_label: str = "") -> bool:
    """
    Append rows in BATCH_SIZE chunks with BATCH_DELAY between chunks.
    Sanitizes nan/inf values before each write.
    """
    if not rows:
        return True
    success = True
    for start in range(0, len(rows), BATCH_SIZE):
        batch = [_sanitize_row(r) for r in rows[start : start + BATCH_SIZE]]
        if not sheet_append_rows(sh, tab_name, batch):
            success = False
        if start + BATCH_SIZE < len(rows):
            time.sleep(BATCH_DELAY)
    return success


# ─────────────────────────────────────────────────────────────────────────────
# UPLOAD SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

def _render_upload_summary(results: list) -> None:
    from collections import Counter
    total_new      = sum(r.get("new_added", 0)        for r in results)
    total_updated  = sum(r.get("existing_updated", 0) for r in results)
    total_orders   = sum(r.get("orders_written", 0)   for r in results)
    total_filtered = sum(r.get("filtered", 0)         for r in results)
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
    st.header("Upload Contact & Order Data")
    st.caption(
        "Accepts Brevo CSV exports, Shopify contact XLSX exports, and "
        "Shopify order CSV exports. File type is detected automatically from "
        "column headers — filenames don't matter."
    )

    if not PD_OK:
        st.error("**pandas** is not installed. Add pandas and openpyxl to requirements.txt.")
        return

    uploaded_files = st.file_uploader(
        "Drop files here (CSV or XLSX)",
        type=["csv", "xlsx"],
        accept_multiple_files=True,
        key="upload_files",
    )

    if not uploaded_files:
        st.info("No files selected yet. Drag one or more files above to begin.")
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
                "file": f.name, "new_added": 0, "existing_updated": 0,
                "orders_written": 0, "filtered": 0, "filter_reasons": {},
            })
            continue

        file_type = detect_file_type(df)
        if file_type is None:
            st.warning(
                f"⚠️ `{f.name}` — unrecognised column layout. "
                "Expected Brevo CSV, Shopify contacts XLSX, or Shopify orders CSV. Skipping."
            )
            session_results.append({
                "file": f.name, "new_added": 0, "existing_updated": 0,
                "orders_written": 0, "filtered": 0, "filter_reasons": {},
            })
            continue

        type_labels = {
            "brevo": "Brevo contacts",
            "shopify_contacts": "Shopify contacts",
            "shopify_orders": "Shopify orders",
        }
        progress.progress(
            pct_base + 0.1 / len(uploaded_files),
            text=f"Processing {type_labels[file_type]}: {f.name}…",
        )

        if file_type in ("brevo", "shopify_contacts"):
            result      = (process_brevo if file_type == "brevo"
                           else process_shopify_contacts)(df, existing_domains)
            file_result = _write_contact_results(
                sh, result, f.name, progress, pct_base, len(uploaded_files)
            )
        else:
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
    Write new master + new_contacts rows, then update existing domains.
    All existing-domain updates fire as a single values_batch_update call —
    one API request regardless of domain count — eliminating 429 errors.
    """
    new_master   = result["new_master_rows"]
    new_contacts = result["new_contact_rows"]
    update_map   = result["update_map"]

    new_added = 0
    if new_master:
        progress.progress(pct_base + 0.4 / n_files,
                          text=f"Writing {len(new_master):,} new companies…")
        if batch_write(sh, SHEET_MASTER, [_master_row_to_list(r) for r in new_master]):
            new_added = len(new_master)

    if new_contacts:
        progress.progress(pct_base + 0.6 / n_files,
                          text=f"Writing {len(new_contacts):,} new_contacts rows…")
        batch_write(sh, SHEET_NEW_CONTACTS, new_contacts)

    existing_updated = 0
    if update_map and GSPREAD_OK:
        progress.progress(pct_base + 0.75 / n_files,
                          text=f"Reading master list for {len(update_map):,} updates…")
        try:
            ws, col, dom_idx = _sheet_index(sh)
            updates_needed   = []

            for domain, updates in update_map.items():
                if domain not in dom_idx:
                    continue
                r = dom_idx[domain]
                for field, value in updates.items():
                    if field in col:
                        updates_needed.append((r + 1, col[field] + 1, value))
                existing_updated += 1

            if updates_needed:
                progress.progress(pct_base + 0.9 / n_files,
                                  text=f"Writing {existing_updated:,} contact updates…")
                _batch_update(ws, updates_needed)

        except Exception as ex:
            st.warning(f"Contact update error: {ex}")

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
    Write order_history rows, then update master_companies spend data.
    All spend updates fire as a single values_batch_update call.
    """
    order_rows    = result["order_rows"]
    domain_totals = result["domain_totals"]

    orders_written = 0
    if order_rows:
        progress.progress(pct_base + 0.4 / n_files,
                          text=f"Writing {len(order_rows):,} order line items…")
        if batch_write(sh, SHEET_ORDERS, order_rows):
            orders_written = len(order_rows)

    if domain_totals and GSPREAD_OK:
        progress.progress(pct_base + 0.6 / n_files,
                          text="Reading master list for spend update…")
        try:
            ws, col, dom_idx = _sheet_index(sh)
            updates_needed   = []

            for domain, stats in domain_totals.items():
                if domain not in dom_idx:
                    continue
                r = dom_idx[domain]
def _sv(v):
                    """Sanitize a single value — replace nan/inf with empty string."""
                    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
                        return ""
                    return v

                for field, value in [
                    ("total_spent",   _sv(round(stats["total_spent"], 2))),
                    ("total_orders",  stats["total_orders"]),
                    ("has_purchases", "True"),
                    ("in_shopify",    "True"),
                    ("last_activity", now_str()),
                ]:
                    if field in col:
                        updates_needed.append((r + 1, col[field] + 1, value))
                if stats["email"] and "best_contact_email" in col:
                    updates_needed.append(
                        (r + 1, col["best_contact_email"] + 1, _sv(stats["email"])))
                if stats["company"] and "company_name" in col:
                    updates_needed.append(
                        (r + 1, col["company_name"] + 1, _sv(stats["company"])))

            if updates_needed:
                progress.progress(
                    pct_base + 0.85 / n_files,
                    text=f"Writing spend updates for {len(domain_totals):,} domains…")
                _batch_update(ws, updates_needed)

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
# PLACEHOLDER TABS
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

    sh = startup()
    if sh is None:
        st.error(
            "**Google Sheets connection failed.** "
            "Check that GOOGLE_SERVICE_ACCOUNT and GSHEET_ID are set in "
            "Streamlit Cloud → Settings → Secrets."
        )
        st.stop()

    tabs = st.tabs(["📤 Upload", "🆕 New Contacts", "👁 Watch List", "📻 HAM"])

    with tabs[0]: tab_upload(sh)
    with tabs[1]: tab_new_contacts(sh)
    with tabs[2]: tab_watch_list(sh)
    with tabs[3]: tab_ham(sh)


if __name__ == "__main__":
    main()
