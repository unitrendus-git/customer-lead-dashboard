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

# Marketplace relay domains — proxy emails from Amazon/eBay/etc. via Shopify.
# Buyer identity is unknown; orders should be filtered, not attributed to a company.
MARKETPLACE_RELAY_DOMAINS = {
    "mail.codisto.com",   # Codisto multichannel connector (Amazon/eBay relay)
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
NC_PAGE_SIZE = 50   # contacts per page in New Contacts tab


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _sanitize_value(v):
    """Replace float nan/inf with '' — Sheets API rejects JSON-non-compliant floats."""
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return ""
    return v


def _sanitize_row(row: list) -> list:
    """Apply _sanitize_value to every element in a row list."""
    return [_sanitize_value(v) for v in row]


def _batch_update(ws, updates_needed: list) -> None:
    """
    Fire a single values_batch_update for a list of (row_1based, col_1based, value)
    tuples. One API call regardless of how many cells — avoids 429 quota errors.
    """
    body = {
        "data": [
            {
                "range":  gspread.utils.rowcol_to_a1(r, c),
                "values": [[_sanitize_value(v)]],
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
    headers = all_vals[0]
    col     = {h: i for i, h in enumerate(headers)}
    dom_col = col.get("domain", 0)
    dom_idx = {all_vals[i][dom_col]: i for i in range(1, len(all_vals))}
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
        if domain in MARKETPLACE_RELAY_DOMAINS:
            filtered += 1; reasons["marketplace relay"] += 1; continue
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
            if domain in MARKETPLACE_RELAY_DOMAINS:
                filtered += 1; reasons["marketplace relay"] += 1
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
# CUSTOMER BACKFILL
# ─────────────────────────────────────────────────────────────────────────────

def _backfill_customers(sh, silent: bool = False) -> None:
    """
    Read order_history tab, aggregate spend per domain, then batch-update
    customer_status = 'customer', total_spent, total_orders, has_purchases
    in BOTH master_companies and new_contacts.
    silent=True suppresses st.info/st.success output (used during auto-trigger after upload).
    """
    with st.spinner("Reading order history…"):
        order_rows = sheet_get_all(sh, SHEET_ORDERS)

    if not order_rows:
        st.warning("No rows found in order_history tab.")
        return

    # Aggregate spend per domain from order_history
    # ORDER_HEADERS: domain, order_number, order_date, product_name, sku,
    #                line_item_price, quantity, order_total, fulfillment_status,
    #                financial_status, billing_company, billing_email
    domain_stats: dict = {}
    for row in order_rows:
        domain = str(row.get("domain", "")).strip().lower()
        if not domain or domain in MARKETPLACE_RELAY_DOMAINS:
            continue
        order_num = str(row.get("order_number", "")).strip()
        try:
            order_total = float(str(row.get("order_total") or 0))
        except (ValueError, TypeError):
            order_total = 0.0

        if domain not in domain_stats:
            domain_stats[domain] = {"total_spent": 0.0, "order_nums": set()}
        if order_num and order_num not in domain_stats[domain]["order_nums"]:
            domain_stats[domain]["total_spent"]  += order_total
            domain_stats[domain]["order_nums"].add(order_num)

    # Finalize order counts
    for d in domain_stats:
        domain_stats[d]["total_orders"] = len(domain_stats[d]["order_nums"])

    n_domains = len(domain_stats)
    if not silent:
        st.info(f"Found {n_domains:,} domains with purchase history. Updating sheets…")

    # ── Update master_companies ──────────────────────────────────────────
    try:
        with st.spinner("Updating master_companies…"):
            ws, col, dom_idx = _sheet_index(sh)
            mc_updates = []
            mc_hit = 0
            for domain, stats in domain_stats.items():
                if domain not in dom_idx:
                    continue
                r = dom_idx[domain]
                spent = _sanitize_value(round(stats["total_spent"], 2))
                for field, value in [
                    ("customer_status", "customer"),
                    ("has_purchases",   "True"),
                    ("total_spent",     spent),
                    ("total_orders",    stats["total_orders"]),
                ]:
                    if field in col:
                        mc_updates.append((r + 1, col[field] + 1, value))
                mc_hit += 1
            if mc_updates:
                _batch_update(ws, mc_updates)
        if not silent:
            st.success(f"✅ master_companies: updated {mc_hit:,} domains")
    except Exception as ex:
        if not silent:
            st.error(f"master_companies update failed: {ex}")
        return

    # ── Update new_contacts ─────────────────────────────────────────────
    try:
        with st.spinner("Updating new_contacts…"):
            nc_ws   = sh.worksheet(SHEET_NEW_CONTACTS)
            nc_vals = nc_ws.get_all_values()
            if not nc_vals:
                st.warning("new_contacts tab is empty.")
                return
            nc_hdrs = nc_vals[0]

            def _col(name):
                return nc_hdrs.index(name) + 1 if name in nc_hdrs else None

            dom_c    = nc_hdrs.index("domain") if "domain" in nc_hdrs else 0
            status_c = _col("customer_status")
            spent_c  = _col("total_spent")
            orders_c = _col("total_orders")

            nc_updates = []
            nc_hit = 0
            for i, row_vals in enumerate(nc_vals[1:], start=2):
                d = row_vals[dom_c].strip().lower()
                if d not in domain_stats:
                    continue
                s = domain_stats[d]
                if status_c:
                    nc_updates.append((i, status_c, "customer"))
                if spent_c:
                    nc_updates.append((i, spent_c, round(s["total_spent"], 2)))
                if orders_c:
                    nc_updates.append((i, orders_c, s["total_orders"]))
                nc_hit += 1

            if nc_updates:
                # Use sheet-name-prefixed ranges to guarantee writes land on new_contacts
                # not the default/first tab (which is master_companies)
                sheet_title = nc_ws.title
                chunk_size = 500
                for chunk_start in range(0, len(nc_updates), chunk_size):
                    chunk = nc_updates[chunk_start:chunk_start + chunk_size]
                    body = {
                        "data": [
                            {
                                "range": f"'{sheet_title}'!{gspread.utils.rowcol_to_a1(r, c)}",
                                "values": [[_sanitize_value(v)]],
                            }
                            for r, c, v in chunk
                        ],
                        "valueInputOption": "RAW",
                    }
                    nc_ws.spreadsheet.values_batch_update(body)
                    time.sleep(0.5)

        if not silent:
            st.success(f"✅ new_contacts: updated {nc_hit:,} domains ({len(nc_updates):,} cells)")
        st.session_state["nc_cache_dirty"] = True
        if not silent:
            st.info("→ Go to New Contacts tab and click \"Reload from Sheet\" to see updated counts.")
    except Exception as ex:
        if not silent:
            st.error(f"new_contacts update failed: {ex}")


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

        # ── Backfill button (shown when no file is queued) ──────────────────────
        st.markdown("---")
        st.markdown("**🔧 Data repair tools**")
        st.caption(
            "Use these if customer status or spend data looks wrong after uploading orders."
        )
        if st.button("🔄 Backfill customer status from order history", key="backfill_customers"):
            _backfill_customers(sh)
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
            # Auto-backfill after every orders upload to sync customer_status
            if file_result.get("orders_written", 0) > 0:
                progress.progress(1.0, text="Syncing customer status…")
                _backfill_customers(sh, silent=True)

        session_results.append(file_result)

    progress.progress(1.0, text="Done.")
    _render_upload_summary(session_results)


def _write_contact_results(sh, result: dict, filename: str,
                            progress, pct_base: float, n_files: int) -> dict:
    """
    Write new master + new_contacts rows, then update existing domains.
    All existing-domain updates fire as a single values_batch_update call.
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
    _sanitize_value() is applied to every cell value before the API call.
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

                spent = _sanitize_value(round(stats["total_spent"], 2))
                for field, value in [
                    ("total_spent",     spent),
                    ("total_orders",    stats["total_orders"]),
                    ("has_purchases",   "True"),
                    ("customer_status", "customer"),
                    ("in_shopify",      "True"),
                    ("last_activity",   now_str()),
                ]:
                    if field in col:
                        updates_needed.append((r + 1, col[field] + 1, value))

                if stats["email"] and "best_contact_email" in col:
                    updates_needed.append(
                        (r + 1, col["best_contact_email"] + 1,
                         _sanitize_value(stats["email"])))
                if stats["company"] and "company_name" in col:
                    updates_needed.append(
                        (r + 1, col["company_name"] + 1,
                         _sanitize_value(stats["company"])))

            if updates_needed:
                progress.progress(
                    pct_base + 0.85 / n_files,
                    text=f"Writing spend updates for {len(domain_totals):,} domains…")
                _batch_update(ws, updates_needed)

            # ── Also promote customer_status + has_purchases in new_contacts ──
            # new_contacts mirrors key fields from master; update them in sync.
            try:
                nc_ws      = sh.worksheet(SHEET_NEW_CONTACTS)
                nc_vals    = nc_ws.get_all_values()
                if nc_vals:
                    nc_hdrs    = nc_vals[0]
                    nc_dom_col = nc_hdrs.index("domain") if "domain" in nc_hdrs else 0
                    nc_status_col = nc_hdrs.index("customer_status") + 1 \
                                    if "customer_status" in nc_hdrs else None
                    nc_spent_col  = nc_hdrs.index("total_spent") + 1 \
                                    if "total_spent" in nc_hdrs else None
                    nc_orders_col = nc_hdrs.index("total_orders") + 1 \
                                    if "total_orders" in nc_hdrs else None
                    nc_updates = []
                    for i, row_vals in enumerate(nc_vals[1:], start=2):
                        d = row_vals[nc_dom_col]
                        if d in domain_totals:
                            s = domain_totals[d]
                            if nc_status_col:
                                nc_updates.append((i, nc_status_col, "customer"))
                            if nc_spent_col:
                                nc_updates.append((i, nc_spent_col,
                                                   round(s["total_spent"], 2)))
                            if nc_orders_col:
                                nc_updates.append((i, nc_orders_col, s["total_orders"]))
                    if nc_updates:
                        _batch_update(nc_ws, nc_updates)
            except Exception as ex:
                st.warning(f"new_contacts customer sync error: {ex}")

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

# ─────────────────────────────────────────────────────────────────────────────
# TAB 2 — NEW CONTACTS
# ─────────────────────────────────────────────────────────────────────────────

def _nc_class_badge(cls: str) -> str:
    """Render an inline HTML color badge for a domain_class value."""
    color = CLASS_COLOR.get(cls, "#546E7A")
    label = CLASS_LABEL.get(cls, cls.title())
    return (
        f'<span style="background:{color};color:#fff;padding:2px 8px;'
        f'border-radius:4px;font-size:0.78rem;font-weight:600;'
        f'letter-spacing:0.03em;">{label}</span>'
    )


def _nc_write_master_flags(sh, domain: str, monitor_val: bool,
                           enrich_val: bool) -> None:
    """
    Write monitor + enrich flags for one domain to master_companies.
    Single _batch_update call — no loop, no 429 risk.
    """
    try:
        ws, col, dom_idx = _sheet_index(sh)
        if domain not in dom_idx:
            return
        r = dom_idx[domain]
        updates = []
        if "monitor" in col:
            updates.append((r + 1, col["monitor"] + 1, str(monitor_val)))
        if "enrich" in col:
            updates.append((r + 1, col["enrich"] + 1, str(enrich_val)))
        if updates:
            _batch_update(ws, updates)
    except Exception as ex:
        st.warning(f"Could not update master_companies for {domain}: {ex}")


def _nc_write_nc_flags(sh, nc_rows_data: list, domain: str,
                       monitor_val: bool, enrich_val: bool) -> None:
    """
    Write monitor + enrich flags back to new_contacts tab for one domain.
    Uses sheet-name-prefixed ranges to guarantee writes land on new_contacts.
    """
    try:
        ws = sh.worksheet(SHEET_NEW_CONTACTS)
        all_vals = ws.get_all_values()
        if not all_vals:
            return
        headers = all_vals[0]
        mon_col = headers.index("monitor") + 1 if "monitor" in headers else None
        enr_col = headers.index("enrich")  + 1 if "enrich"  in headers else None
        dom_col = headers.index("domain")  + 1 if "domain"  in headers else 1
        sheet_title = ws.title

        updates = []
        for i, row_vals in enumerate(all_vals[1:], start=2):
            if row_vals[dom_col - 1] == domain:
                if mon_col:
                    updates.append((i, mon_col, str(monitor_val)))
                if enr_col:
                    updates.append((i, enr_col, str(enrich_val)))
                break
        if updates:
            body = {
                "data": [
                    {
                        "range": f"'{sheet_title}'!{gspread.utils.rowcol_to_a1(r, c)}",
                        "values": [[_sanitize_value(v)]],
                    }
                    for r, c, v in updates
                ],
                "valueInputOption": "RAW",
            }
            ws.spreadsheet.values_batch_update(body)
    except Exception as ex:
        st.warning(f"Could not update new_contacts flags for {domain}: {ex}")


def _nc_mark_reviewed(sh, domains: list) -> bool:
    """
    Set reviewed = True for all given domains in new_contacts.
    Uses sheet-name-prefixed ranges to guarantee writes land on new_contacts.
    """
    try:
        ws = sh.worksheet(SHEET_NEW_CONTACTS)
        all_vals = ws.get_all_values()
        if not all_vals:
            return False
        headers = all_vals[0]
        rev_col = headers.index("reviewed") + 1 if "reviewed" in headers else None
        dom_col = headers.index("domain")   + 1 if "domain"   in headers else 1
        if not rev_col:
            return False
        sheet_title = ws.title

        domain_set = set(domains)
        updates = []
        for i, row_vals in enumerate(all_vals[1:], start=2):
            if row_vals[dom_col - 1] in domain_set:
                updates.append((i, rev_col, "True"))

        if updates:
            body = {
                "data": [
                    {
                        "range": f"'{sheet_title}'!{gspread.utils.rowcol_to_a1(r, c)}",
                        "values": [[v]],
                    }
                    for r, c, v in updates
                ],
                "valueInputOption": "RAW",
            }
            ws.spreadsheet.values_batch_update(body)
        return True
    except Exception as ex:
        st.warning(f"Mark-reviewed error: {ex}")
        return False


def _nc_safe_str(v) -> str:
    """Return string form of v, collapsing None/nan/float-nan to empty string."""
    if v is None:
        return ""
    s = str(v).strip()
    # Reject bare 'nan', 'nan nan', 'nan nan nan', etc.
    tokens = s.lower().split()
    if tokens and all(t == "nan" for t in tokens):
        return ""
    return s


def _nc_safe_bool(v, default: bool = False) -> bool:
    """
    Parse a Sheet boolean string case-insensitively.
    Sheets auto-uppercases USER_ENTERED booleans: 'TRUE'/'FALSE'.
    We also write 'True'/'False' from Python. Handle both.
    """
    if isinstance(v, bool):
        return v
    return str(v).strip().upper() == "TRUE"


def _nc_strip_trailing_nan(s: str) -> str:
    """Strip leading/trailing nan tokens from a name. 'Maxwell nan' -> 'Maxwell'."""
    import re as _re
    s = _re.sub(r'(?i)\s*\bnan\b\s*$', '', s).strip()
    s = _re.sub(r'(?i)^\s*\bnan\b\s*', '', s).strip()
    s = _re.sub(r'  +', ' ', s).strip()
    return s


def _nc_infer_name_from_email(email: str) -> str:
    """
    Extract a best-guess display name from an email local-part.
    jimmylee        -> "Jimmylee"  (single token, title-case)
    jimmy.lee       -> "Jimmy Lee" (dot-separated)
    jimmy_lee       -> "Jimmy Lee" (underscore-separated)
    j.lee / jlee   -> ""           (too short to be useful — suppress)
    jimmy.lee.phd  -> "Jimmy Lee"  (take first two tokens only)
    """
    if not email or "@" not in email:
        return ""
    local = email.split("@")[0].lower()
    # Split on dots, underscores, hyphens, digits
    import re
    parts = [p for p in re.split(r'[._\-0-9]+', local) if len(p) > 1]
    if not parts:
        return ""
    # Suppress if looks like initials only (all single char after split)
    if all(len(p) <= 1 for p in parts):
        return ""
    # Take at most first two meaningful tokens as first/last name
    name_parts = parts[:2]
    return " ".join(p.title() for p in name_parts)


def _nc_completeness(p: dict) -> int:
    """
    Score a prepped contact row 0-6 based on data richness.
    Used for ranking: higher = more actionable.
    """
    score = 0
    if p.get("contact_n_real"):    score += 1
    if p.get("contact_e"):         score += 1
    if p.get("contact_title"):     score += 1
    if p.get("company_real"):      score += 1
    if p.get("tags"):              score += 1
    if p.get("spent_raw", 0) > 0: score += 1
    return score


def _nc_completeness_bar(score: int, max_score: int = 6) -> str:
    """Return an HTML filled/empty dot bar representing completeness."""
    filled = "<span style='color:#0066CC;font-size:0.85rem;'>&#9679;</span>"
    empty  = "<span style='color:#CCC;font-size:0.85rem;'>&#9679;</span>"
    return "".join(filled if i < score else empty for i in range(max_score))


def _nc_load(sh) -> list:
    """
    Load new_contacts from Sheet and cache in session_state.
    Refresh forced by setting st.session_state['nc_cache_dirty'] = True before rerun.
    """
    if (
        "nc_rows_cache" not in st.session_state
        or st.session_state.get("nc_cache_dirty")
    ):
        with st.spinner("Loading contacts from Sheet…"):
            rows = sheet_get_all(sh, SHEET_NEW_CONTACTS)
        st.session_state["nc_rows_cache"] = rows
        st.session_state["nc_cache_dirty"] = False
    return st.session_state["nc_rows_cache"]


def tab_new_contacts(sh) -> None:
    st.header("New Contacts")
    st.caption(
        "All companies in the system — filter to unreviewed, customers, or by segment. "
        "Set monitor and enrich flags, then mark reviewed to clear the queue."
    )

    # ── Load (cached in session_state — filter changes don't re-hit the Sheet) ──
    all_rows   = _nc_load(sh)
    total_rows = len(all_rows)

    if not all_rows:
        st.info("No contacts loaded yet. Upload a file on the Upload tab first.")
        return

    # Pre-compute safe fields for every row once, not inside the render loop
    prepped = []
    for r in all_rows:
        domain        = _nc_safe_str(r.get("domain"))
        company_raw   = _nc_safe_str(r.get("company_name"))
        company_real  = bool(company_raw)
        company       = company_raw or domain
        cls           = _nc_safe_str(r.get("domain_class")) or "commercial"
        status        = _nc_safe_str(r.get("customer_status")) or "prospect"
        tags          = _nc_safe_str(r.get("tags"))
        contact_e     = _nc_safe_str(r.get("best_contact_email"))
        contact_title = _nc_safe_str(r.get("best_contact_title"))

        # Name: use source value if clean; otherwise infer from email
        contact_n_raw  = _nc_strip_trailing_nan(_nc_safe_str(r.get("best_contact_name")))
        contact_n_real = bool(contact_n_raw)
        if contact_n_raw:
            contact_n = contact_n_raw
        else:
            contact_n = _nc_infer_name_from_email(contact_e)

        try:
            spent_raw = float(str(r.get("total_spent") or 0))
        except (ValueError, TypeError):
            spent_raw = 0.0
        try:
            orders_raw = int(float(str(r.get("total_orders") or 0)))
        except (ValueError, TypeError):
            orders_raw = 0

        p = {
            "domain":         domain,
            "company":        company,
            "company_real":   company_real,
            "cls":            cls,
            "status":         status,
            "spent_raw":      spent_raw,
            "orders_raw":     orders_raw,
            "tags":           tags,
            "has_event":      _nc_safe_str(r.get("has_event_tag")).strip() == "True",
            "contact_n":      contact_n,
            "contact_n_real": contact_n_real,
            "contact_e":      contact_e,
            "contact_title":  contact_title,
            "first_seen":     _nc_safe_str(r.get("first_seen")),
            "added_date":     _nc_safe_str(r.get("added_date")),
            "monitor_cur":    _nc_safe_bool(r.get("monitor", "True"),  default=True),
            "enrich_cur":     _nc_safe_bool(r.get("enrich",  "False"), default=False),
            "reviewed":       _nc_safe_bool(r.get("reviewed"),         default=False),
            "_raw":           r,
        }
        p["completeness"] = _nc_completeness(p)
        prepped.append(p)

    available_classes = sorted({p["cls"] for p in prepped if p["cls"]})

    # ── Sidebar filters (no Sheet I/O on change) ──────────────────────────────
    with st.sidebar:
        st.markdown("### 🔍 Filter Contacts")

        show_filter = st.radio(
            "Show",
            ["Unreviewed only", "All contacts", "Customers", "Reviewed only"],
            index=0,
            key="nc_show_filter",
        )

        status_filter = st.radio(
            "Customer status",
            ["All", "Customers", "Prospects"],
            index=0,
            key="nc_status_filter",
        )

        class_filter = st.multiselect(
            "Segment",
            options=available_classes,
            default=[],
            placeholder="All segments",
            key="nc_class_filter",
        )

        has_purchase_only = st.checkbox(
            "Has purchases", value=False, key="nc_has_purchase"
        )
        has_event_only = st.checkbox(
            "Has event tag", value=False, key="nc_has_event"
        )

        st.markdown("---")
        sort_by = st.selectbox(
            "Sort by",
            ["Data completeness (best first)",
             "Date added (newest)", "Date added (oldest)",
             "Company name (A–Z)", "Total spend (high–low)"],
            key="nc_sort",
        )

        st.markdown("---")
        if st.button("🔄 Reload from Sheet", key="nc_reload"):
            st.session_state["nc_cache_dirty"] = True
            st.rerun()

    # ── Apply filters (pure Python — instant) ─────────────────────────────────
    view = prepped

    if show_filter == "Unreviewed only":
        view = [p for p in view if not p["reviewed"]]
    elif show_filter == "Customers":
        view = [p for p in view if p["status"] == "customer"]
    elif show_filter == "Reviewed only":
        view = [p for p in view if p["reviewed"]]

    if status_filter == "Customers":
        view = [p for p in view if p["status"] == "customer"]
    elif status_filter == "Prospects":
        view = [p for p in view if p["status"] != "customer"]

    if class_filter:
        view = [p for p in view if p["cls"] in class_filter]

    if has_purchase_only:
        view = [p for p in view if p["spent_raw"] > 0]

    if has_event_only:
        view = [p for p in view if p["has_event"]]

    # ── Sort ──────────────────────────────────────────────────────────────────
    if sort_by == "Data completeness (best first)":
        view = sorted(view, key=lambda p: -p["completeness"])
    elif sort_by == "Date added (oldest)":
        view = sorted(view, key=lambda p: p["added_date"] or p["first_seen"])
    elif sort_by == "Company name (A–Z)":
        view = sorted(view, key=lambda p: p["company"].lower())
    elif sort_by == "Total spend (high–low)":
        view = sorted(view, key=lambda p: -p["spent_raw"])

    # ── Summary bar ───────────────────────────────────────────────────────────
    n_unreviewed = sum(1 for p in prepped if not p["reviewed"])
    n_customers  = sum(1 for p in prepped if p["status"] == "customer")
    n_with_spend = sum(1 for p in prepped if p["spent_raw"] > 0)

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total in system",  f"{total_rows:,}")
    m2.metric("Unreviewed",       f"{n_unreviewed:,}")
    m3.metric("Customers",        f"{n_customers:,}")
    m4.metric("With purchases",   f"{n_with_spend:,}")
    m5.metric("Showing now",      f"{len(view):,}")

    st.markdown("---")

    if not view:
        st.info("No contacts match the current filters.")
        return

    # ── Bulk action ───────────────────────────────────────────────────────────
    unreviewed_in_view = [p for p in view if not p["reviewed"]]
    if unreviewed_in_view:
        if st.button(
            f"✅ Mark all {len(unreviewed_in_view):,} shown as reviewed",
            key="nc_mark_all",
            type="secondary",
        ):
            domains_to_clear = [p["domain"] for p in unreviewed_in_view if p["domain"]]
            with st.spinner(f"Marking {len(domains_to_clear):,} rows reviewed…"):
                ok = _nc_mark_reviewed(sh, domains_to_clear)
            if ok:
                st.success(f"Marked {len(domains_to_clear):,} contacts reviewed.")
                st.session_state["nc_cache_dirty"] = True
                st.rerun()

    # ── Pagination ────────────────────────────────────────────────────────────
    filter_fp = (
        f"{show_filter}|{status_filter}|{sorted(class_filter)}"
        f"|{has_purchase_only}|{has_event_only}|{sort_by}"
    )
    if st.session_state.get("nc_filter_fp") != filter_fp:
        st.session_state["nc_filter_fp"] = filter_fp
        st.session_state["nc_page"] = 0

    total_pages = max(1, math.ceil(len(view) / NC_PAGE_SIZE))
    page        = min(st.session_state.get("nc_page", 0), total_pages - 1)
    st.session_state["nc_page"] = page
    start_idx   = page * NC_PAGE_SIZE
    end_idx     = start_idx + NC_PAGE_SIZE
    page_view   = view[start_idx:end_idx]

    st.markdown(
        f"**{len(view):,} contact(s)** · sorted by {sort_by.lower()} · "
        f"page {page + 1} of {total_pages} "
        f"({start_idx + 1}–{min(end_idx, len(view))} of {len(view):,})"
    )
    st.markdown("")

    # ── Class color legend ────────────────────────────────────────────────────
    legend_html = " &nbsp; ".join(
        f'<span style="background:{CLASS_COLOR[c]};color:#fff;padding:1px 7px;'
        f'border-radius:3px;font-size:0.73rem;">{CLASS_LABEL[c]}</span>'
        for c in available_classes if c in CLASS_COLOR
    )
    st.markdown(legend_html, unsafe_allow_html=True)
    st.markdown("")

    # ── Per-row expanders (this page only) ────────────────────────────────────
    for idx, p in enumerate(page_view):
        abs_idx       = start_idx + idx
        domain        = p["domain"]
        company       = p["company"]
        cls           = p["cls"]
        status        = p["status"]
        spent_raw     = p["spent_raw"]
        orders_raw    = p["orders_raw"]
        tags          = p["tags"]
        has_event     = p["has_event"]
        contact_n     = p["contact_n"]
        contact_n_real = p["contact_n_real"]
        contact_e     = p["contact_e"]
        contact_title = p["contact_title"]
        first_seen    = p["first_seen"]
        added_date    = p["added_date"]
        monitor_cur   = p["monitor_cur"]
        enrich_cur    = p["enrich_cur"]
        reviewed      = p["reviewed"]

        cls_dot_color = CLASS_COLOR.get(cls, "#546E7A")
        cls_dot = (
            f'<span style="display:inline-block;width:10px;height:10px;'
            f'border-radius:50%;background:{cls_dot_color};margin-right:4px;"></span>'
        )
        spend_pill = (
            f' <span style="background:#1B5E20;color:#fff;padding:1px 6px;'
            f'border-radius:3px;font-size:0.72rem;font-weight:600;">'
            f'{format_currency(spent_raw)}</span>'
        ) if spent_raw > 0 else ""

        event_pill = (
            ' <span style="background:#E65100;color:#fff;padding:1px 6px;'
            'border-radius:3px;font-size:0.72rem;">event</span>'
        ) if has_event else ""

        mon_indicator = " 📡" if monitor_cur else ""
        enr_indicator = " 🔬" if enrich_cur else ""
        rev_indicator = " ✓"  if reviewed    else ""
        completeness_bar = _nc_completeness_bar(p["completeness"])

        label_html = (
            f"{cls_dot}<b>{company}</b> &nbsp;"
            f"<span style='color:#888;font-size:0.85rem;'>{domain}</span>"
            f"{spend_pill}{event_pill}"
            f" &nbsp; {completeness_bar}"
            f"<span style='color:#0066CC;font-size:0.8rem;'>{mon_indicator}</span>"
            f"<span style='color:#6A1B9A;font-size:0.8rem;'>{enr_indicator}</span>"
            f"<span style='color:#2E7D32;font-size:0.8rem;'>{rev_indicator}</span>"
        )
        st.markdown(
            f'<div style="margin-bottom:-10px;padding:4px 2px;font-size:0.88rem;">'
            f'{label_html}</div>',
            unsafe_allow_html=True,
        )

        with st.expander(f"{company}  ·  {domain}", expanded=False):

            badge_html  = _nc_class_badge(cls)
            status_icon = "🟢" if status == "customer" else "⚪"
            rev_badge   = (
                '<span style="background:#4CAF50;color:#fff;padding:1px 7px;'
                'border-radius:3px;font-size:0.75rem;margin-left:8px;">reviewed</span>'
            ) if reviewed else ""
            st.markdown(
                f"{badge_html} &nbsp; {status_icon} &nbsp;"
                f"<span style='font-size:0.85rem;color:#555;'>{status.title()}</span>"
                f"{rev_badge}",
                unsafe_allow_html=True,
            )
            st.markdown("")

            d1, d2, d3 = st.columns(3)
            with d1:
                st.markdown("**Domain**")
                st.code(domain, language=None)
                st.markdown(f"**Added:** {first_seen or added_date or '—'}")
            with d2:
                st.markdown("**Best contact**")
                if contact_n:
                    st.write(contact_n if contact_n_real else f"{contact_n} *(from email)*")
                else:
                    st.write("—")
                if contact_title:
                    st.caption(contact_title)
                if contact_e:
                    st.caption(contact_e)
            with d3:
                st.markdown("**Purchase data**")
                if orders_raw > 0:
                    st.write(f"{format_currency(spent_raw)} across {orders_raw} order(s)")
                else:
                    st.write("No purchases on record")
                if tags:
                    st.caption(f"Tags: {', '.join(t.strip() for t in tags.split(',') if t.strip())}")
                if has_event:
                    st.caption("🎟 Has event tag")

            st.markdown("")

            t1, t2, t3 = st.columns([1, 1, 2])
            with t1:
                new_monitor = st.checkbox(
                    "Monitor", value=monitor_cur,
                    key=f"nc_mon_{abs_idx}_{domain}",
                    help="Add to the active signal-monitoring queue.",
                )
            with t2:
                new_enrich = st.checkbox(
                    "Enrich", value=enrich_cur,
                    key=f"nc_enr_{abs_idx}_{domain}",
                    help="Queue for homepage scrape + ICP classification.",
                )

            if (new_monitor != monitor_cur) or (new_enrich != enrich_cur):
                with st.spinner("Saving flags…"):
                    _nc_write_master_flags(sh, domain, new_monitor, new_enrich)
                    _nc_write_nc_flags(sh, all_rows, domain, new_monitor, new_enrich)
                for cached_p in st.session_state.get("nc_rows_cache", []):
                    if cached_p.get("domain") == domain:
                        cached_p["monitor"] = str(new_monitor)
                        cached_p["enrich"]  = str(new_enrich)
                        break
                st.toast(f"Flags updated for {company}.", icon="✅")

            with t3:
                if not reviewed:
                    if st.button("✅ Mark reviewed", key=f"nc_rev_{abs_idx}_{domain}",
                                 type="primary"):
                        with st.spinner("Marking reviewed…"):
                            _nc_mark_reviewed(sh, [domain])
                        st.session_state["nc_cache_dirty"] = True
                        st.rerun()
                else:
                    st.caption("✓ Already reviewed")

    # ── Prev / Next controls ──────────────────────────────────────────────────
    st.markdown("")
    st.markdown("---")
    nav_l, nav_mid, nav_r = st.columns([1, 3, 1])
    with nav_l:
        if page > 0:
            if st.button("← Previous", key="nc_prev"):
                st.session_state["nc_page"] = page - 1
                st.rerun()
    with nav_mid:
        st.markdown(
            f"<div style='text-align:center;color:#666;font-size:0.85rem;'>"
            f"Page {page + 1} of {total_pages} · "
            f"{start_idx + 1}–{min(end_idx, len(view))} of {len(view):,} contacts"
            f"</div>",
            unsafe_allow_html=True,
        )
    with nav_r:
        if page < total_pages - 1:
            if st.button("Next →", key="nc_next"):
                st.session_state["nc_page"] = page + 1
                st.rerun()


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
