"""
app.py — Customer Lead Dashboard
==================================
UNI-T North America | Signal-driven B2B prospecting dashboard.

Tab build order (deploy and verify each before starting the next):
  1. Upload
  2. New Contacts
  3. Watch List
  4. Signal Feed   (Phase 3)
  5. Outreach Generator (Phase 3)
  6. HAM           (Phase 3)
  7. VV Review Queue (Phase 3)

Do not put credentials in this file. All secrets live in Streamlit Cloud
Settings -> Secrets (TOML format). See CLD_HANDOFF_V2.md for schema.
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

st.set_page_config(
    page_title="UNI-T Customer Lead Dashboard",
    page_icon="📡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

PERSONAL_EMAIL_DOMAINS = {
    "gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "aol.com",
    "icloud.com", "me.com", "mac.com", "live.com", "msn.com",
    "comcast.net", "att.net", "cox.net", "verizon.net", "earthlink.net",
    "sbcglobal.net", "bellsouth.net",
}

MARKETPLACE_RELAY_DOMAINS = {
    "mail.codisto.com",
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

BATCH_SIZE   = 200
BATCH_DELAY  = 1.0
NC_PAGE_SIZE = 50

# Strategic Tier 1 override -- companies always worth monitoring regardless of score.
# Mirrors TIER1_OVERRIDE in build_master_company_list.py.
# Used by the "Apply Tier 1 overrides" repair tool in tab_upload.
TIER1_OVERRIDE = {
    "ti.com", "analog.com", "nxp.com", "onsemi.com", "microchip.com",
    "infineon.com", "st.com", "renesas.com", "mchp.com", "maximintegrated.com",
    "latticesemi.com", "idt.com", "skyworksinc.com", "qorvo.com",
    "wolfspeed.com", "macom.com", "semtech.com",
    "rivian.com", "lucidmotors.com", "borgwarner.com", "vicr.com",
    "teradyne.com", "ni.com", "astronics.com", "spirent.com",
    "qualcomm.com", "broadcom.com", "marvell.com", "amd.com",
    "intel.com", "nvidia.com", "murata.com", "tdk.com", "vishay.com",
    "aptiv.com", "visteon.com", "magna.com", "denso.com", "continental.com",
    "rockwellautomation.com", "abb.com", "yaskawa.com", "danfoss.com",
    "emerson.com", "siemens.com", "schneider-electric.com",
}


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _sanitize_value(v):
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return ""
    return v

def _sanitize_row(row):
    return [_sanitize_value(v) for v in row]

def _batch_update(ws, updates_needed):
    body = {
        "data": [
            {"range": gspread.utils.rowcol_to_a1(r, c), "values": [[_sanitize_value(v)]]}
            for r, c, v in updates_needed
        ],
        "valueInputOption": "USER_ENTERED",
    }
    ws.spreadsheet.values_batch_update(body)

def _sheet_index(sh):
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
# FILE HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def detect_file_type(df):
    cols      = set(df.columns)
    brevo_req = {"EMAIL", "FIRSTNAME", "LASTNAME", "COMPANY"}
    sc_req    = {"Email", "First Name", "Last Name"}
    ord_req   = {"Name", "Financial Status", "Lineitem name", "Lineitem price"}
    if brevo_req.issubset(cols):          return "brevo"
    if sc_req.issubset(cols) and "EMAIL" not in cols: return "shopify_contacts"
    if ord_req.issubset(cols):            return "shopify_orders"
    return None

def read_uploaded_file(f):
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

def classify_domain(domain):
    d = domain.lower().strip()
    if d in DEFENSE_DOMAINS: return "defense"
    for tld in EDUCATION_TLDS:
        if d.endswith(tld): return "education"
    for tld in GOVERNMENT_TLDS:
        if d.endswith(tld): return "government"
    for hint in HAM_SIGNALS:
        if hint in d: return "ham"
    return "commercial"


# ─────────────────────────────────────────────────────────────────────────────
# ROW BUILDERS
# ─────────────────────────────────────────────────────────────────────────────

def _blank_master_row():
    return {h: "" for h in MASTER_HEADERS}

def _master_row_to_list(row):
    return [row.get(h, "") for h in MASTER_HEADERS]

def _new_contact_row(master_row):
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

def _build_master_row(domain, company, email, name, tags, cls, source):
    monitor_default = DEFAULT_MONITOR.get(cls, True)
    suppress        = SUPPRESS_OUTREACH.get(cls, False)
    is_brevo        = source == "brevo"
    is_shopify      = source == "shopify"
    row = _blank_master_row()
    row.update({
        "domain": domain, "company_name": company, "domain_class": cls,
        "customer_status": "prospect", "watch_tier": 2,
        "score": 1, "score_domain": 1, "score_contact": 1, "score_engagement": 1,
        "monitor": str(monitor_default), "enrich": "False",
        "suppress_outreach": str(suppress),
        "brevo_contacts": 1 if is_brevo else 0,
        "shopify_contacts": 1 if is_shopify else 0,
        "total_spent": 0, "total_orders": 0,
        "best_contact_name": name, "best_contact_email": email,
        "tags": tags, "has_event_tag": "True" if tags else "False",
        "in_brevo": str(is_brevo), "in_shopify": str(is_shopify),
        "has_purchases": "False", "first_seen": today_str(),
        "last_activity": now_str(), "enriched": "False",
        "outreach_status": "none",
    })
    if domain in MANUAL_REVIEW_DOMAINS:
        row["notes"] = MANUAL_REVIEW_DOMAINS[domain]
    return row


# ─────────────────────────────────────────────────────────────────────────────
# CONTACT FILE PARSER
# ─────────────────────────────────────────────────────────────────────────────

def _parse_contact_file(df, existing_domains, source):
    from collections import Counter
    if source == "brevo":
        col_email, col_company, col_first, col_last, col_tags = "EMAIL", "COMPANY", "FIRSTNAME", "LASTNAME", "TAGS"
    else:
        col_email, col_company, col_first, col_last, col_tags = "Email", "Company", "First Name", "Last Name", "Tags"

    new_master, new_contacts, update_map = [], [], {}
    filtered, reasons, seen = 0, Counter(), set()

    for _, row in df.iterrows():
        email   = str(row.get(col_email, "")).strip().lower()
        company = str(row.get(col_company, "") or row.get("Address Company", "")).strip()

        if not email or "@" not in email: filtered += 1; reasons["no email"] += 1; continue
        domain = extract_domain(email)
        if not domain: filtered += 1; reasons["no domain"] += 1; continue
        if domain in PERSONAL_EMAIL_DOMAINS:     filtered += 1; reasons["personal email"] += 1; continue
        if domain in MARKETPLACE_RELAY_DOMAINS:  filtered += 1; reasons["marketplace relay"] += 1; continue
        if domain in OWN_DOMAINS:               filtered += 1; reasons["own domain"] += 1; continue
        if domain in seen: continue
        seen.add(domain)

        first = str(row.get(col_first, "")).strip()
        last  = str(row.get(col_last, "")).strip()
        name  = f"{first} {last}".strip()
        tags  = str(row.get(col_tags, "") or "").strip()

        if domain in existing_domains:
            upd = {"last_activity": now_str()}
            upd["in_brevo" if source == "brevo" else "in_shopify"] = "True"
            if tags: upd["tags"] = tags
            update_map[domain] = upd
        else:
            cls        = classify_domain(domain)
            master_row = _build_master_row(domain, company, email, name, tags, cls, source)
            new_master.append(master_row)
            new_contacts.append(_new_contact_row(master_row))
            existing_domains.add(domain)

    return {"new_master_rows": new_master, "new_contact_rows": new_contacts,
            "update_map": update_map, "filtered": filtered, "filter_reasons": reasons}

def process_brevo(df, existing_domains):
    return _parse_contact_file(df, existing_domains, "brevo")

def process_shopify_contacts(df, existing_domains):
    return _parse_contact_file(df, existing_domains, "shopify")


# ─────────────────────────────────────────────────────────────────────────────
# ORDERS PROCESSOR
# ─────────────────────────────────────────────────────────────────────────────

def process_shopify_orders(df, existing_domains):
    from collections import Counter
    order_rows, domain_totals, filtered, reasons = [], {}, 0, Counter()
    cur_email = cur_order = cur_date = cur_company = ""
    cur_total = 0.0
    cur_status = cur_tags = ""

    for _, row in df.iterrows():
        email_raw = str(row.get("Email", "")).strip()
        if email_raw and "@" in email_raw:
            cur_email   = email_raw.lower()
            cur_order   = str(row.get("Name", "")).strip()
            cur_date    = str(row.get("Created at", "") or row.get("Paid at", "")).strip()
            cur_company = str(row.get("Billing Company", "")).strip()
            try:   cur_total = float(str(row.get("Total", 0) or 0).replace(",", ""))
            except: cur_total = 0.0
            cur_status = str(row.get("Financial Status", "")).strip().lower()
            cur_tags   = str(row.get("Tags", "")).strip()

            if cur_status == "refunded": filtered += 1; reasons["refunded order"] += 1; cur_email = ""; continue
            if "amazon" in cur_tags.lower(): filtered += 1; reasons["amazon order"] += 1; cur_email = ""; continue
            domain = extract_domain(cur_email)
            if not domain: cur_email = ""; continue
            if domain in PERSONAL_EMAIL_DOMAINS:    filtered += 1; reasons["personal email"] += 1; cur_email = ""; continue
            if domain in MARKETPLACE_RELAY_DOMAINS: filtered += 1; reasons["marketplace relay"] += 1; cur_email = ""; continue
            if domain in OWN_DOMAINS:              filtered += 1; reasons["own domain"] += 1; cur_email = ""; continue

        if not cur_email: continue
        domain = extract_domain(cur_email)
        if not domain: continue

        try:   qty   = int(float(str(row.get("Lineitem quantity", 1) or 1)))
        except: qty  = 1
        try:   price = float(str(row.get("Lineitem price", 0) or 0).replace(",", ""))
        except: price = 0.0

        order_rows.append([domain, cur_order, cur_date,
            str(row.get("Lineitem name", "")).strip(), str(row.get("Lineitem sku", "")).strip(),
            price, qty, cur_total,
            str(row.get("Fulfillment Status", "")).strip(), cur_status, cur_company, cur_email])

        if domain not in domain_totals:
            domain_totals[domain] = {"total_spent": 0.0, "total_orders": 0,
                "last_date": cur_date, "email": cur_email, "company": cur_company, "counted_orders": set()}
        stats = domain_totals[domain]
        if cur_order not in stats["counted_orders"]:
            stats["total_orders"] += 1
            stats["total_spent"]  += cur_total
            stats["counted_orders"].add(cur_order)
            if cur_date > stats["last_date"]: stats["last_date"] = cur_date

    return {"order_rows": order_rows, "domain_totals": domain_totals,
            "filtered": filtered, "filter_reasons": reasons}


# ─────────────────────────────────────────────────────────────────────────────
# BATCH WRITE
# ─────────────────────────────────────────────────────────────────────────────

def batch_write(sh, tab_name, rows, progress_label=""):
    if not rows: return True
    success = True
    for start in range(0, len(rows), BATCH_SIZE):
        batch = [_sanitize_row(r) for r in rows[start:start + BATCH_SIZE]]
        if not sheet_append_rows(sh, tab_name, batch): success = False
        if start + BATCH_SIZE < len(rows): time.sleep(BATCH_DELAY)
    return success


# ─────────────────────────────────────────────────────────────────────────────
# UPLOAD SUMMARY
# ─────────────────────────────────────────────────────────────────────────────

def _render_upload_summary(results):
    from collections import Counter
    total_new      = sum(r.get("new_added", 0) for r in results)
    total_updated  = sum(r.get("existing_updated", 0) for r in results)
    total_orders   = sum(r.get("orders_written", 0) for r in results)
    total_filtered = sum(r.get("filtered", 0) for r in results)
    merged = Counter()
    for r in results: merged += r.get("filter_reasons", Counter())

    st.success("Upload complete")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("New companies", total_new)
    c2.metric("Existing updated", total_updated)
    c3.metric("Orders processed", total_orders)
    c4.metric("Filtered / skipped", total_filtered)
    if merged:
        with st.expander("Filter breakdown"):
            for reason, count in sorted(merged.items(), key=lambda x: -x[1]):
                st.write(f"- **{reason}** -- {count:,}")


# ─────────────────────────────────────────────────────────────────────────────
# CUSTOMER BACKFILL
# ─────────────────────────────────────────────────────────────────────────────

def _backfill_customers(sh, silent=False):
    with st.spinner("Reading order history..."):
        order_rows = sheet_get_all(sh, SHEET_ORDERS)
    if not order_rows:
        st.warning("No rows found in order_history tab.")
        return

    domain_stats = {}
    for row in order_rows:
        domain    = str(row.get("domain", "")).strip().lower()
        if not domain or domain in MARKETPLACE_RELAY_DOMAINS: continue
        order_num = str(row.get("order_number", "")).strip()
        try:   order_total = float(str(row.get("order_total") or 0))
        except: order_total = 0.0
        if domain not in domain_stats:
            domain_stats[domain] = {"total_spent": 0.0, "order_nums": set()}
        if order_num and order_num not in domain_stats[domain]["order_nums"]:
            domain_stats[domain]["total_spent"] += order_total
            domain_stats[domain]["order_nums"].add(order_num)
    for d in domain_stats:
        domain_stats[d]["total_orders"] = len(domain_stats[d]["order_nums"])

    if not silent:
        st.info(f"Found {len(domain_stats):,} domains with purchase history. Updating sheets...")

    try:
        with st.spinner("Updating master_companies..."):
            ws, col, dom_idx = _sheet_index(sh)
            mc_updates, mc_hit = [], 0
            for domain, stats in domain_stats.items():
                if domain not in dom_idx: continue
                r     = dom_idx[domain]
                spent = _sanitize_value(round(stats["total_spent"], 2))
                for field, value in [("customer_status", "customer"), ("has_purchases", "True"),
                                     ("total_spent", spent), ("total_orders", stats["total_orders"])]:
                    if field in col: mc_updates.append((r + 1, col[field] + 1, value))
                mc_hit += 1
            if mc_updates: _batch_update(ws, mc_updates)
        if not silent: st.success(f"master_companies: updated {mc_hit:,} domains")
    except Exception as ex:
        if not silent: st.error(f"master_companies update failed: {ex}")
        return

    try:
        with st.spinner("Updating new_contacts..."):
            nc_ws   = sh.worksheet(SHEET_NEW_CONTACTS)
            nc_vals = nc_ws.get_all_values()
            if not nc_vals: st.warning("new_contacts tab is empty."); return
            nc_hdrs = nc_vals[0]
            def _col(name): return nc_hdrs.index(name) + 1 if name in nc_hdrs else None
            dom_c    = nc_hdrs.index("domain") if "domain" in nc_hdrs else 0
            status_c = _col("customer_status")
            spent_c  = _col("total_spent")
            orders_c = _col("total_orders")

            nc_updates, nc_hit = [], 0
            for i, row_vals in enumerate(nc_vals[1:], start=2):
                d = row_vals[dom_c].strip().lower()
                if d not in domain_stats: continue
                s = domain_stats[d]
                if status_c: nc_updates.append((i, status_c, "customer"))
                if spent_c:  nc_updates.append((i, spent_c, round(s["total_spent"], 2)))
                if orders_c: nc_updates.append((i, orders_c, s["total_orders"]))
                nc_hit += 1

            if nc_updates:
                sheet_title = nc_ws.title
                for chunk_start in range(0, len(nc_updates), 500):
                    chunk = nc_updates[chunk_start:chunk_start + 500]
                    body  = {"data": [{"range": f"'{sheet_title}'!{gspread.utils.rowcol_to_a1(r, c)}",
                                       "values": [[_sanitize_value(v)]]} for r, c, v in chunk],
                             "valueInputOption": "RAW"}
                    nc_ws.spreadsheet.values_batch_update(body)
                    time.sleep(0.5)

        if not silent: st.success(f"new_contacts: updated {nc_hit:,} domains ({len(nc_updates):,} cells)")
        st.session_state["nc_cache_dirty"] = True
        if not silent: st.info('Go to New Contacts and click "Reload from Sheet" to see updated counts.')
    except Exception as ex:
        if not silent: st.error(f"new_contacts update failed: {ex}")


# ─────────────────────────────────────────────────────────────────────────────
# TIER 1 OVERRIDE REPAIR TOOL
# ─────────────────────────────────────────────────────────────────────────────

def _apply_tier1_overrides(sh) -> None:
    """
    Batch-update watch_tier = 1 in master_companies for every domain in
    TIER1_OVERRIDE that is currently not Tier 1.
    Single values_batch_update call -- no 429 risk.
    """
    with st.spinner("Reading master_companies..."):
        ws, col, dom_idx = _sheet_index(sh)

    if "watch_tier" not in col:
        st.error("watch_tier column not found in master_companies.")
        return

    tier_col = col["watch_tier"] + 1   # 1-based for gspread
    updates  = []
    found    = []

    all_vals = ws.get_all_values()
    headers  = all_vals[0]
    dom_c    = headers.index("domain") if "domain" in headers else 0
    tier_c   = headers.index("watch_tier") if "watch_tier" in headers else None

    for i, row in enumerate(all_vals[1:], start=2):
        domain = row[dom_c].strip().lower()
        if domain in TIER1_OVERRIDE:
            current = row[tier_c].strip() if tier_c is not None else ""
            found.append((domain, current))
            if current != "1":
                updates.append((i, tier_col, 1))

    if not found:
        st.info("None of the Tier 1 override domains were found in the Sheet yet. "
                "Upload contact data first.")
        return

    already = sum(1 for _, t in found if t == "1")
    to_fix  = len(updates)

    st.write(f"Found **{len(found)}** override domains in Sheet -- "
             f"{already} already Tier 1, **{to_fix}** need updating.")

    if to_fix == 0:
        st.success("All override domains are already Tier 1. Nothing to do.")
        return

    with st.spinner(f"Updating {to_fix} domain(s) to Tier 1..."):
        _batch_update(ws, updates)

    st.success(f"Done -- {to_fix} domain(s) set to Tier 1.")
    st.session_state["wl_cache_dirty"] = True
    st.info("Go to Watch List and click Reload from Sheet to see the changes.")



# ─────────────────────────────────────────────────────────────────────────────
# ENRICHMENT
# ─────────────────────────────────────────────────────────────────────────────

# ICP labels Haiku may return -- must match ICP_LABELS in utils.py
_ENRICH_ICP_OPTIONS = [
    "ICP-4 R&D Engineer",
    "ICP-5 Power Electronics",
    "ICP-7 Test Lab",
    "ICP-6 Education",
    "ICP-1 Industrial",
    "ICP-2 Electrician",
    "ICP-3 Solar",
    "Mixed",
    "None",
]

_ENRICH_INDUSTRY_OPTIONS = [
    "Electronics/Semiconductor",
    "Power Electronics/Energy",
    "Automotive/EV",
    "Telecom/RF",
    "Industrial/Manufacturing",
    "Test & Measurement",
    "Education",
    "Government/Research",
    "Other",
]

# Prompt template -- braces doubled for .format() safety
_ENRICH_PROMPT_TMPL = (
    "You are classifying a B2B company for a test and measurement instrument supplier.\n"
    "The company homepage text is provided below.\n\n"
    "Return ONLY a JSON object with these four fields -- no preamble, no markdown fences:\n"
    "{{\n"
    '  \"website_description\": \"One sentence (max 20 words) describing what the company does.\",\n'
    '  \"industry\": \"<one of: {industries}>\",\n'
    '  \"icp_label\": \"<one of: {icps}>\",\n'
    '  \"icp_confidence\": \"<one of: High, Medium, Low>\"\n'
    "}}\n\n"
    "icp_label guidance:\n"
    "- ICP-4 R&D Engineer: electronics design, FPGA, embedded, semiconductor design, hardware startup\n"
    "- ICP-5 Power Electronics: EV, GaN/SiC, motor drives, solar inverter OEM, data-center PSU\n"
    "- ICP-7 Test Lab: contract test lab, consultancy, EMC/EMI, product certification, prototype shop\n"
    "- ICP-6 Education: university, community college, trade school, teaching lab, makerspace\n"
    "- ICP-1 Industrial: plant operations, MRO, manufacturing maintenance, industrial automation\n"
    "- ICP-2 Electrician: electrical contractor, electrical construction, field electrician\n"
    "- ICP-3 Solar: PV installer, solar O&M, renewable energy, BESS/ESS\n"
    "- Mixed: clearly serves multiple ICPs equally\n"
    "- None: no obvious T&M instrument buyer (retail, food, law, HR, etc.)\n\n"
    "Homepage text (first 1500 chars):\n"
    "{text}"
)


def _enrich_fetch_text(domain: str, timeout: int = 8) -> str:
    """Fetch homepage text via trafilatura. Returns empty string on any failure."""
    try:
        import trafilatura
        import requests as _req
        for url in [f"https://www.{domain}", f"https://{domain}"]:
            try:
                resp = _req.get(
                    url, timeout=timeout,
                    headers={"User-Agent": "Mozilla/5.0 (compatible; CLD-Enricher/1.0)"},
                    allow_redirects=True,
                )
                if resp.status_code < 400:
                    extracted = trafilatura.extract(
                        resp.text,
                        include_links=False,
                        include_images=False,
                        include_tables=False,
                    )
                    text = (extracted or "").strip()
                    if text:
                        return text[:1500]
            except Exception:
                continue
        return ""
    except Exception:
        return ""


def _enrich_classify(client, text: str) -> dict:
    """Call Claude Haiku to classify a company from its homepage text."""
    import json as _json
    prompt = _ENRICH_PROMPT_TMPL.format(
        industries=", ".join(_ENRICH_INDUSTRY_OPTIONS),
        icps=", ".join(_ENRICH_ICP_OPTIONS),
        text=text,
    )
    try:
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=256,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = msg.content[0].text.strip()
        # Strip markdown fences if Haiku added them
        if raw.startswith("```"):
            parts = raw.split("```")
            raw = parts[1] if len(parts) > 1 else parts[0]
            if raw.startswith("json"):
                raw = raw[4:]
            raw = raw.strip()
        result = _json.loads(raw)
        desc = str(result.get("website_description", ""))[:120]
        ind  = result.get("industry", "Other")
        icp  = result.get("icp_label", "None")
        conf = result.get("icp_confidence", "Low")
        if ind  not in _ENRICH_INDUSTRY_OPTIONS: ind  = "Other"
        if icp  not in _ENRICH_ICP_OPTIONS:      icp  = "None"
        if conf not in ("High", "Medium", "Low"): conf = "Low"
        return {"website_description": desc, "industry": ind,
                "icp_label": icp, "icp_confidence": conf, "error": ""}
    except Exception as ex:
        return {"website_description": "", "industry": "",
                "icp_label": "", "icp_confidence": "", "error": str(ex)[:120]}


def _run_enrichment(sh, mode: str = "flagged") -> None:
    """
    Enrich company records from their homepage.

    mode:
      "flagged" -- only rows where enrich == "True" and enriched != "True"
      "tier1"   -- all Tier 1 rows where enriched != "True"
      "all"     -- all rows where enriched != "True" (use with caution)
    """
    import time as _time

    client = get_anthropic_client()
    if client is None:
        st.error("Anthropic API client not available. Check ANTHROPIC_API_KEY secret.")
        return

    # ── 1. Load master_companies ───────────────────────────────────────────────
    with st.spinner("Reading master_companies..."):
        ws, col, dom_idx = _sheet_index(sh)

    all_vals = ws.get_all_values()
    if not all_vals:
        st.warning("master_companies is empty.")
        return
    headers = all_vals[0]

    def _hcol(name):
        return headers.index(name) if name in headers else None

    dom_c     = _hcol("domain")
    enrich_c  = _hcol("enrich")
    enrichd_c = _hcol("enriched")
    tier_c    = _hcol("watch_tier")
    cls_c     = _hcol("domain_class")

    # ── 2. Build candidate list ───────────────────────────────────────────────
    candidates = []  # list of (sheet_row_1based, domain)
    for i, row in enumerate(all_vals[1:], start=2):
        domain  = row[dom_c].strip().lower() if dom_c is not None else ""
        if not domain:
            continue
        already = (
            row[enrichd_c].strip().upper() == "TRUE"
            if enrichd_c is not None else False
        )
        if already:
            continue
        enrich_flag = (
            row[enrich_c].strip().upper() == "TRUE"
            if enrich_c is not None else False
        )
        cls = row[cls_c].strip().lower() if cls_c is not None else ""
        try:
            tier_val = int(float(row[tier_c].strip())) if tier_c is not None else 3
        except Exception:
            tier_val = 3

        # Distributors and defense: no outreach, no ICP classification needed
        if cls in ("distributor", "defense"):
            continue

        if mode == "flagged" and not enrich_flag:
            continue
        if mode == "tier1" and tier_val != 1:
            continue
        # mode == 'all' passes everything remaining

        candidates.append((i, domain))

    if not candidates:
        st.info(
            "No domains match the enrichment criteria. "
            "Flag domains with the Enrich checkbox in New Contacts or Watch List, "
            "then re-run."
        )
        return

    est_min = max(1, len(candidates) * 10 // 60)
    st.info(
        f"**{len(candidates):,} domain(s)** queued for enrichment. "
        f"Estimated time: ~{est_min} min "
        f"({len(candidates)} × ~10s each)."
    )

    # ── 3. Per-domain fetch + classify loop ─────────────────────────────
    progress   = st.progress(0.0, text="Starting enrichment...")
    status_box = st.empty()
    results    = []  # (sheet_row, domain, desc, industry, icp, conf)
    errors     = []  # (sheet_row, domain, reason)

    for idx, (sheet_row, domain) in enumerate(candidates):
        pct = idx / len(candidates)
        progress.progress(pct, text=f"Fetching {domain}... ({idx + 1}/{len(candidates)})")
        status_box.markdown(
            f"<span style='font-size:0.82rem;color:#555;'>"
            f"**{idx + 1}/{len(candidates)}** — `{domain}`</span>",
            unsafe_allow_html=True,
        )

        text = _enrich_fetch_text(domain)
        if not text:
            errors.append((sheet_row, domain, "Homepage fetch failed or returned no text"))
            _time.sleep(0.5)
            continue

        result = _enrich_classify(client, text)
        if result["error"]:
            errors.append((sheet_row, domain, result["error"]))
            _time.sleep(0.5)
            continue

        results.append((
            sheet_row, domain,
            result["website_description"],
            result["industry"],
            result["icp_label"],
            result["icp_confidence"],
        ))
        _time.sleep(1.0)  # ~1 req/sec -- well under Haiku rate limits

    progress.progress(1.0, text="Writing results to Sheet...")
    status_box.empty()

    # ── 4. Batch-write successes to master_companies ───────────────────────
    if results:
        updates = []
        for (sheet_row, domain, desc, industry, icp, conf) in results:
            for field, value in [
                ("website_description", desc),
                ("industry",            industry),
                ("icp_label",           icp),
                ("icp_confidence",      conf),
                ("enriched",            "True"),
                ("enrichment_date",     today_str()),
            ]:
                if field in col:
                    updates.append((sheet_row, col[field] + 1, value))
        if updates:
            _batch_update(ws, updates)

    # ── 5. Log errors to enrichment_errors tab ────────────────────────────
    if errors:
        error_rows = [
            [domain, "", today_str(), reason]
            for (_, domain, reason) in errors
        ]
        batch_write(sh, SHEET_ERRORS, error_rows)

    # ── 6. Invalidate Watch List cache ────────────────────────────────────
    st.session_state["wl_cache_dirty"] = True

    # ── 7. Summary ──────────────────────────────────────────────────────────
    st.success(
        f"Enrichment complete: **{len(results):,}** enriched, "
        f"**{len(errors):,}** failed."
    )
    if results:
        icp_counts: dict = {}
        for (_, _, _, _, icp, _) in results:
            icp_counts[icp] = icp_counts.get(icp, 0) + 1
        with st.expander("ICP distribution from this run"):
            for icp_val, cnt in sorted(icp_counts.items(), key=lambda x: -x[1]):
                st.markdown(f"`{icp_val:<30}` {cnt:>4}")
    if errors:
        with st.expander(f"{len(errors)} failed domains"):
            for (_, domain, reason) in errors[:50]:
                st.markdown(f"- `{domain}` — {reason}")
            if len(errors) > 50:
                st.caption(
                    f"... and {len(errors) - 50} more. "
                    "See enrichment_errors tab in Sheet."
                )
    if results:
        st.info(
            'Go to Watch List, enable "Enriched only" filter, '
            "and click Reload from Sheet to see results."
        )


# ─────────────────────────────────────────────────────────────────────────────
# TAB 1 -- UPLOAD
# ─────────────────────────────────────────────────────────────────────────────

def tab_upload(sh):
    st.header("Upload Contact & Order Data")
    st.caption("Accepts Brevo CSV exports, Shopify contact XLSX exports, and Shopify order CSV exports. "
               "File type is detected automatically from column headers.")

    if not PD_OK:
        st.error("**pandas** is not installed. Add pandas and openpyxl to requirements.txt.")
        return

    uploaded_files = st.file_uploader("Drop files here (CSV or XLSX)", type=["csv", "xlsx"],
                                       accept_multiple_files=True, key="upload_files")

    if not uploaded_files:
        st.info("No files selected yet. Drag one or more files above to begin.")
        st.markdown("---")
        st.markdown("**Data repair tools**")
        st.caption("Use these if customer status or spend data looks wrong after uploading orders.")
        if st.button("Backfill customer status from order history", key="backfill_customers"):
            _backfill_customers(sh)

        st.markdown("---")
        st.markdown("**Tier 1 strategic accounts**")
        st.caption(
            "Forces watch_tier=1 for ~47 known ICP-4/5 companies (TI, Qualcomm, "
            "Infineon, etc.) regardless of their engagement score. Run once after "
            "initial data load, or whenever you add domains to TIER1_OVERRIDE."
        )
        if st.button("Apply Tier 1 overrides", key="apply_tier1"):
            _apply_tier1_overrides(sh)

        st.markdown("---")
        st.markdown("**ICP enrichment**")
        st.caption(
            "Fetches each company's homepage, extracts visible text, and uses "
            "Claude Haiku to classify industry, ICP label, and confidence. "
            "Results write to master_companies. Errors log to enrichment_errors tab."
        )
        enrich_mode = st.radio(
            "Enrich which domains?",
            ["Flagged only (enrich = True)", "All Tier 1 (unenriched)", "All unenriched"],
            index=0,
            key="enrich_mode_radio",
            help=(
                "Flagged only: safe for targeted runs. "
                "All Tier 1: enriches all ~418 T1 domains (~70 min). "
                "All unenriched: runs on the full ~4,800 list -- may take 12+ hours."
            ),
        )
        _mode_map = {
            "Flagged only (enrich = True)": "flagged",
            "All Tier 1 (unenriched)": "tier1",
            "All unenriched": "all",
        }
        if st.button("Run enrichment", key="run_enrichment", type="primary"):
            _run_enrichment(sh, mode=_mode_map[enrich_mode])
        return

    st.write(f"**{len(uploaded_files)} file(s) ready to process:**")
    for f in uploaded_files:
        st.write(f"  - `{f.name}` ({f.size:,} bytes)")

    if not st.button("Process uploads", type="primary"):
        return

    existing_domains = sheet_get_domains(sh)
    session_results  = []
    progress         = st.progress(0.0, text="Starting...")

    for file_idx, f in enumerate(uploaded_files):
        pct_base = file_idx / len(uploaded_files)
        progress.progress(pct_base, text=f"Reading {f.name}...")
        df = read_uploaded_file(f)
        if df is None:
            session_results.append({"file": f.name, "new_added": 0, "existing_updated": 0,
                                     "orders_written": 0, "filtered": 0, "filter_reasons": {}})
            continue

        file_type = detect_file_type(df)
        if file_type is None:
            st.warning(f"`{f.name}` -- unrecognised column layout. Skipping.")
            session_results.append({"file": f.name, "new_added": 0, "existing_updated": 0,
                                     "orders_written": 0, "filtered": 0, "filter_reasons": {}})
            continue

        type_labels = {"brevo": "Brevo contacts", "shopify_contacts": "Shopify contacts",
                       "shopify_orders": "Shopify orders"}
        progress.progress(pct_base + 0.1 / len(uploaded_files),
                          text=f"Processing {type_labels[file_type]}: {f.name}...")

        if file_type in ("brevo", "shopify_contacts"):
            result      = (process_brevo if file_type == "brevo" else process_shopify_contacts)(df, existing_domains)
            file_result = _write_contact_results(sh, result, f.name, progress, pct_base, len(uploaded_files))
        else:
            result      = process_shopify_orders(df, existing_domains)
            file_result = _write_order_results(sh, result, f.name, progress, pct_base, len(uploaded_files))
            if file_result.get("orders_written", 0) > 0:
                progress.progress(1.0, text="Syncing customer status...")
                _backfill_customers(sh, silent=True)

        session_results.append(file_result)

    progress.progress(1.0, text="Done.")
    _render_upload_summary(session_results)


def _write_contact_results(sh, result, filename, progress, pct_base, n_files):
    new_master, new_contacts, update_map = result["new_master_rows"], result["new_contact_rows"], result["update_map"]
    new_added = 0
    if new_master:
        progress.progress(pct_base + 0.4 / n_files, text=f"Writing {len(new_master):,} new companies...")
        if batch_write(sh, SHEET_MASTER, [_master_row_to_list(r) for r in new_master]):
            new_added = len(new_master)
    if new_contacts:
        progress.progress(pct_base + 0.6 / n_files, text=f"Writing {len(new_contacts):,} new_contacts rows...")
        batch_write(sh, SHEET_NEW_CONTACTS, new_contacts)

    existing_updated = 0
    if update_map and GSPREAD_OK:
        try:
            ws, col, dom_idx = _sheet_index(sh)
            updates_needed   = []
            for domain, updates in update_map.items():
                if domain not in dom_idx: continue
                r = dom_idx[domain]
                for field, value in updates.items():
                    if field in col: updates_needed.append((r + 1, col[field] + 1, value))
                existing_updated += 1
            if updates_needed: _batch_update(ws, updates_needed)
        except Exception as ex:
            st.warning(f"Contact update error: {ex}")

    return {"file": filename, "new_added": new_added, "existing_updated": existing_updated,
            "orders_written": 0, "filtered": result["filtered"], "filter_reasons": result["filter_reasons"]}


def _write_order_results(sh, result, filename, progress, pct_base, n_files):
    order_rows, domain_totals = result["order_rows"], result["domain_totals"]
    orders_written = 0
    if order_rows:
        progress.progress(pct_base + 0.4 / n_files, text=f"Writing {len(order_rows):,} order line items...")
        if batch_write(sh, SHEET_ORDERS, order_rows): orders_written = len(order_rows)

    if domain_totals and GSPREAD_OK:
        try:
            ws, col, dom_idx = _sheet_index(sh)
            updates_needed   = []
            for domain, stats in domain_totals.items():
                if domain not in dom_idx: continue
                r     = dom_idx[domain]
                spent = _sanitize_value(round(stats["total_spent"], 2))
                for field, value in [("total_spent", spent), ("total_orders", stats["total_orders"]),
                                     ("has_purchases", "True"), ("customer_status", "customer"),
                                     ("in_shopify", "True"), ("last_activity", now_str())]:
                    if field in col: updates_needed.append((r + 1, col[field] + 1, value))
                if stats["email"] and "best_contact_email" in col:
                    updates_needed.append((r + 1, col["best_contact_email"] + 1, _sanitize_value(stats["email"])))
                if stats["company"] and "company_name" in col:
                    updates_needed.append((r + 1, col["company_name"] + 1, _sanitize_value(stats["company"])))
            if updates_needed: _batch_update(ws, updates_needed)

            try:
                nc_ws   = sh.worksheet(SHEET_NEW_CONTACTS)
                nc_vals = nc_ws.get_all_values()
                if nc_vals:
                    nc_hdrs       = nc_vals[0]
                    nc_dom_col    = nc_hdrs.index("domain") if "domain" in nc_hdrs else 0
                    nc_status_col = nc_hdrs.index("customer_status") + 1 if "customer_status" in nc_hdrs else None
                    nc_spent_col  = nc_hdrs.index("total_spent") + 1 if "total_spent" in nc_hdrs else None
                    nc_orders_col = nc_hdrs.index("total_orders") + 1 if "total_orders" in nc_hdrs else None
                    nc_updates    = []
                    for i, row_vals in enumerate(nc_vals[1:], start=2):
                        d = row_vals[nc_dom_col]
                        if d in domain_totals:
                            s = domain_totals[d]
                            if nc_status_col: nc_updates.append((i, nc_status_col, "customer"))
                            if nc_spent_col:  nc_updates.append((i, nc_spent_col, round(s["total_spent"], 2)))
                            if nc_orders_col: nc_updates.append((i, nc_orders_col, s["total_orders"]))
                    if nc_updates: _batch_update(nc_ws, nc_updates)
            except Exception as ex:
                st.warning(f"new_contacts customer sync error: {ex}")
        except Exception as ex:
            st.warning(f"Spend update error: {ex}")

    return {"file": filename, "new_added": 0, "existing_updated": len(domain_totals),
            "orders_written": orders_written, "filtered": result["filtered"],
            "filter_reasons": result["filter_reasons"]}


# ─────────────────────────────────────────────────────────────────────────────
# TAB 2 -- NEW CONTACTS
# ─────────────────────────────────────────────────────────────────────────────

def _nc_class_badge(cls):
    color = CLASS_COLOR.get(cls, "#546E7A")
    label = CLASS_LABEL.get(cls, cls.title())
    return (f'<span style="background:{color};color:#fff;padding:2px 8px;'
            f'border-radius:4px;font-size:0.78rem;font-weight:600;letter-spacing:0.03em;">{label}</span>')

def _nc_write_master_flags(sh, domain, monitor_val, enrich_val):
    try:
        ws, col, dom_idx = _sheet_index(sh)
        if domain not in dom_idx: return
        r       = dom_idx[domain]
        updates = []
        if "monitor" in col: updates.append((r + 1, col["monitor"] + 1, str(monitor_val)))
        if "enrich" in col:  updates.append((r + 1, col["enrich"] + 1, str(enrich_val)))
        if updates: _batch_update(ws, updates)
    except Exception as ex:
        st.warning(f"Could not update master_companies for {domain}: {ex}")

def _nc_write_nc_flags(sh, nc_rows_data, domain, monitor_val, enrich_val):
    try:
        ws          = sh.worksheet(SHEET_NEW_CONTACTS)
        all_vals    = ws.get_all_values()
        if not all_vals: return
        headers     = all_vals[0]
        mon_col     = headers.index("monitor") + 1 if "monitor" in headers else None
        enr_col     = headers.index("enrich") + 1  if "enrich" in headers  else None
        dom_col     = headers.index("domain") + 1  if "domain" in headers  else 1
        sheet_title = ws.title
        updates     = []
        for i, row_vals in enumerate(all_vals[1:], start=2):
            if row_vals[dom_col - 1] == domain:
                if mon_col: updates.append((i, mon_col, str(monitor_val)))
                if enr_col: updates.append((i, enr_col, str(enrich_val)))
                break
        if updates:
            body = {"data": [{"range": f"'{sheet_title}'!{gspread.utils.rowcol_to_a1(r, c)}",
                               "values": [[_sanitize_value(v)]]} for r, c, v in updates],
                    "valueInputOption": "RAW"}
            ws.spreadsheet.values_batch_update(body)
    except Exception as ex:
        st.warning(f"Could not update new_contacts flags for {domain}: {ex}")

def _nc_mark_reviewed(sh, domains):
    try:
        ws          = sh.worksheet(SHEET_NEW_CONTACTS)
        all_vals    = ws.get_all_values()
        if not all_vals: return False
        headers     = all_vals[0]
        rev_col     = headers.index("reviewed") + 1 if "reviewed" in headers else None
        dom_col     = headers.index("domain") + 1   if "domain" in headers   else 1
        if not rev_col: return False
        sheet_title = ws.title
        domain_set  = set(domains)
        updates     = [(i, rev_col, "True") for i, row_vals in enumerate(all_vals[1:], start=2)
                       if row_vals[dom_col - 1] in domain_set]
        if updates:
            body = {"data": [{"range": f"'{sheet_title}'!{gspread.utils.rowcol_to_a1(r, c)}",
                               "values": [[v]]} for r, c, v in updates],
                    "valueInputOption": "RAW"}
            ws.spreadsheet.values_batch_update(body)
        return True
    except Exception as ex:
        st.warning(f"Mark-reviewed error: {ex}")
        return False

def _nc_safe_str(v):
    if v is None: return ""
    s = str(v).strip()
    tokens = s.lower().split()
    if tokens and all(t == "nan" for t in tokens): return ""
    return s

def _nc_safe_bool(v, default=False):
    if isinstance(v, bool): return v
    return str(v).strip().upper() == "TRUE"

def _nc_strip_trailing_nan(s):
    import re as _re
    s = _re.sub(r'(?i)\s*\bnan\b\s*$', '', s).strip()
    s = _re.sub(r'(?i)^\s*\bnan\b\s*', '', s).strip()
    return _re.sub(r'  +', ' ', s).strip()

def _nc_infer_name_from_email(email):
    if not email or "@" not in email: return ""
    import re
    local = email.split("@")[0].lower()
    parts = [p for p in re.split(r'[._\-0-9]+', local) if len(p) > 1]
    if not parts or all(len(p) <= 1 for p in parts): return ""
    return " ".join(p.title() for p in parts[:2])

def _nc_completeness(p):
    score = 0
    if p.get("contact_n_real"): score += 1
    if p.get("contact_e"):      score += 1
    if p.get("contact_title"):  score += 1
    if p.get("company_real"):   score += 1
    if p.get("tags"):           score += 1
    if p.get("spent_raw", 0) > 0: score += 1
    return score

def _nc_completeness_bar(score, max_score=6):
    filled = "<span style='color:#0066CC;font-size:0.85rem;'>&#9679;</span>"
    empty  = "<span style='color:#CCC;font-size:0.85rem;'>&#9679;</span>"
    return "".join(filled if i < score else empty for i in range(max_score))

def _nc_load(sh):
    if "nc_rows_cache" not in st.session_state or st.session_state.get("nc_cache_dirty"):
        with st.spinner("Loading contacts from Sheet..."):
            rows = sheet_get_all(sh, SHEET_NEW_CONTACTS)
        st.session_state["nc_rows_cache"] = rows
        st.session_state["nc_cache_dirty"] = False
    return st.session_state["nc_rows_cache"]


def tab_new_contacts(sh):
    st.header("New Contacts")
    st.caption("All companies in the system -- filter to unreviewed, customers, or by segment. "
               "Set monitor and enrich flags, then mark reviewed to clear the queue.")

    all_rows   = _nc_load(sh)
    total_rows = len(all_rows)
    if not all_rows:
        st.info("No contacts loaded yet. Upload a file on the Upload tab first.")
        return

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
        contact_n_raw = _nc_strip_trailing_nan(_nc_safe_str(r.get("best_contact_name")))
        contact_n_real = bool(contact_n_raw)
        contact_n     = contact_n_raw or _nc_infer_name_from_email(contact_e)
        try:   spent_raw = float(str(r.get("total_spent") or 0))
        except: spent_raw = 0.0
        try:   orders_raw = int(float(str(r.get("total_orders") or 0)))
        except: orders_raw = 0

        p = {
            "domain": domain, "company": company, "company_real": company_real,
            "cls": cls, "status": status, "spent_raw": spent_raw, "orders_raw": orders_raw,
            "tags": tags, "has_event": _nc_safe_str(r.get("has_event_tag")).strip() == "True",
            "contact_n": contact_n, "contact_n_real": contact_n_real,
            "contact_e": contact_e, "contact_title": contact_title,
            "first_seen": _nc_safe_str(r.get("first_seen")), "added_date": _nc_safe_str(r.get("added_date")),
            "monitor_cur": _nc_safe_bool(r.get("monitor", "True"), default=True),
            "enrich_cur":  _nc_safe_bool(r.get("enrich", "False"), default=False),
            "reviewed":    _nc_safe_bool(r.get("reviewed"), default=False),
            "_raw": r,
        }
        p["completeness"] = _nc_completeness(p)
        prepped.append(p)

    available_classes = sorted({p["cls"] for p in prepped if p["cls"]})

    with st.sidebar:
        st.markdown("### Filter Contacts")
        show_filter = st.radio("Show",
            ["Unreviewed only", "All contacts", "Customers", "Reviewed only"], index=0, key="nc_show_filter")
        status_filter = st.radio("Customer status", ["All", "Customers", "Prospects"], index=0, key="nc_status_filter")
        class_filter = st.multiselect("Segment", options=available_classes, default=[],
                                       placeholder="All segments", key="nc_class_filter")
        has_purchase_only = st.checkbox("Has purchases", value=False, key="nc_has_purchase")
        has_event_only    = st.checkbox("Has event tag", value=False, key="nc_has_event")
        st.markdown("---")
        sort_by = st.selectbox("Sort by",
            ["Data completeness (best first)", "Date added (newest)", "Date added (oldest)",
             "Company name (A-Z)", "Total spend (high-low)"], key="nc_sort")
        st.markdown("---")
        if st.button("Reload from Sheet", key="nc_reload"):
            st.session_state["nc_cache_dirty"] = True
            st.rerun()
        sheet_id = get_secret("GSHEET_ID")
        if sheet_id:
            st.markdown(f'<a href="https://docs.google.com/spreadsheets/d/{sheet_id}" target="_blank" '
                        f'style="font-size:0.82rem;color:#0066CC;text-decoration:none;">&#128196; Edit Sheet</a>',
                        unsafe_allow_html=True)

    view = prepped
    if show_filter == "Unreviewed only":    view = [p for p in view if not p["reviewed"]]
    elif show_filter == "Customers":        view = [p for p in view if p["status"] == "customer"]
    elif show_filter == "Reviewed only":    view = [p for p in view if p["reviewed"]]
    if status_filter == "Customers":        view = [p for p in view if p["status"] == "customer"]
    elif status_filter == "Prospects":      view = [p for p in view if p["status"] != "customer"]
    if class_filter:                        view = [p for p in view if p["cls"] in class_filter]
    if has_purchase_only:                   view = [p for p in view if p["spent_raw"] > 0]
    if has_event_only:                      view = [p for p in view if p["has_event"]]

    if sort_by == "Data completeness (best first)": view = sorted(view, key=lambda p: -p["completeness"])
    elif sort_by == "Date added (oldest)":          view = sorted(view, key=lambda p: p["added_date"] or p["first_seen"])
    elif sort_by == "Company name (A-Z)":           view = sorted(view, key=lambda p: p["company"].lower())
    elif sort_by == "Total spend (high-low)":       view = sorted(view, key=lambda p: -p["spent_raw"])

    n_unreviewed = sum(1 for p in prepped if not p["reviewed"])
    n_customers  = sum(1 for p in prepped if p["status"] == "customer")
    n_with_spend = sum(1 for p in prepped if p["spent_raw"] > 0)
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total in system", f"{total_rows:,}")
    m2.metric("Unreviewed",      f"{n_unreviewed:,}")
    m3.metric("Customers",       f"{n_customers:,}")
    m4.metric("With purchases",  f"{n_with_spend:,}")
    m5.metric("Showing now",     f"{len(view):,}")
    st.markdown("---")

    if not view:
        st.info("No contacts match the current filters.")
        return

    unreviewed_in_view = [p for p in view if not p["reviewed"]]
    if unreviewed_in_view:
        if st.button(f"Mark all {len(unreviewed_in_view):,} shown as reviewed", key="nc_mark_all", type="secondary"):
            domains_to_clear = [p["domain"] for p in unreviewed_in_view if p["domain"]]
            with st.spinner(f"Marking {len(domains_to_clear):,} rows reviewed..."):
                ok = _nc_mark_reviewed(sh, domains_to_clear)
            if ok:
                st.success(f"Marked {len(domains_to_clear):,} contacts reviewed.")
                st.session_state["nc_cache_dirty"] = True
                st.rerun()

    filter_fp = f"{show_filter}|{status_filter}|{sorted(class_filter)}|{has_purchase_only}|{has_event_only}|{sort_by}"
    if st.session_state.get("nc_filter_fp") != filter_fp:
        st.session_state["nc_filter_fp"] = filter_fp
        st.session_state["nc_page"] = 0

    total_pages = max(1, math.ceil(len(view) / NC_PAGE_SIZE))
    page        = min(st.session_state.get("nc_page", 0), total_pages - 1)
    st.session_state["nc_page"] = page
    start_idx   = page * NC_PAGE_SIZE
    end_idx     = start_idx + NC_PAGE_SIZE
    page_view   = view[start_idx:end_idx]

    st.markdown(f"**{len(view):,} contact(s)** - sorted by {sort_by.lower()} - "
                f"page {page + 1} of {total_pages} ({start_idx + 1}-{min(end_idx, len(view))} of {len(view):,})")
    st.markdown("")

    legend_html = " &nbsp; ".join(
        f'<span style="background:{CLASS_COLOR[c]};color:#fff;padding:1px 7px;border-radius:3px;font-size:0.73rem;">{CLASS_LABEL[c]}</span>'
        for c in available_classes if c in CLASS_COLOR)
    st.markdown(legend_html, unsafe_allow_html=True)
    st.markdown("")

    for idx, p in enumerate(page_view):
        abs_idx = start_idx + idx
        domain, company, cls, status = p["domain"], p["company"], p["cls"], p["status"]
        spent_raw, orders_raw = p["spent_raw"], p["orders_raw"]
        tags, has_event = p["tags"], p["has_event"]
        contact_n, contact_n_real, contact_e, contact_title = p["contact_n"], p["contact_n_real"], p["contact_e"], p["contact_title"]
        first_seen, added_date = p["first_seen"], p["added_date"]
        monitor_cur, enrich_cur, reviewed = p["monitor_cur"], p["enrich_cur"], p["reviewed"]

        cls_dot = (f'<span style="display:inline-block;width:10px;height:10px;border-radius:50%;'
                   f'background:{CLASS_COLOR.get(cls, "#546E7A")};margin-right:4px;"></span>')
        spend_pill = (f' <span style="background:#1B5E20;color:#fff;padding:1px 6px;border-radius:3px;'
                      f'font-size:0.72rem;font-weight:600;">{format_currency(spent_raw)}</span>') if spent_raw > 0 else ""
        event_pill = (' <span style="background:#E65100;color:#fff;padding:1px 6px;border-radius:3px;'
                      'font-size:0.72rem;">event</span>') if has_event else ""
        mon_indicator = " 📡" if monitor_cur else ""
        enr_indicator = " 🔬" if enrich_cur else ""
        rev_indicator = " ✓"  if reviewed    else ""
        completeness_bar = _nc_completeness_bar(p["completeness"])

        label_html = (f"{cls_dot}<b>{company}</b> &nbsp;"
                      f"<span style='color:#888;font-size:0.85rem;'>{domain}</span>"
                      f"{spend_pill}{event_pill} &nbsp; {completeness_bar}"
                      f"<span style='color:#0066CC;font-size:0.8rem;'>{mon_indicator}</span>"
                      f"<span style='color:#6A1B9A;font-size:0.8rem;'>{enr_indicator}</span>"
                      f"<span style='color:#2E7D32;font-size:0.8rem;'>{rev_indicator}</span>")
        st.markdown(f'<div style="margin-bottom:-10px;padding:4px 2px;font-size:0.88rem;">{label_html}</div>',
                    unsafe_allow_html=True)

        with st.expander(f"{company}  .  {domain}", expanded=False):
            badge_html  = _nc_class_badge(cls)
            status_icon = "🟢" if status == "customer" else "⚪"
            rev_badge   = ('<span style="background:#4CAF50;color:#fff;padding:1px 7px;border-radius:3px;'
                           'font-size:0.75rem;margin-left:8px;">reviewed</span>') if reviewed else ""
            st.markdown(f"{badge_html} &nbsp; {status_icon} &nbsp; "
                        f"<span style='font-size:0.85rem;color:#555;'>{status.title()}</span>{rev_badge}",
                        unsafe_allow_html=True)
            st.markdown("")
            d1, d2, d3 = st.columns(3)
            with d1:
                st.markdown("**Domain**")
                st.code(domain, language=None)
                st.markdown(f"**Added:** {first_seen or added_date or '-'}")
            with d2:
                st.markdown("**Best contact**")
                st.write(contact_n if contact_n_real else f"{contact_n} *(from email)*" if contact_n else "-")
                if contact_title: st.caption(contact_title)
                if contact_e:     st.caption(contact_e)
            with d3:
                st.markdown("**Purchase data**")
                st.write(f"{format_currency(spent_raw)} across {orders_raw} order(s)" if orders_raw > 0 else "No purchases on record")
                if tags:      st.caption(f"Tags: {', '.join(t.strip() for t in tags.split(',') if t.strip())}")
                if has_event: st.caption("Has event tag")
            st.markdown("")
            t1, t2, t3 = st.columns([1, 1, 2])
            with t1:
                new_monitor = st.checkbox("Monitor", value=monitor_cur, key=f"nc_mon_{abs_idx}_{domain}",
                                          help="Add to the active signal-monitoring queue.")
            with t2:
                new_enrich = st.checkbox("Enrich", value=enrich_cur, key=f"nc_enr_{abs_idx}_{domain}",
                                          help="Queue for homepage scrape + ICP classification.")
            if (new_monitor != monitor_cur) or (new_enrich != enrich_cur):
                with st.spinner("Saving flags..."):
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
                    if st.button("Mark reviewed", key=f"nc_rev_{abs_idx}_{domain}", type="primary"):
                        with st.spinner("Marking reviewed..."):
                            _nc_mark_reviewed(sh, [domain])
                        st.session_state["nc_cache_dirty"] = True
                        st.rerun()
                else:
                    st.caption("Already reviewed")

    st.markdown("")
    st.markdown("---")
    nav_l, nav_mid, nav_r = st.columns([1, 3, 1])
    with nav_l:
        if page > 0:
            if st.button("Previous", key="nc_prev"):
                st.session_state["nc_page"] = page - 1
                st.rerun()
    with nav_mid:
        st.markdown(f"<div style='text-align:center;color:#666;font-size:0.85rem;'>"
                    f"Page {page + 1} of {total_pages} - "
                    f"{start_idx + 1}-{min(end_idx, len(view))} of {len(view):,} contacts</div>",
                    unsafe_allow_html=True)
    with nav_r:
        if page < total_pages - 1:
            if st.button("Next", key="nc_next"):
                st.session_state["nc_page"] = page + 1
                st.rerun()


# ─────────────────────────────────────────────────────────────────────────────
# TAB 3 -- WATCH LIST
# ─────────────────────────────────────────────────────────────────────────────

def _wl_render_sidebar(available_classes):
    """
    Render Watch List sidebar. Called inside tab_watch_list so it is scoped
    to the Watch List tab context and does not bleed into New Contacts sidebar.
    """
    with st.sidebar:
        st.markdown("### Filter Watch List")

        st.multiselect("Segment", options=available_classes,
                        default=st.session_state.get("wl_class_filter", []),
                        placeholder="All segments", key="wl_class_filter")

        tier_opts = ["All", "Tier 1", "Tier 2", "Tier 3"]
        st.radio("Watch tier", tier_opts,
                  index=tier_opts.index(st.session_state.get("wl_tier_filter", "All")),
                  key="wl_tier_filter")

        stat_opts = ["All", "Customers", "Prospects"]
        st.radio("Customer status", stat_opts,
                  index=stat_opts.index(st.session_state.get("wl_status_filter", "All")),
                  key="wl_status_filter")

        st.checkbox("Has purchases",        value=st.session_state.get("wl_has_purchase", False), key="wl_has_purchase")
        st.checkbox("Monitor flagged",      value=st.session_state.get("wl_monitor_only", False), key="wl_monitor_only")
        st.checkbox("Enriched only",        value=st.session_state.get("wl_enriched_only", False), key="wl_enriched_only")
        st.checkbox("Exclude distributors", value=st.session_state.get("wl_excl_dist", True),     key="wl_excl_dist")

        st.markdown("---")
        sort_opts = ["Watch tier (1 first)", "Total spend (high-low)", "Company (A-Z)", "Last activity (newest)"]
        cur_sort  = st.session_state.get("wl_sort", "Watch tier (1 first)")
        st.selectbox("Sort by", sort_opts,
                      index=sort_opts.index(cur_sort) if cur_sort in sort_opts else 0,
                      key="wl_sort")

        st.markdown("---")
        if st.button("Reload from Sheet", key="wl_reload"):
            st.session_state["wl_cache_dirty"] = True
            st.rerun()
        sheet_id = get_secret("GSHEET_ID")
        if sheet_id:
            st.markdown(f'<a href="https://docs.google.com/spreadsheets/d/{sheet_id}" target="_blank" '
                        f'style="font-size:0.82rem;color:#0066CC;text-decoration:none;">&#128196; Edit Sheet</a>',
                        unsafe_allow_html=True)


def _wl_load(sh):
    if "wl_rows_cache" not in st.session_state or st.session_state.get("wl_cache_dirty"):
        with st.spinner("Loading Watch List from Sheet..."):
            rows = sheet_get_all(sh, SHEET_MASTER)
        st.session_state["wl_rows_cache"] = rows
        st.session_state["wl_cache_dirty"] = False
    return st.session_state["wl_rows_cache"]

def _wl_safe_str(v):
    if v is None: return ""
    s = str(v).strip()
    tokens = s.lower().split()
    if tokens and all(t == "nan" for t in tokens): return ""
    return s

def _wl_safe_bool(v, default=False):
    if isinstance(v, bool): return v
    return str(v).strip().upper() == "TRUE"

def _wl_safe_float(v):
    try:   return float(str(v or 0))
    except: return 0.0

def _wl_safe_int(v):
    try:   return int(float(str(v or 0)))
    except: return 0

def _wl_strip_nan(s):
    """Remove nan tokens from name strings -- artifact of pandas NaN bleed."""
    import re as _re
    s = _re.sub(r"(?i)\s*\bnan\b\s*$", "", s).strip()
    s = _re.sub(r"(?i)^\s*\bnan\b\s*", "", s).strip()
    return _re.sub(r"  +", " ", s).strip()

def _wl_infer_name(email):
    """Best-guess display name from email local-part."""
    if not email or "@" not in email: return ""
    import re
    local = email.split("@")[0].lower()
    parts = [p for p in re.split(r"[._\-0-9]+", local) if len(p) > 1]
    if not parts or all(len(p) <= 1 for p in parts): return ""
    return " ".join(p.title() for p in parts[:2])

def _wl_name_matches_email(name, email):
    """
    Return True if at least one name token (3+ chars) appears in the email local-part.
    Catches mismatches like 'Niu, Simiao nan' vs 'kevin.wine@rutgers.edu'.
    """
    if not name or not email or "@" not in email: return False
    import re
    local  = email.split("@")[0].lower()
    tokens = [t.lower() for t in re.split(r"[\s,._\-]+", name) if len(t) >= 3]
    return any(tok in local for tok in tokens)

def _wl_tier_color(tier):
    return {1: "#B71C1C", 2: "#E65100", 3: "#546E7A"}.get(tier, "#9E9E9E")

def _wl_tier_badge(tier):
    color = {1: "#B71C1C", 2: "#E65100", 3: "#546E7A"}.get(tier, "#9E9E9E")
    label = {1: "Tier 1",  2: "Tier 2",  3: "Tier 3"}.get(tier, f"Tier {tier}")
    return (f'<span style="background:{color};color:#fff;padding:2px 8px;'
            f'border-radius:4px;font-size:0.78rem;font-weight:600;">{label}</span>')

def _wl_completeness(p):
    score = 0
    if p.get("contact_n"): score += 1
    if p.get("contact_e"): score += 1
    if p.get("contact_t"): score += 1
    if p.get("company"):   score += 1
    if p.get("icp_label"): score += 1
    if p.get("spent", 0) > 0: score += 1
    return score

def _wl_completeness_bar(score, max_score=6):
    filled = "<span style='color:#0066CC;font-size:0.85rem;'>&#9679;</span>"
    empty  = "<span style='color:#CCC;font-size:0.85rem;'>&#9679;</span>"
    return "".join(filled if i < score else empty for i in range(max_score))


def tab_watch_list(sh):
    st.header("Watch List")
    st.caption("All companies in the master list -- filter by segment, tier, and status. "
               "Distributors excluded by default. Expand any row to view the full profile and sales history.")

    all_rows = _wl_load(sh)
    if not all_rows:
        st.info("No data loaded. Upload contacts or orders on the Upload tab first.")
        return

    prepped = []
    for r in all_rows:
        domain  = _wl_safe_str(r.get("domain"))
        company = _wl_safe_str(r.get("company_name")) or domain
        cls     = _wl_safe_str(r.get("domain_class")) or "commercial"
        status  = _wl_safe_str(r.get("customer_status")) or "prospect"
        try:   tier = int(float(str(r.get("watch_tier") or 2)))
        except: tier = 2
        if tier not in (1, 2, 3): tier = 2

        spent     = _wl_safe_float(r.get("total_spent"))
        orders    = _wl_safe_int(r.get("total_orders"))
        monitor   = _wl_safe_bool(r.get("monitor"), default=True)
        enrich    = _wl_safe_bool(r.get("enrich"),  default=False)

        contact_e     = _wl_safe_str(r.get("best_contact_email"))
        contact_t     = _wl_safe_str(r.get("best_contact_title"))
        contact_n_raw = _wl_strip_nan(_wl_safe_str(r.get("best_contact_name")))
        if contact_n_raw and contact_e and not _wl_name_matches_email(contact_n_raw, contact_e):
            contact_n = _wl_infer_name(contact_e) or contact_n_raw
        else:
            contact_n = contact_n_raw or _wl_infer_name(contact_e)

        icp_label = _wl_safe_str(r.get("icp_label"))
        icp_conf  = _wl_safe_str(r.get("icp_confidence"))
        enriched  = _wl_safe_bool(r.get("enriched"), default=False)
        notes     = _wl_safe_str(r.get("notes"))
        last_act  = _wl_safe_str(r.get("last_activity"))

        p = {"domain": domain, "company": company, "cls": cls, "status": status,
             "tier": tier, "spent": spent, "orders": orders, "monitor": monitor, "enrich": enrich,
             "contact_n": contact_n, "contact_e": contact_e, "contact_t": contact_t,
             "icp_label": icp_label, "icp_conf": icp_conf, "enriched": enriched,
             "notes": notes, "last_act": last_act, "_raw": r}
        p["completeness"] = _wl_completeness(p)
        prepped.append(p)

    available_classes = sorted({p["cls"] for p in prepped if p["cls"]})

    _wl_render_sidebar(available_classes)

    class_filter      = st.session_state.get("wl_class_filter", [])
    tier_filter       = st.session_state.get("wl_tier_filter", "All")
    status_filter     = st.session_state.get("wl_status_filter", "All")
    has_purchase_only = st.session_state.get("wl_has_purchase", False)
    monitor_only      = st.session_state.get("wl_monitor_only", False)
    enriched_only     = st.session_state.get("wl_enriched_only", False)
    excl_dist         = st.session_state.get("wl_excl_dist", True)
    sort_by           = st.session_state.get("wl_sort", "Watch tier (1 first)")

    view = prepped
    if excl_dist:     view = [p for p in view if p["cls"] != "distributor"]
    if class_filter:  view = [p for p in view if p["cls"] in class_filter]
    if tier_filter != "All":
        try:
            tier_num = int(tier_filter.split()[1])
            view = [p for p in view if p["tier"] == tier_num]
        except (IndexError, ValueError):
            pass
    if status_filter == "Customers":   view = [p for p in view if p["status"] == "customer"]
    elif status_filter == "Prospects": view = [p for p in view if p["status"] != "customer"]
    if has_purchase_only: view = [p for p in view if p["spent"] > 0]
    if monitor_only:      view = [p for p in view if p["monitor"]]
    if enriched_only:     view = [p for p in view if p["enriched"]]

    if sort_by == "Watch tier (1 first)":    view = sorted(view, key=lambda p: (p["tier"], -p["spent"]))
    elif sort_by == "Company (A-Z)":         view = sorted(view, key=lambda p: p["company"].lower())
    elif sort_by == "Last activity (newest)": view = sorted(view, key=lambda p: p["last_act"] or "", reverse=True)
    else:                                     view = sorted(view, key=lambda p: -p["spent"])

    n_tier1    = sum(1 for p in prepped if p["tier"] == 1)
    n_customer = sum(1 for p in prepped if p["status"] == "customer")
    n_purchase = sum(1 for p in prepped if p["spent"] > 0)
    total_rev  = sum(p["spent"] for p in prepped)

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Total companies", f"{len(prepped):,}")
    m2.metric("Tier 1 Active",   f"{n_tier1:,}")
    m3.metric("Customers",       f"{n_customer:,}")
    m4.metric("With purchases",  f"{n_purchase:,}")
    m5.metric("Total revenue",   format_currency(total_rev))

    st.markdown("---")

    if not view:
        st.info("No companies match the current filters.")
        return

    t_counts = {t: sum(1 for p in view if p["tier"] == t) for t in [1, 2, 3]}
    active_qt = st.session_state.get("wl_quick_tier")
    qcols = st.columns(4)
    for i, (lbl, val) in enumerate([("All tiers", None), ("Tier 1", 1), ("Tier 2", 2), ("Tier 3", 3)]):
        cnt = len(view) if val is None else t_counts.get(val, 0)
        with qcols[i]:
            if st.button(f"{lbl}  {cnt:,}", key=f"wl_qt_{i}",
                          type="primary" if active_qt == val else "secondary",
                          use_container_width=True):
                st.session_state["wl_quick_tier"] = val
                st.session_state["wl_page"] = 0
                st.rerun()

    qt = st.session_state.get("wl_quick_tier")
    if qt is not None:
        view = [p for p in view if p["tier"] == qt]

    filter_fp = (f"{sorted(class_filter)}|{tier_filter}|{status_filter}"
                 f"|{has_purchase_only}|{monitor_only}|{enriched_only}|{excl_dist}|{sort_by}|{qt}")
    if st.session_state.get("wl_filter_fp") != filter_fp:
        st.session_state["wl_filter_fp"] = filter_fp
        st.session_state["wl_page"] = 0

    total_pages = max(1, math.ceil(len(view) / NC_PAGE_SIZE))
    page        = min(st.session_state.get("wl_page", 0), total_pages - 1)
    st.session_state["wl_page"] = page
    start_idx   = page * NC_PAGE_SIZE
    end_idx     = start_idx + NC_PAGE_SIZE
    page_view   = view[start_idx:end_idx]

    st.markdown(f"**{len(view):,} compan{'y' if len(view) == 1 else 'ies'}** - "
                f"sorted by {sort_by.lower()} - page {page + 1} of {total_pages} "
                f"({start_idx + 1}-{min(end_idx, len(view))} of {len(view):,})")
    st.markdown("")

    legend_html = " &nbsp; ".join(
        f'<span style="background:{CLASS_COLOR[c]};color:#fff;padding:1px 7px;border-radius:3px;font-size:0.73rem;">{CLASS_LABEL[c]}</span>'
        for c in available_classes if c in CLASS_COLOR)
    st.markdown(legend_html, unsafe_allow_html=True)
    st.markdown("")

    for idx, p in enumerate(page_view):
        abs_idx = start_idx + idx
        domain, company, cls, tier = p["domain"], p["company"], p["cls"], p["tier"]
        spent, orders, monitor, enrich = p["spent"], p["orders"], p["monitor"], p["enrich"]

        cls_dot = (f'<span style="display:inline-block;width:10px;height:10px;border-radius:50%;'
                   f'background:{CLASS_COLOR.get(cls, "#546E7A")};margin-right:4px;"></span>')
        tier_color = _wl_tier_color(tier)
        tier_pill  = (f' <span style="background:{tier_color};color:#fff;padding:1px 6px;'
                      f'border-radius:3px;font-size:0.72rem;font-weight:600;">T{tier}</span>')
        spend_pill = (f' <span style="background:#1B5E20;color:#fff;padding:1px 6px;'
                      f'border-radius:3px;font-size:0.72rem;font-weight:600;">{format_currency(spent)}</span>') if spent > 0 else ""
        mon_indicator    = " 📡" if monitor else ""
        completeness_bar = _wl_completeness_bar(p["completeness"])

        label_html = (f"{cls_dot}<b>{company}</b> &nbsp;"
                      f"<span style='color:#888;font-size:0.85rem;'>{domain}</span>"
                      f"{tier_pill}{spend_pill} &nbsp; {completeness_bar}"
                      f"<span style='color:#0066CC;font-size:0.8rem;'>{mon_indicator}</span>")
        st.markdown(f'<div style="margin-bottom:-10px;padding:4px 2px;font-size:0.88rem;">{label_html}</div>',
                    unsafe_allow_html=True)

        with st.expander(f"{company}  .  {domain}", expanded=False):

            seg_color  = CLASS_COLOR.get(cls, "#546E7A")
            seg_label  = CLASS_LABEL.get(cls, cls.title())
            badge_html = (f'<span style="background:{seg_color};color:#fff;padding:2px 8px;'
                          f'border-radius:4px;font-size:0.78rem;font-weight:600;">{seg_label}</span>')
            status_icon = "🟢" if p["status"] == "customer" else "⚪"
            st.markdown(f"{badge_html} &nbsp; {_wl_tier_badge(tier)} &nbsp; {status_icon} &nbsp; "
                        f"<span style='font-size:0.85rem;color:#555;'>{p['status'].title()}</span>",
                        unsafe_allow_html=True)
            st.markdown("")

            d1, d2, d3 = st.columns(3)
            with d1:
                st.markdown("**Domain**")
                st.code(domain, language=None)
                if p["last_act"]: st.caption(f"Last activity: {p['last_act'][:10]}")
            with d2:
                st.markdown("**Best contact**")
                st.write(p["contact_n"] or "-")
                if p["contact_t"]: st.caption(p["contact_t"])
                if p["contact_e"]: st.caption(p["contact_e"])
            with d3:
                st.markdown("**ICP classification**")
                if p["icp_label"]:
                    conf_color = {"High": "#2E7D32", "Medium": "#E65100", "Low": "#B71C1C"}.get(p["icp_conf"], "#555")
                    st.markdown(f"{p['icp_label']} &nbsp; "
                                f'<span style="color:{conf_color};font-size:0.8rem;">{p["icp_conf"] or ""}</span>',
                                unsafe_allow_html=True)
                else:
                    st.write("Not yet enriched")

            st.markdown("")
            t1, t2 = st.columns([1, 1])
            with t1:
                new_monitor = st.checkbox("Monitor", value=monitor, key=f"wl_mon_{abs_idx}_{domain}",
                                          help="Add to active signal-monitoring queue.")
            with t2:
                new_enrich = st.checkbox("Enrich", value=enrich, key=f"wl_enr_{abs_idx}_{domain}",
                                          help="Queue for homepage scrape + ICP classification.")

            if (new_monitor != monitor) or (new_enrich != enrich):
                with st.spinner("Saving flags..."):
                    try:
                        ws, col, dom_idx_map = _sheet_index(sh)
                        if domain in dom_idx_map:
                            r_idx       = dom_idx_map[domain]
                            flag_updates = []
                            if "monitor" in col: flag_updates.append((r_idx + 1, col["monitor"] + 1, str(new_monitor)))
                            if "enrich" in col:  flag_updates.append((r_idx + 1, col["enrich"] + 1, str(new_enrich)))
                            if flag_updates: _batch_update(ws, flag_updates)
                    except Exception as ex:
                        st.warning(f"Could not save flags: {ex}")
                for cached_p in st.session_state.get("wl_rows_cache", []):
                    if cached_p.get("domain") == domain:
                        cached_p["monitor"] = str(new_monitor)
                        cached_p["enrich"]  = str(new_enrich)
                        break
                st.toast(f"Flags updated for {company}.", icon="✅")

            st.markdown("")
            new_notes = st.text_area("Notes", value=p["notes"], key=f"wl_notes_{abs_idx}_{domain}",
                                      height=72, placeholder="Add notes about this company...")
            if st.button("Save notes", key=f"wl_save_notes_{abs_idx}_{domain}"):
                with st.spinner("Saving..."):
                    try:
                        ws, col, dom_idx_map = _sheet_index(sh)
                        if domain in dom_idx_map and "notes" in col:
                            r_idx = dom_idx_map[domain]
                            _batch_update(ws, [(r_idx + 1, col["notes"] + 1, new_notes)])
                            for cached_p in st.session_state.get("wl_rows_cache", []):
                                if cached_p.get("domain") == domain:
                                    cached_p["notes"] = new_notes
                                    break
                            st.toast("Notes saved.", icon="✅")
                    except Exception as ex:
                        st.warning(f"Could not save notes: {ex}")

            st.markdown("---")
            st.markdown("**Sales History**")
            if orders > 0:
                c1, c2 = st.columns(2)
                c1.metric("Total spent",  format_currency(spent))
                c2.metric("Total orders", orders)
                with st.spinner("Loading order details..."):
                    order_data = [row for row in sheet_get_all(sh, SHEET_ORDERS)
                                  if _wl_safe_str(row.get("domain")) == domain]
                if order_data:
                    for o in order_data[:20]:
                        o_date    = _wl_safe_str(o.get("order_date"))[:10]
                        o_product = _wl_safe_str(o.get("product_name")) or "-"
                        o_price   = _wl_safe_float(o.get("line_item_price"))
                        o_qty     = _wl_safe_int(o.get("quantity"))
                        st.markdown(f"<span style='color:#888;font-size:0.8rem;'>{o_date}</span> &nbsp; "
                                    f"{o_product} &nbsp; "
                                    f"<span style='color:#2E7D32;font-size:0.8rem;'>x{o_qty} &nbsp; {format_currency(o_price)}</span>",
                                    unsafe_allow_html=True)
                    if len(order_data) > 20:
                        st.caption(f"... and {len(order_data) - 20} more line items. Open Sheet for full history.")
            else:
                st.caption("No purchases on record.")

            st.markdown("---")
            st.caption("Signal monitoring coming in Phase 3.")
            st.markdown("---")
            st.caption("Outreach generator coming in Phase 3.")

    st.markdown("")
    st.markdown("---")
    nav_l, nav_mid, nav_r = st.columns([1, 3, 1])
    with nav_l:
        if page > 0:
            if st.button("Previous", key="wl_prev"):
                st.session_state["wl_page"] = page - 1
                st.rerun()
    with nav_mid:
        st.markdown(f"<div style='text-align:center;color:#666;font-size:0.85rem;'>"
                    f"Page {page + 1} of {total_pages} - "
                    f"{start_idx + 1}-{min(end_idx, len(view))} of {len(view):,} companies</div>",
                    unsafe_allow_html=True)
    with nav_r:
        if page < total_pages - 1:
            if st.button("Next", key="wl_next"):
                st.session_state["wl_page"] = page + 1
                st.rerun()


def tab_ham(sh):
    st.header("HAM Radio")
    st.info("Coming soon -- HAM segment signal feed and outreach queue.")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    st.markdown(
        """
        <div style="background:#0066CC;padding:14px 24px 10px 24px;border-radius:6px;margin-bottom:20px;">
            <span style="color:#FFFFFF;font-size:1.5rem;font-weight:700;letter-spacing:0.02em;">
                UNI-T Customer Lead Dashboard
            </span>
            <span style="color:#B3D1F5;font-size:0.9rem;margin-left:16px;">
                Signal-driven B2B prospecting - North America
            </span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    sh = startup()
    if sh is None:
        st.error("**Google Sheets connection failed.** "
                 "Check that GOOGLE_SERVICE_ACCOUNT and GSHEET_ID are set in "
                 "Streamlit Cloud - Settings - Secrets.")
        st.stop()

    tabs = st.tabs(["Upload", "New Contacts", "Watch List", "HAM"])
    with tabs[0]: tab_upload(sh)
    with tabs[1]: tab_new_contacts(sh)
    with tabs[2]: tab_watch_list(sh)
    with tabs[3]: tab_ham(sh)


if __name__ == "__main__":
    main()
