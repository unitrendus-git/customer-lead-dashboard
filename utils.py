"""
utils.py — Customer Lead Dashboard
====================================
Shared foundation module. Every other script imports from here.
Handles:
  - Streamlit secrets access
  - Google Sheets connection + tab initialization
  - Common read/write helpers for all six CLD tabs
  - Anthropic client
  - Segment rules (domain_class → monitor/outreach behavior)

Deploy: Streamlit Cloud — secrets stored in app Settings → Secrets.
Never put credentials in code.
"""

import json
import datetime
import streamlit as st

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_OK = True
except ImportError:
    GSPREAD_OK = False

try:
    import anthropic
    ANTHROPIC_OK = True
except ImportError:
    ANTHROPIC_OK = False

# ─────────────────────────────────────────────────────────────────────────────
# SECRETS
# ─────────────────────────────────────────────────────────────────────────────

def get_secret(key: str, fallback: str = "") -> str:
    """Read a value from Streamlit secrets. Returns fallback if missing."""
    try:
        return st.secrets[key]
    except Exception:
        return fallback


# ─────────────────────────────────────────────────────────────────────────────
# SHEET TAB NAMES + HEADERS
# ─────────────────────────────────────────────────────────────────────────────

SHEET_MASTER       = "master_companies"
SHEET_NEW_CONTACTS = "new_contacts"
SHEET_ORDERS       = "order_history"
SHEET_SIGNALS      = "signals_log"
SHEET_VV           = "vv_log"
SHEET_OUTREACH     = "outreach_log"
SHEET_ERRORS       = "enrichment_errors"

MASTER_HEADERS = [
    "domain", "company_name", "domain_class", "distributor_parent",
    "customer_status", "watch_tier", "score", "score_domain",
    "score_contact", "score_engagement",
    "monitor", "enrich", "suppress_outreach",
    "brevo_contacts", "shopify_contacts",
    "total_spent", "total_orders",
    "best_contact_name", "best_contact_email", "best_contact_title",
    "category_of_interest", "tags", "has_event_tag",
    "in_brevo", "in_shopify", "has_purchases",
    "first_seen", "last_activity",
    "enriched", "enrichment_date",
    "website_description", "industry", "icp_label", "icp_confidence",
    "vv_last_visit", "vv_pages_visited",
    "last_signal_date", "last_signal_summary",
    "outreach_status", "outreach_last_date",
    "notes",
]

NEW_CONTACTS_HEADERS = [
    "domain", "company_name", "domain_class", "customer_status",
    "total_spent", "total_orders",
    "brevo_contacts", "tags", "has_event_tag",
    "best_contact_name", "best_contact_email", "best_contact_title",
    "first_seen", "monitor", "enrich",
    "added_date", "reviewed",
]

ORDER_HEADERS = [
    "domain", "order_number", "order_date", "product_name",
    "sku", "line_item_price", "quantity", "order_total",
    "fulfillment_status", "financial_status",
    "billing_company", "billing_email",
]

SIGNALS_HEADERS = [
    "domain", "company_name", "signal_type", "signal_date",
    "signal_headline", "signal_url", "relevance_score", "actioned",
]

VV_HEADERS = [
    "domain", "company_name", "visit_date", "pages_visited",
    "identified_name", "identified_email", "identified_title",
    "pages_list", "classification", "actioned",
]

OUTREACH_HEADERS = [
    "domain", "company_name", "contact_name", "contact_email",
    "send_date", "subject", "signal_context",
    "icp_label", "brevo_message_id", "status", "reply_received",
]

ERRORS_HEADERS = [
    "domain", "company_name", "error_date", "error_reason",
]

ALL_TABS = [
    (SHEET_MASTER,       MASTER_HEADERS),
    (SHEET_NEW_CONTACTS, NEW_CONTACTS_HEADERS),
    (SHEET_ORDERS,       ORDER_HEADERS),
    (SHEET_SIGNALS,      SIGNALS_HEADERS),
    (SHEET_VV,           VV_HEADERS),
    (SHEET_OUTREACH,     OUTREACH_HEADERS),
    (SHEET_ERRORS,       ERRORS_HEADERS),
]


# ─────────────────────────────────────────────────────────────────────────────
# GOOGLE SHEETS CONNECTION
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_resource(ttl=300)
def gsheet_connect():
    """
    Authenticate with Google Sheets via service account.
    Cached for 5 minutes — avoids re-auth on every Streamlit rerun.
    Returns gspread client or None on failure.
    """
    if not GSPREAD_OK:
        st.warning("gspread not installed. Run: pip install gspread google-auth")
        return None

    sa_raw = get_secret("GOOGLE_SERVICE_ACCOUNT")
    if not sa_raw:
        st.warning("GOOGLE_SERVICE_ACCOUNT secret not found.")
        return None

    try:
        sa_info = json.loads(sa_raw)
        # Fix literal \n in private key (common when pasting into Streamlit secrets)
        pk = sa_info.get("private_key", "")
        if pk and r"\n" in pk:
            sa_info["private_key"] = pk.replace(r"\n", chr(10))

        scopes = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
        return gspread.authorize(creds)

    except json.JSONDecodeError as ex:
        st.warning(f"Sheets: bad JSON in GOOGLE_SERVICE_ACCOUNT (pos {ex.pos}): {ex.msg}")
        return None
    except Exception as ex:
        st.warning(f"Google Sheets auth failed: {ex}")
        return None


def gsheet_open():
    """Open the CLD sheet by ID from secrets. Returns worksheet or None."""
    gc = gsheet_connect()
    if not gc:
        return None
    try:
        return gc.open_by_key(get_secret("GSHEET_ID"))
    except Exception as ex:
        st.warning(f"Could not open CLD sheet: {ex}")
        return None


def gsheet_ensure_tabs(sh):
    """
    Create all CLD tabs with correct headers if they don't exist yet.
    Runs once per session (cached in st.session_state).
    Call this at app startup before any tab reads/writes.
    """
    if st.session_state.get("cld_tabs_verified"):
        return

    try:
        existing = [ws.title for ws in sh.worksheets()]
        for title, headers in ALL_TABS:
            if title not in existing:
                ws = sh.add_worksheet(title=title, rows=5000, cols=len(headers))
                ws.append_row(headers, value_input_option="RAW")
        st.session_state["cld_tabs_verified"] = True
    except Exception as ex:
        st.warning(f"Tab initialization error: {ex}")


# ─────────────────────────────────────────────────────────────────────────────
# COMMON SHEET HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def sheet_get_all(sh, tab_name: str) -> list[dict]:
    """
    Read all rows from a tab. Returns list of dicts keyed by header row.
    Returns empty list on any error.
    """
    try:
        ws = sh.worksheet(tab_name)
        return ws.get_all_records()
    except Exception as ex:
        st.warning(f"Could not read {tab_name}: {ex}")
        return []


def sheet_append_rows(sh, tab_name: str, rows: list[list]) -> bool:
    """
    Append a list of rows to a tab. Each row is a plain list of values.
    Returns True on success, False on failure.
    """
    if not rows:
        return True
    try:
        ws = sh.worksheet(tab_name)
        ws.append_rows(rows, value_input_option="USER_ENTERED")
        return True
    except Exception as ex:
        st.warning(f"Could not append to {tab_name}: {ex}")
        return False


def sheet_update_cell(sh, tab_name: str, row_index: int,
                      col_name: str, value) -> bool:
    """
    Update a single cell by row index (1-based, excluding header) and column name.
    Finds the column position from the header row automatically.
    Returns True on success.
    """
    try:
        ws = sh.worksheet(tab_name)
        headers = ws.row_values(1)
        if col_name not in headers:
            st.warning(f"Column '{col_name}' not found in {tab_name}")
            return False
        col_index = headers.index(col_name) + 1  # gspread is 1-based
        ws.update_cell(row_index + 1, col_index, value)  # +1 to skip header
        return True
    except Exception as ex:
        st.warning(f"Could not update {tab_name}[{row_index}].{col_name}: {ex}")
        return False


def sheet_update_row(sh, tab_name: str, match_col: str,
                     match_val: str, updates: dict) -> bool:
    """
    Find a row where match_col == match_val and update multiple columns at once.
    updates = {"col_name": value, ...}
    Returns True if row found and updated, False otherwise.
    """
    try:
        ws = sh.worksheet(tab_name)
        headers = ws.row_values(1)
        if match_col not in headers:
            return False

        col_idx = headers.index(match_col) + 1
        col_values = ws.col_values(col_idx)

        # Find matching row (skip header at index 0)
        for i, val in enumerate(col_values[1:], start=2):
            if val == match_val:
                for col_name, new_val in updates.items():
                    if col_name in headers:
                        update_col = headers.index(col_name) + 1
                        ws.update_cell(i, update_col, new_val)
                return True
        return False
    except Exception as ex:
        st.warning(f"Could not update row in {tab_name}: {ex}")
        return False


def sheet_get_domains(sh) -> set:
    """
    Return the set of all domains currently in master_companies.
    Used for deduplication during upload.
    """
    try:
        ws = sh.worksheet(SHEET_MASTER)
        all_domains = ws.col_values(1)  # domain is column 1
        return set(all_domains[1:])     # skip header
    except Exception:
        return set()


# ─────────────────────────────────────────────────────────────────────────────
# ANTHROPIC CLIENT
# ─────────────────────────────────────────────────────────────────────────────

@st.cache_resource
def get_anthropic_client():
    """Return Anthropic client. Cached for session lifetime."""
    if not ANTHROPIC_OK:
        return None
    api_key = get_secret("ANTHROPIC_API_KEY")
    if not api_key:
        return None
    return anthropic.Anthropic(api_key=api_key)


# ─────────────────────────────────────────────────────────────────────────────
# SEGMENT RULES
# ─────────────────────────────────────────────────────────────────────────────

# Which signal types each domain_class gets monitored for
MONITOR_SIGNALS = {
    "commercial":  ["news", "hiring", "vv_visit", "funding"],
    "education":   ["news", "hiring", "vv_visit"],
    "ham":         ["news", "vv_visit"],          # ARRL/QRZ instead of Google News
    "government":  ["news", "vv_visit"],
    "defense":     ["vv_visit"],                  # VV only — market intelligence
    "distributor": [],                            # No monitoring
}

# Default monitor toggle per domain_class
DEFAULT_MONITOR = {
    "commercial":  True,
    "education":   True,
    "ham":         True,
    "government":  True,
    "defense":     False,   # VV parsed but not in active monitor queue
    "distributor": False,
}

# Whether outreach generation is available per domain_class
OUTREACH_ENABLED = {
    "commercial":  True,
    "education":   True,
    "ham":         True,
    "government":  True,
    "defense":     False,   # suppress_outreach always True
    "distributor": False,
}

# ICP template to use for outreach generation per domain_class
OUTREACH_TEMPLATE = {
    "commercial":  "icp_matched",   # Uses icp_label from enrichment
    "education":   "icp6_education",
    "ham":         "ham",
    "government":  "government",
    "defense":     None,
    "distributor": None,
}

# Display label per domain_class
CLASS_LABEL = {
    "commercial":  "Commercial",
    "education":   "Education",
    "ham":         "HAM Radio",
    "government":  "Government",
    "defense":     "Defense",
    "distributor": "Distributor",
}

# Color coding for domain_class badges in the UI
CLASS_COLOR = {
    "commercial":  "#0066CC",
    "education":   "#2E7D32",
    "ham":         "#6A1B9A",
    "government":  "#E65100",
    "defense":     "#B71C1C",
    "distributor": "#546E7A",
}

# Special domains requiring manual review — never auto-generate outreach
MANUAL_REVIEW_DOMAINS = {
    "tektronix.com": "Competitor account — likely engineer personal/skunkworks purchase.",
}

# Known low-ICP-fit domains — can be pre-flagged before enrichment runs
LOW_ICP_DOMAINS = {
    "wsscwater.com": "Water utility — no T&M instrument buyer apparent.",
    "jax.org":       "Biomedical research — likely outside addressable market.",
    "varian.com":    "Medical systems — outside addressable market.",
    "trace3.com":    "IT reseller — outside addressable market.",
}

# ─────────────────────────────────────────────────────────────────────────────
# ENRICHMENT LABELS
# ─────────────────────────────────────────────────────────────────────────────

INDUSTRY_LABELS = [
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

ICP_LABELS = [
    "ICP-4 R&D Engineer",
    "ICP-5 Power Electronics",
    "ICP-6 Education",
    "ICP-7 Test Lab",
    "ICP-1 Industrial",
    "ICP-2 Electrician",
    "ICP-3 Solar",
    "Mixed",
    "None",
]

ICP_CONFIDENCE = ["High", "Medium", "Low"]


# ─────────────────────────────────────────────────────────────────────────────
# UTILITY HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def now_str() -> str:
    """Current UTC datetime as ISO string."""
    return datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S")


def today_str() -> str:
    """Today's date as YYYY-MM-DD string."""
    return datetime.datetime.utcnow().strftime("%Y-%m-%d")


def extract_domain(email: str) -> str:
    """
    Extract domain from an email address.
    Returns empty string if email is blank or malformed.
    """
    if not email or "@" not in email:
        return ""
    return email.strip().lower().split("@")[-1]


def format_currency(value) -> str:
    """Format a numeric value as $X,XXX.XX for display."""
    try:
        return f"${float(value):,.2f}"
    except (ValueError, TypeError):
        return "$0.00"


def tier_badge(tier) -> str:
    """Return a display label for a watch tier value."""
    labels = {1: "Tier 1 — Active", 2: "Tier 2 — Passive", 3: "Tier 3 — List"}
    try:
        return labels.get(int(tier), f"Tier {tier}")
    except (ValueError, TypeError):
        return "Unknown"
