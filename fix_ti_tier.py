"""
fix_ti_tier.py -- one-shot script to set watch_tier=1 for all TIER1_OVERRIDE
domains that currently have watch_tier != 1 in master_companies.

Run from the Customer Lead Dashboard folder:
    python fix_ti_tier.py

Requires: pip install gspread google-auth
Secrets: set GOOGLE_SERVICE_ACCOUNT env var (JSON string), or edit KEY_FILE below.
"""

import json
import os
import sys
import gspread
from google.oauth2.service_account import Credentials

# ── CONFIG ────────────────────────────────────────────────────────────────────
# Edit KEY_FILE if running without the env var

GSHEET_ID = "1oW08gVUcBFMyyumiv6Wl8jGVbvcDmvO-9BIM2e_lN04"
KEY_FILE  = r"C:\Users\pjsto\Private\Keys\google_service_account.json"

TIER1_OVERRIDE = {
    # Semiconductor design (ICP-4 / ICP-5 core)
    "ti.com", "analog.com", "nxp.com", "onsemi.com", "microchip.com",
    "infineon.com", "st.com", "renesas.com", "mchp.com", "maximintegrated.com",
    "latticesemi.com", "idt.com", "skyworksinc.com", "qorvo.com",
    "wolfspeed.com", "macom.com", "semtech.com",
    # EV / Power Electronics OEMs (ICP-5 core)
    "rivian.com", "lucidmotors.com", "borgwarner.com", "vicr.com",
    # ATE / Test engineering (ICP-7 / ICP-4)
    "teradyne.com", "ni.com", "astronics.com", "spirent.com",
    # Electronics design (ICP-4 broad)
    "qualcomm.com", "broadcom.com", "marvell.com", "amd.com",
    "intel.com", "nvidia.com", "murata.com", "tdk.com", "vishay.com",
    # Automotive electronics (ICP-4 / ICP-5)
    "aptiv.com", "visteon.com", "magna.com", "denso.com", "continental.com",
    # Industrial drives / automation (ICP-1 / ICP-5)
    "rockwellautomation.com", "abb.com", "yaskawa.com", "danfoss.com",
    "emerson.com", "siemens.com", "schneider-electric.com",
}

# ── CONNECT ───────────────────────────────────────────────────────────────────

sa_raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT")
if sa_raw:
    sa_info = json.loads(sa_raw)
    print("Using GOOGLE_SERVICE_ACCOUNT env var.")
else:
    with open(KEY_FILE) as f:
        sa_info = json.load(f)
    print(f"Using key file: {KEY_FILE}")

pk = sa_info.get("private_key", "")
if pk and r"\n" in pk:
    sa_info["private_key"] = pk.replace(r"\n", "\n")

scopes = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]
creds = Credentials.from_service_account_info(sa_info, scopes=scopes)
gc    = gspread.authorize(creds)
sh    = gc.open_by_key(GSHEET_ID)
ws    = sh.worksheet("master_companies")

print("Reading master_companies...")
all_vals = ws.get_all_values()
headers  = all_vals[0]
col      = {h: i for i, h in enumerate(headers)}
dom_col  = col.get("domain", 0)
tier_col = col.get("watch_tier")

if tier_col is None:
    sys.exit("ERROR: 'watch_tier' column not found in master_companies headers.")

# Find rows that need updating
updates = []
matched = []
for i, row in enumerate(all_vals[1:], start=2):
    domain = row[dom_col].strip().lower()
    if domain in TIER1_OVERRIDE:
        current_tier = row[tier_col].strip()
        matched.append((domain, current_tier))
        if current_tier != "1":
            updates.append((i, tier_col + 1, 1))  # gspread is 1-based

print(f"\nOverride domains found in Sheet: {len(matched)}")
for domain, tier in sorted(matched):
    mark = " <-- will update" if tier != "1" else " (already T1)"
    print(f"  {domain}: tier={tier}{mark}")

if not updates:
    print("\nNothing to update -- all override domains are already Tier 1.")
    sys.exit(0)

print(f"\nUpdating {len(updates)} row(s) to Tier 1...")
body = {
    "data": [
        {
            "range":  gspread.utils.rowcol_to_a1(r, c),
            "values": [[v]],
        }
        for r, c, v in updates
    ],
    "valueInputOption": "USER_ENTERED",
}
ws.spreadsheet.values_batch_update(body)
print(f"Done. {len(updates)} row(s) updated to watch_tier=1.")
print("\nReload from Sheet in the dashboard to see the changes.")
