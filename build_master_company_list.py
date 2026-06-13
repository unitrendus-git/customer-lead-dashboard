"""
build_master_company_list.py
Customer Lead Dashboard — Phase 1
Merges Brevo + Shopify exports into a scored master company list.

Usage (run from Customer Lead Dashboard folder):
    python build_master_company_list.py

Outputs:
    master_company_list.csv   — one row per unique corporate domain
    filter_log.csv            — every contact and why it was kept/excluded

Re-run after dropping additional Shopify batches into sources/ folder.
The script is additive — all SHOPIFY_GLOB matches are loaded together.
"""

import csv
import glob
import sys
from collections import defaultdict
import re
import openpyxl

# ── CONFIG ────────────────────────────────────────────────────────────────────

BREVO_CSV      = "sources/7938800-6a2c13d4756f52fcae9420da-DhxkDJ.csv"
SHOPIFY_GLOB   = "sources/Export_*.xlsx"   # picks up all batches
OUTPUT_MASTER  = "master_company_list.csv"
OUTPUT_LOG     = "filter_log.csv"

# ── FILTER SETS ───────────────────────────────────────────────────────────────

FREE_DOMAINS = {
    'gmail.com','yahoo.com','hotmail.com','outlook.com','icloud.com',
    'aol.com','msn.com','live.com','me.com','mac.com','protonmail.com',
    'comcast.net','att.net','verizon.net','sbcglobal.net','cox.net',
    'earthlink.net','bellsouth.net','ymail.com','googlemail.com',
    'qq.com','163.com','126.com','sina.com','sohu.com',
    'mail.com','gmx.com','zoho.com','fastmail.com','tutanota.com',
    'rocketmail.com','inbox.com','lavabit.com',
}

OWN_DOMAINS = {'uni-trendus.com', 'uni-trend.com'}

MARKETPLACE_TAGS = {'amazon', 'ebay', 'walmart'}

# ── DISTRIBUTOR REGISTRY ──────────────────────────────────────────────────────
# Format: domain -> parent group name
# Any domain where ANY Shopify contact has premier/authorized tag will be
# added automatically at runtime. These are the known/manual additions.

DISTRIBUTOR_REGISTRY = {
    # TestEquity family
    'testequity.com':       'TestEquity',
    'hiscoinc.com':         'TestEquity',
    'tequipment.net':       'TestEquity',
    'conres.com':           'TestEquity',
    # ITM family
    'itm.com':              'ITM',
    'globaltestsupply.com': 'ITM',
    'reedinstruments.com':  'ITM',
    # Fotronic / Test Equipment Depot
    'fotronic.com':         'Fotronic',
    # Stratatek
    'stratatek.com':        'Stratatek',
    # Sper Scientific / Sanwa America
    'sperscientific.com':   'Sper Scientific',
    'sperdirect.com':       'Sper Scientific',
    # Others
    'anncode.com':          'Anncode',
    'astutegroup.com':      'Astute Electronics',
    'atm1.com':             'ATM',
    'duncaninstr.com':      'Duncan Instruments',
    'electro5.com':         'Electro5',
    'jameco.com':           'Jameco',
    'kohen.ca':             'Ko-Hen Electronics',
    'masterelectronics.com':'Master Electronics',
    'newark.com':           'Newark / Farnell',
    'sierratechnical.com':  'Sierra Technical',
    'simpleydirecto.mx':    'Simpley Directo',
    'skycraftsurplus.com':  'Skycraft Surplus',
    'solarbotics.com':      'Solarbotics',
    'testrack.com':         'TestRack Inc',
    'transcat.com':         'Transcat',
    'valuetronics.com':     'Valuetronics',
    # Manually added -- confirmed distributors without authorized tag
    'electrorent.com':      'Electro Rent',
    # Major national distributors not yet seen in data
    'digikey.com':          'Digi-Key',
    'mouser.com':           'Mouser',
    'arrow.com':            'Arrow',
    'avnet.com':            'Avnet',
    'ttiinc.com':           'TTI',
    'grainger.com':         'Grainger',
    'rs-online.com':        'RS Components',
    'element14.com':        'Element14',
    'farnell.com':          'Farnell',
    'mcmaster.com':         'McMaster-Carr',
    'alliedelec.com':       'Allied Electronics',
    'heilind.com':          'Heilind',
    'futureelectronics.com':'Future Electronics',
}

# ── DEFENSE DOMAINS ───────────────────────────────────────────────────────────
# Monitored but outreach suppressed

DEFENSE_DOMAINS = {
    'lmco.com', 'l3harris.com', 'baesystems.com', 'northropgrumman.com',
    'raytheon.com', 'rtx.com', 'boeing.com', 'generaldynamics.com',
    'leidos.com', 'saic.com', 'bah.com', 'mda.ca', 'curtisswright.com',
    'ultra.group', 'leonardodrs.com', 'textron.com', 'sierra-nevada.com',
    'spacex.com', 'anduril.com', 'palantir.com',
    'ngc.com', 'harris.com',   # Northrop Grumman and L3Harris legacy domains
}

# ── TIER 1 STRATEGIC OVERRIDE ────────────────────────────────────────────────
# Domains that always receive Tier 1 regardless of scoring.
# These are definitively ICP-4 (R&D Engineer) / ICP-5 (Power Electronics) /
# ICP-7 (Test Lab) companies where any contact is high-value.
# Add new domains here as you encounter them; re-run build to apply.

TIER1_OVERRIDE = {
    # Semiconductor design (ICP-4 / ICP-5 core)
    "ti.com",              # Texas Instruments
    "analog.com",          # Analog Devices
    "nxp.com",             # NXP Semiconductors
    "onsemi.com",          # onsemi (formerly ON Semi)
    "microchip.com",       # Microchip Technology
    "infineon.com",        # Infineon Technologies
    "st.com",              # STMicroelectronics
    "renesas.com",         # Renesas Electronics
    "mchp.com",            # Microchip alt domain
    "maximintegrated.com", # Maxim (now ADI)
    "latticesemi.com",     # Lattice Semiconductor
    "idt.com",             # IDT (Renesas)
    "skyworksinc.com",     # Skyworks Solutions
    "qorvo.com",           # Qorvo (RF/power)
    "wolfspeed.com",       # Wolfspeed (SiC/GaN -- ICP-5)
    "macom.com",           # MACOM Technology
    "semtech.com",         # Semtech
    # EV / Power Electronics OEMs (ICP-5 core)
    "rivian.com",          # Rivian
    "lucidmotors.com",     # Lucid Motors
    "borgwarner.com",      # BorgWarner (EV powertrain)
    "vicr.com",            # Vicor (power components)
    # ATE / Test engineering (ICP-7 / ICP-4)
    "teradyne.com",        # Teradyne
    "ni.com",              # NI (National Instruments)
    "astronics.com",       # Astronics Test Systems
    "spirent.com",         # Spirent
    # Electronics design (ICP-4 broad)
    "qualcomm.com",        # Qualcomm
    "broadcom.com",        # Broadcom
    "marvell.com",         # Marvell Technology
    "amd.com",             # AMD (including Xilinx)
    "intel.com",           # Intel
    "nvidia.com",          # NVIDIA
    "murata.com",          # Murata Manufacturing
    "tdk.com",             # TDK
    "vishay.com",          # Vishay
    # Automotive electronics (ICP-4 / ICP-5)
    "aptiv.com",           # Aptiv
    "visteon.com",         # Visteon
    "magna.com",           # Magna International
    "denso.com",           # Denso
    "continental.com",     # Continental AG
    # Industrial drives / automation (ICP-1 / ICP-5)
    "rockwellautomation.com",  # Rockwell Automation
    "abb.com",                 # ABB
    "yaskawa.com",             # Yaskawa
    "danfoss.com",             # Danfoss
    "emerson.com",             # Emerson
    "siemens.com",             # Siemens
    "schneider-electric.com",  # Schneider Electric
}

# ── HAM RADIO DOMAINS ─────────────────────────────────────────────────────────
# Tracked as a distinct segment -- sold at 4 ham festivals/year + ARRL promo.
# Tagged domain_class = 'ham'. Excluded from commercial outreach generation.
# Show clustering confirmed: Feb, May, Aug, Nov/Dec spikes in order data.

HAM_DOMAINS_EXPLICIT = {
    'arrl.net', 'usham.org', 'qrz.com', 'eham.net', 'arrl.org',
    'hamradioplayground.com', 'irvingarc.org', 'orlando220.org',
}

# Amateur radio callsign pattern (US prefixes W/K/N/A + digit + suffix)
HAM_CALLSIGN_PATTERN = re.compile(
    r'^(w|k|n|aa|ab|ac|ad|ae|af|ag|ai|aj|ak|'
    r'wa|wb|wc|wd|we|wf|wg|wi|wj|wk|wl|wm|wn|wo|wp|wq|wr|ws|wt|wu|wv|ww|wx|wy|wz|'
    r'ka|kb|kc|kd|ke|kf|kg|ki|kj|kk|kl|km|kn|ko|kp|kq|kr|ks|kt|ku|kv|kw|kx|ky|kz)'
    r'\d[a-z0-9]{1,4}\.(com|net|org|us|radio)$',
    re.IGNORECASE
)

HAM_COMPANY_PATTERN = re.compile(
    r'\b(ham radio|hamfest|ham club|amateur radio|arrl)\b',
    re.IGNORECASE
)

# Domains that look like HAM but are not -- explicitly excluded
HAM_FALSE_POSITIVES = {
    # cs.com removed: confirmed HAM (retiree, bought post-show March 2025)
    # formfactor.com: semiconductor test company, Brevo ham tag is data entry error
    'formfactor.com', 'k104fm.com',
    'snhu.edu', 'byu.edu', 'durham-tech.net', 'rdu.com',
    'windhambrannon.com', 'mtd.org', 'live.ca', 'illinois.edu',
    'durhamusa.com', 'uabmc.edu', 'chatham.edu', 'caes.com',
    'chathamcountync.gov', 'uagm.edu', 'yrh.com', 'champaero.com',
    'portcityair.com', 'unh.edu', 'binghamton.edu', 'hamptonu.edu',
    'bhm.k12.al.us', 'dcicomm.com',
}

def is_ham_domain(domain, tags_set, company_name):
    if domain in HAM_FALSE_POSITIVES:
        return False
    if domain in HAM_DOMAINS_EXPLICIT:
        return True
    if 'ham' in tags_set:
        return True
    if HAM_CALLSIGN_PATTERN.match(domain):
        return True
    if HAM_COMPANY_PATTERN.search(company_name or ''):
        return True
    return False

# ── SCORING ───────────────────────────────────────────────────────────────────

ENGINEER_TITLE_KEYWORDS = [
    'engineer', 'technician', 'tech', 'scientist', 'physicist',
    'researcher', 'r&d', 'lab', 'laboratory', 'hardware', 'firmware',
    'embedded', 'electrical', 'electronic', 'rf', 'signal', 'power',
    'test', 'measurement', 'metrology', 'calibration', 'quality',
    'procurement', 'purchasing', 'buyer', 'sourcing',
    'director', 'manager', 'lead', 'principal', 'senior', 'chief',
    'professor', 'instructor', 'faculty', 'phd', 'postdoc',
]

EVENT_TAGS = {
    'designcon 2024', 'auto test expo 2024', 'ieee',
    'automotive testing expo', 'ims2026', 'tradeshow',
    'zoominfo', 't&mt',
}

# ── HELPERS ───────────────────────────────────────────────────────────────────

def get_domain(email):
    email = str(email or '').strip().lower()
    if '@' in email:
        d = email.split('@')[1].strip()
        if '.' in d and len(d) > 3:
            return d
    return None

def classify_domain(domain, all_distributor_domains):
    """Priority: distributor > edu > gov > defense > commercial"""
    if domain in all_distributor_domains:
        return 'distributor'
    # Canadian university domains that don't end in .edu
    EDU_EQUIVALENT = {
        'mcmaster.ca', 'utoronto.ca', 'ubc.ca', 'ualberta.ca',
        'uwaterloo.ca', 'queensu.ca', 'yorku.ca', 'uoguelph.ca',
    }
    if domain.endswith('.edu') or domain.endswith('.ac.uk') or domain in EDU_EQUIVALENT:
        return 'education'
    if domain.endswith('.gov') or domain.endswith('.mil'):
        return 'government'
    if domain in DEFENSE_DOMAINS:
        return 'defense'
    return 'commercial'

def has_engineer_title(title):
    title = str(title or '').lower()
    return any(kw in title for kw in ENGINEER_TITLE_KEYWORDS)

def parse_tags(tag_str):
    if not tag_str:
        return set()
    return {t.strip().lower() for t in str(tag_str).split(',') if t.strip()}

def score_company(domain_class, has_title, in_shopify,
                  has_purchases, in_brevo, has_event_tag):
    # Dimension 1 -- domain tier
    if domain_class == 'commercial':
        d = 3
    elif domain_class in ('education', 'government'):
        d = 2
    else:  # distributor, defense, ham
        d = 1

    # Dimension 2 -- contact quality
    if has_title:
        c = 3
    elif in_brevo or in_shopify:
        c = 2
    else:
        c = 1

    # Dimension 3 -- engagement signal
    if in_shopify and has_purchases and in_brevo:
        e = 3
    elif (in_shopify and has_purchases) or has_event_tag:
        e = 2
    else:
        e = 1

    total = d + c + e
    return total, {'score_domain': d, 'score_contact': c, 'score_engagement': e}

def score_to_tier(score, domain=""):
    """Convert 3-dimension score to watch tier. Domain overrides take priority."""
    if domain in TIER1_OVERRIDE:
        return 1
    if score >= 7: return 1
    if score >= 4: return 2
    return 3

# ── LOAD BREVO ────────────────────────────────────────────────────────────────

print("Loading Brevo CSV...")
brevo_contacts = []
with open(BREVO_CSV, encoding='utf-8', errors='replace') as f:
    reader = csv.DictReader(f)
    for row in reader:
        brevo_contacts.append(row)
print(f"  {len(brevo_contacts):,} contacts")

# ── LOAD SHOPIFY (all batches) ─────────────────────────────────────────────────

print("Loading Shopify XLSX batches...")
shopify_data    = []
shopify_headers = None
shopify_files   = sorted(glob.glob(SHOPIFY_GLOB))

if not shopify_files:
    print("  WARNING: no Shopify export files found matching", SHOPIFY_GLOB)
else:
    for fpath in shopify_files:
        wb = openpyxl.load_workbook(fpath, read_only=True)
        ws = wb['Customers']
        rows = list(ws.iter_rows(values_only=True))
        if shopify_headers is None:
            shopify_headers = rows[0]
        shopify_data.extend(rows[1:])
        print(f"  {fpath}: {len(rows)-1:,} records")

print(f"  Total Shopify records: {len(shopify_data):,}")

def sf(row, name):
    try:
        return row[shopify_headers.index(name)]
    except (ValueError, IndexError):
        return None

# ── PASS 1: PROPAGATE DISTRIBUTOR TAGS BY DOMAIN ──────────────────────────────

print("\nPass 1: Propagating distributor tags by domain...")
auto_dist = {}
for row in shopify_data:
    domain = get_domain(sf(row, 'Email'))
    if not domain or domain in FREE_DOMAINS or domain in OWN_DOMAINS:
        continue
    tags = parse_tags(sf(row, 'Tags'))
    if any('premier' in t or 'authorized' in t for t in tags):
        if domain not in DISTRIBUTOR_REGISTRY:
            company = str(sf(row, 'Address Company') or '').strip()
            auto_dist[domain] = company or domain

all_dist = dict(DISTRIBUTOR_REGISTRY)
new_auto = {d: p for d, p in auto_dist.items() if d not in all_dist}
all_dist.update(new_auto)

print(f"  {len(DISTRIBUTOR_REGISTRY)} registry entries")
print(f"  {len(new_auto)} new auto-detected distributor domains: {list(new_auto.keys())}")
print(f"  {len(all_dist)} total distributor domains")

# ── PASS 2: BUILD DOMAIN RECORDS ──────────────────────────────────────────────

print("\nPass 2: Building domain records...")

companies  = defaultdict(lambda: {
    'domain': '', 'company_name': '', 'domain_class': '',
    'distributor_parent': '',
    'brevo_contacts': 0, 'shopify_contacts': 0,
    'has_purchases': False, 'total_spent': 0.0, 'total_orders': 0,
    'best_contact_name': '', 'best_contact_email': '', 'best_contact_title': '',
    'has_engineer_title': False,
    'tags': set(), 'category_of_interest': '',
    'has_event_tag': False,
    'in_brevo': False, 'in_shopify': False, 'shopify_linked': False,
    'first_seen': '', 'last_activity': '',
})
filter_log = []

# --- Brevo ---
brevo_out = 0
for row in brevo_contacts:
    email  = row.get('EMAIL', '').strip()
    domain = get_domain(email)

    if not domain:
        filter_log.append({'source':'brevo','email':email,'reason':'no_email'}); brevo_out+=1; continue
    if domain in FREE_DOMAINS:
        filter_log.append({'source':'brevo','email':email,'reason':'free_domain'}); brevo_out+=1; continue
    if domain in OWN_DOMAINS:
        filter_log.append({'source':'brevo','email':email,'reason':'own_domain'}); brevo_out+=1; continue

    co = companies[domain]
    co['domain']         = domain
    co['in_brevo']       = True
    co['brevo_contacts'] += 1

    name = (row.get('COMPANY') or '').strip()
    if len(name) > len(co['company_name']):
        co['company_name'] = name

    title = (row.get('JOB_TITLE') or '').strip()
    if title and not co['has_engineer_title']:
        co['best_contact_name']  = f"{row.get('FIRSTNAME','')} {row.get('LASTNAME','')}".strip()
        co['best_contact_email'] = email
        co['best_contact_title'] = title
        if has_engineer_title(title):
            co['has_engineer_title'] = True
    elif not co['best_contact_email']:
        co['best_contact_name']  = f"{row.get('FIRSTNAME','')} {row.get('LASTNAME','')}".strip()
        co['best_contact_email'] = email

    tags = parse_tags(row.get('TAGS',''))
    co['tags'] |= tags
    if tags & EVENT_TAGS:
        co['has_event_tag'] = True

    cat = (row.get('CATEGORY_OF_INTEREST') or '').strip()
    if cat and not co['category_of_interest']:
        co['category_of_interest'] = cat

    if str(row.get('SHOPIFY','')).strip().lower() == 'yes':
        co['shopify_linked'] = True

    added    = (row.get('ADDED_TIME') or '').strip()
    modified = (row.get('MODIFIED_TIME') or '').strip()
    if added    and (not co['first_seen']    or added    < co['first_seen']):    co['first_seen']    = added
    if modified and (not co['last_activity'] or modified > co['last_activity']): co['last_activity'] = modified

print(f"  Brevo: {brevo_out:,} excluded, {len(brevo_contacts)-brevo_out:,} mapped")

# --- Shopify ---
shopify_out  = 0
seen_shopify = set()

for row in shopify_data:
    email  = str(sf(row, 'Email') or '').strip().lower()
    domain = get_domain(email)
    tags   = parse_tags(sf(row, 'Tags'))

    if not domain:
        filter_log.append({'source':'shopify','email':email,'reason':'no_email'}); shopify_out+=1; continue
    if domain in FREE_DOMAINS:
        filter_log.append({'source':'shopify','email':email,'reason':'free_domain'}); shopify_out+=1; continue
    if domain in OWN_DOMAINS:
        filter_log.append({'source':'shopify','email':email,'reason':'own_domain'}); shopify_out+=1; continue
    if tags & MARKETPLACE_TAGS:
        filter_log.append({'source':'shopify','email':email,'reason':'marketplace_buyer'}); shopify_out+=1; continue

    co = companies[domain]
    co['domain']     = domain
    co['in_shopify'] = True

    if email not in seen_shopify:
        seen_shopify.add(email)
        co['shopify_contacts'] += 1

    name = str(sf(row, 'Address Company') or '').strip()
    if len(name) > len(co['company_name']):
        co['company_name'] = name

    top_row = sf(row, 'Top Row')
    if top_row is True or str(top_row).strip().lower() in ('true','1','yes'):
        spent  = float(sf(row, 'Total Spent')  or 0)
        orders = int(sf(row,   'Total Orders') or 0)
        co['total_spent']  += spent
        co['total_orders'] += orders
        if spent > 0:
            co['has_purchases'] = True

    co['tags'] |= tags

    created = str(sf(row, 'Created At') or '').strip()
    updated = str(sf(row, 'Updated At') or '').strip()
    if created and (not co['first_seen']    or created < co['first_seen']):    co['first_seen']    = created
    if updated and (not co['last_activity'] or updated > co['last_activity']): co['last_activity'] = updated

print(f"  Shopify: {shopify_out:,} excluded, {len(shopify_data)-shopify_out:,} rows processed")

# ── PASS 3: CLASSIFY, SCORE, OUTPUT ───────────────────────────────────────────

print(f"\nPass 3: Classifying and scoring {len(companies):,} domains...")
output_rows = []

for domain, co in companies.items():
    co['domain_class']       = classify_domain(domain, set(all_dist.keys()))
    co['distributor_parent'] = all_dist.get(domain, '')
    # Override: HAM segment check (after distributor/edu/gov/defense)
    if co['domain_class'] == 'commercial':
        tags_set = co.get('tags', set())
        company_name = co.get('company_name', '')
        if is_ham_domain(domain, tags_set, company_name):
            co['domain_class'] = 'ham'
            if 'ham' not in co['tags']:
                co['tags'] |= {'ham'}

    score, breakdown = score_company(
        domain_class  = co['domain_class'],
        has_title     = co['has_engineer_title'],
        in_shopify    = co['in_shopify'],
        has_purchases = co['has_purchases'],
        in_brevo      = co['in_brevo'],
        has_event_tag = co['has_event_tag'],
    )
    tier = score_to_tier(score, domain=domain)

    if co['has_purchases'] and co['total_spent'] > 0:
        if   co['total_spent'] >= 5000: status = 'customer_high_value'
        elif co['total_spent'] >= 500:  status = 'customer'
        else:                           status = 'customer_low_value'
    elif co['in_shopify']:     status = 'prospect_shopify'
    elif co['shopify_linked']: status = 'prospect_linked'
    else:                      status = 'prospect'

    output_rows.append({
        'domain':               domain,
        'company_name':         co['company_name'],
        'domain_class':         co['domain_class'],
        'distributor_parent':   co['distributor_parent'],
        'customer_status':      status,
        'watch_tier':           tier,
        'score':                score,
        'score_domain':         breakdown['score_domain'],
        'score_contact':        breakdown['score_contact'],
        'score_engagement':     breakdown['score_engagement'],
        'suppress_outreach':    co['domain_class'] == 'defense',
        'brevo_contacts':       co['brevo_contacts'],
        'shopify_contacts':     co['shopify_contacts'],
        'total_spent':          round(co['total_spent'], 2),
        'total_orders':         co['total_orders'],
        'best_contact_name':    co['best_contact_name'],
        'best_contact_email':   co['best_contact_email'],
        'best_contact_title':   co['best_contact_title'],
        'category_of_interest': co['category_of_interest'],
        'tags':                 '|'.join(sorted(co['tags'] - {'','nan'})),
        'has_event_tag':        co['has_event_tag'],
        'in_brevo':             co['in_brevo'],
        'in_shopify':           co['in_shopify'],
        'has_purchases':        co['has_purchases'],
        'first_seen':           co['first_seen'],
        'last_activity':        co['last_activity'],
        # Phase 2 enrichment columns -- populated by enrich_companies.py
        'enriched':             False,
        'enrichment_date':      '',
        'website_description':  '',
        'icp_label':            '',
        'icp_confidence':       '',
        # Phase 2 signal columns -- populated by signal_monitor.py
        'vv_last_visit':        '',
        'vv_pages_visited':     '',
        'last_signal_date':     '',
        'last_signal_summary':  '',
        # Phase 3 outreach columns
        'outreach_status':      'none',
        'outreach_last_date':   '',
        'notes':                '',
    })

output_rows.sort(key=lambda r: (r['watch_tier'], -r['total_spent']))

# ── WRITE ─────────────────────────────────────────────────────────────────────

print(f"\nWriting {OUTPUT_MASTER}...")
with open(OUTPUT_MASTER, 'w', newline='', encoding='utf-8') as f:
    writer = csv.DictWriter(f, fieldnames=list(output_rows[0].keys()))
    writer.writeheader()
    writer.writerows(output_rows)

print(f"Writing {OUTPUT_LOG}...")
with open(OUTPUT_LOG, 'w', newline='', encoding='utf-8') as f:
    writer = csv.DictWriter(f, fieldnames=['source','email','reason'])
    writer.writeheader()
    writer.writerows(filter_log)

# ── SUMMARY ───────────────────────────────────────────────────────────────────

from collections import Counter
cc = Counter(r['domain_class']    for r in output_rows)
tc = Counter(r['watch_tier']      for r in output_rows)
sc = Counter(r['customer_status'] for r in output_rows)
dist_rows     = [r for r in output_rows if r['domain_class'] == 'distributor']
family_counts = Counter(r['distributor_parent'] for r in dist_rows)

# Count how many domains were Tier 1 via override vs scoring
override_t1 = sum(1 for r in output_rows if r['domain'] in TIER1_OVERRIDE)
scored_t1   = tc[1] - override_t1

print(f"""
+======================================================+
|         MASTER COMPANY LIST -- BUILD SUMMARY         |
+======================================================+
|  Shopify batches loaded:      {len(shopify_files):>6}                |
|  Total unique domains:        {len(output_rows):>6,}                |
+======================================================+
|  BY CLASSIFICATION                                   |
|    Commercial:                {cc['commercial']:>6,}                |
|    Education:                 {cc['education']:>6,}                |
|    Government:                {cc['government']:>6,}                |
|    Defense:                   {cc['defense']:>6,}                |
|    HAM Radio:                 {cc['ham']:>6,}                |
|    Distributor:               {cc['distributor']:>6,}                |
+======================================================+
|  BY WATCH TIER                                       |
|    Tier 1 (active monitor):   {tc[1]:>6,}                |
|      via strategic override:  {override_t1:>6,}                |
|      via score >=7:           {scored_t1:>6,}                |
|    Tier 2 (passive monitor):  {tc[2]:>6,}                |
|    Tier 3 (list only):        {tc[3]:>6,}                |
+======================================================+
|  BY CUSTOMER STATUS                                  |
|    High-value customer:       {sc['customer_high_value']:>6,}                |
|    Customer:                  {sc['customer']:>6,}                |
|    Low-value customer:        {sc['customer_low_value']:>6,}                |
|    Prospect (Shopify):        {sc['prospect_shopify']:>6,}                |
|    Prospect (Brevo-linked):   {sc['prospect_linked']:>6,}                |
|    Prospect:                  {sc['prospect']:>6,}                |
+======================================================+""")

print(f"""
Output files:
  {OUTPUT_MASTER}  ({len(output_rows):,} rows)
  {OUTPUT_LOG}

To add more Tier 1 strategic accounts:
  Edit TIER1_OVERRIDE in this script and re-run.
""")
