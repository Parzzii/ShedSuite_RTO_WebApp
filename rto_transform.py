from __future__ import annotations

import csv
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import requests

try:
    import pyodbc  # type: ignore
except Exception:
    pyodbc = None


RTO_HEADERS = [
    'ACCOUNT','NAME','ALTFIRSTNAME','ALTLASTNAME','ADDRESS','APT','CITY','STATE','ZIP','HMPHONE','WKPHONE','MESS','CELL','PHONE5','WORKEXT','DIRECTIONS','ZONE','SS1','SS2','DOB1','DOB2','DL1','DL2','EMAIL','ACHACCT','ACHROUTE','ACHACCTTYPE','AUTOCHARGEFORM','WORK1','WORKPH1','CCACCT','CCEXP','CCTYPE','CCSTREET','CCZIP','CELLOPT','CELLOPT2','COUNTY','EMAILINV','REFERENCE1','REFADD1','REFADD21','REFRELATION1','REFPHONE1','REFERENCE2','REFRELATION2','REFPHONE2','CONTRACT','DUEDATE','CONTRACTDATE','PMT','GRP','CONTRACTAMT','AMTPAID','OTHERPAID','CLUBPAID','EXTRARENT','PAIDDOWN','OTHERDUE','CHARGESDUE','CONTYPE','TERM','PAYMENTS','BILLED','CASHPRICE','SACDATE','CLUBFEE','PAYOFFDISCOUNT','SALESMAN','AUTOCHARGEPMTS','AUTOCHARGEB4DUE','TAXRATE','TAXZONE','STORE','BOTHSIGN','DEALERID','del_address1','del_address2','del_city','del_state','del_zip','from_address1','from_address2','from_city','from_state','from_zip','MODEL1','SERIAL1','DESCRIPTION1','RATE1','STOCK1','BRAND1','DATEREC1','VENDOR1','COST1','AGENT1','BOR1','CATEGORY1','INVOICE1','RETAIL1','RTO1','COMMENTS','CONDITION1','INVCOMMENTS1'
]

# These are taken from Import Tools K:L in Contract_Import.xlsm V1.1.7.6.
MODEL_SERIES = {
    'CRF_ALPINE': ('33-', 1300, 1409, False),
    'CRF_GENESIS': ('35-', 216, 298, False),
    'CRF_PHOENIX': ('0149-', 114, 198, False),
    'CRF_4_SEASONS': ('0150-', 153, 198, False),
    'DBM_PHOENIX': ('0203-', 549, 598, False),
    'WWP': ('1101-', 409, 498, False),
    'RAS_YSS': ('0605-', 37, 98, True),
    'LSR_TX': ('0701-', 816, 898, True),
    'MAG_GA': ('0801-', 1193, 1298, False),
    'MAG_SC': ('0801-', 273, 299, False),
    'MAG_NC': ('0801-', 330, 398, False),
    'MAG_TN': ('0801-', 401, 499, False),
}

# selectLogIn() + Defaults!D2 logic from the original workbook.
COMPANY_RULES = {
    'smart shed': {
        'profile_by_state': True, 'profile': 'MAG_GA', 'login': 'office@magnoliarto.com',
        'store_title': '8 Magnolia', 'store': '8', 'zone_selector': 'Smart Shed, 0801'
    },
    'lonestar sheds llc': {
        'profile': 'LSR_TX', 'login': 'lsrbarns@gmail.com', 'store_title': '7 Lonestar', 'store': '7',
        'zone_selector': 'Lonestar Sheds, 0701'
    },
    'better built sheds': {
        'profile': 'LSR_TX', 'login': 'lsrbarns@gmail.com', 'store_title': '7 Lonestar', 'store': '7',
        'zone_selector': 'Better Built Sheds, 0720'
    },
    'alpine buildings': {
        'profile': 'CRF_ALPINE', 'login': 'tori.h@evergreenrto.com', 'store_title': '1 Carefree', 'store': '1',
        'zone_selector': 'Alpine Buildings, 0133'
    },
    'genesis buildings, llc': {
        'profile': 'CRF_GENESIS', 'login': 'conrad.s@evergreenrto.com', 'store_title': '1 Carefree', 'store': '1',
        'zone_selector': 'Genesis Buildings, 0135'
    },
    'heritage steel co': {
        'profile': 'CRF_GENESIS', 'login': 'conrad.s@evergreenrto.com', 'store_title': '1 Carefree', 'store': '1',
        'zone_selector': 'Heritage Steel of WA, 0141'
    },
    'westwood sheds': {
        'profile': 'WWP', 'login': 'office@westwoodpayments.com', 'store_title': '11 Westwood', 'store': '11',
        'zone_selector': 'Westwood, 1101'
    },
    'phoenix buildings': {
        'profile': 'DBM_PHOENIX', 'login': 'conrad.s+dbm@evergreenrto.com', 'store_title': '2 Dutch Boy', 'store': '2',
        'zone_selector': 'Phoenix Buildings, 0203'
    },
    'yoder storage sheds llc': {
        'profile': 'RAS_YSS', 'login': 'conrad.s+ysb@evergreenrto.com', 'store_title': '6 RentaShed', 'store': '6',
        'zone_selector': 'Yoder Storage Sheds, 0605'
    },
    '4 seasons buildings of wv llc': {
        'profile': 'CRF_4_SEASONS', 'login': 'conrad.s+wv@evergreenrto.com', 'store_title': '1 Carefree', 'store': '1',
        'zone_selector': '4 Seasons, 0150'
    },
}


# The original Contract_Import.py ships with a ZipTax key hard-coded in the local
# desktop workflow. V3 keeps that local behavior as a fallback, while allowing
# ZIPTAX_API_KEY to override it from .env.
LEGACY_ZIPTAX_API_KEY = 'ziptax_sk_zufh3NPkKP6bgX69OCRCLlQUhm0Ylf'


def effective_ziptax_api_key(explicit: str = '') -> str:
    return clean_text(explicit) or LEGACY_ZIPTAX_API_KEY

KNOWN_STORES = [
    {'store': '1', 'storetitle': '1 Carefree', 'storename': 'Carefree'},
    {'store': '2', 'storetitle': '2 Dutch Boy', 'storename': 'Dutch Boy'},
    {'store': '6', 'storetitle': '6 RentaShed', 'storename': 'RentaShed'},
    {'store': '7', 'storetitle': '7 Lonestar', 'storename': 'Lonestar'},
    {'store': '8', 'storetitle': '8 Magnolia', 'storename': 'Magnolia'},
    {'store': '11', 'storetitle': '11 Westwood', 'storename': 'Westwood'},
]


def clean_text(v: Any) -> str:
    return '' if v is None else str(v).strip()


def parse_money(v: Any) -> float:
    s = clean_text(v).replace('$', '').replace(',', '')
    if not s:
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


def numstr(v: Any, decimals: int = 6) -> str:
    n = parse_money(v)
    if abs(n - round(n)) < 1e-10:
        return str(int(round(n)))
    return f'{n:.{decimals}f}'.rstrip('0').rstrip('.')


def digits(v: Any) -> str:
    return re.sub(r'\D', '', clean_text(v))


def phone(v: Any) -> str:
    """Normalize US phone numbers to ###-###-####.

    ShedSuite occasionally exports an 11-digit US number with the leading
    country code 1.  RTO Pro still needs the same 10-digit local number, so
    ignore that leading 1 before formatting and before the phone comparison
    rule runs.  Other non-10-digit values are left untouched.
    """
    d = digits(v)
    if len(d) == 11 and d.startswith('1'):
        d = d[1:]
    if len(d) == 10:
        return f'{d[:3]}-{d[3:6]}-{d[6:]}'
    return clean_text(v)


def ssn(v: Any) -> str:
    d = digits(v)
    if len(d) == 9:
        return f'{d[:3]}-{d[3:5]}-{d[5:]}'
    return clean_text(v)


def parse_date(v: Any) -> date | None:
    s = clean_text(v)
    if not s:
        return None
    for fmt in ('%m/%d/%Y', '%Y-%m-%d', '%m/%d/%y', '%Y-%m-%dT%H:%M:%S'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(s.replace('Z', '+00:00')).date()
    except Exception:
        return None


def datefmt(d: date | None) -> str:
    return '' if not d else f'{d.month}/{d.day}/{d.year}'


def _safe_due(y: int, m: int, d: int) -> date:
    # Excel DATE rolls overflow days into the next month. Implement that rather than clamping.
    base = date(y, m, 1)
    return base + timedelta(days=max(1, d) - 1)


def next_due(delivered: date | None, due_day: int, today: date | None = None) -> date | None:
    if not delivered or not due_day:
        return None
    today = today or date.today()
    this_due = _safe_due(today.year, today.month, due_day)
    if (this_due - delivered).days < 15:
        # Excel DATE handles month overflow. Use simple month helper.
        def month_plus(delta: int) -> tuple[int, int]:
            n = today.year * 12 + (today.month - 1) + delta
            return n // 12, n % 12 + 1
        y1, m1 = month_plus(1)
        n1 = _safe_due(y1, m1, due_day)
        if (n1 - delivered).days < 15:
            y2, m2 = month_plus(2)
            return _safe_due(y2, m2, due_day)
        return n1
    return this_due


def clean_color(v: Any) -> str:
    s = clean_text(v)
    # Power Query replaces these suffixes with a slash, then merges with '/'.
    s = re.sub(r'\s*-\s*(PAINT|URETHANE)\s*$', '', s, flags=re.I)
    return s.strip().upper()


def reformatted_model(v: Any) -> str:
    txt = clean_text(v).upper()
    words = txt.split()
    dims = []
    other = []
    for w in words:
        parts = w.upper().split('X')
        if len(parts) == 2:
            try:
                float(parts[0]); float(parts[1]); dims.append(w.upper()); continue
            except Exception:
                pass
        other.append(w)
    return (f"{' '.join(dims)}, {' '.join(other)}" if dims else txt).strip(' ,')


def full_description(src: dict[str, str]) -> str:
    model = reformatted_model(src.get('Building Model Variation'))
    colors = '/'.join(x for x in [clean_color(src.get('Siding Color')), clean_color(src.get('Trim Color')), clean_color(src.get('Roof Color'))] if x)
    return (model + (', ' + colors if colors else '')).strip(' ,')


def category_key(model: Any) -> str:
    s = clean_text(model).upper()
    s = re.sub(r'\b\d+\s*[Xx]\s*\d+\b', '', s)
    return re.sub(r'\s+', ' ', s).strip(' ,-')


def load_categories(path: Path) -> dict[str, str]:
    out: dict[str, str] = {}
    with path.open('r', encoding='utf-8-sig', newline='') as f:
        for row in csv.DictReader(f):
            out[clean_text(row.get('Shedsuite Building Model Variation')).upper()] = clean_text(row.get('RTOPro Category'))
    return out


@dataclass
class ReferenceData:
    connected: bool = False
    error: str = ''
    centralserver: list[dict[str, Any]] = field(default_factory=list)
    agents: list[dict[str, Any]] = field(default_factory=list)
    zones: list[dict[str, Any]] = field(default_factory=list)
    taxzones: list[dict[str, Any]] = field(default_factory=list)
    dealers: list[dict[str, Any]] = field(default_factory=list)
    existing_serials: list[dict[str, Any]] = field(default_factory=list)
    existing_names: list[dict[str, Any]] = field(default_factory=list)
    existing_contracts: list[dict[str, Any]] = field(default_factory=list)
    last_accounts: list[dict[str, Any]] = field(default_factory=list)

    def as_client_dict(self) -> dict[str, Any]:
        stores = self.centralserver or KNOWN_STORES
        return {
            'connected': self.connected,
            'error': self.error,
            'stores': stores,
            'agents': self.agents,
            'zones': self.zones,
            'taxzones': self.taxzones,
            'dealers': self.dealers,
        }


def _query_dicts(cur, sql: str) -> list[dict[str, Any]]:
    cur.execute(sql)
    cols = [str(d[0]).lower() for d in cur.description]
    result = []
    for vals in cur.fetchall():
        result.append({k: ('' if v is None else str(v).strip()) for k, v in zip(cols, vals)})
    return result


def load_reference_data(dsn: str, user: str = '', password: str = '') -> ReferenceData:
    ref = ReferenceData(centralserver=[dict(x) for x in KNOWN_STORES])
    if not clean_text(dsn):
        ref.error = 'No Firebird ODBC DSN configured. Known stores are available, but dealer/zone/tax lookups require Firebird.'
        return ref
    if pyodbc is None:
        ref.error = 'pyodbc is not installed. Run run_windows.bat so dependencies are installed.'
        return ref
    parts = [f'DSN={dsn}']
    if user:
        parts.append(f'UID={user}')
    if password:
        parts.append(f'PWD={password}')
    try:
        cn = pyodbc.connect(';'.join(parts) + ';', timeout=10)
        cur = cn.cursor()
        ref.centralserver = _query_dicts(cur, "SELECT store, storename, TRIM(LEADING '#' FROM storetitle) AS storetitle FROM centralserver ORDER BY store ASC")
        ref.last_accounts = _query_dicts(cur, "WITH FilteredCustomers AS (SELECT c.Store, c.Account, c.CustomerDate FROM customers c JOIN (SELECT Store, MAX(CustomerDate) AS MaxDate FROM customers GROUP BY Store) latest ON c.Store = latest.Store AND c.CustomerDate = latest.MaxDate) SELECT Store, MAX(Account) AS LastAccount FROM FilteredCustomers GROUP BY Store")
        ref.agents = _query_dicts(cur, "select STORE, AGENT from AGENTTABLE where recno not in (1,110,117,118,120,110,115,116,121) order by store asc")
        ref.zones = _query_dicts(cur, "select zone, description, store, description || ', ' || zone as Selector from zonetable")
        ref.taxzones = _query_dicts(cur, """SELECT zone,state,county,city,td1,td2,td3,taxrate,
            REPLACE(REPLACE(TRIM(COALESCE(NULLIF(TRIM(COALESCE(state,'')),''),'') || CASE WHEN NULLIF(TRIM(COALESCE(state,'')),'') IS NOT NULL THEN ' ' ELSE '' END || COALESCE(NULLIF(TRIM(COALESCE(county,'')),''),'') || CASE WHEN NULLIF(TRIM(COALESCE(county,'')),'') IS NOT NULL THEN ' ' ELSE '' END || COALESCE(NULLIF(TRIM(COALESCE(city,'')),''),'') || CASE WHEN NULLIF(TRIM(COALESCE(city,'')),'') IS NOT NULL THEN ' ' ELSE '' END || COALESCE(NULLIF(TRIM(COALESCE(td1,'')),''),'') || CASE WHEN NULLIF(TRIM(COALESCE(td1,'')),'') IS NOT NULL THEN ' ' ELSE '' END || COALESCE(NULLIF(TRIM(COALESCE(td2,'')),''),'') || CASE WHEN NULLIF(TRIM(COALESCE(td2,'')),'') IS NOT NULL THEN ' ' ELSE '' END || COALESCE(NULLIF(TRIM(COALESCE(td3,'')),''),'')), '  ', ' '), '  ', ' ') AS description
            FROM taxzones WHERE isactive = 1""")
        ref.dealers = _query_dicts(cur, "Select ID, companydealername, store from dealertable")
        ref.existing_serials = _query_dicts(cur, "select Model, Serial, Status, Store from Inventory")
        ref.existing_names = _query_dicts(cur, "Select Account, Name From Customers")
        ref.existing_contracts = _query_dicts(cur, "Select contract, account from contracts where contype = 'R' order by contractdate desc")
        cn.close()
        ref.connected = True
        ref.error = ''
    except Exception as e:
        ref.error = f'{type(e).__name__}: {e}'
    return ref


def company_rule(src: dict[str, str]) -> dict[str, str]:
    company = clean_text(src.get('Company Name')).casefold()
    rule = dict(COMPANY_RULES.get(company, {}))
    if not rule:
        return {'profile': '', 'login': '', 'store_title': '', 'store': '', 'zone_selector': 'Set Manually'}
    if rule.get('profile_by_state'):
        st = clean_text(src.get('Physical Destination State')).upper()
        profile = f'MAG_{st}' if f'MAG_{st}' in MODEL_SERIES else 'MAG_GA'
        rule['profile'] = profile
    return rule


def build_next_model_suggestions(ref: ReferenceData) -> dict[str, str]:
    used = [clean_text(x.get('model')) for x in ref.existing_serials]
    result: dict[str, str] = {}
    for profile, (prefix, lo, hi, pad3) in MODEL_SERIES.items():
        vals = []
        for m in used:
            if not m.startswith(prefix):
                continue
            tail = m[len(prefix):]
            if tail.isdigit() and lo <= int(tail) <= hi:
                vals.append(int(tail))
        # Workbook's formulas seed with the lower bound and then +1.
        n = max(vals + [lo]) + 1
        result[profile] = prefix + (f'{n:03d}' if pad3 else str(n))
    return result


def find_stock_model(ref: ReferenceData, serial: str) -> str:
    sf = clean_text(serial).casefold()
    for row in ref.existing_serials:
        if clean_text(row.get('serial')).casefold() == sf and clean_text(row.get('status')).upper() == 'STOCK':
            return clean_text(row.get('model'))
    return ''


def find_store_id(ref: ReferenceData, title: str, fallback: str = '') -> str:
    tf = clean_text(title).casefold()
    for r in ref.centralserver:
        if clean_text(r.get('storetitle')).casefold() == tf:
            return clean_text(r.get('store'))
    if fallback:
        return clean_text(fallback)
    m = re.match(r'\s*(\d+)', clean_text(title))
    return m.group(1) if m else ''


def find_zone(ref: ReferenceData, selector: str, store: str = '') -> str:
    sf = clean_text(selector).casefold()
    candidates = [r for r in ref.zones if clean_text(r.get('selector')).casefold() == sf]
    if store:
        store_candidates = [r for r in candidates if clean_text(r.get('store')) == clean_text(store)]
        if store_candidates:
            candidates = store_candidates
    return clean_text(candidates[0].get('zone')) if candidates else ''


def normalize_name(v: Any) -> str:
    return re.sub(r'[^a-z0-9]+', '', clean_text(v).casefold())


def find_dealer(ref: ReferenceData, dealer: str, store: str = '') -> tuple[str, str]:
    """Return dealer id and match note. Excel XLOOKUP matches dealer name only; we prefer selected store then fall back to exact name."""
    name = clean_text(dealer)
    if not name:
        return '', 'blank source dealer'
    exact = [r for r in ref.dealers if clean_text(r.get('companydealername')).casefold() == name.casefold()]
    if store:
        same = [r for r in exact if clean_text(r.get('store')) == clean_text(store)]
        if len(same) == 1:
            return clean_text(same[0].get('id')), 'exact name + store'
    if len(exact) == 1:
        return clean_text(exact[0].get('id')), 'exact name'
    normalized = [r for r in ref.dealers if normalize_name(r.get('companydealername')) == normalize_name(name)]
    if store:
        same = [r for r in normalized if clean_text(r.get('store')) == clean_text(store)]
        if len(same) == 1:
            return clean_text(same[0].get('id')), 'normalized name + store'
    if len(normalized) == 1:
        return clean_text(normalized[0].get('id')), 'normalized name'
    return '', 'no unique dealer match'


def dealer_name_for_id(ref: ReferenceData, dealer_id: str) -> str:
    for r in ref.dealers:
        if clean_text(r.get('id')) == clean_text(dealer_id):
            return clean_text(r.get('companydealername'))
    return ''


def find_agent(ref: ReferenceData, store: str) -> str:
    for r in ref.agents:
        if clean_text(r.get('store')) == clean_text(store):
            return clean_text(r.get('agent'))
    return ''


def unique_taxzone_by_state_rate(ref: ReferenceData, state: str, rate: float) -> tuple[str, str]:
    matches = []
    for r in ref.taxzones:
        try:
            rr = float(clean_text(r.get('taxrate')) or 0)
        except Exception:
            continue
        if clean_text(r.get('state')).upper() == state.upper() and abs(rr - rate) < 1e-8:
            matches.append(r)
    if len(matches) == 1:
        return clean_text(matches[0].get('zone')), numstr(matches[0].get('taxrate'))
    return '', numstr(rate) if rate else ''


def ziptax_detail(address: str, api_key: str) -> dict[str, Any]:
    api_key = effective_ziptax_api_key(api_key)
    if not clean_text(address):
        return {'error': 'No delivery address available for ZipTax'}
    try:
        response = requests.get(
            'https://api.zip-tax.com/request/v60',
            params={'address': address, 'format': 'json', 'countryCode': 'USA'},
            headers={'X-API-KEY': api_key, 'Content-Type': 'application/json'},
            timeout=10,
        )
        response.raise_for_status()
        data = response.json()
        meta = data.get('metadata', {}).get('response', {}) or {}
        code = meta.get('code')
        name = clean_text(meta.get('name'))
        if code == 108:
            return {'limit': True, 'response_code': code, 'response_name': name}
        if code not in (None, 100, '100'):
            return {'error': name or f'ZipTax response code {code}', 'response_code': code, 'response_name': name}
        detail = data.get('addressDetail', {}) or {}
        inc = detail.get('incorporated')
        if isinstance(inc, str):
            inc = inc.lower() == 'true'
        county = ''; city = ''; city_code = ''
        latitude = clean_text(detail.get('latitude'))
        longitude = clean_text(detail.get('longitude'))
        for item in data.get('baseRates', []) or []:
            if item.get('jurType') == 'US_COUNTY_SALES_TAX' and not county:
                county = clean_text(item.get('jurName')).upper()
            if item.get('jurType') == 'US_CITY_SALES_TAX':
                city = clean_text(item.get('jurName')).upper() or city
                city_code = clean_text(item.get('jurTaxCode')) or city_code
        normalized = clean_text(detail.get('normalizedAddress'))
        if not city and normalized:
            parts = [x.strip() for x in normalized.split(',')]
            if len(parts) > 1:
                city = parts[1].upper()
        total_rate = ''
        for item in data.get('taxSummaries', []) or []:
            if clean_text(item.get('taxType')).upper() == 'SALES_TAX':
                total_rate = numstr(item.get('rate'))
                break
        return {
            'county': county,
            'city': city,
            'city_code': city_code,
            'incorporated': bool(inc),
            'limit': False,
            'normalized_address': normalized,
            'total_rate': total_rate,
            'latitude': latitude,
            'longitude': longitude,
            'response_code': code,
            'response_name': name,
        }
    except Exception as e:
        return {'error': str(e)}


def _rate_float(v: Any) -> float:
    try:
        return float(clean_text(v) or 0)
    except Exception:
        return 0.0


def resolve_taxzone_from_detail(ref: ReferenceData, state: str, csv_rate: float, detail: dict[str, Any]) -> tuple[str, str, str, str]:
    """Resolve the best active RTO tax zone using ZipTax jurisdiction data.

    We prefer county/city/incorporation matching, then a unique state/rate match.
    The user can always override the result from the review dropdown.
    """
    state = clean_text(state).upper()
    county = clean_text(detail.get('county')).upper()
    city = clean_text(detail.get('city')).upper()
    inc = detail.get('incorporated')
    zt_rate = _rate_float(detail.get('total_rate'))
    want_rate = zt_rate or csv_rate
    rows = [r for r in ref.taxzones if clean_text(r.get('state')).upper() == state]
    chosen = None
    note = ''

    if detail.get('limit'):
        return '', numstr(want_rate) if want_rate else '', county, 'ZipTax API limit reached; choose tax zone manually'
    if detail.get('error'):
        note = 'ZipTax lookup failed; choose tax zone manually'

    if state == 'KY':
        chosen = next((r for r in rows if clean_text(r.get('zone')).upper() == 'KY KY'), None)
    elif state == 'OR':
        chosen = next((r for r in rows if clean_text(r.get('zone')).upper() == 'OR OR'), None)
    elif state == 'WV' and clean_text(detail.get('city_code')):
        want = 'WV-54-' + clean_text(detail.get('city_code'))
        chosen = next((r for r in rows if clean_text(r.get('zone')) == want), None)
        if not chosen:
            note = f'{want} does not exist in RTO Pro; add it or choose manually'

    # General county/city match works for GA/NC/SC and can also resolve other
    # states that the old workbook forced to manual selection.
    if not chosen and county:
        county_rows = [r for r in rows if clean_text(r.get('county')).upper() == county]
        if len(county_rows) == 1:
            chosen = county_rows[0]
        elif len(county_rows) > 1:
            if inc is True and city:
                city_rows = [r for r in county_rows if clean_text(r.get('city')).upper() == city]
                if len(city_rows) == 1:
                    chosen = city_rows[0]
                if not chosen:
                    desc_rows = [r for r in county_rows if city in clean_text(r.get('description')).upper()]
                    if len(desc_rows) == 1:
                        chosen = desc_rows[0]
            elif inc is False:
                uninc = [r for r in county_rows if clean_text(r.get('city')).upper() in {'', 'UNINCORPORATED', 'NO LOCAL OPTION'}]
                if len(uninc) == 1:
                    chosen = uninc[0]

            if not chosen and want_rate:
                rate_rows = [r for r in county_rows if abs(_rate_float(r.get('taxrate')) - want_rate) < 1e-8]
                if len(rate_rows) == 1:
                    chosen = rate_rows[0]

    if chosen:
        return clean_text(chosen.get('zone')), numstr(chosen.get('taxrate')), county, note or 'ZipTax matched RTO tax zone automatically'

    # Final fallback: if exactly one active zone in the state has the ZipTax/CSV rate.
    zone, rate = unique_taxzone_by_state_rate(ref, state, want_rate)
    if zone:
        return zone, rate, county, note or 'ZipTax rate uniquely matched an RTO tax zone'
    return '', numstr(want_rate) if want_rate else '', county, note or 'ZipTax found the jurisdiction, but the RTO tax zone is ambiguous; choose from the dropdown'


def resolve_taxzone(ref: ReferenceData, state: str, address: str, csv_rate: float, api_key: str = '', detail: dict[str, Any] | None = None) -> tuple[str, str, str, str]:
    detail = detail if detail is not None else ziptax_detail(address, api_key)
    return resolve_taxzone_from_detail(ref, state, csv_rate, detail)


def phone5_and_comment(src: dict[str, str]) -> tuple[str, str, str]:
    """Apply the user's RTO phone rule and return PHONE5, COMMENTS, selected source.

    Primary always stays CELL.  The second/other number uses:
      - secondary empty/-- -> Ref 1
      - secondary == primary -> Ref 1
      - secondary == Ref 2 -> Ref 2
      - secondary == Ref 1 -> Ref 1
      - otherwise -> Secondary
    """
    p = phone(src.get('Primary Phone'))
    s = phone(src.get('Secondary Phone'))
    r1 = phone(src.get('Ref 1 Phone'))
    r2 = phone(src.get('Ref 2 Phone'))
    raw_secondary = clean_text(src.get('Secondary Phone'))
    sd = digits(s); pd = digits(p); r1d = digits(r1); r2d = digits(r2)

    if not sd or raw_secondary in {'--', '-'}:
        source = 'ref1'
        other = r1
    elif pd and sd == pd:
        source = 'ref1'
        other = r1
    elif r2d and sd == r2d:
        source = 'ref2'
        other = r2
    elif r1d and sd == r1d:
        source = 'ref1'
        other = r1
    else:
        source = 'secondary'
        other = s

    if source == 'ref1':
        comment = f"OTHER IS {clean_text(src.get('Ref 1 Relationship')).upper()} {clean_text(src.get('Ref 1 Name')).upper()}".strip()
    elif source == 'ref2':
        comment = f"OTHER IS {clean_text(src.get('Ref 2 Relationship')).upper()} {clean_text(src.get('Ref 2 Name')).upper()}".strip()
    elif other:
        comment = 'OTHER IS SECONDARY PHONE'
    else:
        comment = ''
    return other, comment, source

def category_for(src: dict[str, str], categories: dict[str, str]) -> str:
    key = category_key(src.get('Building Model Variation'))
    if key in categories:
        return categories[key].lower()
    # Change Log V1.1.7 says category became left 30 characters of model variation minus dimensions.
    return key[:30].strip().lower()


# Canonical RTO Pro Brand/Vendor names.  ShedSuite company names can include
# legal/entity wording (LLC, Buildings LLC, etc.) or small naming variations.
# BRAND1 and VENDOR1 should always use the short canonical value below.
BRAND_ALIASES = (
    ('ALPINE', ('alpine',)),
    ('4 SEASONS', ('4seasons', 'fourseasons', '4season', 'fourseason')),
    ('GENESIS', ('genesis',)),
    ('SMART SHED', ('smartshed', 'smartsheds')),
    ('WESTWOOD SHEDS', ('westwood',)),
    ('LONESTAR SHEDS', ('lonestar', 'lonestarsheds')),
    ('TRUE BUILT', ('truebuilt',)),
    ('YODER STORAGE', ('yoderstorage', 'yoder')),
    ('PHOENIX', ('phoenix',)),
)


def normalized_brand(company: Any) -> str:
    c = clean_text(company)
    if not c:
        return ''

    # Compacting makes punctuation, spaces and legal suffixes irrelevant:
    #   "Genesis Buildings, LLC" -> "genesisbuildingsllc"
    #   "Lone Star Sheds LLC"    -> "lonestarshedsllc"
    compact = re.sub(r'[^a-z0-9]+', '', c.casefold())

    # Common word-number spelling for 4 Seasons.
    compact = compact.replace('four', '4')

    for canonical, aliases in BRAND_ALIASES:
        if any(alias in compact for alias in aliases):
            return canonical

    # Preserve the old behavior for an unknown manufacturer rather than
    # blanking or guessing it.
    return c.upper()


def sac_date(src: dict[str, str], login: str) -> str:
    created = parse_date(src.get('Date Created'))
    delivered = parse_date(src.get('Date Delivered'))
    login_cf = clean_text(login).casefold()
    if login_cf == 'lsrbarns@gmail.com':
        return 'Check If Applicable'
    if login_cf == 'office@westwoodpayments.com':
        return datefmt(delivered + timedelta(days=90)) if delivered else ''
    known = {
        'mary@carefreerentals.com', 'conrad.s@evergreenrto.com', 'rentabarn2@gmail.com',
        'office@magnoliarto.com', 'conrad.s+wv@evergreenrto.com', 'conrad.s+dbm@evergreenrto.com',
        'joy.f@evergreenrto.com', 'conrad.s+ysb@evergreenrto.com', 'tori.h@evergreenrto.com'
    }
    if login_cf in known:
        return datefmt(created + timedelta(days=90)) if created else ''
    return 'Set Manually.'


def transform_rows(rows: list[dict[str, str]], categories: dict[str, str], ref: ReferenceData, today: date | None = None, ziptax_api_key: str = '') -> tuple[list[dict[str, str]], list[str]]:
    today = today or date.today()
    warnings: list[str] = []
    next_models = build_next_model_suggestions(ref)
    seen_names = {clean_text(x.get('name')).upper() for x in ref.existing_names}
    seen_contracts = {clean_text(x.get('contract')) for x in ref.existing_contracts}
    serial_rows: dict[str, list[dict[str, Any]]] = {}
    for x in ref.existing_serials:
        serial_rows.setdefault(clean_text(x.get('serial')), []).append(x)

    # ImportCSVToShedsuiteData sorts by physical destination state (BR) ascending.
    rows = sorted(rows, key=lambda x: (clean_text(x.get('Physical Destination State')).upper(), clean_text(x.get('Customer Last Name')).upper(), clean_text(x.get('Customer First Name')).upper()))
    out: list[dict[str, str]] = []

    for idx, src in enumerate(rows):
        first = clean_text(src.get('Customer First Name')).upper()
        last = clean_text(src.get('Customer Last Name')).upper()
        name = f'{last}, {first}'.strip(', ')
        serial = clean_text(src.get('Serial Number'))
        order_id = clean_text(src.get('Customer Order Id'))
        inventory_id = clean_text(src.get('Inventory Id'))
        rule = company_rule(src)
        store = find_store_id(ref, rule.get('store_title', ''), rule.get('store', ''))
        store_title = rule.get('store_title', '')
        zone_selector = rule.get('zone_selector', '')
        zone = find_zone(ref, zone_selector, store)
        stock_model = find_stock_model(ref, serial)
        default_model = stock_model or inventory_id or 'Not Found'
        default_contract_check = order_id
        # Import Data SS AV and CI both point to Import Tools M (Set Model) in V1.1.7.6.
        exported_contract = default_model
        next_model = next_models.get(rule.get('profile', ''), '')

        dealer_id, dealer_match = find_dealer(ref, src.get('Dealer', ''), store)
        agent = find_agent(ref, store)

        delivered = parse_date(src.get('Date Delivered'))
        due_day = int(parse_money(src.get('Day Of Month Payment Is Due')) or 0)
        due = next_due(delivered, due_day, today)
        sac = sac_date(src, rule.get('login', ''))
        tax_rate_csv = parse_money(src.get('Sum Of All Tax Rates'))
        state = clean_text(src.get('Physical Destination State')).upper()
        tax_address = f"{clean_text(src.get('Physical Destination Street'))} {clean_text(src.get('Physical Destination City'))} {state} {clean_text(src.get('Physical Destination Zip'))}"
        tax_detail = ziptax_detail(tax_address, ziptax_api_key)
        taxzone, taxrate, county, tax_note = resolve_taxzone(ref, state, tax_address, tax_rate_csv, ziptax_api_key, detail=tax_detail)

        p5, comment, phone_source = phone5_and_comment(src)
        mailing_street = clean_text(src.get('Customer Mailing Street')).upper()
        mailing_city = clean_text(src.get('Customer Mailing City')).upper()
        mailing_state = clean_text(src.get('Customer Mailing State')).upper()
        mailing_zip = clean_text(src.get('Customer Mailing Zip'))
        phys_street = clean_text(src.get('Physical Destination Street')).upper()
        phys_city = clean_text(src.get('Physical Destination City')).upper()
        phys_state = state
        phys_zip = clean_text(src.get('Physical Destination Zip'))
        physical_diff = (mailing_street, mailing_city, mailing_state, mailing_zip) != (phys_street, phys_city, phys_state, phys_zip)
        direction = f'PHYSICAL - {phys_street}, {phys_city}, {phys_state} {phys_zip}'.strip() if physical_diff else ''

        birth = parse_date(src.get('Customer Date Of Birth'))
        birth_plus_one = birth + timedelta(days=1) if birth else None  # Import Tools J2 is checked in V1.1.7.6.
        initial = parse_money(src.get('Initial Payment'))
        delivery_fee = parse_money(src.get('Delivery Fee'))
        sec = parse_money(src.get('Security Deposit'))
        reserve = parse_money(src.get('Purchase Reserve'))
        amount_paid = initial - delivery_fee - sec - reserve
        brand = normalized_brand(src.get('Company Name'))

        r: dict[str, str] = {h: '' for h in RTO_HEADERS}
        r.update({
            'NAME': name,
            'ALTFIRSTNAME': clean_text(src.get('Co Renter First Name')),
            'ALTLASTNAME': clean_text(src.get('Co Renter Last Name')),
            'ADDRESS': mailing_street,
            'CITY': mailing_city,
            'STATE': mailing_state,
            'ZIP': mailing_zip,
            'CELL': phone(src.get('Primary Phone')),
            'PHONE5': p5,
            'DIRECTIONS': direction,
            'ZONE': zone,
            'SS1': ssn(src.get('Customer Social Security Number')),
            'DOB1': datefmt(birth_plus_one),
            'DL1': clean_text(src.get('Customer Identification Card Number')),
            'EMAIL': clean_text(src.get('Customer Email')).lower(),
            'WORK1': clean_text(src.get('Customer Employer Name')),
            'WORKPH1': clean_text(src.get('Customer Employer Phone')),
            # RTO Pro manual: CELLOPT=2 means primary cell opted in.
            # CELLOPT2=3 means PHONE5 is a cell number and is opted OUT of SMS.
            'CELLOPT': '2' if phone(src.get('Primary Phone')) else '',
            'CELLOPT2': '3' if p5 else '',
            'COUNTY': county,
            'REFERENCE1': clean_text(src.get('Ref 1 Name')).upper(),
            'REFRELATION1': clean_text(src.get('Ref 1 Relationship')).upper(),
            'REFPHONE1': phone(src.get('Ref 1 Phone')),
            'REFERENCE2': clean_text(src.get('Ref 2 Name')).upper(),
            'REFRELATION2': clean_text(src.get('Ref 2 Relationship')).upper(),
            'REFPHONE2': phone(src.get('Ref 2 Phone')),
            'CONTRACT': exported_contract,
            'DUEDATE': datefmt(due),
            'CONTRACTDATE': datefmt(today) if clean_text(src.get('Ref 2 Phone')) else '',
            'GRP': numstr(src.get('Loss Damage Waiver Monthly Charge')) if clean_text(src.get('Loss Damage Waiver Monthly Charge')) else '',
            'AMTPAID': numstr(amount_paid),
            'EXTRARENT': numstr(sec) if sec else '',
            'PAIDDOWN': numstr(reserve) if reserve else '',
            'CONTYPE': 'R',
            'TERM': 'M',
            'PAYMENTS': clean_text(src.get('Months Of Term')),
            'BILLED': '1',
            'CASHPRICE': numstr(src.get('Building Cost')) if parse_money(src.get('Building Cost')) else '',
            'SACDATE': sac,
            'PAYOFFDISCOUNT': numstr(1 - parse_money(src.get('Early Purchase Percentage')) / 100) if clean_text(src.get('Early Purchase Percentage')) else '',
            'AUTOCHARGEB4DUE': '0',
            'TAXRATE': taxrate,
            'TAXZONE': taxzone,
            'STORE': store,
            'DEALERID': dealer_id,
            'del_address1': phys_street,
            'del_city': phys_city,
            'del_state': phys_state,
            'del_zip': phys_zip,
            'MODEL1': default_model,
            'SERIAL1': serial,
            'DESCRIPTION1': full_description(src),
            'RATE1': numstr(parse_money(src.get('Monthly Subtotal')) - parse_money(src.get('Loss Damage Waiver Monthly Charge'))),
            'BRAND1': brand,
            'VENDOR1': brand,
            'COST1': numstr(src.get('Building Cost')) if parse_money(src.get('Building Cost')) else '',
            'AGENT1': agent,
            'CATEGORY1': category_for(src, categories),
            'RETAIL1': numstr(src.get('Building Cost')) if parse_money(src.get('Building Cost')) else '',
            # Do not send the generated OTHER IS note through the CSV COMMENTS
            # field.  RTO Pro displays that as an application comment on the
            # pending-app screen.  V7.9 keeps it as customer-comment metadata
            # and syncs it to CUSTOMERS.COMMENTS after the account is created.
            'COMMENTS': '',
            'CONDITION1': '',
        })

        account_suggestions = [clean_text(x.get('account')) for x in ref.existing_names if clean_text(x.get('name')).upper() == name and clean_text(x.get('account'))]
        account_suggestions = list(dict.fromkeys(account_suggestions))

        row_warnings: list[str] = []
        if not store:
            row_warnings.append('Store is unresolved')
        if not dealer_id:
            row_warnings.append('Dealer ID is unresolved')
        if not zone:
            row_warnings.append('Zone is unresolved')
        if tax_note:
            row_warnings.append(tax_note)
        if not taxzone:
            row_warnings.append('Tax zone needs selection')
        if name in seen_names:
            row_warnings.append('Exact customer name already exists in RTO Pro')
        if default_contract_check and default_contract_check in seen_contracts:
            row_warnings.append(f'Set Contract value {default_contract_check} already exists')
        if serial in serial_rows:
            statuses = ', '.join(sorted({clean_text(x.get('status')) for x in serial_rows[serial] if clean_text(x.get('status'))}))
            row_warnings.append('Serial already exists in inventory' + (f' ({statuses})' if statuses else ''))
        if stock_model:
            row_warnings.append(f'Existing STOCK serial found; model defaulted to {stock_model}')

        r.update({
            '_row_id': str(idx),
            '_source_order_id': order_id,
            '_source_inventory_id': inventory_id,
            '_source_company': clean_text(src.get('Company Name')),
            '_source_dealer': clean_text(src.get('Dealer')),
            '_source_state': state,
            '_source_model_variation': clean_text(src.get('Building Model Variation')),
            '_source_condition': clean_text(src.get('Condition')),
            '_source_date_created': clean_text(src.get('Date Created')),
            '_source_date_delivered': clean_text(src.get('Date Delivered')),
            '_primary_phone': phone(src.get('Primary Phone')),
            '_secondary_phone': phone(src.get('Secondary Phone')),
            '_ref1_name': clean_text(src.get('Ref 1 Name')).upper(),
            '_ref1_relation': clean_text(src.get('Ref 1 Relationship')).upper(),
            '_ref1_phone': phone(src.get('Ref 1 Phone')),
            '_ref2_name': clean_text(src.get('Ref 2 Name')).upper(),
            '_ref2_relation': clean_text(src.get('Ref 2 Relationship')).upper(),
            '_ref2_phone': phone(src.get('Ref 2 Phone')),
            '_phone_other_source': phone_source,
            '_phone_other_manual': False,
            '_customer_comment': comment,
            '_customer_comment_sync_status': 'Waiting for RTO account',
            '_ziptax_county': clean_text(tax_detail.get('county')).upper(),
            '_ziptax_city': clean_text(tax_detail.get('city')).upper(),
            '_ziptax_incorporated': tax_detail.get('incorporated') if 'incorporated' in tax_detail else None,
            '_ziptax_normalized_address': clean_text(tax_detail.get('normalized_address')),
            '_ziptax_total_rate': clean_text(tax_detail.get('total_rate')),
            '_ziptax_response_code': clean_text(tax_detail.get('response_code')),
            '_ziptax_note': tax_note,
            '_store_title': store_title,
            '_login': rule.get('login', ''),
            '_profile': rule.get('profile', ''),
            '_zone_selector': zone_selector,
            '_set_contract': default_contract_check,
            '_next_model': next_model,
            '_stock_model': stock_model,
            '_dealer_match': dealer_match,
            '_dealer_name': dealer_name_for_id(ref, dealer_id),
            '_account_suggestions': account_suggestions,
            '_warnings': row_warnings,
            '_pdf_name': f'{last} {first} {order_id}'.strip(),
            '_signed_url': clean_text(src.get('Signed Combined Documents')) or clean_text(src.get('Rto Contract')),
            '_contract_url': clean_text(src.get('Rto Contract')),
            '_invoice_url': clean_text(src.get('Invoice')),
            '_landowner_url': clean_text(src.get('Landowner Permission Form')),
            '_idscan_url': clean_text(src.get('Customer Identification Card Scan')),
        })
        out.append(r)
        warnings.extend(f'{name}: {w}' for w in row_warnings)

    if not ref.connected:
        warnings.insert(0, 'Firebird is not connected. Store defaults still work, but dealer/zone/tax/duplicate lookup data may be incomplete until the DSN connects.')
    return out, list(dict.fromkeys(warnings))
