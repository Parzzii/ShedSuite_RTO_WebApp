from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import fitz

try:
    from pypdf import PdfReader
except Exception:  # optional fallback
    PdfReader = None

try:
    import pdfplumber
except Exception:  # optional fallback
    pdfplumber = None


SECTION_NAMES = {
    'customer information': 'customer',
    'unit information': 'unit',
    'contract information': 'contract',
    'co-renter': 'corenter',
    'co renter': 'corenter',
    'additional contact 1': 'ref1',
    'additional contact 2': 'ref2',
    'employer/source of income': 'employer',
    'employer / source of income': 'employer',
    'landlord': 'landlord',
}


# V7.20 PDF EPO schedules.  These are the provider/term percentages supplied by
# the business and intentionally feed the exact same ShedSuite source field
# (Early Purchase Percentage) used by rto_transform.py.  The existing transform
# remains authoritative for converting the percentage to RTO Pro PAYOFFDISCOUNT.
PDF_EPO_SCHEDULES = {
    'DBM': {24: 65, 36: 60, 48: 55, 60: 45},
    'CHOICE CAPITAL': {24: 65, 36: 60, 48: 55, 60: 45},
    'RENTABARN': {24: 70, 36: 60, 48: 50, 60: 45},
    'X-GEN': {24: 70, 36: 60, 48: 55, 54: 55, 60: 50, 72: 45},
}


def _canonical_epo_provider(value: str) -> str:
    raw = re.sub(r'\s+', ' ', str(value or '')).strip().casefold()
    compact = re.sub(r'[^a-z0-9]+', '', raw)
    if 'choicecapital' in compact:
        return 'CHOICE CAPITAL'
    if 'rentabarn' in compact or 'wolfvalley' in compact or compact == 'wvb' or compact.startswith('wvb'):
        return 'RENTABARN'
    if compact == 'dbm' or compact.startswith('dbm'):
        return 'DBM'
    if 'xgen' in compact:
        return 'X-GEN'
    return str(value or '').strip().upper()


def _term_months(value: str) -> int | None:
    m = re.search(r'\b(\d{1,3})\b', str(value or ''))
    return int(m.group(1)) if m else None


def pdf_epo_percentage(provider: str, term: str) -> str:
    """Return the configured PDF EPO percentage for provider + term.

    Empty string means the combination is not configured and must not be
    guessed.  Keeping this as a public helper also lets app.py recalculate the
    value if the user manually changes the RTO term during review.
    """
    canonical = _canonical_epo_provider(provider)
    months = _term_months(term)
    if months is None:
        return ''
    value = PDF_EPO_SCHEDULES.get(canonical, {}).get(months)
    return str(value) if value is not None else ''


def _clean(value: Any) -> str:
    return re.sub(r'\s+', ' ', '' if value is None else str(value)).strip()


def _key(value: Any) -> str:
    s = _clean(value).casefold()
    s = s.replace('&', ' and ')
    s = re.sub(r'[^a-z0-9]+', ' ', s)
    return re.sub(r'\s+', ' ', s).strip()


def _money(value: str) -> str:
    s = _clean(value)
    m = re.search(r'-?\$?\s*([0-9][0-9,]*(?:\.\d+)?)', s)
    return m.group(1).replace(',', '') if m else s


def _percentage(value: str) -> str:
    s = _clean(value)
    m = re.search(r'(-?\d+(?:\.\d+)?)\s*%?', s)
    return m.group(1) if m else ''


def _date_from_value(value: str) -> str:
    """Extract an explicit m/d/yyyy date from a mixed label value.

    RentaBarn's 90 Days SAC row is commonly formatted like
    ``YES (10/13/2026)``.  That parenthesized date is authoritative; do not
    recompute it from Contract Date.
    """
    s = _clean(value)
    m = re.search(r'\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b', s)
    if not m:
        return ''
    return f'{int(m.group(1)):02d}/{int(m.group(2)):02d}/{m.group(3)}'


def _normalize_size(value: str) -> str:
    """Normalize shed dimensions to the compact RTO form (e.g. 10x12)."""
    s = _clean(value).lower().replace('×', 'x')
    # Feet/inch marks and words are display noise for MODEL/DESCRIPTION mapping.
    s = s.replace('′', '').replace('″', '').replace("'", '').replace('\"', '')
    s = re.sub(r'\b(?:ft|feet|foot|in|inch|inches)\b', '', s, flags=re.I)
    nums = re.findall(r'\d+(?:\.\d+)?', s)
    if len(nums) >= 2:
        return 'x'.join(nums)
    # If the source is already a compact non-standard dimension, at least strip spaces.
    return re.sub(r'\s+', '', s)


def _yes(value: str) -> bool:
    return _key(value) in {'yes', 'y', 'true', '1', 'on'}


def _amount_in_parentheses(value: str) -> str:
    s = _clean(value)
    matches = re.findall(r'\$\s*([0-9][0-9,]*(?:\.\d+)?)', s)
    if matches:
        return matches[-1].replace(',', '')
    return _money(s)


def _split_name(value: str) -> tuple[str, str]:
    s = _clean(value)
    if not s:
        return '', ''
    if ',' in s:
        last, first = [x.strip() for x in s.split(',', 1)]
        return first, last
    parts = s.split()
    if len(parts) == 1:
        return parts[0], ''
    return ' '.join(parts[:-1]), parts[-1]


def _split_city_state_zip(value: str) -> tuple[str, str, str]:
    s = _clean(value)
    if not s:
        return '', '', ''
    # Common forms: "Mocksville, TX 75149" or "Mocksville TX 75149".
    m = re.match(r'^(.*?)(?:,\s*|\s+)([A-Za-z]{2})\s+(\d{5}(?:-\d{4})?)$', s)
    if m:
        return m.group(1).strip(' ,'), m.group(2).upper(), m.group(3)
    m = re.match(r'^(.*?)(?:,\s*|\s+)([A-Za-z]{2})$', s)
    if m:
        return m.group(1).strip(' ,'), m.group(2).upper(), ''
    return s, '', ''


def _canonical_company(value: str) -> str:
    raw = _clean(value)
    compact = re.sub(r'[^a-z0-9]+', '', raw.casefold()).replace('four', '4')
    rules = [
        (('alpine',), 'Alpine Buildings'),
        (('4seasons', '4season'), '4 Seasons Buildings of WV LLC'),
        (('genesis',), 'Genesis Buildings, LLC'),
        (('smartshed', 'smartsheds'), 'Smart Shed'),
        (('westwood',), 'Westwood Sheds'),
        (('lonestar',), 'Lonestar Sheds LLC'),
        (('yoderstorage', 'yoder'), 'Yoder Storage Sheds LLC'),
        (('phoenix',), 'Phoenix Buildings'),
    ]
    for aliases, canonical in rules:
        if any(a in compact for a in aliases):
            return canonical
    return raw


def _table_scope(rows: list[list[str]]) -> str:
    keys = {_key(cell) for row in rows for cell in row if _clean(cell)}
    joined = ' | '.join(keys)
    # Prefer explicit section headings when the PDF table finder keeps the
    # blue title row as part of the table.
    for heading, scope in SECTION_NAMES.items():
        if _key(heading) in keys:
            return scope
    if 'new or used' in keys or 'manufacturer' in keys or 'serial' in keys or 'serial number' in keys:
        return 'unit'
    contract_markers = {
        '90 days sac', 'agreement', 'agreement number', 'dealer', 'salesperson',
        'rto terms', 'contract date', 'county of delivery', 'city tax',
        'county tax', 'tax code', 'total tax', 'pmt before tax', 'ldw',
        'total monthly pmt', 'contract price', 'security deposit',
        'purchase reserve', 'initial payment type', 'auto pay',
        'paperless billing', 'order type',
    }
    if any(marker in keys for marker in contract_markers) or '90 days sac' in joined or 'agreement' in joined:
        return 'contract'
    if 'mailing address' in keys or 'physical address' in keys or 'paperless bill email' in keys:
        return 'customer'
    if 'customer owns land' in joined:
        return 'landlord'
    if 'due day of month' in joined or 'source of income' in joined:
        return 'employer'
    # Contact tables are usually tiny and contain Relationship.
    if 'relationship' in keys and 'name' in keys and 'phone' in keys:
        return 'contact'
    if 'dob' in keys and 'name' in keys and 'phone' in keys:
        return 'corenter'
    return ''


def _pairs_from_table(rows: list[list[Any]]) -> list[tuple[str, str]]:
    out: list[tuple[str, str]] = []
    for raw_row in rows:
        row = [_clean(x) for x in (raw_row or [])]
        row = [x for x in row if x]
        if len(row) < 2:
            continue
        # Most source tables are label/value. If a finder merges multiple columns,
        # consume cells in pairs rather than throwing the extra data away.
        for i in range(0, len(row) - 1, 2):
            if row[i] and row[i + 1]:
                out.append((row[i], row[i + 1]))
    return out


def _extract_tables(doc: fitz.Document) -> list[tuple[str, list[tuple[str, str]]]]:
    """Extract label/value pairs without mixing side-by-side sections.

    Some real RentaBarn PDFs are detected by PyMuPDF as one four-column table
    (Customer label/value on the left, Unit/Contract label/value on the right).
    The old parser assigned one scope to the entire table, which could shift
    nearly every numeric field.  Track the section heading independently for
    each label/value column pair instead.
    """
    result: list[tuple[str, list[tuple[str, str]]]] = []
    contact_index = 0

    for page in doc:
        try:
            finder = page.find_tables()
            tables = getattr(finder, 'tables', []) or []
        except Exception:
            tables = []

        for table in tables:
            try:
                raw_matrix = table.extract() or []
            except Exception:
                continue
            matrix = [[_clean(x) for x in (row or [])] for row in raw_matrix]
            if not matrix:
                continue

            ncols = max((len(row) for row in matrix), default=0)
            if ncols <= 2:
                scope = _table_scope(matrix)
                pairs = _pairs_from_table(matrix)
                if scope == 'contact':
                    contact_index += 1
                    scope = 'ref1' if contact_index == 1 else 'ref2'
                if pairs:
                    result.append((scope, pairs))
                continue

            # Multi-section / side-by-side table. Keep empty cells so columns do
            # not collapse and cause label/value drift.
            pair_count = (ncols + 1) // 2
            pair_scopes = [''] * pair_count
            buckets: dict[str, list[tuple[str, str]]] = {}

            for raw_row in matrix:
                row = list(raw_row) + [''] * (ncols - len(raw_row))
                for pair_i in range(pair_count):
                    label_i = pair_i * 2
                    value_i = label_i + 1
                    label = _clean(row[label_i]) if label_i < len(row) else ''
                    value = _clean(row[value_i]) if value_i < len(row) else ''

                    # A section heading can occupy either cell of a pair.
                    heading_scope = ''
                    for candidate in (label, value):
                        ck = _key(candidate)
                        if ck in SECTION_NAMES:
                            heading_scope = SECTION_NAMES[ck]
                            break
                    if heading_scope:
                        if heading_scope == 'contact':
                            contact_index += 1
                            heading_scope = 'ref1' if contact_index == 1 else 'ref2'
                        pair_scopes[pair_i] = heading_scope
                        continue

                    if not label or not value:
                        continue

                    scope = pair_scopes[pair_i]
                    if not scope:
                        # Infer only from this pair's label/value—not all columns.
                        scope = _table_scope([[label, value]])
                        if scope == 'contact':
                            contact_index += 1
                            scope = 'ref1' if contact_index == 1 else 'ref2'
                        if scope:
                            pair_scopes[pair_i] = scope

                    buckets.setdefault(scope, []).append((label, value))

            for scope, pairs in buckets.items():
                if pairs:
                    result.append((scope, pairs))

    return result


def _extract_text_sections_from_text(text: str) -> list[tuple[str, list[tuple[str, str]]]]:
    """Parse known label/value sections from a raw text extraction."""
    lines = [_clean(x) for x in str(text or '').splitlines() if _clean(x)]
    known_labels = {
        'name','mailing address','physical address','city state zip','cell phone','secondary phone','dob','ssn','dl','dl number','email','paperless bill email',
        'new or used','manufacturer','serial','serial number','model','model number','style','unit type','size','base color','trim color','roof color','cash price','description','side fees','unit options',
        '90 days sac','dealer','salesperson','agreement','agreement number','rto terms','contract date','county of delivery','city tax','county tax','tax code','total tax','pmt before tax','ldw','total monthly pmt','contract price','security deposit','purchase reserve','initial payment type','auto pay','paperless billing','order type',
        'phone','relationship','address','employer','employer name','due day of month','customer owns land'
    }
    result: list[tuple[str, list[tuple[str, str]]]] = []
    scope = ''
    buckets: dict[str, list[tuple[str, str]]] = {}
    i = 0
    while i < len(lines):
        line = lines[i]
        lk = _key(line)
        if lk in SECTION_NAMES:
            scope = SECTION_NAMES[lk]
            buckets.setdefault(scope, [])
            i += 1
            continue
        matched_section = next((v for k, v in SECTION_NAMES.items() if lk == _key(k)), None)
        if matched_section:
            scope = matched_section
            buckets.setdefault(scope, [])
            i += 1
            continue

        if lk in known_labels and i + 1 < len(lines):
            nxt = lines[i + 1]
            nk = _key(nxt)
            if nk not in known_labels and nk not in SECTION_NAMES:
                buckets.setdefault(scope, []).append((line, nxt))
                i += 2
                continue

        matched = False
        for label in sorted(known_labels, key=len, reverse=True):
            pattern = re.compile(r'^' + re.escape(label).replace(r'\\ ', r'\\s+') + r'\\s*[:#-]?\\s+(.+)$', re.I)
            m = pattern.match(line)
            if m and _clean(m.group(1)) and _clean(m.group(1)) not in {'#', ':', '-'}:
                buckets.setdefault(scope, []).append((label, m.group(1)))
                matched = True
                break
        if matched:
            i += 1
            continue
        i += 1
    for s, pairs in buckets.items():
        if pairs:
            result.append((s, pairs))
    return result


def _extract_text_sections(doc: fitz.Document) -> list[tuple[str, list[tuple[str, str]]]]:
    return _extract_text_sections_from_text('\n'.join(page.get_text('text') for page in doc))


# Labels expected in each visual section.  This gives the parser a layout-aware
# path that does not depend on PyMuPDF's table finder.  Real provider PDFs can
# expose the visible table as loose positioned text rather than a true PDF table.
_SPATIAL_LABELS: dict[str, list[tuple[str, tuple[str, ...]]]] = {
    'customer': [
        ('Name', ('Name',)), ('Mailing Address', ('Mailing Address',)),
        ('Physical Address', ('Physical Address',)), ('City/State/Zip', ('City/State/Zip', 'City State Zip')),
        ('Cell Phone', ('Cell Phone',)), ('Secondary Phone', ('Secondary Phone',)),
        ('DOB', ('DOB',)), ('SSN', ('SSN',)), ('DL #', ('DL #', 'DL Number', 'DL')),
        ('Email', ('Email',)), ('Paperless Bill Email', ('Paperless Bill Email',)),
    ],
    'unit': [
        ('New or Used', ('New or Used',)), ('Manufacturer', ('Manufacturer',)),
        ('Serial #', ('Serial #', 'Serial Number', 'Serial')), ('Model #', ('Model #', 'Model Number', 'Model')),
        ('Style', ('Style',)), ('Unit Type', ('Unit Type',)), ('Size', ('Size',)),
        ('Base Color', ('Base Color',)), ('Trim Color', ('Trim Color',)), ('Roof Color', ('Roof Color',)),
        ('Cash Price', ('Cash Price',)), ('Description', ('Description',)), ('Side Fees', ('Side Fees',)),
        ('Unit Options', ('Unit Options',)),
    ],
    'contract': [
        ('90 Days SAC', ('90 Days SAC',)), ('Dealer', ('Dealer',)), ('Salesperson', ('Salesperson',)),
        ('Agreement #', ('Agreement #', 'Agreement Number', 'Agreement')), ('RTO Terms', ('RTO Terms',)),
        ('Contract Date', ('Contract Date',)), ('County of Delivery', ('County of Delivery',)),
        ('City Tax', ('City Tax',)), ('County Tax', ('County Tax',)), ('Tax Code', ('Tax Code',)),
        ('Total Tax', ('Total Tax',)), ('PMT Before Tax', ('PMT Before Tax',)), ('LDW', ('LDW',)),
        ('Total Monthly PMT', ('Total Monthly PMT',)), ('Contract Price', ('Contract Price',)),
        ('Cash Price', ('Cash Price',)), ('Security Deposit', ('Security Deposit',)),
        ('Purchase Reserve', ('Purchase Reserve',)), ('Initial Payment Type', ('Initial Payment Type',)),
        ('Auto Pay', ('Auto Pay',)), ('Paperless Billing', ('Paperless Billing',)), ('Order Type', ('Order Type',)),
    ],
    'corenter': [('Name', ('Name',)), ('Phone', ('Phone',)), ('DOB', ('DOB',)), ('Email', ('Email',))],
    'ref1': [('Name', ('Name',)), ('Phone', ('Phone',)), ('Relationship', ('Relationship',)), ('Email', ('Email',))],
    'ref2': [('Name', ('Name',)), ('Phone', ('Phone',)), ('Relationship', ('Relationship',)), ('Email', ('Email',))],
    'employer': [
        ('Name', ('Name', 'Employer Name')), ('Phone', ('Phone',)), ('Address', ('Address',)),
        ('City/State/Zip', ('City/State/Zip', 'City State Zip')), ('Due Day of Month', ('Due Day of Month',)),
    ],
    'landlord': [
        ('Name', ('Name',)), ('Phone', ('Phone',)), ('Address', ('Address',)),
        ('City/State/Zip', ('City/State/Zip', 'City State Zip')), ('Customer Owns Land', ('Customer Owns Land',)),
    ],
}


def _section_regions(page: fitz.Page) -> list[tuple[str, fitz.Rect]]:
    """Locate the visible blue section headings and derive their column regions."""
    hits: list[tuple[str, fitz.Rect, int]] = []
    half = page.rect.width / 2.0
    for heading, scope in SECTION_NAMES.items():
        # Avoid duplicate aliases for the same printed heading where possible.
        try:
            rects = page.search_for(heading)
        except Exception:
            rects = []
        for rect in rects:
            col = 0 if ((rect.x0 + rect.x1) / 2.0) < half else 1
            hits.append((scope, fitz.Rect(rect), col))
    if not hits:
        return []
    # De-duplicate near-identical heading hits caused by aliases/case folding.
    dedup: list[tuple[str, fitz.Rect, int]] = []
    for scope, rect, col in sorted(hits, key=lambda x: (x[2], x[1].y0, x[1].x0)):
        if any(s == scope and c == col and abs(r.y0 - rect.y0) < 3 for s, r, c in dedup):
            continue
        dedup.append((scope, rect, col))

    regions: list[tuple[str, fitz.Rect]] = []
    for col in (0, 1):
        col_hits = [(s, r) for s, r, c in dedup if c == col]
        col_hits.sort(key=lambda x: x[1].y0)
        x0 = page.rect.x0 if col == 0 else half
        x1 = half if col == 0 else page.rect.x1
        for idx, (scope, rect) in enumerate(col_hits):
            y0 = max(page.rect.y0, rect.y1 - 1)
            y1 = col_hits[idx + 1][1].y0 - 1 if idx + 1 < len(col_hits) else page.rect.y1
            if y1 > y0:
                regions.append((scope, fitz.Rect(x0, y0, x1, y1)))
    return regions


def _same_line_value(page: fitz.Page, label_rect: fitz.Rect, region: fitz.Rect) -> str:
    """Return text printed to the right of a label on the same visual row."""
    cy = (label_rect.y0 + label_rect.y1) / 2.0
    tol = max(5.0, label_rect.height * 0.72)
    try:
        words = page.get_text('words', clip=region) or []
    except Exception:
        words = []
    vals = []
    for w in words:
        x0, y0, x1, y1, text = w[:5]
        wy = (y0 + y1) / 2.0
        if x0 >= label_rect.x1 - 0.5 and abs(wy - cy) <= tol:
            vals.append((x0, _clean(text)))
    vals.sort(key=lambda x: x[0])
    value = _clean(' '.join(t for _, t in vals if t))
    # search_for('Agreement #') / search_for('Model #') can occasionally match
    # only the word part and leave the printed '#' as the first value token.
    value = re.sub(r'^(?:#|:|-)\s*', '', value).strip()
    return value


def _extract_spatial_pairs(doc: fitz.Document) -> list[tuple[str, list[tuple[str, str]]]]:
    """Geometry-based reader for provider PDFs with positioned text but no table objects.

    This is intentionally tried before table extraction.  It reads each label's
    value from the same printed row inside its visual section, preventing the
    left/right columns from being interleaved by the PDF text extractor.
    """
    out: list[tuple[str, list[tuple[str, str]]]] = []
    for page in doc:
        regions = _section_regions(page)
        if not regions:
            continue
        for scope, region in regions:
            specs = _SPATIAL_LABELS.get(scope, [])
            if not specs:
                continue
            pairs: list[tuple[str, str]] = []
            used_y: list[float] = []
            for canonical, aliases in specs:
                found_value = ''
                found_rect = None
                for alias in aliases:
                    try:
                        hits = page.search_for(alias)
                    except Exception:
                        hits = []
                    for rect in hits:
                        center = fitz.Point((rect.x0 + rect.x1) / 2.0, (rect.y0 + rect.y1) / 2.0)
                        if not region.contains(center):
                            continue
                        # Skip a hit if it is almost exactly the same baseline as
                        # an already consumed alias for this canonical field.
                        if any(abs(rect.y0 - y) < 2 for y in used_y):
                            continue
                        value = _same_line_value(page, rect, region)
                        if value:
                            found_value = value
                            found_rect = rect
                            break
                    if found_value:
                        break
                if found_value:
                    pairs.append((canonical, found_value))
                    if found_rect is not None:
                        used_y.append(found_rect.y0)
            if pairs:
                out.append((scope, pairs))
    return out


def _fallback_text_candidates(path: Path, doc: fitz.Document) -> list[tuple[str, str]]:
    """Collect text using multiple PDF engines; different generators expose text differently."""
    candidates: list[tuple[str, str]] = []
    try:
        text = '\n'.join(page.get_text('text', sort=True) for page in doc)
        if _clean(text):
            candidates.append(('PyMuPDF', text))
    except Exception:
        pass
    try:
        blocks = []
        for page in doc:
            for b in page.get_text('blocks', sort=True) or []:
                if len(b) >= 5 and _clean(b[4]):
                    blocks.append(_clean(b[4]))
        text = '\n'.join(blocks)
        if _clean(text):
            candidates.append(('PyMuPDF blocks', text))
    except Exception:
        pass
    if PdfReader is not None:
        try:
            reader = PdfReader(str(path))
            text = '\n'.join((page.extract_text() or '') for page in reader.pages)
            if _clean(text):
                candidates.append(('pypdf', text))
        except Exception:
            pass
    if pdfplumber is not None:
        try:
            with pdfplumber.open(str(path)) as pdf:
                text = '\n'.join((page.extract_text(x_tolerance=2, y_tolerance=3) or '') for page in pdf.pages)
            if _clean(text):
                candidates.append(('pdfplumber', text))
        except Exception:
            pass
    # Preserve unique candidates only.
    uniq: list[tuple[str, str]] = []
    seen = set()
    for engine, text in candidates:
        marker = _clean(text)
        if marker and marker not in seen:
            seen.add(marker)
            uniq.append((engine, text))
    return uniq


def _text_quality(text: str) -> int:
    t = _key(text)
    score = min(len(t), 2000) // 20
    for phrase in ('customer information','unit information','contract information','manufacturer','agreement','pmt before tax','rentabarn','choice capital'):
        if phrase in t:
            score += 100
    return score


def _extract_pdfplumber_tables(path: Path) -> list[tuple[str, list[tuple[str, str]]]]:
    if pdfplumber is None:
        return []
    result: list[tuple[str, list[tuple[str, str]]]] = []
    contact_index = 0
    try:
        with pdfplumber.open(str(path)) as pdf:
            for page in pdf.pages:
                try:
                    tables = page.extract_tables() or []
                except Exception:
                    tables = []
                for raw in tables:
                    matrix = [[_clean(x) for x in (row or [])] for row in (raw or [])]
                    if not matrix:
                        continue
                    ncols = max((len(row) for row in matrix), default=0)
                    if ncols <= 2:
                        scope = _table_scope(matrix)
                        pairs = _pairs_from_table(matrix)
                        if scope == 'contact':
                            contact_index += 1
                            scope = 'ref1' if contact_index == 1 else 'ref2'
                        if pairs:
                            result.append((scope, pairs))
                    else:
                        # Pair columns independently to avoid left/right mixing.
                        pair_count = (ncols + 1) // 2
                        scopes = [''] * pair_count
                        buckets: dict[str, list[tuple[str, str]]] = {}
                        for raw_row in matrix:
                            row = list(raw_row) + [''] * (ncols - len(raw_row))
                            for pair_i in range(pair_count):
                                li, vi = pair_i * 2, pair_i * 2 + 1
                                label = _clean(row[li]) if li < len(row) else ''
                                value = _clean(row[vi]) if vi < len(row) else ''
                                for candidate in (label, value):
                                    ck = _key(candidate)
                                    if ck in SECTION_NAMES:
                                        sc = SECTION_NAMES[ck]
                                        if sc == 'contact':
                                            contact_index += 1
                                            sc = 'ref1' if contact_index == 1 else 'ref2'
                                        scopes[pair_i] = sc
                                        break
                                if label and value and scopes[pair_i]:
                                    buckets.setdefault(scopes[pair_i], []).append((label, value))
                        result.extend((sc, pairs) for sc, pairs in buckets.items() if pairs)
    except Exception:
        return []
    return result

def _assign(fields: dict[str, dict[str, str]], scope: str, label: str, value: str) -> None:
    lk = _key(label)
    value = _clean(value)
    if not lk or not value:
        return
    fields.setdefault(scope or 'unknown', {})[lk] = value


def _get(fields: dict[str, dict[str, str]], scope: str, *labels: str) -> str:
    bucket = fields.get(scope, {})
    for label in labels:
        v = bucket.get(_key(label), '')
        if v:
            return v
    return ''


def _sum_rates(*values: str) -> str:
    total = 0.0
    found = False
    for value in values:
        p = _percentage(value)
        if not p:
            continue
        try:
            total += float(p)
            found = True
        except Exception:
            pass
    # ShedSuite uses percentage points (e.g. 7.25), not 0.0725.
    return (f'{total:.6f}'.rstrip('0').rstrip('.')) if found else ''




def _detect_pdf_provider(text: str, filename: str = '') -> str:
    raw_text = _clean(text).casefold()
    compact = re.sub(r'[^a-z0-9]+', '', raw_text)
    if 'choicecapital' in compact:
        return 'CHOICE CAPITAL'
    # RentaBarn contracts may identify the provider as Wolfvalley/Wolf Valley
    # or WVB.  They all intentionally use the same RentaBarn EPO schedule.
    if 'rentabarn' in compact or 'wolfvalley' in compact or re.search(r'\bwvb\b', raw_text, flags=re.I):
        return 'RENTABARN'
    if re.search(r'(?<![a-z0-9])dbm(?![a-z0-9])', raw_text, flags=re.I):
        return 'DBM'
    if 'xgen' in compact:
        return 'X-GEN'

    # Some provider logos render as artwork instead of selectable text, so use
    # filename tokens only as a fallback after inspecting PDF text itself.
    # of selectable text. Their detailed-information filenames include the CF
    # source token (e.g. PDF-01_CF_6-60376_DetailedInformationPage_...).
    # Use this only as a fallback after inspecting the PDF text itself.
    name = Path(str(filename or '')).name
    if re.search(r'(?:^|[_\-])CF(?:[_\-]|$)', name, flags=re.I):
        return 'CHOICE CAPITAL'
    if (re.search(r'(?:^|[_\-])(?:RB|WVB)(?:[_\-]|$)', name, flags=re.I)
            or 'rentabarn' in name.casefold() or 'wolfvalley' in re.sub(r'[^a-z0-9]+', '', name.casefold())):
        return 'RENTABARN'
    if re.search(r'(?:^|[_\-])DBM(?:[_\-]|$)', name, flags=re.I):
        return 'DBM'
    if re.search(r'(?:^|[_\-])X[ _\-]?GEN(?:[_\-]|$)', name, flags=re.I) or 'xgen' in re.sub(r'[^a-z0-9]+', '', name.casefold()):
        return 'X-GEN'
    return ''


def _value_right_of_label(doc: fitz.Document, label: str) -> str:
    """Read the value visually to the right of an exact PDF label.

    This is used as an authoritative fallback for critical payment fields.
    Table extraction can occasionally combine the left/right halves of the
    form; locating the printed label rectangle and reading words on the same
    baseline avoids pulling a neighboring numeric value.
    """
    for page in doc:
        try:
            hits = page.search_for(label)
            words = page.get_text('words') or []
        except Exception:
            continue
        for rect in hits:
            same_line = []
            cy = (rect.y0 + rect.y1) / 2.0
            for w in words:
                x0, y0, x1, y1, text = w[:5]
                wy = (y0 + y1) / 2.0
                # PDF table text can be a few points off vertically, so use a
                # forgiving baseline window but only accept words to the right.
                if x0 >= rect.x1 - 1 and abs(wy - cy) <= max(6.0, rect.height * 0.8):
                    same_line.append((x0, _clean(text)))
            same_line.sort(key=lambda x: x[0])
            value = _clean(' '.join(t for _, t in same_line if t))
            if value:
                return value
    return ''

def _looks_supported(text: str) -> bool:
    t = _key(text)
    signatures = [
        ('customer information', 'unit information', 'contract information'),
        ('new or used', 'manufacturer', 'agreement'),
        ('rentabarn', 'customer information', 'contract information'),
    ]
    return any(all(x in t for x in sig) for sig in signatures)


def parse_contract_pdf(path: Path) -> tuple[dict[str, str], dict[str, str], list[str]]:
    """Parse a RentaBarn-style contract PDF into the ShedSuite source schema.

    Returns (source_row, metadata, warnings). The source row intentionally uses
    the same keys as a ShedSuite CSV row so the existing V7.x transform/review
    pipeline remains authoritative.
    """
    warnings: list[str] = []
    doc = fitz.open(path)
    try:
        text_candidates = _fallback_text_candidates(path, doc)
        full_text = max(text_candidates, key=lambda x: _text_quality(x[1]))[1] if text_candidates else ''

        # V7.11: geometry first.  A real browser-generated provider PDF may have
        # perfectly selectable text yet expose no usable table structure.
        extracted = _extract_spatial_pairs(doc)
        extracted += _extract_tables(doc)
        extracted += _extract_pdfplumber_tables(path)
        for _engine, candidate_text in text_candidates:
            extracted += _extract_text_sections_from_text(candidate_text)

        if not extracted and not _clean(full_text):
            raise ValueError('PDF reader could not extract text from this file. It may be an image-only/scanned PDF; please upload the original downloaded PDF.')

        # Do not reject solely on one engine's raw-text signature anymore.
        # Accept when the spatial/table readers find recognizable form labels.
        recognized_pairs = sum(len(pairs) for scope, pairs in extracted if scope in SECTION_NAMES.values())
        if not _looks_supported(full_text) and recognized_pairs < 4:
            raise ValueError('PDF text was found, but the provider form fields could not be recognized. Upload the original RentaBarn/Choice Capital detailed-information PDF.')

        pdf_provider = _detect_pdf_provider(full_text, path.name)

        fields: dict[str, dict[str, str]] = {}
        for scope, pairs in extracted:
            for label, value in pairs:
                bucket = fields.setdefault(scope or 'unknown', {})
                base_key = _key(label)
                if not base_key:
                    continue
                if base_key not in bucket:
                    bucket[base_key] = _clean(value)
                else:
                    n = 2
                    while f'{base_key} {n}' in bucket:
                        n += 1
                    bucket[f'{base_key} {n}'] = _clean(value)

        customer_name = _get(fields, 'customer', 'Name')
        first, last = _split_name(customer_name)
        mail_city, mail_state, mail_zip = _split_city_state_zip(_get(fields, 'customer', 'City/State/Zip', 'City State Zip'))
        physical = _get(fields, 'customer', 'Physical Address')
        physical_csz = _get(fields, 'customer', 'City/State/Zip 2', 'City State Zip 2')
        phys_city, phys_state, phys_zip = _split_city_state_zip(physical_csz)
        mailing = _get(fields, 'customer', 'Mailing Address')

        co_name = _get(fields, 'corenter', 'Name')
        co_first, co_last = _split_name(co_name)
        ref1_name = _get(fields, 'ref1', 'Name')
        ref2_name = _get(fields, 'ref2', 'Name')

        manufacturer_raw = _get(fields, 'unit', 'Manufacturer')
        company = _canonical_company(manufacturer_raw)
        serial = _get(fields, 'unit', 'Serial #', 'Serial Number', 'Serial')
        model = _get(fields, 'unit', 'Model #', 'Model Number', 'Model')
        style = _get(fields, 'unit', 'Style')
        unit_type = _get(fields, 'unit', 'Unit Type')
        size = _normalize_size(_get(fields, 'unit', 'Size'))
        desc = _get(fields, 'unit', 'Description')
        # RentaBarn commonly appends the generic word "Shed" (e.g.
        # "Side Utility Shed") while the existing category map uses
        # "SIDE UTILITY". Strip only that generic trailing word so the old
        # category mapping can keep working without a second category system.
        style_for_mapping = re.sub(r'\s+SHED\s*$', '', style, flags=re.I).strip()
        model_variation = ' '.join(x for x in [size, style_for_mapping or unit_type or desc] if x).strip()
        if not model_variation:
            model_variation = style or unit_type or desc

        agreement = _get(fields, 'contract', 'Agreement #', 'Agreement Number', 'Agreement')
        dealer = _get(fields, 'contract', 'Dealer')
        salesperson = _get(fields, 'contract', 'Salesperson')
        contract_date = _get(fields, 'contract', 'Contract Date')
        sac_raw = _get(fields, 'contract', '90 Days SAC')
        sac_date = _date_from_value(sac_raw)
        term = _get(fields, 'contract', 'RTO Terms')
        epo_percentage = pdf_epo_percentage(pdf_provider, term)
        city_tax = _get(fields, 'contract', 'City Tax')
        county_tax = _get(fields, 'contract', 'County Tax')
        # Critical payment values: prefer the value printed on the same visual
        # row as the label. This prevents a four-column PDF table from shifting
        # PMT Before Tax onto LDW/Total Tax/another neighboring number.
        ldw = _money(_value_right_of_label(doc, 'LDW') or _get(fields, 'contract', 'LDW'))
        pmt_before_tax = _money(_value_right_of_label(doc, 'PMT Before Tax') or _get(fields, 'contract', 'PMT Before Tax'))
        total_monthly_pmt = _money(_value_right_of_label(doc, 'Total Monthly PMT') or _get(fields, 'contract', 'Total Monthly PMT'))
        total_tax = _money(_value_right_of_label(doc, 'Total Tax') or _get(fields, 'contract', 'Total Tax'))
        cash_price = _money(_get(fields, 'unit', 'Cash Price') or _get(fields, 'contract', 'Cash Price'))
        security = _money(_value_right_of_label(doc, 'Security Deposit') or _get(fields, 'contract', 'Security Deposit'))
        reserve = _money(_value_right_of_label(doc, 'Purchase Reserve') or _get(fields, 'contract', 'Purchase Reserve'))
        initial = _amount_in_parentheses(_get(fields, 'contract', 'Initial Payment Type'))
        condition = _get(fields, 'unit', 'New or Used')
        paperless_billing = _get(fields, 'contract', 'Paperless Billing')

        emp_city, emp_state, emp_zip = _split_city_state_zip(_get(fields, 'employer', 'City/State/Zip', 'City State Zip'))
        due_day = _get(fields, 'employer', 'Due Day of Month')

        # If physical address is not separately populated, the contract is still
        # usable: fall back to mailing address to keep ZipTax/review functional.
        if not physical:
            physical = mailing
        if not phys_city:
            phys_city, phys_state, phys_zip = mail_city, mail_state, mail_zip

        source: dict[str, str] = {
            'Serial Number': serial,
            'Customer Order Id': agreement or f'PDF-{path.stem}',
            'Inventory Id': model,
            'Customer First Name': first,
            'Customer Last Name': last,
            'Company Name': company,
            'Dealer': dealer,
            'Customer Mailing Street': mailing,
            'Customer Mailing City': mail_city,
            'Customer Mailing State': mail_state,
            'Customer Mailing Zip': mail_zip,
            'Physical Destination Street': physical,
            'Physical Destination City': phys_city,
            'Physical Destination State': phys_state,
            'Physical Destination Zip': phys_zip,
            'Primary Phone': _get(fields, 'customer', 'Cell Phone'),
            'Secondary Phone': _get(fields, 'customer', 'Secondary Phone'),
            'Customer Date Of Birth': _get(fields, 'customer', 'DOB'),
            'Customer Social Security Number': _get(fields, 'customer', 'SSN'),
            'Customer Identification Card Number': _get(fields, 'customer', 'DL #', 'DL Number', 'DL'),
            'Customer Email': _get(fields, 'customer', 'Email') or _get(fields, 'customer', 'Paperless Bill Email'),
            'Co Renter First Name': co_first,
            'Co Renter Last Name': co_last,
            'Co Renter Phone': _get(fields, 'corenter', 'Phone'),
            'Co Renter Date Of Birth': _get(fields, 'corenter', 'DOB'),
            'Co Renter Email': _get(fields, 'corenter', 'Email'),
            'Ref 1 Name': ref1_name,
            'Ref 1 Phone': _get(fields, 'ref1', 'Phone'),
            'Ref 1 Relationship': _get(fields, 'ref1', 'Relationship'),
            'Ref 2 Name': ref2_name,
            'Ref 2 Phone': _get(fields, 'ref2', 'Phone'),
            'Ref 2 Relationship': _get(fields, 'ref2', 'Relationship'),
            'Customer Employer Name': _get(fields, 'employer', 'Name', 'Employer Name'),
            'Customer Employer Phone': _get(fields, 'employer', 'Phone'),
            'Building Model Variation': model_variation,
            # Preserve the raw PDF style fields as extra evidence for V7.12's
            # approved-category matcher. These are harmless for CSV transforms
            # and make vague descriptions much easier to classify correctly.
            'Description': desc,
            'Unit Type': unit_type,
            'Style': style,
            'Siding Color': _get(fields, 'unit', 'Base Color'),
            'Trim Color': _get(fields, 'unit', 'Trim Color'),
            'Roof Color': _get(fields, 'unit', 'Roof Color'),
            'Condition': condition,
            'Date Created': contract_date,
            'Date Delivered': '',
            'Day Of Month Payment Is Due': due_day,
            'Loss Damage Waiver Monthly Charge': ldw,
            'Initial Payment': initial,
            'Delivery Fee': '0',
            'Security Deposit': security,
            'Purchase Reserve': reserve,
            'Months Of Term': term,
            'Building Cost': cash_price,
            'Early Purchase Percentage': epo_percentage,
            'Monthly Subtotal': pmt_before_tax,
            'Sum Of All Tax Rates': _sum_rates(city_tax, county_tax),
            # No ShedSuite URLs exist for this source. The original uploaded PDF
            # itself is copied into Combined_Files by app.py.
            'Signed Combined Documents': '',
            'Rto Contract': '',
            'Invoice': '',
            'Landowner Permission Form': '',
            'Customer Identification Card Scan': '',
            'Co Renter Identification Card Scan': '',
            '_import_source_type': 'pdf',
            '_import_pdf_name': path.name,
            '_pdf_provider': pdf_provider,
            '_pdf_epo_percentage': epo_percentage,
            '_pdf_epo_term_months': str(_term_months(term) or ''),
            '_pdf_salesperson_source': salesperson,
            '_pdf_contract_price': _money(_get(fields, 'contract', 'Contract Price')),
            '_pdf_sac_date': sac_date,
            '_pdf_sac_source': sac_raw,
            '_pdf_agreement_number': agreement,
            '_pdf_pmt_before_tax': pmt_before_tax,
            '_pdf_total_monthly_pmt': total_monthly_pmt,
            '_pdf_total_tax': total_tax,
            '_pdf_ldw': ldw,
            '_pdf_security_deposit': security,
            '_pdf_purchase_reserve': reserve,
            '_pdf_county_of_delivery': _get(fields, 'contract', 'County of Delivery'),
            '_pdf_tax_code': _get(fields, 'contract', 'Tax Code'),
            '_pdf_order_type': _get(fields, 'contract', 'Order Type'),
            '_pdf_auto_pay': _get(fields, 'contract', 'Auto Pay'),
            '_pdf_paperless_billing': paperless_billing,
            '_pdf_email_invoice': '1' if _yes(paperless_billing) else ('0' if _key(paperless_billing) in {'no', 'n', 'false', '0', 'off'} else ''),
            '_pdf_landlord_name': _get(fields, 'landlord', 'Name'),
            '_pdf_landlord_phone': _get(fields, 'landlord', 'Phone'),
            '_pdf_landlord_address': _get(fields, 'landlord', 'Address'),
            '_pdf_landlord_city_state_zip': _get(fields, 'landlord', 'City/State/Zip', 'City State Zip'),
            '_pdf_customer_owns_land': _get(fields, 'landlord', 'Customer Owns Land'),
            '_pdf_employer_address': _get(fields, 'employer', 'Address'),
            '_pdf_employer_city': emp_city,
            '_pdf_employer_state': emp_state,
            '_pdf_employer_zip': emp_zip,
        }

        required = {
            'customer name': customer_name,
            'serial number': serial,
            'agreement number': agreement,
            'manufacturer': manufacturer_raw,
        }
        for label, value in required.items():
            if not value:
                warnings.append(f'{path.name}: PDF parser could not find {label}; review that field manually.')
        if not dealer:
            warnings.append(f'{path.name}: PDF parser could not find dealer; choose the dealer/store on review.')
        if not model:
            warnings.append(f'{path.name}: PDF parser could not find Model #; review MODEL1 manually.')
        if not sac_date:
            warnings.append(f'{path.name}: PDF parser could not find the date inside 90 Days SAC; review SACDATE manually.')
        if not pmt_before_tax:
            warnings.append(f'{path.name}: PDF parser could not find PMT Before Tax; review RATE1 manually.')
        if not total_monthly_pmt:
            warnings.append(f'{path.name}: PDF parser could not find Total Monthly PMT; verify the final RTO payment manually.')
        if pdf_provider and _term_months(term) and not epo_percentage:
            warnings.append(
                f'{path.name}: no configured EPO percentage for {pdf_provider} at {_term_months(term)} months; review PAYOFFDISCOUNT manually.'
            )
        if salesperson:
            warnings.append(f'{path.name}: salesperson extracted from PDF: {salesperson}.')

        metadata = {
            'provider': (pdf_provider.title() + ' PDF') if pdf_provider else 'Contract PDF',
            'reader_engines': ', '.join(engine for engine, _ in text_candidates) or 'PyMuPDF spatial',
            'filename': path.name,
            'agreement': agreement,
            'manufacturer_original': manufacturer_raw,
            'manufacturer_mapped': company,
            'salesperson': salesperson,
            'condition': condition,
        }
        return source, metadata, warnings
    finally:
        doc.close()
