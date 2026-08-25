from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import fitz


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
    if '90 days sac' in joined or 'agreement' in joined or 'salesperson' in keys or 'order type' in keys:
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
                matrix = table.extract() or []
            except Exception:
                continue
            matrix = [[_clean(x) for x in (row or [])] for row in matrix]
            scope = _table_scope(matrix)
            pairs = _pairs_from_table(matrix)
            if scope == 'contact':
                contact_index += 1
                scope = 'ref1' if contact_index == 1 else 'ref2'
            if pairs:
                result.append((scope, pairs))
    return result


def _extract_text_sections(doc: fitz.Document) -> list[tuple[str, list[tuple[str, str]]]]:
    # Fallback for PDFs whose borders are not detectable as tables. It recognizes
    # the known section headings and both "Label Value" and alternating Label/Value
    # text extraction orders.
    text = '\n'.join(page.get_text('text') for page in doc)
    lines = [_clean(x) for x in text.splitlines() if _clean(x)]
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
        # Some headings include punctuation / casing around the section name.
        matched_section = next((v for k, v in SECTION_NAMES.items() if lk == _key(k)), None)
        if matched_section:
            scope = matched_section
            buckets.setdefault(scope, [])
            i += 1
            continue

        # Try "known label + value" on one line.
        matched = False
        for label in sorted(known_labels, key=len, reverse=True):
            pattern = re.compile(r'^' + re.escape(label).replace(r'\ ', r'\s+') + r'\s*[:#-]?\s+(.+)$', re.I)
            m = pattern.match(line)
            if m and _clean(m.group(1)):
                buckets.setdefault(scope, []).append((label, m.group(1)))
                matched = True
                break
        if matched:
            i += 1
            continue

        # Alternating extraction: label on one line, value on the next.
        if lk in known_labels and i + 1 < len(lines):
            nxt = lines[i + 1]
            nk = _key(nxt)
            if nk not in known_labels and nk not in SECTION_NAMES:
                buckets.setdefault(scope, []).append((line, nxt))
                i += 2
                continue
        i += 1
    for s, pairs in buckets.items():
        if pairs:
            result.append((s, pairs))
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
        full_text = '\n'.join(page.get_text('text') for page in doc)
        if not _clean(full_text):
            raise ValueError('PDF has no selectable text. Export/download the original digital PDF instead of a scan or screenshot.')
        if not _looks_supported(full_text):
            raise ValueError('PDF is not recognized as the supported RentaBarn-style contract format.')

        extracted = _extract_tables(doc)
        if not extracted:
            extracted = _extract_text_sections(doc)
        else:
            # Text fallback can recover cells table detection missed. Add it after
            # table values so table extraction wins when both are present.
            extracted += _extract_text_sections(doc)

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
        size = _get(fields, 'unit', 'Size')
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
        term = _get(fields, 'contract', 'RTO Terms')
        city_tax = _get(fields, 'contract', 'City Tax')
        county_tax = _get(fields, 'contract', 'County Tax')
        ldw = _money(_get(fields, 'contract', 'LDW'))
        pmt_before_tax = _money(_get(fields, 'contract', 'PMT Before Tax'))
        cash_price = _money(_get(fields, 'unit', 'Cash Price') or _get(fields, 'contract', 'Cash Price'))
        security = _money(_get(fields, 'contract', 'Security Deposit'))
        reserve = _money(_get(fields, 'contract', 'Purchase Reserve'))
        initial = _amount_in_parentheses(_get(fields, 'contract', 'Initial Payment Type'))
        condition = _get(fields, 'unit', 'New or Used')

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
            'Early Purchase Percentage': '',
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
            '_pdf_salesperson_source': salesperson,
            '_pdf_contract_price': _money(_get(fields, 'contract', 'Contract Price')),
            '_pdf_total_monthly_pmt': _money(_get(fields, 'contract', 'Total Monthly PMT')),
            '_pdf_total_tax': _money(_get(fields, 'contract', 'Total Tax')),
            '_pdf_county_of_delivery': _get(fields, 'contract', 'County of Delivery'),
            '_pdf_tax_code': _get(fields, 'contract', 'Tax Code'),
            '_pdf_order_type': _get(fields, 'contract', 'Order Type'),
            '_pdf_auto_pay': _get(fields, 'contract', 'Auto Pay'),
            '_pdf_paperless_billing': _get(fields, 'contract', 'Paperless Billing'),
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
        if salesperson:
            warnings.append(f'{path.name}: salesperson extracted from PDF: {salesperson}.')

        metadata = {
            'provider': 'RentaBarn PDF',
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
