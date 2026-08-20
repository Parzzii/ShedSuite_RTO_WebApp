from __future__ import annotations

import asyncio
import base64
import csv
import html as html_lib
import json
import os
import re
import shutil
import subprocess
import uuid
import xml.dom.minidom
from pathlib import Path
from datetime import datetime
from urllib.parse import unquote
from xml.etree.ElementTree import Element, SubElement, tostring
from zipfile import ZipFile, ZIP_DEFLATED

import fitz

try:
    import pyodbc
except Exception:
    pyodbc = None

from dotenv import load_dotenv
from playwright.async_api import async_playwright, expect
from flask import Flask, flash, redirect, render_template, request, send_file, url_for, abort, jsonify

from rto_transform import (
    RTO_HEADERS,
    KNOWN_STORES,
    ReferenceData,
    load_categories,
    load_reference_data,
    transform_rows,
    effective_ziptax_api_key,
    ziptax_detail,
    resolve_taxzone,
)
from pdf_tools import download_and_combine, extract_pdf_fields, safe_filename
from delivery_certificate import append_delivery_certificates


BASE = Path(__file__).resolve().parent
WORK = BASE / 'work' / 'jobs'
WORK.mkdir(parents=True, exist_ok=True)

AUTH_DIR = BASE / 'auth_state'
AUTH_DIR.mkdir(parents=True, exist_ok=True)

SHEDSUITE_BASE_URL = 'https://app.shedsuite.com/rto/order/'
MODEL_TRACKER_FILE = BASE / 'model_series.json'

# The web app keeps DB passwords in memory only for the life of this local
# Python process. They are never written into the job folder.
JOB_DB_PASSWORDS: dict[str, str] = {}

# These are the SAME contract-number ranges used by the Excel workbook's
# "Next model" formulas. This matters because Magnolia GA/SC/NC/TN all use
# prefix 0801 but are separate sequences.
MODEL_DB_RULES = {
    'crf_alpine':               {'prefix': '33',   'min': 1300, 'max': 1409},
    'crf_genesis':              {'prefix': '35',   'min': 216,  'max': 298},
    'crf_phoenix':              {'prefix': '0149', 'min': 114,  'max': 198},
    'crf_4_seasons':            {'prefix': '0150', 'min': 153,  'max': 198},
    'dbm_phoenix':              {'prefix': '0203', 'min': 549,  'max': 598},
    'wwp':                      {'prefix': '1101', 'min': 409,  'max': 498},
    'ras_yss':                  {'prefix': '0605', 'min': 37,   'max': 98},
    'lsr_tx':                   {'prefix': '0701', 'min': 816,  'max': 898},
    'magnolia_smart_shed_ga':   {'prefix': '0801', 'min': 1193, 'max': 1298},
    'magnolia_smart_shed_sc':   {'prefix': '0801', 'min': 273,  'max': 299},
    'magnolia_smart_shed_nc':   {'prefix': '0801', 'min': 330,  'max': 398},
    'magnolia_smart_shed_tn':   {'prefix': '0801', 'min': 401,  'max': 499},
}

DEFAULT_MODEL_SERIES = {
    'magnolia_smart_shed_ga': {'label': 'Magnolia / Smart Shed — GA', 'prefix': '0801', 'last_used': '', 'example': '0801-1223'},
    'magnolia_smart_shed_sc': {'label': 'Magnolia / Smart Shed — SC', 'prefix': '0801', 'last_used': '', 'example': '0801-283'},
    'magnolia_smart_shed_nc': {'label': 'Magnolia / Smart Shed — NC', 'prefix': '0801', 'last_used': '', 'example': '0801-331'},
    'magnolia_smart_shed_tn': {'label': 'Magnolia / Smart Shed — TN', 'prefix': '0801', 'last_used': '', 'example': '0801-403'},
    'crf_genesis': {'label': 'CRF / Genesis', 'prefix': '35', 'last_used': '', 'example': '35-219'},
    'crf_alpine': {'label': 'CRF / Alpine', 'prefix': '33', 'last_used': '', 'example': '33-1322'},
    'crf_phoenix': {'label': 'CRF / Phoenix', 'prefix': '0149', 'last_used': '', 'example': '0149-119'},
    'crf_4_seasons': {'label': 'CRF / 4 Seasons', 'prefix': '0150', 'last_used': '', 'example': '0150-159'},
    'dbm_phoenix': {'label': 'DBM Phoenix', 'prefix': '0203', 'last_used': '0203-584', 'example': ''},
    'wwp': {'label': 'WWP', 'prefix': '1101', 'last_used': '1101-437', 'example': ''},
    'ras_yss': {'label': 'RAS YSS', 'prefix': '0605', 'last_used': '0605-049', 'example': ''},
    'lsr_tx': {'label': 'LSR TX', 'prefix': '0701', 'last_used': '0701-832', 'example': ''},
}

load_dotenv(BASE / '.env')

app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'local-rto-webapp-v2')


# ---------------------------------------------------------------------------
# ShedSuite browser/file helpers
# ---------------------------------------------------------------------------

def shedsuite_auth_path(email: str) -> Path:
    """Store one Playwright login session per ShedSuite email address."""
    safe = re.sub(r'[^a-zA-Z0-9._-]', '_', (email or '').strip().lower())
    if not safe:
        safe = 'default'
    return AUTH_DIR / f'shedsuite_auth_{safe}.json'


def extract_shedsuite_file_url(url: str) -> str:
    """Unwrap ShedSuite viewer URLs when they contain the real https URL."""
    if not url:
        return ''
    value = unquote(str(url))
    first = value.find('https')
    if first == -1:
        return value
    second = value.find('https', first + 5)
    return value[second:] if second != -1 else value[first:]


def _delivery_name(name: str) -> bool:
    n = (name or '').strip().lower()
    return 'delivery' in n and ('certificate' in n or 'cert' in n)


def _bytes_to_pdf(data: bytes, target: Path) -> bool:
    if not data:
        return False
    target.parent.mkdir(parents=True, exist_ok=True)
    if data[:5] == b'%PDF-':
        target.write_bytes(data)
        return True
    try:
        image_doc = fitz.open(stream=data)
        pdf_bytes = image_doc.convert_to_pdf()
        image_doc.close()
        pdf_doc = fitz.open('pdf', pdf_bytes)
        pdf_doc.save(target)
        pdf_doc.close()
        return target.exists() and target.stat().st_size > 0
    except Exception:
        return False


async def _request_to_pdf(context, url: str, target: Path) -> bool:
    url = extract_shedsuite_file_url(url)
    if not url or url.startswith('blob:'):
        return False
    try:
        response = await context.request.get(url, timeout=60000)
        if not response.ok:
            return False
        return _bytes_to_pdf(await response.body(), target)
    except Exception:
        return False


async def _blob_page_to_pdf(page, target: Path) -> bool:
    try:
        data_url = await page.evaluate("""async () => {
            const r = await fetch(window.location.href);
            const b = await r.blob();
            return await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = () => resolve(reader.result);
                reader.onerror = reject;
                reader.readAsDataURL(b);
            });
        }""")
        if not data_url or ',' not in data_url:
            return False
        return _bytes_to_pdf(base64.b64decode(data_url.split(',', 1)[1]), target)
    except Exception:
        return False


async def download_delivery_certificate_for_order(page, context, order_id: str, target: Path) -> tuple[bool, str]:
    """Capture Delivery Certificate bytes inside the authenticated ShedSuite browser."""
    popup_task = None
    download_task = None
    responses = []
    def remember_response(resp):
        responses.append(resp)
    try:
        await page.goto(SHEDSUITE_BASE_URL + str(order_id), wait_until='domcontentloaded', timeout=90000)
        await page.wait_for_selector('span[class^="fileName"]', timeout=90000)
        spans = await page.query_selector_all('span[class^="fileName"]')
        target_span = None
        target_name = ''
        for span in spans:
            name = ((await span.text_content()) or '').strip()
            if _delivery_name(name):
                target_span = span
                target_name = name
                break
        if target_span is None:
            return False, 'Delivery Certificate was not listed on the ShedSuite order page.'

        try:
            href = await target_span.evaluate("el => (el.closest('a') && el.closest('a').href) || ''")
            if href and await _request_to_pdf(context, href, target):
                return True, f'{target_name} (authenticated link)'
        except Exception:
            pass

        page.on('response', remember_response)
        popup_task = asyncio.create_task(page.wait_for_event('popup'))
        download_task = asyncio.create_task(page.wait_for_event('download'))
        old_url = page.url
        await target_span.click()
        done, pending = await asyncio.wait([popup_task, download_task], timeout=20, return_when=asyncio.FIRST_COMPLETED)

        if download_task in done:
            try:
                download = await download_task
                temp = target.with_suffix('.download')
                await download.save_as(str(temp))
                raw = temp.read_bytes() if temp.exists() else b''
                temp.unlink(missing_ok=True)
                if _bytes_to_pdf(raw, target):
                    return True, f'{target_name} (browser download)'
            except Exception:
                pass

        if popup_task in done:
            try:
                popup = await popup_task
                try:
                    await popup.wait_for_load_state('domcontentloaded', timeout=30000)
                except Exception:
                    pass
                popup_url = popup.url
                if popup_url.startswith(('blob:', 'data:')):
                    saved = await _blob_page_to_pdf(popup, target)
                else:
                    saved = await _request_to_pdf(context, popup_url, target)
                    if not saved:
                        saved = await _blob_page_to_pdf(popup, target)
                try:
                    await popup.close()
                except Exception:
                    pass
                if saved:
                    return True, f'{target_name} (popup/viewer)'
            except Exception:
                pass

        await page.wait_for_timeout(2500)
        for resp in reversed(responses):
            try:
                ctype = str((resp.headers or {}).get('content-type', '')).lower()
                url = str(resp.url or '')
                if 'pdf' in ctype or 'delivery' in url.lower() or 'certificate' in url.lower():
                    if _bytes_to_pdf(await resp.body(), target):
                        return True, f'{target_name} (network response)'
            except Exception:
                continue

        if page.url != old_url:
            if page.url.startswith(('blob:', 'data:')):
                if await _blob_page_to_pdf(page, target):
                    return True, f'{target_name} (same-tab blob)'
            elif await _request_to_pdf(context, page.url, target):
                return True, f'{target_name} (same-tab viewer)'

        return False, f'{target_name} was found and clicked, but the PDF bytes could not be captured.'
    except Exception as e:
        return False, f'Delivery Certificate error: {e}'
    finally:
        try:
            page.remove_listener('response', remember_response)
        except Exception:
            pass
        for task in (popup_task, download_task):
            if task and not task.done():
                task.cancel()


async def get_links_for_order(page, order_id: str) -> list[tuple[str, str]]:
    """Collect normal ShedSuite file URLs. Delivery Certificate is captured separately."""
    try:
        await page.goto(SHEDSUITE_BASE_URL + str(order_id), wait_until='domcontentloaded', timeout=90000)
        await page.wait_for_selector('span[class^="fileName"]', timeout=90000)
        spans = await page.query_selector_all('span[class^="fileName"]')
        files = []
        for span in spans:
            file_name = ''
            try:
                file_name = ((await span.text_content()) or '').strip()
                lower_name = file_name.lower()
                if 'signed combined documents' in lower_name or _delivery_name(file_name):
                    continue
                try:
                    href = await span.evaluate("el => (el.closest('a') && el.closest('a').href) || ''")
                    if href and href.startswith('http'):
                        files.append((file_name, extract_shedsuite_file_url(href)))
                        continue
                except Exception:
                    pass
                popup_task = asyncio.create_task(page.wait_for_event('popup'))
                download_task = asyncio.create_task(page.wait_for_event('download'))
                await span.click()
                done, pending = await asyncio.wait([popup_task, download_task], timeout=15, return_when=asyncio.FIRST_COMPLETED)
                for task in pending:
                    task.cancel()
                file_url = ''
                if popup_task in done:
                    popup = await popup_task
                    try:
                        await popup.wait_for_load_state('domcontentloaded', timeout=30000)
                    except Exception:
                        pass
                    file_url = extract_shedsuite_file_url(popup.url)
                    try:
                        await popup.close()
                    except Exception:
                        pass
                elif download_task in done:
                    file_url = extract_shedsuite_file_url((await download_task).url)
                if file_url:
                    files.append((file_name, file_url))
            except Exception as e:
                print(f'ShedSuite order {order_id}: could not collect {file_name}: {e}')
        return files
    except Exception as e:
        print(f'ShedSuite order {order_id}: failed to open Files list: {e}')
        return []


async def get_shedsuite_files(login_email: str, login_password: str, order_ids: list[str], delivery_dir: Path):
    """Log in once, collect normal links, and save Delivery Certificates locally."""
    auth_path = shedsuite_auth_path(login_email or 'manual')
    delivery_dir.mkdir(parents=True, exist_ok=True)
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context_kwargs = {
            'viewport': {'width': 1100, 'height': 850},
            'accept_downloads': True,
            'user_agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36',
        }
        if auth_path.exists():
            context_kwargs['storage_state'] = str(auth_path)
        context = await browser.new_context(**context_kwargs)
        context.set_default_navigation_timeout(90000)
        context.set_default_timeout(15000)
        page = await context.new_page()
        await page.goto('https://app.shedsuite.com/', wait_until='domcontentloaded', timeout=90000)
        await page.wait_for_timeout(2500)
        login_button = page.get_by_role('button', name='Log in')
        if await login_button.is_visible():
            if login_email and login_password:
                await page.get_by_label('Email').fill(login_email)
                await page.get_by_label('Password').fill(login_password)
                await login_button.click()
            else:
                print('Log into the open ShedSuite Chrome window. The program will continue automatically after login.')
            await expect(page).to_have_url(re.compile(r'https://app\.shedsuite\.com/rto/.*'), timeout=300000)
            await context.storage_state(path=str(auth_path))

        results = {}
        delivery_paths = {}
        delivery_notes = {}
        for order_id in order_ids:
            order_id = str(order_id or '').strip()
            if not order_id:
                continue
            results[order_id] = await get_links_for_order(page, order_id)
            delivery_target = delivery_dir / f'{safe_filename(order_id)}_Delivery_Certificate.pdf'
            ok, note = await download_delivery_certificate_for_order(page, context, order_id, delivery_target)
            delivery_notes[order_id] = note
            if ok and delivery_target.exists():
                delivery_paths[order_id] = str(delivery_target)
                print(f'ShedSuite order {order_id}: Delivery Certificate saved.')
            else:
                print(f'ShedSuite order {order_id}: {note}')
        await context.storage_state(path=str(auth_path))
        await context.close()
        await browser.close()
        return results, delivery_paths, delivery_notes


def append_pdf(target: Path, extra_pdf: Path) -> bool:
    """Append Delivery Certificate after every existing page in target."""
    try:
        if not extra_pdf.exists():
            return False
        if not target.exists():
            shutil.copy2(extra_pdf, target)
            return True
        main = fitz.open(target)
        extra = fitz.open(extra_pdf)
        main.insert_pdf(extra)
        temp = target.with_suffix('.merged.pdf')
        main.save(temp)
        extra.close()
        main.close()
        temp.replace(target)
        return True
    except Exception as e:
        print(f'Could not append Delivery Certificate: {e}')
        return False

def is_invoice(name: str) -> bool:
    n = (name or '').strip().lower()
    return 'invoice' in n


def is_customer_data_sheet(name: str) -> bool:
    n = (name or '').strip().lower()
    return (
        'customer data sheet' in n
        or 'customer datasheet' in n
        or n == 'data sheet'
    )


def is_delivery_certificate(name: str) -> bool:
    return 'delivery certificate' in (name or '').strip().lower()


def is_id_document(name: str) -> bool:
    n = (name or '').strip().lower()
    return any(
        token in n
        for token in [
            'identification card',
            'identification',
            'id card',
            'driver license',
            "driver's license",
            'drivers license',
        ]
    )


def is_contract(name: str) -> bool:
    n = (name or '').strip().lower()

    # These are not the rental contract itself.
    if is_delivery_certificate(n):
        return False
    if 'landowner' in n or 'permission' in n:
        return False

    return (
        'rto contract' in n
        or 'rental purchase agreement' in n
        or 'rental contract' in n
        or n == 'contract'
        or n.endswith(' contract')
    )


def dedupe_urls(items: list[tuple[str, str]]) -> list[tuple[str, str]]:
    """Preserve order while removing duplicate URLs."""
    seen = set()
    result = []

    for name, url in items:
        url = str(url or '').strip()
        if not url or url in seen:
            continue

        seen.add(url)
        result.append((name, url))

    return result


def build_ordered_document_list(
    srcrow: dict[str, str],
    shedsuite_files: list[tuple[str, str]],
) -> list[tuple[str, str]]:
    """
    Build the final document packet in EXACTLY this group order:

      1. Invoice
      2. Customer Data Sheet
      3. Contract
      4. Any other files
      5. ID
      6. Delivery Certificate

    Direct CSV URLs are preferred for Invoice, Contract and ID because they
    are already present in the report. Customer Data Sheet and Delivery
    Certificate are taken from the ShedSuite order page.
    """
    all_files = dedupe_urls(shedsuite_files)

    invoices = [(n, u) for n, u in all_files if is_invoice(n)]
    data_sheets = [(n, u) for n, u in all_files if is_customer_data_sheet(n)]
    contracts = [(n, u) for n, u in all_files if is_contract(n)]
    ids = [(n, u) for n, u in all_files if is_id_document(n)]
    delivery_certs = [(n, u) for n, u in all_files if is_delivery_certificate(n)]

    classified_urls = {
        u
        for group in [invoices, data_sheets, contracts, ids, delivery_certs]
        for _, u in group
    }

    others = [
        (n, u)
        for n, u in all_files
        if u not in classified_urls
    ]

    # CSV URLs are clean direct links. Put them first in their respective
    # categories, then deduplicate against anything found on the order page.
    invoice_csv = (srcrow.get('Invoice') or '').strip()
    contract_csv = (srcrow.get('Rto Contract') or '').strip()
    id_csv = (srcrow.get('Customer Identification Card Scan') or '').strip()
    co_id_csv = (srcrow.get('Co Renter Identification Card Scan') or '').strip()
    landowner_csv = (srcrow.get('Landowner Permission Form') or '').strip()

    if invoice_csv:
        invoices.insert(0, ('Invoice', invoice_csv))

    if contract_csv:
        contracts.insert(0, ('RTO Contract', contract_csv))

    # Landowner permission belongs in "any other files".
    if landowner_csv:
        others.append(('Landowner Permission Form', landowner_csv))

    # Primary ID, then co-renter ID if present.
    if id_csv:
        ids.insert(0, ('Customer Identification Card Scan', id_csv))

    if co_id_csv:
        ids.append(('Co Renter Identification Card Scan', co_id_csv))

    ordered = (
        dedupe_urls(invoices)
        + dedupe_urls(data_sheets)
        + dedupe_urls(contracts)
        + dedupe_urls(others)
        + dedupe_urls(ids)
        + dedupe_urls(delivery_certs)
    )

    return dedupe_urls(ordered)


# ---------------------------------------------------------------------------
# DOB and model-series helpers
# ---------------------------------------------------------------------------

def exact_csv_date(value: str) -> str:
    value = str(value or '').strip()
    if not value:
        return ''
    m = re.fullmatch(r'(\d{1,2})/(\d{1,2})/(\d{4})', value)
    if m:
        return f'{int(m.group(1)):02d}/{int(m.group(2)):02d}/{m.group(3)}'
    m = re.match(r'^(\d{4})-(\d{2})-(\d{2})', value)
    if m:
        return f'{m.group(2)}/{m.group(3)}/{m.group(1)}'
    return value


def fix_dobs_from_source(rows, source_rows):
    by_order = {str(x.get('Customer Order Id', '')).strip(): x for x in source_rows}
    for r in rows:
        src = by_order.get(str(r.get('_source_order_id', '')).strip(), {})
        dob1 = exact_csv_date(src.get('Customer Date Of Birth') or src.get('Customer DOB') or '')
        dob2 = exact_csv_date(src.get('Co Renter Date Of Birth') or src.get('Co Renter DOB') or '')
        if dob1:
            r['DOB1'] = dob1
        if dob2:
            r['DOB2'] = dob2
        r['_source_dob1'] = dob1
        r['_source_dob2'] = dob2


def _slug(value):
    return re.sub(r'[^a-z0-9]+', '_', str(value or '').strip().lower()).strip('_') or 'unknown'


def model_series_key(src):
    # Pick the correct company/location-specific model-number sequence.
    company = str(src.get('Company Name', '') or '').strip()
    dealer = str(src.get('Dealer', '') or '').strip()
    combined = f'{company} {dealer}'.lower()
    state = str(
        src.get('Company State')
        or src.get('Physical Destination State')
        or src.get('Customer Mailing State')
        or ''
    ).strip().upper()

    # DBM Phoenix MUST be checked before CRF Phoenix.
    if 'dbm' in combined and 'phoenix' in combined:
        return 'dbm_phoenix'

    if re.search(r'\bwwp\b', combined):
        return 'wwp'

    if re.search(r'\bras\b', combined) and re.search(r'\byss\b', combined):
        return 'ras_yss'

    if re.search(r'\blsr\b', combined) and (
        state == 'TX'
        or re.search(r'\btx\b', combined)
        or 'texas' in combined
    ):
        return 'lsr_tx'

    if ('smart shed' in combined or 'smart-shed' in combined or 'magnolia' in combined) and state in {'GA', 'SC', 'NC', 'TN'}:
        return f'magnolia_smart_shed_{state.lower()}'

    if 'genesis' in combined:
        return 'crf_genesis'
    if 'alpine' in combined:
        return 'crf_alpine'
    if 'phoenix' in combined and 'crf' in combined:
        return 'crf_phoenix'
    if '4 seasons' in combined or 'four seasons' in combined:
        return 'crf_4_seasons'

    # Any future company still gets its own editable tracker card.
    return 'company_' + _slug(company or dealer)


def load_model_tracker():
    tracker = {k: dict(v) for k, v in DEFAULT_MODEL_SERIES.items()}
    if MODEL_TRACKER_FILE.exists():
        try:
            saved = json.loads(MODEL_TRACKER_FILE.read_text(encoding='utf-8'))
            if isinstance(saved, dict):
                for key, value in saved.items():
                    if isinstance(value, dict):
                        tracker.setdefault(key, {})
                        tracker[key].update(value)
        except Exception:
            pass
    return tracker


def save_model_tracker(tracker):
    MODEL_TRACKER_FILE.write_text(json.dumps(tracker, indent=2), encoding='utf-8')


def _suffix_number(model, prefix):
    model = str(model or '').strip()
    prefix = str(prefix or '').strip()
    if not model:
        return None
    m = re.fullmatch(re.escape(prefix) + r'-(\d+)', model) if prefix else re.search(r'(\d+)$', model)
    return int(m.group(1)) if m else None


def next_model(last_used, prefix):
    last_used = str(last_used or '').strip()
    prefix = str(prefix or '').strip()

    if not last_used:
        return f'{prefix}-?' if prefix else '?'

    if prefix:
        match = re.fullmatch(re.escape(prefix) + r'-(\d+)', last_used)
    else:
        match = re.search(r'(\d+)$', last_used)

    if not match:
        return f'{prefix}-?' if prefix else '?'

    suffix_text = match.group(1)
    next_number = str(int(suffix_text) + 1).zfill(len(suffix_text))
    return f'{prefix}-{next_number}' if prefix else next_number


def update_model_tracker(rows, source_rows):
    tracker = load_model_tracker()
    by_order = {str(x.get('Customer Order Id', '')).strip(): x for x in source_rows}
    for r in rows:
        src = by_order.get(str(r.get('_source_order_id', '')).strip(), {})
        if not src:
            continue
        key = model_series_key(src)
        if key not in tracker:
            label = str(src.get('Company Name') or src.get('Dealer') or 'Other company').strip()
            tracker[key] = {'label': label, 'prefix': '', 'last_used': '', 'example': ''}
        prefix = str(tracker[key].get('prefix', '') or '')
        model = str(r.get('MODEL1', '') or '').strip()
        new_n = _suffix_number(model, prefix)
        old_n = _suffix_number(tracker[key].get('last_used', ''), prefix)
        if new_n is not None and (old_n is None or new_n > old_n):
            tracker[key]['last_used'] = model
            old_n = new_n

        # Your existing V2/V3 transform already had a Firebird-derived
        # "Next model" concept. If that helper is present under any metadata
        # key containing both NEXT and MODEL, use it to infer the real last-used
        # model without depending on the exact metadata key name.
        for meta_key, meta_value in r.items():
            mk = str(meta_key).lower()
            if 'next' not in mk or 'model' not in mk:
                continue
            suggestion = str(meta_value or '').strip()
            suggested_n = _suffix_number(suggestion, prefix)
            if suggested_n is None or suggested_n <= 0:
                continue
            inferred_n = suggested_n - 1
            inferred = f'{prefix}-{inferred_n}' if prefix else str(inferred_n)
            current_n = _suffix_number(tracker[key].get('last_used', ''), prefix)
            if current_n is None or inferred_n > current_n:
                tracker[key]['last_used'] = inferred
            break
    save_model_tracker(tracker)


# ---------------------------------------------------------------------------
# Live RTO Pro / Firebird refresh helpers
# ---------------------------------------------------------------------------

def _job_config(jp: Path) -> dict:
    p = jp / 'job_config.json'
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding='utf-8'))
    except Exception:
        return {}


def _job_db_credentials(jp: Path) -> tuple[str, str, str]:
    cfg = _job_config(jp)
    dsn = str(cfg.get('dsn') or os.getenv('RTO_DSN', 'Firebird DSN')).strip()
    user = str(cfg.get('db_user') or os.getenv('RTO_DB_USER', '')).strip()
    password = JOB_DB_PASSWORDS.get(jp.name, '') or os.getenv('RTO_DB_PASSWORD', '')
    return dsn, user, password


def _open_rto_connection(dsn: str, user: str, password: str):
    if pyodbc is None:
        raise RuntimeError('pyodbc is not installed. Run: pip install pyodbc')
    parts = [f'DSN={dsn}']
    if user:
        parts.append(f'UID={user}')
    if password:
        parts.append(f'PWD={password}')
    return pyodbc.connect(';'.join(parts) + ';', timeout=15)


def fetch_rto_contract_accounts(dsn: str, user: str, password: str) -> list[tuple[str, str]]:
    """Read the same contracts data that the Excel ExistingContracts query uses."""
    conn = _open_rto_connection(dsn, user, password)
    try:
        cur = conn.cursor()
        cur.execute('SELECT CONTRACT, ACCOUNT FROM CONTRACTS')
        result = []
        for contract, account in cur.fetchall():
            c = str(contract or '').strip()
            if not c:
                continue
            if account is None:
                a = ''
            else:
                try:
                    a = str(int(float(str(account).strip())))
                except Exception:
                    a = str(account).strip()
            result.append((c, a))
        return result
    finally:
        conn.close()


def refresh_model_tracker_from_contracts(contract_accounts: list[tuple[str, str]]) -> dict:
    """Refresh last-used model/contract numbers from LIVE RTO Pro data."""
    tracker = load_model_tracker()
    contracts = [str(c or '').strip() for c, _ in contract_accounts if str(c or '').strip()]
    refreshed_at = datetime.now().strftime('%Y-%m-%d %I:%M:%S %p')

    for key, item in tracker.items():
        rule = MODEL_DB_RULES.get(key)
        prefix = str((rule or {}).get('prefix') or item.get('prefix') or '').strip()
        if not prefix:
            continue

        low = (rule or {}).get('min')
        high = (rule or {}).get('max')
        best_contract = ''
        best_n = None
        pattern = re.compile(r'^' + re.escape(prefix) + r'-(\d+)$')

        for contract in contracts:
            m = pattern.match(contract)
            if not m:
                continue
            n = int(m.group(1))
            if low is not None and n < low:
                continue
            if high is not None and n > high:
                continue
            if best_n is None or n > best_n:
                best_n = n
                best_contract = contract

        if best_contract:
            item['last_used'] = best_contract
            item['source'] = 'RTO Pro / Firebird'
            item['refreshed_at'] = refreshed_at

    save_model_tracker(tracker)
    return tracker


def refresh_job_from_rto(jp: Path, rows: list[dict]) -> dict:
    """Refresh reference lists, model series, and newly assigned account numbers."""
    dsn, user, password = _job_db_credentials(jp)

    ref = load_reference_data(dsn, user, password)
    refs = ref.as_client_dict()
    (jp / 'refs.json').write_text(json.dumps(refs, indent=2), encoding='utf-8')

    contract_accounts = fetch_rto_contract_accounts(dsn, user, password)
    refresh_model_tracker_from_contracts(contract_accounts)

    account_by_contract = {
        str(c).strip(): str(a).strip()
        for c, a in contract_accounts
        if str(a).strip()
    }

    matched = 0
    for r in rows:
        account = ''
        matched_contract = ''
        for candidate in [r.get('CONTRACT', ''), r.get('MODEL1', '')]:
            candidate = str(candidate or '').strip()
            if candidate and candidate in account_by_contract:
                account = account_by_contract[candidate]
                matched_contract = candidate
                break
        if account:
            r['ACCOUNT'] = account
            r['_rto_account_found'] = account
            r['_rto_contract_found'] = matched_contract
            matched += 1

    save_job_rows(jp, rows)

    (jp / 'last_rto_refresh.json').write_text(
        json.dumps(
            {
                'refreshed_at': datetime.now().isoformat(),
                'contracts_read': len(contract_accounts),
                'accounts_matched': matched,
            },
            indent=2,
        ),
        encoding='utf-8',
    )

    return {
        'contracts_read': len(contract_accounts),
        'accounts_matched': matched,
        'refs': refs,
    }


def rto_images_root() -> Path:
    configured = os.getenv('RTO_IMAGES_ROOT', '').strip()
    if configured:
        return Path(configured)

    if os.getenv('RTO_TEST_MODE', 'false').lower() == 'true':
        return Path(r'C:\RTO Pro\TestDB\Images')

    return Path(r'\\evg-s01\Company Data\RTO Pro\Evergreen\Images')


def copy_pdf_to_rto_imaging(src: Path, account: str) -> Path:
    """Reproduce the Excel copy_pdfs naming/folder behavior."""
    account = str(account or '').strip()

    if not account:
        raise RuntimeError('RTO account number is blank.')

    try:
        account_int = int(float(account))
    except Exception:
        raise RuntimeError(f'Invalid RTO account number: {account}')

    root = rto_images_root()
    target_folder = root / str(account_int)
    target_folder.mkdir(parents=True, exist_ok=True)

    existing_files = [
        p
        for p in target_folder.iterdir()
        if p.is_file()
    ]

    sequence = len(existing_files) + 1
    filename = f'{str(account_int).zfill(10)}-{sequence:02d}.pdf'
    target = target_folder / filename

    shutil.copy2(src, target)
    return target


def workflow_panel_html(job: str, rows: list[dict], refs: dict) -> str:
    connected = bool(refs.get('connected'))

    status_text = (
        'Firebird connected'
        if connected
        else (
            'Firebird not connected'
            + (
                ': ' + str(refs.get('error'))
                if refs.get('error')
                else ''
            )
        )
    )

    status_bg = '#dcfce7' if connected else '#fee2e2'
    status_fg = '#166534' if connected else '#991b1b'

    pdf_rows = []
    available_count = 0

    for i, r in enumerate(rows):
        name = html_lib.escape(str(r.get('NAME') or f'Contract {i + 1}'))
        order_id = html_lib.escape(str(r.get('_source_order_id', '') or ''))
        available = bool(r.get('_pdf_available'))

        if available:
            available_count += 1
            open_url = url_for('pdf', job=job, index=i)
            download_url = open_url + '?download=1'

            manifest = r.get('_pdf_manifest', [])
            if isinstance(manifest, list):
                manifest_text = ' → '.join(str(x) for x in manifest)
            else:
                manifest_text = str(manifest or '')

            manifest_text = html_lib.escape(manifest_text)
            pages = html_lib.escape(str(r.get('_pdf_pages', '') or ''))

            buttons = (
                f'<a href="{open_url}" target="_blank" '
                'style="padding:6px 9px;background:#2563eb;color:white;'
                'border-radius:6px;text-decoration:none">Open PDF</a> '
                f'<a href="{download_url}" '
                'style="padding:6px 9px;background:#475569;color:white;'
                'border-radius:6px;text-decoration:none">Download</a>'
            )

            detail = (pages + ' pages · ' if pages else '') + manifest_text
        else:
            buttons = (
                '<span style="color:#b91c1c;font-weight:700">'
                'No combined PDF</span>'
            )
            detail = 'Re-process the CSV with Download PDFs enabled.'

        pdf_rows.append(
            '<div style="display:grid;'
            'grid-template-columns:minmax(180px,1fr) '
            'minmax(180px,.6fr) minmax(260px,2fr);'
            'gap:10px;align-items:center;padding:9px 0;'
            'border-bottom:1px solid #e5e7eb">'
            f'<div><b>{name}</b>'
            f'<div style="font-size:11px;color:#64748b">'
            f'Order {order_id}</div></div>'
            f'<div>{buttons}</div>'
            f'<div style="font-size:11px;color:#475569">'
            f'{detail}</div>'
            '</div>'
        )

    return f"""
<section style="max-width:1500px;margin:16px auto;padding:0 16px;
                font-family:Arial,sans-serif">
  <div style="background:white;border:1px solid #cbd5e1;
              border-radius:12px;box-shadow:0 2px 8px rgba(15,23,42,.08);
              overflow:hidden">
    <div style="background:#0f172a;color:white;padding:13px 15px;
                display:flex;justify-content:space-between;gap:10px;
                align-items:center">
      <div>
        <div style="font-size:16px;font-weight:800">RTO Pro Workflow</div>
        <div style="font-size:11px;color:#cbd5e1">
          Same sequence as Excel: CSV → RTO Pro → refresh DB → PDF imaging
        </div>
      </div>
      <span style="background:{status_bg};color:{status_fg};font-size:11px;
                   font-weight:800;padding:5px 9px;border-radius:999px">
        {html_lib.escape(status_text)}
      </span>
    </div>

    <div style="padding:14px;display:flex;gap:9px;flex-wrap:wrap;
                align-items:center">
      <form method="post"
            action="{url_for('run_rto', job=job)}"
            onsubmit="return confirm('Launch RTO Pro and import this CSV?')"
            style="margin:0">
        <button type="submit"
                style="padding:9px 13px;background:#16a34a;color:white;
                       border:0;border-radius:7px;font-weight:800;cursor:pointer">
          1. Import to RTO Pro
        </button>
      </form>

      <form method="post"
            action="{url_for('refresh_rto', job=job)}"
            style="margin:0">
        <button type="submit"
                style="padding:9px 13px;background:#2563eb;color:white;
                       border:0;border-radius:7px;font-weight:800;cursor:pointer">
          2. Refresh RTO Data
        </button>
      </form>

      <form method="post"
            action="{url_for('import_pdfs_to_rto', job=job)}"
            onsubmit="return confirm('Only click this AFTER RTO Pro confirms the customer import. Continue?')"
            style="margin:0">
        <button type="submit"
                style="padding:9px 13px;background:#7c3aed;color:white;
                       border:0;border-radius:7px;font-weight:800;cursor:pointer">
          3. Import PDFs to RTO Imaging
        </button>
      </form>

      <a href="{url_for('download', job=job, kind='csv')}"
         style="padding:9px 13px;background:#475569;color:white;
                border-radius:7px;text-decoration:none;font-weight:700">
        Download RTO CSV
      </a>
    </div>

    <div style="padding:0 14px 13px;color:#475569;font-size:12px">
      After RTO Pro shows its successful import confirmation, click
      <b>Refresh RTO Data</b>. That reads the updated Firebird tables,
      fills assigned account numbers, and refreshes last-used model numbers.
      Then click <b>Import PDFs to RTO Imaging</b>.
    </div>
  </div>
</section>

<section style="max-width:1500px;margin:16px auto;padding:0 16px;
                font-family:Arial,sans-serif">
  <details open
           style="background:white;border:1px solid #cbd5e1;
                  border-radius:12px;box-shadow:0 2px 8px rgba(15,23,42,.08);
                  overflow:hidden">
    <summary style="cursor:pointer;background:#f8fafc;padding:12px 14px;
                    font-weight:800">
      Combined PDFs — {available_count}/{len(rows)} ready
    </summary>
    <div style="padding:4px 14px 12px">
      {''.join(pdf_rows)}
    </div>
  </details>
</section>
"""


def inject_review_panels(
    page_html: str,
    source_rows: list[dict],
    rows: list[dict],
    job: str,
    refs: dict,
) -> str:

    panels = (
        workflow_panel_html(job, rows, refs)
        + model_tracker_panel_html(source_rows)
    )

    match = re.search(
        r'<body[^>]*>',
        page_html,
        flags=re.IGNORECASE,
    )

    if match:
        insert_at = match.end()
        return page_html[:insert_at] + panels + page_html[insert_at:]

    return panels + page_html


def model_tracker_panel_html(source_rows):
    tracker = load_model_tracker()
    active = {model_series_key(x) for x in source_rows}

    # Keep every known company above, and automatically add new companies from future CSV files.
    for src in source_rows:
        key = model_series_key(src)
        if key not in tracker:
            tracker[key] = {
                'label': str(src.get('Company Name') or src.get('Dealer') or key),
                'prefix': '',
                'last_used': '',
                'example': '',
            }
    save_model_tracker(tracker)

    cards = []
    for key, item in tracker.items():
        label = html_lib.escape(str(item.get('label', key)))
        prefix = html_lib.escape(str(item.get('prefix', '') or ''))
        last_used = html_lib.escape(str(item.get('last_used', '') or ''))
        nxt = html_lib.escape(next_model(item.get('last_used', ''), item.get('prefix', '')))
        example = html_lib.escape(str(item.get('example', '') or ''))
        source_note = html_lib.escape(str(item.get('source', 'Manual/fallback') or 'Manual/fallback'))
        refreshed_at = html_lib.escape(str(item.get('refreshed_at', '') or ''))

        if key in active:
            card_style = 'border:2px solid #2563eb;background:#eff6ff;'
            badge = '<span style="font-size:10px;font-weight:700;background:#2563eb;color:white;padding:2px 6px;border-radius:999px">USED BY THIS CSV</span>'
        else:
            card_style = 'border:1px solid #d1d5db;background:white;'
            badge = ''

        example_html = (
            '<div style="margin-top:7px;font-size:11px;color:#6b7280"><b>Reference example:</b> ' + example + '</div>'
            if example else ''
        )

        cards.append(
            '<div data-series="' + html_lib.escape(key) + '" style="border-radius:9px;padding:11px;' + card_style + '">' +
            '<div style="display:flex;justify-content:space-between;align-items:center;gap:8px;margin-bottom:8px">' +
            '<div style="font-weight:750;font-size:13px">' + label + '</div>' + badge + '</div>' +
            '<div style="display:grid;grid-template-columns:72px 1fr;gap:7px;align-items:center;font-size:12px">' +
            '<span style="color:#475569">Prefix</span>' +
            '<input class="model-prefix" value="' + prefix + '" style="width:100%;padding:6px 7px;border:1px solid #cbd5e1;border-radius:6px;background:#f8fafc">' +
            '<span style="color:#475569">Last used</span>' +
            '<input class="model-last" value="' + last_used + '" placeholder="enter actual last used" style="width:100%;padding:6px 7px;border:1px solid #cbd5e1;border-radius:6px">' +
            '<span style="color:#475569">Next</span>' +
            '<strong class="model-next" style="font-size:14px;color:#0f172a">' + nxt + '</strong>' +
            '</div>' + example_html + '<div style="margin-top:6px;font-size:10px;color:#64748b">Source: ' + source_note + ((' · ' + refreshed_at) if refreshed_at else '') + '</div></div>'
        )

    panel = r'''
<section id="model-series-panel" style="max-width:1500px;margin:16px auto;padding:0 16px;font-family:Arial,sans-serif">
  <div style="background:white;border:1px solid #cbd5e1;border-radius:12px;box-shadow:0 2px 8px rgba(15,23,42,.08);overflow:hidden">
    <div style="display:flex;align-items:center;justify-content:space-between;gap:14px;background:#111827;color:white;padding:12px 14px">
      <div>
        <div style="font-size:15px;font-weight:800">Last Used Model Numbers</div>
        <div style="font-size:11px;color:#cbd5e1;margin-top:2px">Live RTO Pro / Firebird data is preferred. Manual values are only a fallback.</div>
      </div>
      <button id="model-series-toggle" type="button" onclick="toggleModelSeries()" style="background:#334155;color:white;border:1px solid #64748b;border-radius:7px;padding:6px 11px;cursor:pointer;white-space:nowrap">Minimize</button>
    </div>

    <div id="model-series-content">
      <div style="padding:13px;display:grid;grid-template-columns:repeat(auto-fit,minmax(245px,1fr));gap:10px">
''' + ''.join(cards) + r'''
      </div>
      <div style="padding:0 13px 13px 13px;display:flex;justify-content:flex-end;gap:8px">
        <button type="button" onclick="saveModelSeries()" style="padding:8px 14px;background:#2563eb;color:white;border:0;border-radius:7px;cursor:pointer;font-weight:700">Save manual fallback</button>
      </div>
    </div>
  </div>
</section>

<script>
function calcNext(prefix, lastUsed) {
    const p = (prefix || '').trim();
    const l = (lastUsed || '').trim();
    const match = l.match(/(\d+)$/);
    if (!match) return p ? p + '-?' : '?';
    const oldSuffix = match[1];
    const nextValue = String(parseInt(oldSuffix, 10) + 1).padStart(oldSuffix.length, '0');
    return p ? p + '-' + nextValue : nextValue;
}

function setModelSeriesMinimized(minimized) {
    const content = document.getElementById('model-series-content');
    const button = document.getElementById('model-series-toggle');
    if (!content || !button) return;
    content.style.display = minimized ? 'none' : 'block';
    button.textContent = minimized ? 'Expand' : 'Minimize';
    localStorage.setItem('modelTrackerMinimized', minimized ? 'yes' : 'no');
}

function toggleModelSeries() {
    const content = document.getElementById('model-series-content');
    if (!content) return;
    setModelSeriesMinimized(content.style.display !== 'none');
}

function initializeModelSeries() {
    document.querySelectorAll('#model-series-panel [data-series]').forEach(el => {
        const p = el.querySelector('.model-prefix');
        const l = el.querySelector('.model-last');
        const n = el.querySelector('.model-next');
        const refresh = () => n.textContent = calcNext(p.value, l.value);
        p.addEventListener('input', refresh);
        l.addEventListener('input', refresh);
        refresh();
    });
    setModelSeriesMinimized(localStorage.getItem('modelTrackerMinimized') === 'yes');
}

async function saveModelSeries() {
    const data = {};
    document.querySelectorAll('#model-series-panel [data-series]').forEach(el => {
        data[el.dataset.series] = {
            prefix: el.querySelector('.model-prefix').value.trim(),
            last_used: el.querySelector('.model-last').value.trim()
        };
    });
    const r = await fetch('/api/model-tracker', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify(data)
    });
    alert(r.ok ? 'Model numbers saved.' : 'Could not save model numbers.');
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeModelSeries);
} else {
    initializeModelSeries();
}
</script>
'''
    return panel


def inject_model_tracker(page_html, source_rows):
    # Put the card in normal document flow at the TOP of the review page.
    # No position:fixed, so it does not float over the table.
    panel = model_tracker_panel_html(source_rows)
    match = re.search(r'<body[^>]*>', page_html, flags=re.IGNORECASE)
    if match:
        insert_at = match.end()
        return page_html[:insert_at] + panel + page_html[insert_at:]
    return panel + page_html


# ---------------------------------------------------------------------------
# Existing app helpers
# ---------------------------------------------------------------------------

def job_path(job_id: str) -> Path:
    if not job_id or not all(c.isalnum() for c in job_id):
        raise ValueError('invalid job id')
    p = (WORK / job_id).resolve()
    if WORK.resolve() not in p.parents:
        raise ValueError('invalid job path')
    return p


def read_csv_file(path: Path) -> list[dict[str, str]]:
    with path.open('r', encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f))


def write_import_csv(rows: list[dict], path: Path):
    # Keep the downloadable/job copy UTF-8 with BOM for Excel friendliness.
    # RTO Pro itself gets a separate ANSI/CP1252 staging copy immediately
    # before import (see stage_rto_import).
    with path.open('w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fieldnames=RTO_HEADERS, extrasaction='ignore')
        w.writeheader()
        for r in rows:
            w.writerow({h: r.get(h, '') for h in RTO_HEADERS})


def validate_rto_runtime(rows: list[dict]) -> list[str]:
    """Catch the RTO Pro field limits that can block a live import."""
    errors = []
    for i, r in enumerate(rows, 1):
        if not str(r.get('NAME') or '').strip():
            errors.append(f'Row {i}: NAME is blank.')
        if not str(r.get('STORE') or '').strip():
            errors.append(f'Row {i}: STORE is blank.')
        if len(str(r.get('CONTRACT') or '')) > 10:
            errors.append(f'Row {i}: CONTRACT exceeds 10 characters.')
        if len(str(r.get('MODEL1') or '')) > 15:
            errors.append(f'Row {i}: MODEL1 exceeds 15 characters.')
    return errors


def _write_rto_runtime_csv(rows: list[dict], path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open(
        'w', encoding='cp1252', errors='replace', newline=''
    ) as f:
        w = csv.DictWriter(
            f,
            fieldnames=RTO_HEADERS,
            extrasaction='ignore',
            lineterminator='\r\n',
        )
        w.writeheader()
        for r in rows:
            w.writerow({h: r.get(h, '') for h in RTO_HEADERS})


def stage_rto_import(rows: list[dict], jp: Path, exe: Path) -> Path:
    """Stage the V3 rows as the writable ANSI app.csv RTO Pro expects.

    This is the proven V5.3 launch behavior: try an explicit override, then
    the RTO folder, LocalAppData, TEMP, and finally the job folder.
    """
    jp.mkdir(parents=True, exist_ok=True)
    configured = os.getenv('RTO_IMPORT_STAGING', '').strip()
    local_app = os.getenv('LOCALAPPDATA', '').strip()
    temp_root = os.getenv('TEMP', '').strip()

    candidates: list[Path] = []
    if configured:
        candidates.append(Path(configured))
    candidates.append(exe.parent / 'app.csv')
    if local_app:
        candidates.append(Path(local_app) / 'ShedSuiteRTO' / 'app.csv')
    if temp_root:
        candidates.append(Path(temp_root) / 'ShedSuiteRTO' / 'app.csv')
    candidates.append(jp / 'app.csv')

    seen = set()
    failures = []
    for staging in candidates:
        key = str(staging).lower()
        if key in seen:
            continue
        seen.add(key)
        if '-store' in staging.name.lower():
            failures.append(f'{staging}: filename contains -store')
            continue
        if staging.suffix.lower() != '.csv':
            failures.append(f'{staging}: not a .csv file')
            continue
        try:
            _write_rto_runtime_csv(rows, staging)
            if not staging.exists() or staging.stat().st_size < 10:
                raise OSError('file was not created')
            with staging.open(
                'r', encoding='cp1252', errors='replace', newline=''
            ) as f:
                actual_headers = next(csv.reader(f), [])
            if actual_headers != RTO_HEADERS:
                raise RuntimeError(
                    'header row does not match the RTO import fields'
                )
            shutil.copy2(staging, jp / 'RTOProImport_LAUNCHED.csv')
            return staging
        except (PermissionError, OSError) as exc:
            failures.append(f'{staging}: {exc}')
            continue

    raise RuntimeError(
        'Could not create a writable RTO app.csv. Tried: ' + ' | '.join(failures)
    )


def write_review_csv(rows: list[dict], path: Path):
    fields = [
        'NAME',
        '_source_order_id',
        'SERIAL1',
        '_source_company',
        '_source_dealer',
        'STORE',
        '_store_title',
        'DEALERID',
        '_dealer_name',
        'MODEL1',
        'CONTRACT',
        '_set_contract',
        'ZONE',
        'TAXZONE',
        'TAXRATE',
        'SALESMAN',
        'AGENT1',
        'INVOICE1',
        'DUEDATE',
        'SACDATE',
        'CATEGORY1',
    ]
    with path.open('w', encoding='utf-8-sig', newline='') as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction='ignore')
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, '') for k in fields})


def write_xmls(rows: list[dict], folder: Path, pdf_folder: Path):
    folder.mkdir(parents=True, exist_ok=True)

    for old in folder.glob('*'):
        if old.is_file():
            old.unlink()

    for r in rows:
        root = Element('application')

        # V1.1.7.6 XML exporter begins with Import Data SS column B,
        # intentionally omitting ACCOUNT.
        for h in RTO_HEADERS[1:]:
            v = str(r.get(h, '') or '').strip()
            if not v:
                continue
            e = SubElement(root, h)
            e.text = v

        pretty = xml.dom.minidom.parseString(
            tostring(root, encoding='unicode')
        ).toprettyxml(indent='  ')

        base = safe_filename(r.get('_pdf_name') or r.get('NAME') or 'contract')
        (folder / f'app{base}.xml').write_text(pretty, encoding='utf-8')

        src = pdf_folder / f'{base}.pdf'
        if src.exists():
            shutil.copy2(src, folder / f'app{base}.pdf')


def build_package(jp: Path):
    target = jp / 'RTO_Import_Package.zip'

    with ZipFile(target, 'w', ZIP_DEFLATED) as z:
        for name in ['RTOProImport.csv', 'Review_Report.csv', 'PDF_Import_Report.csv', 'RTOProImport_LAUNCHED.csv', 'Run_RTO_Import.cmd', 'RTO_Launch_Command.txt']:
            p = jp / name
            if p.exists():
                z.write(p, name)

        for dirname in ['XML_Files', 'Combined_Files']:
            root = jp / dirname
            if root.exists():
                for p in root.rglob('*'):
                    if p.is_file():
                        z.write(p, p.relative_to(jp))


def save_job_rows(jp: Path, rows: list[dict]):
    (jp / 'rows.json').write_text(
        json.dumps(rows, indent=2),
        encoding='utf-8',
    )
    write_import_csv(rows, jp / 'RTOProImport.csv')
    write_review_csv(rows, jp / 'Review_Report.csv')
    write_xmls(rows, jp / 'XML_Files', jp / 'Combined_Files')
    build_package(jp)


def load_job(job: str):
    jp = job_path(job)

    if not (jp / 'rows.json').exists():
        abort(404)

    rows = json.loads(
        (jp / 'rows.json').read_text(encoding='utf-8')
    )

    refs = (
        json.loads((jp / 'refs.json').read_text(encoding='utf-8'))
        if (jp / 'refs.json').exists()
        else {
            'stores': KNOWN_STORES,
            'dealers': [],
            'zones': [],
            'taxzones': [],
            'agents': [],
            'connected': False,
            'error': '',
        }
    )

    warnings = (
        json.loads((jp / 'warnings.json').read_text(encoding='utf-8'))
        if (jp / 'warnings.json').exists()
        else []
    )

    return jp, rows, refs, warnings


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get('/')
def index():
    return render_template(
        'index.html',
        dsn=os.getenv('RTO_DSN', 'Firebird DSN'),
    )


@app.post('/process')
def process():
    f = request.files.get('csv_file')

    if not f or not f.filename:
        flash('Choose the ShedSuite contracts CSV.', 'error')
        return redirect(url_for('index'))

    job = uuid.uuid4().hex[:12]
    jp = job_path(job)
    jp.mkdir(parents=True, exist_ok=True)

    src = jp / 'input.csv'
    f.save(src)

    try:
        source_rows = read_csv_file(src)
    except Exception as e:
        flash(f'Could not read the CSV: {e}', 'error')
        return redirect(url_for('index'))

    required = {
        'Serial Number',
        'Customer Order Id',
        'Customer First Name',
        'Customer Last Name',
        'Company Name',
        'Dealer',
    }

    if not source_rows or not required.issubset(source_rows[0].keys()):
        flash(
            'This does not match the ShedSuite RTO contracts report format.',
            'error',
        )
        return redirect(url_for('index'))

    dsn = request.form.get('dsn', '').strip()
    db_user = request.form.get('db_user', '').strip()
    db_password = request.form.get('db_password', '')
    JOB_DB_PASSWORDS[job] = db_password

    ref = load_reference_data(dsn, db_user, db_password)

    categories = load_categories(
        BASE / 'DataMaps' / 'Categories.csv'
    )

    rows, warnings = transform_rows(
        source_rows,
        categories,
        ref,
        ziptax_api_key=effective_ziptax_api_key(
            os.getenv('ZIPTAX_API_KEY', '')
        ),
    )

    # Keep DOB as the exact calendar date written in the ShedSuite CSV.
    fix_dobs_from_source(rows, source_rows)

    # Read LAST-USED numbers from the live RTO Pro contracts table.
    # Do not count this pending CSV as "used" until RTO Pro actually imports it.
    try:
        live_contracts = fetch_rto_contract_accounts(
            dsn,
            db_user,
            db_password,
        )
        refresh_model_tracker_from_contracts(
            live_contracts
        )
    except Exception as e:
        print(
            'Live model-number refresh skipped: '
            + str(e)
        )

    # Keep DSN name only for later refresh/copy actions.
    # Never persist DB password.
    (jp / 'job_config.json').write_text(
        json.dumps(
            {
                'dsn': dsn,
                'db_user': db_user,
            },
            indent=2,
        ),
        encoding='utf-8',
    )

    refs = ref.as_client_dict()

    (jp / 'refs.json').write_text(
        json.dumps(refs, indent=2),
        encoding='utf-8',
    )

    (jp / 'source_rows.json').write_text(
        json.dumps(source_rows, indent=2),
        encoding='utf-8',
    )

    pdf_dir = jp / 'Combined_Files'
    pdf_dir.mkdir(exist_ok=True)

    pdf_errors = []

    for r in rows:
        r['_pdf_available'] = False
        r['_pdf_salesman'] = ''
        r['_pdf_coordinates'] = ''
        r['_pdf_latitude'] = ''
        r['_pdf_longitude'] = ''
        r['_pdf_invoice'] = ''
        r['_pdf_manifest'] = []
        r['_pdf_pages'] = ''

    # -----------------------------------------------------------------------
    # PDF DOWNLOAD / COMBINE
    #
    # Final combined PDF order:
    #   1. Invoice
    #   2. Customer Data Sheet
    #   3. Contract
    #   4. Any other files
    #   5. ID
    #   6. Delivery Certificate
    # -----------------------------------------------------------------------
    if request.form.get('download_pdfs') == 'yes':
        by_order = {
            str(x.get('Customer Order Id', '')).strip(): x
            for x in source_rows
        }
        pdf_mode = request.form.get('pdf_mode', 'contract')

        # V6 keeps V3's fast direct-link behavior for the normal packet.
        # Chromium is NOT used to crawl every ShedSuite file anymore.
        for r in rows:
            order_id = str(r.get('_source_order_id', '')).strip()
            srcrow = by_order.get(order_id, {})
            ordered_documents = build_ordered_document_list(srcrow, [])

            if pdf_mode != 'all':
                ordered_documents = [
                    (name, url)
                    for name, url in ordered_documents
                    if is_invoice(name) or is_contract(name)
                ]

            urls = [url for _name, url in ordered_documents]
            base = safe_filename(r['_pdf_name'])
            target = pdf_dir / f'{base}.pdf'

            if urls:
                errs = download_and_combine(urls, target)
                pdf_errors.extend([f'{r["NAME"]}: {x}' for x in errs])

            r['_pdf_manifest'] = [name for name, _url in ordered_documents]
            r['_pdf_available'] = target.exists()

        # V5.3 capability transplanted into V3: open ShedSuite only for the
        # Delivery Certificate and append it LAST.  The module honors V3's
        # `_login` company mapping so it does not search unrelated companies.
        try:
            cert_errors = append_delivery_certificates(
                rows, source_rows, pdf_dir, BASE
            )
            pdf_errors.extend(cert_errors)
        except Exception as e:
            pdf_errors.append('Delivery Certificate lookup failed: ' + str(e))

        # Run V3's existing PDF extraction against the final packet.
        for r in rows:
            base = safe_filename(r['_pdf_name'])
            target = pdf_dir / f'{base}.pdf'

            if r.get('_delivery_certificate') == 'Downloaded':
                r['_pdf_manifest'].append('Delivery Certificate')

            if not target.exists():
                r['_pdf_available'] = False
                continue

            fields = extract_pdf_fields(target)

            r['_pdf_salesman'] = fields.get('salesman', '')
            r['_pdf_coordinates'] = fields.get('coordinates', '')
            r['_pdf_latitude'] = fields.get('latitude', '')
            r['_pdf_longitude'] = fields.get('longitude', '')
            r['_pdf_invoice'] = fields.get('invoice', '')

            if fields.get('salesman'):
                r['SALESMAN'] = fields['salesman']

            if fields.get('coordinates'):
                existing_direction = r.get('DIRECTIONS', '')
                if fields['coordinates'] not in existing_direction:
                    r['DIRECTIONS'] = (
                        fields['coordinates']
                        + (('      ' + existing_direction) if existing_direction else '')
                    )

            if fields.get('invoice'):
                r['INVOICE1'] = fields['invoice']

            r['_pdf_available'] = True
            try:
                _doc = fitz.open(target)
                r['_pdf_pages'] = len(_doc)
                _doc.close()
            except Exception:
                r['_pdf_pages'] = ''

    warnings.extend(
        'PDF: ' + x
        for x in pdf_errors
    )

    (jp / 'warnings.json').write_text(
        json.dumps(warnings, indent=2),
        encoding='utf-8',
    )

    save_job_rows(jp, rows)

    return redirect(
        url_for(
            'review',
            job=job,
        )
    )


@app.get('/review/<job>')
def review(job):
    jp, rows, refs, warnings = load_job(job)

    source_rows = (
        json.loads(
            (jp / 'source_rows.json').read_text(
                encoding='utf-8'
            )
        )
        if (jp / 'source_rows.json').exists()
        else []
    )

    source_by_order = {
        str(x.get('Customer Order Id', '')).strip(): x
        for x in source_rows
    }

    sources = [
        source_by_order.get(
            str(r.get('_source_order_id', '')).strip(),
            {},
        )
        for r in rows
    ]

    page_html = render_template(
        'review.html',
        job=job,
        rows=rows,
        refs=refs,
        warnings=warnings,
        headers=RTO_HEADERS,
        sources=sources,
    )
    return inject_review_panels(page_html, source_rows, rows, job, refs)


@app.post('/save/<job>')
def save(job):
    jp, old_rows, refs, warnings = load_job(job)

    raw = request.form.get('rows_payload', '')

    try:
        incoming = json.loads(raw)

        if not isinstance(incoming, list) or len(incoming) != len(old_rows):
            raise ValueError('row count changed')

    except Exception as e:
        flash(
            f'Could not save edits: {e}',
            'error',
        )
        return redirect(
            url_for(
                'review',
                job=job,
            )
        )

    metadata = [
        k
        for k in old_rows[0].keys()
        if k.startswith('_')
    ] if old_rows else []

    rows = []

    for old, new in zip(old_rows, incoming):
        r = {
            h: str(
                new.get(
                    h,
                    old.get(h, ''),
                )
                or ''
            ).strip()
            for h in RTO_HEADERS
        }

        for k in metadata:
            r[k] = new.get(
                k,
                old.get(k, ''),
            )

        # Save UI helper metadata as well.
        for k in [
            '_store_title',
            '_dealer_name',
            '_set_contract',
            '_warnings',
        ]:
            if k in new:
                r[k] = new[k]

        rows.append(r)

    source_rows = json.loads((jp / 'source_rows.json').read_text(encoding='utf-8')) if (jp / 'source_rows.json').exists() else []
    save_job_rows(jp, rows)

    flash(
        'All row edits were saved. CSV, XML and ZIP were rebuilt '
        'from the edited values.',
        'ok',
    )

    return redirect(
        url_for(
            'review',
            job=job,
        )
    )


@app.post('/api/model-tracker')
def api_model_tracker():
    incoming = request.get_json(silent=True) or {}
    if not isinstance(incoming, dict):
        return jsonify({'ok': False, 'error': 'Invalid model tracker data'}), 400
    tracker = load_model_tracker()
    for key, values in incoming.items():
        if not isinstance(values, dict):
            continue
        tracker.setdefault(key, {'label': key, 'prefix': '', 'last_used': '', 'example': ''})
        tracker[key]['prefix'] = str(values.get('prefix', tracker[key].get('prefix', '')) or '').strip()
        tracker[key]['last_used'] = str(values.get('last_used', tracker[key].get('last_used', '')) or '').strip()
    save_model_tracker(tracker)
    return jsonify({'ok': True})


@app.post('/api/ziptax/<job>/<int:index>')
def api_ziptax(job, index):
    jp, rows, refs, _ = load_job(job)

    if index < 0 or index >= len(rows):
        abort(404)

    payload = request.get_json(silent=True) or {}

    street = str(
        payload.get('street', '')
        or ''
    ).strip()

    city = str(
        payload.get('city', '')
        or ''
    ).strip()

    state = str(
        payload.get('state', '')
        or ''
    ).strip().upper()

    zipcode = str(
        payload.get('zip', '')
        or ''
    ).strip()

    csv_rate = payload.get(
        'csv_rate',
        rows[index].get('_ziptax_total_rate')
        or rows[index].get('TAXRATE')
        or 0,
    )

    try:
        csv_rate = float(
            str(csv_rate or '0')
            .replace('%', '')
            .strip()
        )

        if csv_rate > 1:
            csv_rate /= 100.0

    except Exception:
        csv_rate = 0.0

    address = f'{street} {city} {state} {zipcode}'.strip()

    detail = ziptax_detail(
        address,
        effective_ziptax_api_key(
            os.getenv(
                'ZIPTAX_API_KEY',
                '',
            )
        ),
    )

    ref = ReferenceData(
        taxzones=refs.get(
            'taxzones',
            [],
        )
    )

    zone, rate, county, note = resolve_taxzone(
        ref,
        state,
        address,
        csv_rate,
        effective_ziptax_api_key(
            os.getenv(
                'ZIPTAX_API_KEY',
                '',
            )
        ),
        detail=detail,
    )

    return jsonify(
        {
            'ok': not bool(detail.get('error'))
            and not bool(detail.get('limit')),
            'zone': zone,
            'rate': rate,
            'county': county,
            'city': detail.get('city', ''),
            'incorporated': detail.get('incorporated'),
            'normalized_address': detail.get(
                'normalized_address',
                '',
            ),
            'total_rate': detail.get('total_rate', ''),
            'response_code': detail.get(
                'response_code',
                '',
            ),
            'note': note,
            'error': detail.get('error', ''),
        }
    )


@app.get('/download/<job>/<kind>')
def download(job, kind):
    jp = job_path(job)

    mapping = {
        'csv': jp / 'RTOProImport.csv',
        'review': jp / 'Review_Report.csv',
        'zip': jp / 'RTO_Import_Package.zip',
    }

    p = mapping.get(kind)

    if not p or not p.exists():
        abort(404)

    return send_file(
        p,
        as_attachment=True,
    )


@app.get('/pdf/<job>/<int:index>')
def pdf(job, index):
    jp, rows, _, _ = load_job(job)

    if index < 0 or index >= len(rows):
        abort(404)

    p = (
        jp
        / 'Combined_Files'
        / f"{safe_filename(rows[index].get('_pdf_name', 'contract'))}.pdf"
    )

    if not p.exists():
        abort(404)

    return send_file(
        p,
        as_attachment=request.args.get('download') == '1',
        download_name=p.name,
    )


@app.post('/refresh-rto/<job>')
def refresh_rto(job):
    jp, rows, _, _ = load_job(job)

    try:
        result = refresh_job_from_rto(
            jp,
            rows,
        )

        flash(
            f"RTO data refreshed: "
            f"{result['contracts_read']} contracts read; "
            f"{result['accounts_matched']} account number(s) matched. "
            "Last-used model numbers were refreshed from Firebird.",
            'ok',
        )

    except Exception as e:
        flash(
            'Could not refresh RTO Pro / Firebird data: '
            + str(e)
            + '. If this app was restarted, re-enter the CSV/database '
              'login or set RTO_DB_PASSWORD in .env.',
            'error',
        )

    return redirect(
        url_for(
            'review',
            job=job,
        )
    )


@app.post('/import-pdfs-to-rto/<job>')
def import_pdfs_to_rto(job):
    jp, rows, _, _ = load_job(job)

    try:
        # Refresh first so the account numbers and model tracker reflect the
        # import RTO Pro just completed.
        result = refresh_job_from_rto(
            jp,
            rows,
        )

    except Exception as e:
        flash(
            'PDF import stopped because RTO data could not be refreshed: '
            + str(e),
            'error',
        )

        return redirect(
            url_for(
                'review',
                job=job,
            )
        )

    copied = 0
    errors = []
    report = []

    for i, r in enumerate(rows):
        account = str(
            r.get(
                'ACCOUNT',
                '',
            )
            or ''
        ).strip()

        pdf_path = (
            jp
            / 'Combined_Files'
            / f"{safe_filename(r.get('_pdf_name', 'contract'))}.pdf"
        )

        entry = {
            'NAME': r.get('NAME', ''),
            'CONTRACT': r.get('CONTRACT', ''),
            'ACCOUNT': account,
            'SOURCE_PDF': str(pdf_path),
            'TARGET_PDF': '',
            'STATUS': '',
        }

        if not pdf_path.exists():
            msg = (
                f"{r.get('NAME', i + 1)}: "
                'combined PDF not found'
            )

            errors.append(msg)
            entry['STATUS'] = msg
            report.append(entry)
            continue

        if not account:
            msg = (
                f"{r.get('NAME', i + 1)}: "
                'RTO account number not found after refresh'
            )

            errors.append(msg)
            entry['STATUS'] = msg
            report.append(entry)
            continue

        try:
            target = copy_pdf_to_rto_imaging(
                pdf_path,
                account,
            )

            r['_pdf_imaging_path'] = str(
                target
            )

            entry['TARGET_PDF'] = str(
                target
            )

            entry['STATUS'] = 'Copied'
            copied += 1

        except Exception as e:
            msg = (
                f"{r.get('NAME', i + 1)}: "
                + str(e)
            )

            errors.append(msg)
            entry['STATUS'] = msg

        report.append(entry)

    report_path = (
        jp
        / 'PDF_Import_Report.csv'
    )

    with report_path.open(
        'w',
        encoding='utf-8-sig',
        newline='',
    ) as f:

        fields = [
            'NAME',
            'CONTRACT',
            'ACCOUNT',
            'SOURCE_PDF',
            'TARGET_PDF',
            'STATUS',
        ]

        w = csv.DictWriter(
            f,
            fieldnames=fields,
        )

        w.writeheader()
        w.writerows(report)

    save_job_rows(
        jp,
        rows,
    )

    if copied:
        flash(
            f'{copied} combined PDF(s) imported to '
            f'RTO Pro document imaging at {rto_images_root()}. '
            f"The database refresh matched "
            f"{result['accounts_matched']} account number(s).",
            'ok',
        )

    if errors:
        flash(
            'PDF imaging warnings: '
            + ' | '.join(
                errors[:8]
            ),
            'error',
        )

    return redirect(
        url_for(
            'review',
            job=job,
        )
    )


@app.post('/run-rto/<job>')
def run_rto(job):
    jp, rows, _, _ = load_job(job)

    errors = validate_rto_runtime(rows)
    if errors:
        flash(
            'RTO import blocked: ' + ' | '.join(errors[:8]),
            'error',
        )
        return redirect(url_for('review', job=job))

    exe = Path(os.getenv('RTO_EXE', r'C:\RTOwin\RTO-win.exe'))

    try:
        # V5.3 capability: always prepare the writable app.csv first so the
        # file is still available even when launching RTO Pro is disabled.
        staging = stage_rto_import(rows, jp, exe)
        build_package(jp)
    except Exception as e:
        flash('Could not prepare the RTO import: ' + str(e), 'error')
        return redirect(url_for('review', job=job))

    if os.getenv('ENABLE_RTO_LAUNCH', 'true').lower() != 'true':
        flash(
            f'RTO file is ready at {staging}, but launch is disabled in .env.',
            'error',
        )
        return redirect(url_for('review', job=job))

    if os.name != 'nt' or not exe.exists():
        flash(
            f'RTO Pro executable not found: {exe}. '
            f'The import file is ready at {staging}.',
            'error',
        )
        return redirect(url_for('review', job=job))

    try:
        # Exact V5.3 command shape. STORE is already a CSV field.
        command_parts = [str(exe), '-importcust', str(staging)]
        command_line = f'"{exe}" -importcust "{staging}"'

        (jp / 'Run_RTO_Import.cmd').write_text(
            '@echo off\r\n'
            + f'cd /d "{exe.parent}"\r\n'
            + command_line + '\r\n',
            encoding='cp1252',
        )
        (jp / 'RTO_Launch_Command.txt').write_text(
            command_line + '\n', encoding='utf-8'
        )

        subprocess.Popen(
            command_parts,
            cwd=str(exe.parent),
            shell=False,
        )
        build_package(jp)

        flash(
            f'RTO Pro import launched with app.csv at {staging}. '
            'After RTO Pro confirms the import, click Refresh RTO Data, '
            'then Import PDFs to RTO Imaging.',
            'ok',
        )
    except Exception as e:
        flash('RTO Pro could not be launched: ' + str(e), 'error')

    return redirect(url_for('review', job=job))


if __name__ == '__main__':
    # Local browser application only.
    app.run(
        host='127.0.0.1',
        port=5050,
        debug=False,
    )
