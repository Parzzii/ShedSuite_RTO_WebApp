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
import threading
import time
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
    phone,
    build_next_model_suggestions,
    MODEL_SERIES,
)
from pdf_tools import download_and_combine, extract_pdf_fields, safe_filename
from delivery_certificate import append_delivery_certificates
from pdf_contract_import import parse_contract_pdf


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

# V7 background Delivery Certificate workers. Each job gets its own lock so
# browser progress can be persisted without overwriting edits made in the UI.
JOB_LOCKS: dict[str, threading.RLock] = {}
CERT_ACTIVE: set[str] = set()
CERT_ACTIVE_LOCK = threading.Lock()

def _job_lock(job: str) -> threading.RLock:
    if job not in JOB_LOCKS:
        JOB_LOCKS[job] = threading.RLock()
    return JOB_LOCKS[job]

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

# Used-building contract rules. RTO Pro does not allow a contract number to be
# reused, while a used building must keep its original MODEL1.  V6.3.2 therefore
# keeps MODEL1 untouched. ShedSuite uses MODEL1 + suffix; PDF contracts use
# Agreement # + suffix because Agreement # is their contract number.
USED_CONTRACT_SUFFIXES = ['U'] + [chr(c) for c in range(ord('A'), ord('Z') + 1) if chr(c) != 'U']


# V7.5 inventory-only receiving. RTO Pro's inventory receiver is a different
# import path from -importcust, and uses these field names with -r.
INVENTORY_HEADERS = [
    'model', 'serial', 'stock', 'daterec', 'cost', 'balance', 'rbvbook',
    'description', 'invoice', 'brand', 'category', 'vendor', 'agent', 'bor',
    'rentalretail', 'quantity', 'store', 'retail', 'rto',
]

# Company/model-series choices reuse the exact numbering profiles already used
# by the contract workflow. Brand/vendor values are the canonical V6.3.3 names.
INVENTORY_PROFILES = {
    'MAG_GA': {'label': 'SMART SHED — GA', 'brand': 'SMART SHED', 'store': '8', 'tracker': 'magnolia_smart_shed_ga'},
    'MAG_SC': {'label': 'SMART SHED — SC', 'brand': 'SMART SHED', 'store': '8', 'tracker': 'magnolia_smart_shed_sc'},
    'MAG_NC': {'label': 'SMART SHED — NC', 'brand': 'SMART SHED', 'store': '8', 'tracker': 'magnolia_smart_shed_nc'},
    'MAG_TN': {'label': 'SMART SHED — TN', 'brand': 'SMART SHED', 'store': '8', 'tracker': 'magnolia_smart_shed_tn'},
    'CRF_ALPINE': {'label': 'ALPINE', 'brand': 'ALPINE', 'store': '1', 'tracker': 'crf_alpine'},
    'CRF_GENESIS': {'label': 'GENESIS', 'brand': 'GENESIS', 'store': '1', 'tracker': 'crf_genesis'},
    'CRF_PHOENIX': {'label': 'PHOENIX — Carefree', 'brand': 'PHOENIX', 'store': '1', 'tracker': 'crf_phoenix'},
    'CRF_4_SEASONS': {'label': '4 SEASONS', 'brand': '4 SEASONS', 'store': '1', 'tracker': 'crf_4_seasons'},
    'DBM_PHOENIX': {'label': 'PHOENIX — Dutch Boy', 'brand': 'PHOENIX', 'store': '2', 'tracker': 'dbm_phoenix'},
    'WWP': {'label': 'WESTWOOD SHEDS', 'brand': 'WESTWOOD SHEDS', 'store': '11', 'tracker': 'wwp'},
    'RAS_YSS': {'label': 'YODER STORAGE', 'brand': 'YODER STORAGE', 'store': '6', 'tracker': 'ras_yss'},
    'LSR_TX': {'label': 'LONESTAR SHEDS', 'brand': 'LONESTAR SHEDS', 'store': '7', 'tracker': 'lsr_tx'},
    # No model-number sequence for TRUE BUILT exists in the original workbook,
    # so V7.5 will not invent one. The model stays manually editable.
    'TRUE_BUILT': {'label': 'TRUE BUILT — manual model', 'brand': 'TRUE BUILT', 'store': '', 'tracker': ''},
    'CUSTOM': {'label': 'OTHER / CUSTOM', 'brand': '', 'store': '', 'tracker': ''},
}


def _contract_key(value) -> str:
    return str(value or '').strip().upper()


def used_model_history(model: str, existing_contracts) -> list[str]:
    """Return historical RTO contracts that identify MODEL1 as previously used."""
    base = _contract_key(model)
    if not base:
        return []
    matches = []
    for item in existing_contracts or []:
        if isinstance(item, dict):
            contract = item.get('contract', '')
        elif isinstance(item, (tuple, list)):
            contract = item[0] if item else ''
        else:
            contract = item
        c = _contract_key(contract)
        if not c:
            continue
        if c == base:
            matches.append(str(contract).strip())
            continue
        if c.startswith(base):
            suffix = c[len(base):]
            if suffix in USED_CONTRACT_SUFFIXES:
                matches.append(str(contract).strip())
    return list(dict.fromkeys(matches))


def next_used_contract(model: str, unavailable_contracts) -> tuple[str, str]:
    """Choose U first, then A/B/C... while respecting RTO's 10-char limit."""
    base = str(model or '').strip()
    if not base:
        return '', ''
    unavailable = {_contract_key(x) for x in (unavailable_contracts or []) if _contract_key(x)}
    for suffix in USED_CONTRACT_SUFFIXES:
        candidate = f'{base}{suffix}'
        if len(candidate) <= 10 and _contract_key(candidate) not in unavailable:
            return candidate, suffix
    return '', ''


def apply_used_building_defaults(rows: list[dict], source_rows: list[dict], ref: ReferenceData) -> None:
    """Auto-detect prior-use buildings and reserve a unique suffixed contract."""
    source_by_order = {
        str(x.get('Customer Order Id', '') or '').strip(): x
        for x in source_rows
    }
    existing = [
        str(x.get('contract', '') or '').strip()
        for x in (ref.existing_contracts or [])
        if str(x.get('contract', '') or '').strip()
    ]
    unavailable = list(existing)

    for row in rows:
        model = str(row.get('MODEL1', '') or '').strip()
        order_id = str(row.get('_source_order_id', '') or '').strip()
        src = source_by_order.get(order_id, {})
        source_type = str(src.get('_import_source_type', row.get('_source_type', 'csv')) or 'csv').lower()
        # ShedSuite contracts intentionally use MODEL1 as their base number.
        # RentaBarn PDF contracts use Agreement # as CONTRACT, so a used PDF
        # building (if one needs a suffix) must suffix the agreement—not MODEL1.
        base_contract = order_id if source_type == 'pdf' and order_id else model
        condition = str(src.get('Condition', '') or row.get('_source_condition', '') or '').strip()
        condition_low = condition.casefold()
        history = used_model_history(model, existing)

        condition_detected = any(
            token in condition_low
            for token in ('used', 'pre-owned', 'preowned', 'repo', 'repossessed')
        )
        detected = bool(history or condition_detected)

        reasons = []
        if history:
            reasons.append('RTO Pro history: ' + ', '.join(history[:4]))
        if condition_detected:
            source_label = 'PDF condition' if str(src.get('_import_source_type', '') or '').lower() == 'pdf' else 'ShedSuite condition'
            reasons.append(source_label + ': ' + condition)

        row['_used_detected'] = detected
        row['_used_building'] = detected
        row['_used_detection_reason'] = ' · '.join(reasons)
        row['_used_base_contract'] = base_contract
        row['_used_contract_suffix'] = ''
        row['_used_suffix_manual'] = False
        row['_used_contract_error'] = ''

        if detected:
            contract, suffix = next_used_contract(base_contract, unavailable)
            if contract:
                row['CONTRACT'] = contract
                row['_used_contract_suffix'] = suffix
                unavailable.append(contract)
                row.setdefault('_warnings', []).append(
                    f'Used building detected; MODEL1 kept as {model} and CONTRACT set to {contract}'
                )
            else:
                row['_used_contract_error'] = (
                    'No available used-building suffix fits the 10-character RTO contract limit.'
                )
                row.setdefault('_warnings', []).append(row['_used_contract_error'])
        else:
            # Reserve the normal contract so another used row in the same upload
            # cannot accidentally choose it.
            current = str(row.get('CONTRACT', '') or '').strip()
            if current:
                unavailable.append(current)


def add_contract_history_to_refs(refs: dict, existing_contracts) -> dict:
    """Expose only contract identifiers needed by the browser-side suffix picker."""
    refs = dict(refs or {})
    vals = []
    for item in existing_contracts or []:
        if isinstance(item, dict):
            value = item.get('contract', '')
        elif isinstance(item, (tuple, list)):
            value = item[0] if item else ''
        else:
            value = item
        value = str(value or '').strip()
        if value:
            vals.append(value)
    refs['existing_contracts'] = list(dict.fromkeys(vals))
    return refs


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



def _format_series_model(profile: str, number: int) -> str:
    rule = MODEL_SERIES.get(profile)
    if not rule:
        return ''
    prefix, _lo, _hi, pad3 = rule
    return prefix + (f'{number:03d}' if pad3 else str(number))


def inventory_profile_payload(ref: ReferenceData) -> list[dict]:
    """Build V7.5 company choices with live inventory-based model suggestions."""
    tracker = load_model_tracker()
    live_next = build_next_model_suggestions(ref) if ref.connected else {}
    models = [str(x.get('model', '') or '').strip() for x in (ref.existing_serials or [])]
    agents_by_store = {
        str(x.get('store', '') or '').strip(): str(x.get('agent', '') or '').strip()
        for x in (ref.agents or [])
        if str(x.get('store', '') or '').strip()
    }
    stores = ref.centralserver or KNOWN_STORES
    store_titles = {
        str(x.get('store', '') or '').strip(): str(x.get('storetitle', '') or x.get('storename', '') or '').strip()
        for x in stores
    }

    payload = []
    for profile, meta in INVENTORY_PROFILES.items():
        latest = ''
        suggested = ''
        if profile in MODEL_SERIES:
            prefix, lo, hi, _pad3 = MODEL_SERIES[profile]
            vals = []
            for model in models:
                if not model.startswith(prefix):
                    continue
                tail = model[len(prefix):]
                if tail.isdigit() and lo <= int(tail) <= hi:
                    vals.append(int(tail))
            if vals:
                latest = _format_series_model(profile, max(vals))
            if ref.connected:
                suggested = str(live_next.get(profile, '') or '').strip()

        tracker_key = meta.get('tracker', '')
        tracked = tracker.get(tracker_key, {}) if tracker_key else {}
        if not latest:
            latest = str(tracked.get('last_used', '') or '').strip()
        if not suggested and latest:
            tracker_prefix = str(tracked.get('prefix', '') or '').strip()
            suggested = next_model(latest, tracker_prefix)
            if suggested.endswith('-?') or suggested == '?':
                suggested = ''

        store = str(meta.get('store', '') or '').strip()
        payload.append({
            'profile': profile,
            'label': meta.get('label', profile),
            'brand': meta.get('brand', ''),
            'vendor': meta.get('brand', ''),
            'store': store,
            'store_title': store_titles.get(store, ''),
            'agent': agents_by_store.get(store, ''),
            'latest_model': latest,
            'suggested_model': suggested,
        })
    return payload


def _inventory_refs_path(jp: Path) -> Path:
    return jp / 'inventory_refs.json'


def _inventory_rows_path(jp: Path) -> Path:
    return jp / 'inventory_rows.json'


def load_inventory_job(job: str) -> tuple[Path, list[dict], dict]:
    jp = job_path(job)
    if not jp.exists():
        abort(404)
    try:
        rows = json.loads(_inventory_rows_path(jp).read_text(encoding='utf-8'))
    except Exception:
        rows = []
    try:
        refs = json.loads(_inventory_refs_path(jp).read_text(encoding='utf-8'))
    except Exception:
        refs = {'profiles': [], 'stores': KNOWN_STORES, 'agents': [], 'connected': False, 'error': ''}
    return jp, rows if isinstance(rows, list) else [], refs if isinstance(refs, dict) else {}


def _inventory_number(value) -> str:
    text = str(value or '').strip().replace('$', '').replace(',', '')
    if not text:
        return ''
    try:
        number = float(text)
        return str(int(number)) if number.is_integer() else ('%.2f' % number).rstrip('0').rstrip('.')
    except Exception:
        return text


def normalize_inventory_rows(rows: list[dict]) -> list[dict]:
    out = []
    for raw in rows or []:
        if not isinstance(raw, dict):
            continue
        row = {h: str(raw.get(h, '') or '').strip() for h in INVENTORY_HEADERS}
        row['company_profile'] = str(raw.get('company_profile', '') or '').strip()
        row['model'] = row['model'][:15]
        row['serial'] = row['serial'][:30]
        row['stock'] = row['stock'][:15]
        row['description'] = row['description'][:250]
        row['invoice'] = row['invoice'][:15]
        row['brand'] = row['brand'].upper()[:15]
        row['vendor'] = row['vendor'].upper()[:15]
        row['category'] = (row['category'] or 'NONE')[:31]
        row['agent'] = row['agent'][:30]
        row['store'] = row['store'][:10]
        row['daterec'] = row['daterec'] or datetime.now().strftime('%m/%d/%Y')
        row['quantity'] = row['quantity'] or '1'
        row['rentalretail'] = (row['rentalretail'] or 'RENTAL').upper()
        row['bor'] = (row['bor'] or 'NO').upper()
        row['cost'] = _inventory_number(row['cost'])
        row['retail'] = _inventory_number(row['retail'])
        row['rto'] = _inventory_number(row['rto'])
        # Let RTO Pro default BALANCE/RBVBOOK to cost when blank, matching its
        # documented receiver behavior.
        row['balance'] = _inventory_number(row['balance'])
        row['rbvbook'] = _inventory_number(row['rbvbook'])
        out.append(row)
    return out


def validate_inventory_rows(rows: list[dict], refs: dict | None = None) -> list[str]:
    errors = []
    seen_pairs = set()
    existing_pairs = set((refs or {}).get('existing_pairs', []) or [])
    for i, r in enumerate(rows, 1):
        model = str(r.get('model', '') or '').strip()
        serial = str(r.get('serial', '') or '').strip()
        store = str(r.get('store', '') or '').strip()
        label = f'Inventory {i}'
        if not str(r.get('company_profile', '') or '').strip():
            errors.append(f'{label}: choose a company.')
        if not model:
            errors.append(f'{label}: Model is required.')
        if not serial:
            errors.append(f'{label}: Serial is required.')
        if not store:
            errors.append(f'{label}: Store is required.')
        if len(model) > 15:
            errors.append(f'{label}: Model exceeds 15 characters.')
        if len(serial) > 30:
            errors.append(f'{label}: Serial exceeds 30 characters.')
        if r.get('stock') and not str(r.get('stock')).isdigit():
            errors.append(f'{label}: Stock must be numeric when entered.')
        if model and serial:
            pair = f'{model.upper()}|{serial.upper()}'
            if pair in seen_pairs:
                errors.append(f'{label}: {model} / {serial} is duplicated in this batch.')
            seen_pairs.add(pair)
            if pair in existing_pairs:
                errors.append(f'{label}: {model} / {serial} already exists in RTO Pro inventory.')
    return list(dict.fromkeys(errors))


def write_inventory_csv(rows: list[dict], path: Path, encoding: str = 'utf-8-sig') -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', encoding=encoding, errors='replace', newline='') as f:
        w = csv.DictWriter(f, fieldnames=INVENTORY_HEADERS, extrasaction='ignore', lineterminator='\r\n')
        w.writeheader()
        for r in rows:
            w.writerow({h: r.get(h, '') for h in INVENTORY_HEADERS})


def save_inventory_job(jp: Path, rows: list[dict], refs: dict) -> list[dict]:
    clean_rows = normalize_inventory_rows(rows)
    errors = validate_inventory_rows(clean_rows, refs)
    if errors:
        raise ValueError(' | '.join(errors[:10]))
    _inventory_rows_path(jp).write_text(json.dumps(clean_rows, indent=2), encoding='utf-8')
    write_inventory_csv(clean_rows, jp / 'InventoryImport.csv')
    with ZipFile(jp / 'Inventory_Import_Package.zip', 'w', ZIP_DEFLATED) as z:
        z.write(jp / 'InventoryImport.csv', 'InventoryImport.csv')
        for name in ['InventoryImport_LAUNCHED.csv', 'Run_RTO_Inventory_Import.cmd', 'RTO_Inventory_Launch_Command.txt']:
            p = jp / name
            if p.exists():
                z.write(p, name)
    return clean_rows


def stage_inventory_import(rows: list[dict], jp: Path, exe: Path) -> Path:
    configured = os.getenv('RTO_INVENTORY_STAGING', '').strip()
    local_app = os.getenv('LOCALAPPDATA', '').strip()
    temp_root = os.getenv('TEMP', '').strip()
    candidates = []
    if configured:
        candidates.append(Path(configured))
    candidates.append(exe.parent / 'inventory_import.csv')
    if local_app:
        candidates.append(Path(local_app) / 'ShedSuiteRTO' / 'inventory_import.csv')
    if temp_root:
        candidates.append(Path(temp_root) / 'ShedSuiteRTO' / 'inventory_import.csv')
    candidates.append(jp / 'inventory_import_runtime.csv')

    failures = []
    seen = set()
    for staging in candidates:
        key = str(staging).lower()
        if key in seen:
            continue
        seen.add(key)
        try:
            write_inventory_csv(rows, staging, encoding='cp1252')
            with staging.open('r', encoding='cp1252', errors='replace', newline='') as f:
                headers = next(csv.reader(f), [])
            if headers != INVENTORY_HEADERS:
                raise RuntimeError('inventory header row does not match RTO Pro receiver fields')
            shutil.copy2(staging, jp / 'InventoryImport_LAUNCHED.csv')
            return staging
        except (PermissionError, OSError, RuntimeError) as exc:
            failures.append(f'{staging}: {exc}')
    raise RuntimeError('Could not create a writable inventory import CSV. Tried: ' + ' | '.join(failures))


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


def _prepare_phone_and_customer_comment(row: dict) -> None:
    """Keep SMS flags and the generated OTHER IS note in the intended places.

    RTO Pro documents CELLOPT2=3 as: PHONE5 is a cell number and is opted out
    of SMS.  The generated OTHER IS note is intentionally kept out of the
    import COMMENTS field because that field is shown as an application comment
    during pending-web-application review.
    """
    other_phone = phone(row.get('PHONE5', ''))
    row['PHONE5'] = other_phone
    row['CELLOPT2'] = '3' if other_phone else ''

    primary_phone = phone(row.get('CELL', ''))
    row['CELL'] = primary_phone
    if primary_phone and not str(row.get('CELLOPT', '') or '').strip():
        row['CELLOPT'] = '2'

    app_comment = str(row.get('COMMENTS', '') or '').strip()
    customer_comment = str(row.get('_customer_comment', '') or '').strip()
    if not customer_comment and re.match(r'^OTHER\s+IS\b', app_comment, flags=re.I):
        row['_customer_comment'] = app_comment
        customer_comment = app_comment
    if customer_comment and re.match(r'^OTHER\s+IS\b', app_comment, flags=re.I):
        row['COMMENTS'] = ''


def _normal_customer_comment(existing: str, generated: str) -> str:
    """Merge our generated OTHER IS note without destroying existing comments."""
    existing = re.sub(r'\s+', ' ', str(existing or '')).strip()
    generated = re.sub(r'\s+', ' ', str(generated or '')).strip()
    if not generated:
        return existing
    if generated.casefold() in existing.casefold():
        return existing

    # Replace only an older generated OTHER IS segment. Preserve unrelated notes.
    parts = [x.strip() for x in existing.split('|') if x.strip()]
    parts = [x for x in parts if not re.match(r'^OTHER\s+IS\b', x, flags=re.I)]
    parts.append(generated)
    return ' | '.join(parts)


def sync_customer_comments_to_rto(
    dsn: str, user: str, password: str, rows: list[dict]
) -> dict:
    """Write generated OTHER IS notes to the normal RTO customer comment field.

    RTO Pro's documented import field COMMENTS is used on the pending application
    screen.  Once the import has assigned an ACCOUNT, this routine updates the
    normal CUSTOMERS.COMMENTS field so the note appears with customer maintenance
    directly below Directions.
    """
    targets = [
        r for r in rows
        if str(r.get('_customer_comment', '') or '').strip()
        and str(r.get('ACCOUNT', '') or '').strip()
    ]
    if not targets:
        return {'synced': 0, 'skipped': 0, 'errors': []}

    conn = _open_rto_connection(dsn, user, password)
    synced = 0
    skipped = 0
    errors: list[str] = []
    try:
        cur = conn.cursor()
        cur.execute(
            "SELECT TRIM(RDB$FIELD_NAME) FROM RDB$RELATION_FIELDS "
            "WHERE RDB$RELATION_NAME='CUSTOMERS'"
        )
        columns = {str(x[0] or '').strip().upper() for x in cur.fetchall()}
        if 'COMMENTS' not in columns:
            raise RuntimeError('RTO CUSTOMERS table does not expose a COMMENTS field.')

        for row in targets:
            account = str(row.get('ACCOUNT', '') or '').strip()
            store = str(row.get('STORE', '') or '').strip()
            generated = str(row.get('_customer_comment', '') or '').strip()
            name = str(row.get('NAME', '') or account).strip()
            try:
                if store and 'STORE' in columns:
                    cur.execute(
                        'SELECT COMMENTS FROM CUSTOMERS WHERE ACCOUNT=? AND STORE=?',
                        account, store,
                    )
                else:
                    cur.execute('SELECT COMMENTS FROM CUSTOMERS WHERE ACCOUNT=?', account)
                rec = cur.fetchone()
                if rec is None:
                    skipped += 1
                    row['_customer_comment_sync_status'] = 'RTO customer not found yet'
                    continue

                existing = '' if rec[0] is None else str(rec[0]).strip()
                merged = _normal_customer_comment(existing, generated)
                if merged != existing:
                    if store and 'STORE' in columns:
                        cur.execute(
                            'UPDATE CUSTOMERS SET COMMENTS=? WHERE ACCOUNT=? AND STORE=?',
                            merged, account, store,
                        )
                    else:
                        cur.execute(
                            'UPDATE CUSTOMERS SET COMMENTS=? WHERE ACCOUNT=?',
                            merged, account,
                        )
                row['_customer_comment_sync_status'] = 'Synced to RTO customer comments'
                synced += 1
            except Exception as exc:
                row['_customer_comment_sync_status'] = 'Comment sync failed'
                errors.append(f'{name}: {exc}')
        conn.commit()
        return {'synced': synced, 'skipped': skipped, 'errors': errors}
    except Exception:
        try:
            conn.rollback()
        except Exception:
            pass
        raise
    finally:
        conn.close()


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
    contract_accounts = fetch_rto_contract_accounts(dsn, user, password)
    refs = add_contract_history_to_refs(ref.as_client_dict(), contract_accounts)
    (jp / 'refs.json').write_text(json.dumps(refs, indent=2), encoding='utf-8')

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

    # V7.9: the generated OTHER IS note belongs in the normal customer comment
    # field below Directions, not in the pending application's comments box.
    comment_sync = sync_customer_comments_to_rto(dsn, user, password, rows)

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
        'customer_comments_synced': int(comment_sync.get('synced', 0)),
        'customer_comments_skipped': int(comment_sync.get('skipped', 0)),
        'customer_comment_errors': list(comment_sync.get('errors', [])),
        'refs': refs,
    }


def ensure_used_contracts_available(jp: Path, rows: list[dict], refs: dict) -> list[str]:
    """Recheck V6.3.2 used suffixes immediately before launching RTO Pro.

    If Firebird is reachable, this uses the newest contract list. Otherwise it
    falls back to the contract list captured when the job was created/refreshed.
    """
    live_contracts = []
    try:
        dsn, user, password = _job_db_credentials(jp)
        live_contracts = [c for c, _a in fetch_rto_contract_accounts(dsn, user, password)]
    except Exception:
        live_contracts = list((refs or {}).get('existing_contracts') or [])

    unavailable = {_contract_key(x) for x in live_contracts if _contract_key(x)}
    changes = []

    # Normal rows reserve their exact contract numbers first.
    for row in rows:
        used = row.get('_used_building') is True or str(row.get('_used_building', '')).lower() in ('1', 'true', 'yes', 'on')
        if not used:
            c = _contract_key(row.get('CONTRACT', ''))
            if c:
                unavailable.add(c)

    for row in rows:
        used = row.get('_used_building') is True or str(row.get('_used_building', '')).lower() in ('1', 'true', 'yes', 'on')
        if not used:
            continue

        model = str(row.get('MODEL1', '') or '').strip()
        base_contract = str(row.get('_used_base_contract', '') or '').strip() or model
        current = str(row.get('CONTRACT', '') or '').strip()
        current_key = _contract_key(current)
        model_key = _contract_key(base_contract)
        suffix = current_key[len(model_key):] if model_key and current_key.startswith(model_key) else ''
        current_valid = (
            bool(current)
            and current_key not in unavailable
            and bool(suffix)
            and bool(re.fullmatch(r'[A-Z0-9]+', suffix))
            and len(current) <= 10
        )

        if current_valid:
            row['_used_contract_suffix'] = suffix
            row['_used_contract_error'] = ''
            unavailable.add(current_key)
            continue

        candidate, new_suffix = next_used_contract(base_contract, unavailable)
        if not candidate:
            row['_used_contract_error'] = 'No available used-building suffix fits the 10-character RTO contract limit.'
            continue

        if current != candidate:
            changes.append(f"{row.get('NAME', 'Contract')}: {current or '(blank)'} → {candidate}")
        row['CONTRACT'] = candidate
        row['_used_contract_suffix'] = new_suffix
        row['_used_suffix_manual'] = False
        row['_used_contract_error'] = ''
        unavailable.add(_contract_key(candidate))

    return changes


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
    status_class = 'good' if connected else 'bad'

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
                f'<a class="workflow-mini-button pdf-open" href="{open_url}" '
                'target="_blank" rel="noopener">Open PDF</a>'
                f'<a class="workflow-mini-button pdf-download" href="{download_url}">Download</a>'
            )
            detail = (pages + ' pages · ' if pages else '') + manifest_text
        else:
            buttons = '<span class="combined-pdf-missing">No combined PDF</span>'
            detail = 'Re-process the CSV with Download PDFs enabled.'

        pdf_rows.append(
            '<div class="combined-pdf-row">'
            f'<div class="combined-pdf-name"><b>{name}</b>'
            f'<span>Order {order_id}</span></div>'
            f'<div class="combined-pdf-actions">{buttons}</div>'
            f'<div class="combined-pdf-detail">{detail}</div>'
            '</div>'
        )

    return f"""
<section id="rto-workflow-panel" class="injected-workflow-panel">
  <div class="workflow-card">
    <div class="workflow-head">
      <div>
        <div class="workflow-title">RTO Pro Workflow</div>
        <div class="workflow-subtitle">
          Same sequence as Excel: CSV → RTO Pro → refresh DB → PDF imaging
        </div>
      </div>
      <span class="workflow-db-status {status_class}">
        {html_lib.escape(status_text)}
      </span>
    </div>

    <div class="workflow-actions">
      <form method="post"
            action="{url_for('run_rto', job=job)}"
            onsubmit="return confirm('Launch RTO Pro and import this CSV?')">
        <button class="workflow-button workflow-import" type="submit">
          1. Import to RTO Pro
        </button>
      </form>

      <form method="post"
            action="{url_for('refresh_rto', job=job)}">
        <button class="workflow-button workflow-refresh" type="submit">
          2. Refresh RTO Data + Sync Comment
        </button>
      </form>

      <form method="post"
            action="{url_for('import_pdfs_to_rto', job=job)}"
            onsubmit="return confirm('Only click this AFTER RTO Pro confirms the customer import. Continue?')">
        <button class="workflow-button workflow-imaging" type="submit">
          3. Import PDFs to RTO Imaging
        </button>
      </form>

      <a class="workflow-button workflow-download"
         href="{url_for('download', job=job, kind='csv')}">
        Download RTO CSV
      </a>
    </div>

    <div class="workflow-help">
      After RTO Pro shows its successful import confirmation, click
      <b>Refresh RTO Data</b>. That reads the updated Firebird tables,
      fills assigned account numbers, refreshes last-used model numbers, and syncs the
      <b>OTHER IS...</b> note into the normal customer comment box below Directions.
      Then click <b>Import PDFs to RTO Imaging</b>.
    </div>
  </div>
</section>

<section id="combined-pdfs-panel" class="injected-workflow-panel">
  <details class="combined-pdfs-card" open>
    <summary class="combined-pdfs-summary">
      <span>Combined PDFs</span>
      <strong>{available_count}/{len(rows)} ready</strong>
    </summary>
    <div class="combined-pdfs-list">
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


def read_csv_file_with_headers(path: Path) -> tuple[list[dict[str, str]], list[str]]:
    """Read one ShedSuite CSV and preserve its header list even when it has no rows."""
    with path.open('r', encoding='utf-8-sig', newline='') as f:
        reader = csv.DictReader(f)
        rows = list(reader)
        return rows, list(reader.fieldnames or [])


def _safe_input_csv_name(filename: str, index: int) -> str:
    """Create a stable local filename without trusting the browser-supplied path."""
    original = Path(str(filename or '')).name
    stem = safe_filename(Path(original).stem or f'input_{index}')
    return f'{index:02d}_{stem}.csv'


def _safe_input_pdf_name(filename: str, index: int) -> str:
    """Create a stable local PDF filename without trusting the browser path."""
    original = Path(str(filename or '')).name
    stem = safe_filename(Path(original).stem or f'contract_{index}')
    return f'{index:02d}_{stem}.pdf'


def merge_contract_inputs(files, jp: Path, required: set[str]) -> tuple[list[dict[str, str]], dict, list[str]]:
    """Merge ShedSuite CSVs and supported contract PDFs into one source-row list.

    CSV rows remain authoritative if the same Customer Order/Agreement number is
    supplied twice, because CSV rows contain the ShedSuite document links needed
    for Delivery Certificates. PDF rows are converted into the same source schema
    used by rto_transform.py and then enter the normal review/export pipeline.
    """
    csv_files = []
    pdf_files = []
    unsupported = []
    for uploaded in files:
        suffix = Path(str(uploaded.filename or '')).suffix.lower()
        if suffix == '.csv':
            csv_files.append(uploaded)
        elif suffix == '.pdf':
            pdf_files.append(uploaded)
        else:
            unsupported.append(Path(str(uploaded.filename or '')).name or 'unnamed file')
    if unsupported:
        raise ValueError('Unsupported file type: ' + ', '.join(unsupported) + '. Choose CSV or PDF files.')

    merged: list[dict[str, str]] = []
    warnings: list[str] = []
    summary = {
        'file_count': len(csv_files) + len(pdf_files),
        'csv_count': len(csv_files),
        'pdf_count': len(pdf_files),
        'row_count': 0,
        'duplicates_skipped': 0,
        'files': [],
    }

    if csv_files:
        csv_rows, csv_summary, csv_warnings = merge_shedsuite_csv_files(
            csv_files, jp / 'Input_CSVs', required
        )
        for row in csv_rows:
            row['_import_source_type'] = 'csv'
        merged.extend(csv_rows)
        warnings.extend(csv_warnings)
        summary['duplicates_skipped'] += int(csv_summary.get('duplicates_skipped', 0) or 0)
        for item in csv_summary.get('files', []):
            summary['files'].append(dict(item, type='csv'))

    seen = {
        str(r.get('Customer Order Id', '') or '').strip().casefold():
        str(r.get('_import_csv_name', '') or r.get('_import_pdf_name', '') or 'another input')
        for r in merged
        if str(r.get('Customer Order Id', '') or '').strip()
    }

    input_pdf_dir = jp / 'Input_PDFs'
    input_pdf_dir.mkdir(parents=True, exist_ok=True)
    for index, uploaded in enumerate(pdf_files, 1):
        original_name = Path(str(uploaded.filename or '')).name or f'contract_{index}.pdf'
        stored_name = _safe_input_pdf_name(original_name, index)
        saved_path = input_pdf_dir / stored_name
        uploaded.save(saved_path)
        try:
            row, metadata, pdf_warnings = parse_contract_pdf(saved_path)
        except Exception as exc:
            warnings.append(f'{original_name}: PDF skipped ({exc})')
            summary['files'].append({
                'name': original_name, 'stored_name': stored_name, 'type': 'pdf',
                'rows': 0, 'accepted': 0, 'duplicates_skipped': 0,
                'provider': 'Unrecognized PDF', 'error': str(exc),
            })
            continue

        row['_import_source_type'] = 'pdf'
        row['_import_pdf_name'] = original_name
        row['_import_pdf_stored_name'] = stored_name
        order_id = str(row.get('Customer Order Id', '') or '').strip()
        order_key = order_id.casefold()
        if order_key and order_key in seen:
            summary['duplicates_skipped'] += 1
            warnings.append(
                f'Duplicate contract/agreement {order_id} in {original_name} was skipped; '
                f'already loaded from {seen[order_key]}.'
            )
            continue
        if order_key:
            seen[order_key] = original_name
        merged.append(row)
        warnings.extend(pdf_warnings)
        summary['files'].append({
            'name': original_name, 'stored_name': stored_name, 'type': 'pdf',
            'rows': 1, 'accepted': 1, 'duplicates_skipped': 0,
            'provider': metadata.get('provider', 'Contract PDF'),
        })

    summary['row_count'] = len(merged)
    if not merged and pdf_files:
        pdf_errors = [x for x in warnings if 'PDF skipped' in x]
        if pdf_errors:
            raise ValueError(pdf_errors[0])
    return merged, summary, warnings


def merge_shedsuite_csv_files(files, input_dir: Path, required: set[str]) -> tuple[list[dict[str, str]], dict, list[str]]:
    """Save, validate and merge multiple ShedSuite contract-report CSV files.

    Duplicate ShedSuite Customer Order IDs are kept only once. This makes
    overlapping reports safe and also prevents selecting the same CSV twice
    from producing duplicate RTO contracts. The first occurrence wins.
    """
    input_dir.mkdir(parents=True, exist_ok=True)
    merged: list[dict[str, str]] = []
    seen_order_ids: dict[str, str] = {}
    duplicate_messages: list[str] = []
    file_summaries: list[dict] = []

    for index, uploaded in enumerate(files, 1):
        original_name = Path(str(uploaded.filename or '')).name or f'input_{index}.csv'
        stored_name = _safe_input_csv_name(original_name, index)
        saved_path = input_dir / stored_name
        uploaded.save(saved_path)

        try:
            rows, headers = read_csv_file_with_headers(saved_path)
        except Exception as exc:
            raise ValueError(f'{original_name}: could not read CSV ({exc})') from exc

        missing = sorted(required - set(headers))
        if missing:
            raise ValueError(
                f'{original_name}: this does not match the ShedSuite RTO contracts report format. '
                f'Missing column(s): {", ".join(missing)}'
            )

        accepted = 0
        duplicates = 0
        for raw in rows:
            row = dict(raw or {})
            row['_import_csv_name'] = original_name
            order_id = str(row.get('Customer Order Id', '') or '').strip()
            order_key = order_id.casefold()

            if order_key and order_key in seen_order_ids:
                duplicates += 1
                duplicate_messages.append(
                    f'Duplicate ShedSuite order {order_id} in {original_name} was skipped; '
                    f'already loaded from {seen_order_ids[order_key]}.'
                )
                continue

            if order_key:
                seen_order_ids[order_key] = original_name
            merged.append(row)
            accepted += 1

        file_summaries.append({
            'name': original_name,
            'stored_name': stored_name,
            'rows': len(rows),
            'accepted': accepted,
            'duplicates_skipped': duplicates,
        })

    summary = {
        'file_count': len(file_summaries),
        'row_count': len(merged),
        'duplicates_skipped': sum(x['duplicates_skipped'] for x in file_summaries),
        'files': file_summaries,
    }
    return merged, summary, duplicate_messages


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
    """Catch the RTO Pro field limits and V6.3.1 used-building rules."""
    errors = []
    seen_contracts: set[str] = set()
    for i, r in enumerate(rows, 1):
        name = str(r.get('NAME') or f'Row {i}').strip()
        contract = str(r.get('CONTRACT') or '').strip()
        model = str(r.get('MODEL1') or '').strip()
        if not str(r.get('NAME') or '').strip():
            errors.append(f'Row {i}: NAME is blank.')
        if not str(r.get('STORE') or '').strip():
            errors.append(f'Row {i}: STORE is blank.')
        if len(contract) > 10:
            errors.append(f'Row {i}: CONTRACT exceeds 10 characters.')
        if len(model) > 15:
            errors.append(f'Row {i}: MODEL1 exceeds 15 characters.')
        if contract:
            ck = _contract_key(contract)
            if ck in seen_contracts:
                errors.append(f'{name}: CONTRACT {contract} is duplicated in this import.')
            seen_contracts.add(ck)

        used = r.get('_used_building') is True or str(r.get('_used_building', '')).lower() in ('1', 'true', 'yes', 'on')
        if used:
            if not model:
                errors.append(f'{name}: used building needs a MODEL1.')
            if not contract:
                errors.append(f'{name}: used building needs a suffixed CONTRACT.')
            elif _contract_key(contract) == _contract_key(model):
                errors.append(f'{name}: used building CONTRACT must have a suffix; MODEL1 must stay unchanged.')
            elif not _contract_key(contract).startswith(_contract_key(model)):
                errors.append(f'{name}: used building CONTRACT must be MODEL1 plus a suffix.')
            else:
                suffix = _contract_key(contract)[len(_contract_key(model)):]
                if not suffix or not re.fullmatch(r'[A-Z0-9]+', suffix):
                    errors.append(f'{name}: used building suffix {suffix or "(blank)"} must contain only letters and numbers.')
            if str(r.get('_used_contract_error') or '').strip():
                errors.append(f'{name}: {r.get("_used_contract_error")}')
    return list(dict.fromkeys(errors))


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
        '_source_type',
        '_source_file_name',
        '_source_csv_name',
        '_source_pdf_name',
        'SERIAL1',
        '_source_company',
        '_source_dealer',
        'STORE',
        '_store_title',
        'DEALERID',
        '_dealer_name',
        'MODEL1',
        'CONTRACT',
        '_used_building',
        '_used_detected',
        '_used_contract_suffix',
        '_used_suffix_manual',
        '_used_detection_reason',
        '_used_contract_error',
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


def _ensure_row_uids(rows: list[dict]) -> bool:
    """Give every contract a stable UI id so sorting never depends on row index."""
    changed = False
    seen: set[str] = set()
    for r in rows:
        uid = str(r.get('_row_uid', '') or '').strip()
        if not uid or uid in seen:
            uid = uuid.uuid4().hex
            r['_row_uid'] = uid
            changed = True
        seen.add(uid)
    return changed


def save_job_rows(jp: Path, rows: list[dict]):
    for row in rows:
        _prepare_phone_and_customer_comment(row)
    _ensure_row_uids(rows)
    # V7 serializes writes with the certificate worker so edits and background
    # PDF completion cannot overwrite one another.
    with _job_lock(jp.name):
        (jp / 'rows.json').write_text(
            json.dumps(rows, indent=2),
            encoding='utf-8',
        )
        write_import_csv(rows, jp / 'RTOProImport.csv')
        write_review_csv(rows, jp / 'Review_Report.csv')
        write_xmls(rows, jp / 'XML_Files', jp / 'Combined_Files')
        build_package(jp)


def _certificate_status_path(jp: Path) -> Path:
    return jp / 'delivery_certificate_status.json'


def _read_certificate_status(jp: Path) -> dict:
    p = _certificate_status_path(jp)
    if not p.exists():
        return {'active': False, 'complete': True, 'items': {}, 'updated_at': ''}
    try:
        data = json.loads(p.read_text(encoding='utf-8'))
        return data if isinstance(data, dict) else {'active': False, 'complete': True, 'items': {}}
    except Exception:
        return {'active': False, 'complete': False, 'items': {}, 'error': 'Could not read certificate status.'}


def _write_certificate_status(jp: Path, data: dict) -> None:
    data = dict(data or {})
    data['updated_at'] = datetime.now().isoformat(timespec='seconds')
    temp = _certificate_status_path(jp).with_suffix('.tmp')
    temp.write_text(json.dumps(data, indent=2), encoding='utf-8')
    temp.replace(_certificate_status_path(jp))


def _initialize_certificate_status(jp: Path, rows: list[dict], active: bool) -> None:
    items = {}
    queued_any = False
    for r in rows:
        oid = str(r.get('_source_order_id', '') or '').strip()
        if not oid:
            continue
        is_pdf_source = str(r.get('_source_type', '') or '').lower() == 'pdf'
        should_queue = bool(active and not is_pdf_source)
        queued_any = queued_any or should_queue
        if is_pdf_source:
            detail = 'Uploaded contract PDF; no ShedSuite Delivery Certificate lookup needed.'
        elif should_queue:
            detail = 'Waiting for background Chromium…'
        else:
            detail = 'Delivery Certificate was not requested.'
        items[oid] = {
            'status': 'queued' if should_queue else 'not_requested',
            'name': str(r.get('NAME', '') or oid),
            'detail': detail,
        }
    _write_certificate_status(jp, {
        'active': bool(active and queued_any),
        'complete': not bool(active and queued_any),
        # On Windows the worker uses a normal Chromium window kept minimized/
        # off-screen because ShedSuite is more reliable than in true headless mode.
        'headless': os.name != 'nt',
        'browser_mode': 'headless' if os.name != 'nt' else 'background',
        'items': items,
    })


def _extract_ready_certificate_fields(jp: Path, row: dict) -> dict[str, str]:
    """Re-extract PDF fields after the Delivery Certificate has been appended.

    V7.1 extracted coordinates before Chromium added the Delivery Certificate,
    which meant the review page could stay blank even though the final PDF had
    coordinates. V7.2 calls this immediately when each certificate is ready.
    """
    oid = str(row.get('_source_order_id', '') or '').strip()
    target = jp / 'Combined_Files' / f"{safe_filename(row.get('_pdf_name') or row.get('NAME') or oid)}.pdf"
    if not target.exists():
        return {}
    try:
        return extract_pdf_fields(target)
    except Exception as exc:
        print(f'Could not refresh PDF fields for {oid or row.get("NAME", "contract")}: {exc}')
        return {}


def _merge_certificate_results(job: str, worker_rows: list[dict], cert_warnings: list[str]) -> None:
    jp = job_path(job)
    with _job_lock(job):
        if not (jp / 'rows.json').exists():
            return
        current = json.loads((jp / 'rows.json').read_text(encoding='utf-8'))
        worker_by_order = {str(r.get('_source_order_id', '') or '').strip(): r for r in worker_rows}
        current_orders = set()
        for r in current:
            oid = str(r.get('_source_order_id', '') or '').strip()
            current_orders.add(oid)
            wr = worker_by_order.get(oid)
            if not wr:
                continue
            if wr.get('_delivery_certificate') in {'Downloaded', 'Missing'}:
                for key in ['_delivery_certificate', '_delivery_certificate_method', '_delivery_certificate_name']:
                    if key in wr:
                        r[key] = wr.get(key, '')
            if wr.get('_delivery_certificate') == 'Downloaded':
                manifest = list(r.get('_pdf_manifest') or [])
                if 'Delivery Certificate' not in manifest:
                    manifest.append('Delivery Certificate')
                r['_pdf_manifest'] = manifest
                target = jp / 'Combined_Files' / f"{safe_filename(r.get('_pdf_name') or r.get('NAME') or oid)}.pdf"
                r['_pdf_available'] = target.exists()

                # V7.2: Delivery Certificate often contains the coordinates.
                # Re-extract after it is appended and merge those server-owned
                # fields into the latest row without touching the user's other edits.
                fields = _extract_ready_certificate_fields(jp, wr)
                coord_map = {
                    '_pdf_coordinates': 'coordinates',
                    '_pdf_latitude': 'latitude',
                    '_pdf_longitude': 'longitude',
                }
                for meta_key, field_key in coord_map.items():
                    value = str(fields.get(field_key, '') or '').strip()
                    if value:
                        r[meta_key] = value
                        wr[meta_key] = value
                coords = str(fields.get('coordinates', '') or '').strip()
                if coords:
                    existing_direction = str(r.get('DIRECTIONS', '') or '')
                    if coords not in existing_direction:
                        r['DIRECTIONS'] = coords + (('      ' + existing_direction) if existing_direction else '')

                if target.exists():
                    try:
                        doc = fitz.open(target)
                        r['_pdf_pages'] = len(doc)
                        doc.close()
                    except Exception:
                        pass

        # A contract may have been discarded while Chromium was working. Remove
        # any packet the worker created for it so it cannot leak into Full ZIP.
        for oid, wr in worker_by_order.items():
            if oid in current_orders:
                continue
            target = jp / 'Combined_Files' / f"{safe_filename(wr.get('_pdf_name') or wr.get('NAME') or oid)}.pdf"
            try:
                if target.exists():
                    target.unlink()
            except Exception:
                pass

        if cert_warnings:
            wp = jp / 'warnings.json'
            try:
                warnings = json.loads(wp.read_text(encoding='utf-8')) if wp.exists() else []
            except Exception:
                warnings = []
            for w in cert_warnings:
                msg = 'PDF: ' + str(w)
                if msg not in warnings:
                    warnings.append(msg)
            wp.write_text(json.dumps(warnings, indent=2), encoding='utf-8')
        save_job_rows(jp, current)


def _run_certificate_worker(job: str, only_order_ids: set[str] | None = None) -> None:
    jp = job_path(job)
    try:
        with _job_lock(job):
            if not (jp / 'rows.json').exists():
                return
            rows = json.loads((jp / 'rows.json').read_text(encoding='utf-8'))
            source_rows = json.loads((jp / 'source_rows.json').read_text(encoding='utf-8')) if (jp / 'source_rows.json').exists() else []
            if only_order_ids:
                rows = [r for r in rows if str(r.get('_source_order_id', '') or '').strip() in only_order_ids]
            else:
                # Initial V7.6.2 run: queue only ShedSuite/CSV contracts. Imported
                # PDFs are already complete packets and must never enter the
                # ShedSuite browser lookup.
                rows = [
                    r for r in rows
                    if str(r.get('_source_type', 'csv') or 'csv').lower() != 'pdf'
                    and str(r.get('_delivery_certificate', '') or '').lower() in {'queued', ''}
                ]
            status = _read_certificate_status(jp)
            status['active'] = bool(rows)
            status['complete'] = not bool(rows)
            browser_headless = os.name != 'nt'
            status['headless'] = browser_headless
            status['browser_mode'] = 'headless' if browser_headless else 'background'
            status.setdefault('items', {})
            for r in rows:
                oid = str(r.get('_source_order_id', '') or '').strip()
                if oid:
                    status['items'][oid] = {
                        'status': 'queued',
                        'name': str(r.get('NAME', '') or oid),
                        'detail': 'Waiting for background Chromium…',
                    }
            _write_certificate_status(jp, status)

        if not rows:
            return

        def still_present(oid: str) -> bool:
            try:
                with _job_lock(job):
                    latest = json.loads((jp / 'rows.json').read_text(encoding='utf-8'))
                    st = _read_certificate_status(jp)
                    if str(st.get('items', {}).get(oid, {}).get('status', '')).lower() == 'skipped':
                        return False
                return any(str(r.get('_source_order_id', '') or '').strip() == oid for r in latest)
            except Exception:
                return False

        def progress(oid: str, state: str, detail: str, row: dict) -> None:
            # When the certificate is ready, it has already been appended to the
            # combined PDF by delivery_certificate.py. Extract coordinates now so
            # the browser can receive them while the remaining certificates run.
            fields = _extract_ready_certificate_fields(jp, row) if state == 'ready' else {}
            coords = str(fields.get('coordinates', '') or '').strip()
            latitude = str(fields.get('latitude', '') or '').strip()
            longitude = str(fields.get('longitude', '') or '').strip()
            if coords:
                row['_pdf_coordinates'] = coords
                row['_pdf_latitude'] = latitude
                row['_pdf_longitude'] = longitude

            with _job_lock(job):
                st = _read_certificate_status(jp)
                st.setdefault('items', {})
                existing_state = str(st['items'].get(oid, {}).get('status', '')).lower()
                # A discard/skip made in the UI wins over a late browser result.
                if existing_state in {'discarded', 'skipped'}:
                    return

                # Persist the fresh PDF metadata into the latest row immediately.
                # This makes refreshes safe even before the whole Chromium batch ends.
                if coords and (jp / 'rows.json').exists():
                    try:
                        latest_rows = json.loads((jp / 'rows.json').read_text(encoding='utf-8'))
                        latest = next((x for x in latest_rows if str(x.get('_source_order_id', '') or '').strip() == oid), None)
                        if latest is not None:
                            latest['_pdf_coordinates'] = coords
                            latest['_pdf_latitude'] = latitude
                            latest['_pdf_longitude'] = longitude
                            existing_direction = str(latest.get('DIRECTIONS', '') or '')
                            if coords not in existing_direction:
                                latest['DIRECTIONS'] = coords + (('      ' + existing_direction) if existing_direction else '')
                            temp_rows = (jp / 'rows.json').with_suffix('.tmp')
                            temp_rows.write_text(json.dumps(latest_rows, indent=2), encoding='utf-8')
                            temp_rows.replace(jp / 'rows.json')
                    except Exception as exc:
                        print(f'Could not persist live coordinates for {oid}: {exc}')

                st['active'] = True
                st['complete'] = False
                status_detail = str(detail or '')
                if coords and state == 'ready':
                    status_detail = (status_detail + ' Coordinates extracted automatically.').strip()
                st['items'][oid] = {
                    'status': state,
                    'name': str(row.get('NAME', '') or oid),
                    'detail': status_detail,
                    'coordinates': coords,
                    'latitude': latitude,
                    'longitude': longitude,
                }
                _write_certificate_status(jp, st)

        # On Windows run Chromium as a normal browser kept minimized/off-screen.
        # ShedSuite has proven more reliable this way than in true headless mode,
        # and the Flask review page remains fully editable while it runs.
        browser_headless = os.name != 'nt'
        warnings = append_delivery_certificates(
            rows, source_rows, jp / 'Combined_Files', BASE,
            progress=progress, should_process=still_present, headless=browser_headless
        )
        _merge_certificate_results(job, rows, warnings)

        with _job_lock(job):
            st = _read_certificate_status(jp)
            st['active'] = False
            st['complete'] = True
            st['finished_at'] = datetime.now().isoformat(timespec='seconds')
            _write_certificate_status(jp, st)
    except Exception as exc:
        try:
            with _job_lock(job):
                st = _read_certificate_status(jp)
                st['active'] = False
                st['complete'] = True
                st['error'] = str(exc)
                _write_certificate_status(jp, st)
        except Exception:
            pass
    finally:
        with CERT_ACTIVE_LOCK:
            CERT_ACTIVE.discard(job)


def _start_certificate_worker(job: str, only_order_ids: set[str] | None = None) -> bool:
    with CERT_ACTIVE_LOCK:
        if job in CERT_ACTIVE:
            return False
        CERT_ACTIVE.add(job)
    thread = threading.Thread(
        target=_run_certificate_worker,
        args=(job, only_order_ids),
        daemon=True,
        name=f'delivery-cert-{job}',
    )
    thread.start()
    return True


def load_job(job: str):
    jp = job_path(job)

    if not (jp / 'rows.json').exists():
        abort(404)

    with _job_lock(job):
        rows = json.loads(
            (jp / 'rows.json').read_text(encoding='utf-8')
        )
        # V7.7 adds stable row ids for collapse/reorder actions. Older jobs can
        # still be opened safely; only rows.json needs the metadata update.
        if _ensure_row_uids(rows):
            temp_rows = (jp / 'rows.json').with_suffix('.tmp')
            temp_rows.write_text(json.dumps(rows, indent=2), encoding='utf-8')
            temp_rows.replace(jp / 'rows.json')

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



@app.post('/inventory/start')
def inventory_start():
    job = uuid.uuid4().hex[:12]
    jp = job_path(job)
    jp.mkdir(parents=True, exist_ok=True)

    dsn = request.form.get('dsn', '').strip()
    db_user = request.form.get('db_user', '').strip()
    db_password = request.form.get('db_password', '')
    JOB_DB_PASSWORDS[job] = db_password

    ref = load_reference_data(dsn, db_user, db_password)
    profiles = inventory_profile_payload(ref)
    refs = ref.as_client_dict()
    refs.update({
        'profiles': profiles,
        'existing_pairs': [
            f"{str(x.get('model', '') or '').strip().upper()}|{str(x.get('serial', '') or '').strip().upper()}"
            for x in (ref.existing_serials or [])
            if str(x.get('model', '') or '').strip() and str(x.get('serial', '') or '').strip()
        ],
    })

    (jp / 'job_config.json').write_text(
        json.dumps({'dsn': dsn, 'db_user': db_user, 'mode': 'inventory'}, indent=2),
        encoding='utf-8',
    )
    _inventory_refs_path(jp).write_text(json.dumps(refs, indent=2), encoding='utf-8')
    _inventory_rows_path(jp).write_text('[]', encoding='utf-8')
    write_inventory_csv([], jp / 'InventoryImport.csv')
    return redirect(url_for('inventory_review', job=job))


@app.get('/inventory/<job>')
def inventory_review(job):
    jp, rows, refs = load_inventory_job(job)
    return render_template(
        'inventory.html',
        job=job,
        rows=rows,
        refs=refs,
        inventory_headers=INVENTORY_HEADERS,
    )


@app.post('/api/inventory/<job>/save')
def inventory_save(job):
    jp, _old_rows, refs = load_inventory_job(job)
    data = request.get_json(silent=True) or {}
    rows = data.get('rows', [])
    if not isinstance(rows, list):
        return jsonify({'ok': False, 'error': 'Inventory payload must be a list.'}), 400
    try:
        clean_rows = save_inventory_job(jp, rows, refs)
    except ValueError as e:
        return jsonify({'ok': False, 'error': str(e)}), 400
    return jsonify({'ok': True, 'count': len(clean_rows)})


@app.get('/inventory/<job>/download')
def inventory_download(job):
    jp, _rows, _refs = load_inventory_job(job)
    p = jp / 'InventoryImport.csv'
    if not p.exists():
        abort(404)
    return send_file(p, as_attachment=True, download_name='InventoryImport.csv')


@app.get('/inventory/<job>/package')
def inventory_package(job):
    jp, rows, refs = load_inventory_job(job)
    try:
        save_inventory_job(jp, rows, refs)
    except ValueError as e:
        flash('Inventory package blocked: ' + str(e), 'error')
        return redirect(url_for('inventory_review', job=job))
    return send_file(jp / 'Inventory_Import_Package.zip', as_attachment=True)


@app.post('/inventory/<job>/run-rto')
def inventory_run_rto(job):
    jp, rows, refs = load_inventory_job(job)
    try:
        rows = save_inventory_job(jp, rows, refs)
    except ValueError as e:
        flash('Inventory import blocked: ' + str(e), 'error')
        return redirect(url_for('inventory_review', job=job))

    exe = Path(os.getenv('RTO_EXE', r'C:\RTOwin\RTO-win.exe'))
    try:
        staging = stage_inventory_import(rows, jp, exe)
    except Exception as e:
        flash('Could not prepare the inventory import: ' + str(e), 'error')
        return redirect(url_for('inventory_review', job=job))

    if os.getenv('ENABLE_RTO_LAUNCH', 'true').lower() != 'true':
        flash(f'Inventory file is ready at {staging}, but RTO launch is disabled in .env.', 'error')
        return redirect(url_for('inventory_review', job=job))

    if os.name != 'nt' or not exe.exists():
        flash(f'RTO Pro executable not found: {exe}. The inventory file is ready at {staging}.', 'error')
        return redirect(url_for('inventory_review', job=job))

    try:
        command_parts = [str(exe), '-r', str(staging)]
        command_line = f'"{exe}" -r "{staging}"'
        (jp / 'Run_RTO_Inventory_Import.cmd').write_text(
            '@echo off\r\n' + f'cd /d "{exe.parent}"\r\n' + command_line + '\r\n',
            encoding='cp1252',
        )
        (jp / 'RTO_Inventory_Launch_Command.txt').write_text(command_line + '\n', encoding='utf-8')
        subprocess.Popen(command_parts, cwd=str(exe.parent), shell=False)
        save_inventory_job(jp, rows, refs)
        flash(f'RTO Pro inventory receiver launched for {len(rows)} item(s).', 'ok')
    except Exception as e:
        flash('RTO Pro inventory receiver could not be launched: ' + str(e), 'error')
    return redirect(url_for('inventory_review', job=job))


@app.post('/process')
def process():
    # V7.6 accepts any mix of ShedSuite CSVs and supported contract PDFs.
    files = [f for f in request.files.getlist('contract_files') if f and f.filename]
    # Backward compatibility for bookmarked/older V7 pages.
    if not files:
        files = [f for f in request.files.getlist('csv_files') if f and f.filename]
    if not files:
        legacy = request.files.get('csv_file')
        if legacy and legacy.filename:
            files = [legacy]

    if not files:
        flash('Choose one or more ShedSuite CSV or contract PDF files.', 'error')
        return redirect(url_for('index'))

    job = uuid.uuid4().hex[:12]
    jp = job_path(job)
    jp.mkdir(parents=True, exist_ok=True)

    required = {
        'Serial Number',
        'Customer Order Id',
        'Customer First Name',
        'Customer Last Name',
        'Company Name',
        'Dealer',
    }

    try:
        source_rows, input_summary, input_warnings = merge_contract_inputs(
            files, jp, required
        )
    except ValueError as e:
        shutil.rmtree(jp, ignore_errors=True)
        flash(str(e), 'error')
        return redirect(url_for('index'))

    if not source_rows:
        shutil.rmtree(jp, ignore_errors=True)
        flash('The selected files contain no contract rows.', 'error')
        return redirect(url_for('index'))

    (jp / 'input_files.json').write_text(
        json.dumps(input_summary, indent=2), encoding='utf-8'
    )

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

    # Keep the original input filename/source type on every transformed row.
    source_by_order_for_file = {
        str(x.get('Customer Order Id', '') or '').strip(): x
        for x in source_rows
    }
    for r in rows:
        src_for_file = source_by_order_for_file.get(
            str(r.get('_source_order_id', '') or '').strip(), {}
        )
        source_type = str(src_for_file.get('_import_source_type', 'csv') or 'csv').lower()
        source_name = str(
            src_for_file.get('_import_pdf_name', '')
            or src_for_file.get('_import_csv_name', '')
            or ''
        )
        r['_source_type'] = source_type
        r['_source_file_name'] = source_name
        r['_source_csv_name'] = str(src_for_file.get('_import_csv_name', '') or '')
        r['_source_pdf_name'] = str(src_for_file.get('_import_pdf_name', '') or '')
        if source_type == 'pdf':
            # RentaBarn PDF fields are authoritative for these labeled values.
            # Do not reuse ShedSuite-derived calculations when the PDF prints an
            # explicit agreement/payment amount.
            agreement = str(src_for_file.get('_pdf_agreement_number', '') or src_for_file.get('Customer Order Id', '') or '').strip()
            if agreement:
                r['CONTRACT'] = agreement
            if str(src_for_file.get('_pdf_salesperson_source', '') or '').strip():
                r['SALESMAN'] = str(src_for_file.get('_pdf_salesperson_source', '') or '').strip()
            if str(src_for_file.get('_pdf_county_of_delivery', '') or '').strip() and not str(r.get('COUNTY', '') or '').strip():
                r['COUNTY'] = str(src_for_file.get('_pdf_county_of_delivery', '') or '').strip().upper()
            if str(src_for_file.get('_pdf_contract_price', '') or '').strip():
                r['CONTRACTAMT'] = str(src_for_file.get('_pdf_contract_price', '') or '').strip()

            # RTO Pro expects the individual inventory RATE when inventory is
            # included and PMT should be 0.00.  RentaBarn's "PMT Before Tax"
            # is already the rent-before-tax amount; LDW is a separate row and
            # must NOT be subtracted from it a second time.
            pdf_pmt_before_tax = str(src_for_file.get('_pdf_pmt_before_tax', '') or '').strip()
            if pdf_pmt_before_tax:
                r['RATE1'] = pdf_pmt_before_tax
                r['PMT'] = '0.00'
            pdf_ldw = str(src_for_file.get('_pdf_ldw', '') or '').strip()
            if pdf_ldw:
                r['GRP'] = pdf_ldw
            pdf_security = str(src_for_file.get('_pdf_security_deposit', '') or '').strip()
            if pdf_security:
                r['EXTRARENT'] = pdf_security
            pdf_reserve = str(src_for_file.get('_pdf_purchase_reserve', '') or '').strip()
            if pdf_reserve:
                r['PAIDDOWN'] = pdf_reserve

            # 90 Days SAC is printed as e.g. YES (10/13/2026). Use that exact
            # date rather than recomputing 90 days from another field.
            pdf_sac = str(src_for_file.get('_pdf_sac_date', '') or '').strip()
            if pdf_sac:
                r['SACDATE'] = pdf_sac

            # RTO Pro EMAILINV: 1 = invoice by email only, 0 = not email-only.
            pdf_email_invoice = str(src_for_file.get('_pdf_email_invoice', '') or '').strip()
            if pdf_email_invoice in {'0', '1'}:
                r['EMAILINV'] = pdf_email_invoice

            if str(src_for_file.get('Date Created', '') or '').strip():
                r['CONTRACTDATE'] = str(src_for_file.get('Date Created', '') or '').strip()
            r['CONDITION1'] = str(src_for_file.get('Condition', '') or '').strip().upper()

            # Expose exact PDF values in the review UI for visual verification.
            for meta_key in [
                '_pdf_sac_date', '_pdf_sac_source', '_pdf_agreement_number',
                '_pdf_pmt_before_tax', '_pdf_total_monthly_pmt', '_pdf_total_tax',
                '_pdf_ldw', '_pdf_security_deposit', '_pdf_purchase_reserve',
                '_pdf_paperless_billing', '_pdf_email_invoice',
            ]:
                r[meta_key] = str(src_for_file.get(meta_key, '') or '').strip()
    warnings.extend('Input: ' + x for x in input_warnings)
    if input_summary.get('file_count', 0) > 1:
        parts = []
        if input_summary.get('csv_count'):
            parts.append(f"{input_summary['csv_count']} CSV")
        if input_summary.get('pdf_count'):
            parts.append(f"{input_summary['pdf_count']} PDF")
        warnings.append(
            f"Input: merged {' + '.join(parts) or str(input_summary['file_count']) + ' file(s)'} into "
            f"{input_summary['row_count']} unique contracts"
            + (f"; skipped {input_summary['duplicates_skipped']} duplicate(s)."
               if input_summary.get('duplicates_skipped') else '.')
        )

    # Keep DOB as the exact calendar date written in the ShedSuite CSV.
    fix_dobs_from_source(rows, source_rows)

    # V6.3.1 used-building workflow. This is deliberately post-transform so the
    # original V3 mapping/model logic stays authoritative.
    apply_used_building_defaults(rows, source_rows, ref)

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

    refs = add_contract_history_to_refs(ref.as_client_dict(), ref.existing_contracts)

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
    # ShedSuite CSV rows keep the existing packet order and background Delivery
    # Certificate workflow. Imported contract PDFs are already the source packet,
    # so they are copied directly into Combined_Files and never sent to Chromium.
    # -----------------------------------------------------------------------
    by_order = {
        str(x.get('Customer Order Id', '')).strip(): x
        for x in source_rows
    }
    download_requested = request.form.get('download_pdfs') == 'yes'
    pdf_mode = request.form.get('pdf_mode', 'contract')
    delivery_order_ids: set[str] = set()

    for r in rows:
        order_id = str(r.get('_source_order_id', '')).strip()
        srcrow = by_order.get(order_id, {})
        source_type = str(srcrow.get('_import_source_type', r.get('_source_type', 'csv')) or 'csv').lower()
        base = safe_filename(r['_pdf_name'])
        target = pdf_dir / f'{base}.pdf'

        if source_type == 'pdf':
            stored_name = str(srcrow.get('_import_pdf_stored_name', '') or '').strip()
            original_pdf = jp / 'Input_PDFs' / stored_name if stored_name else None
            if original_pdf and original_pdf.exists():
                try:
                    shutil.copy2(original_pdf, target)
                    r['_pdf_manifest'] = ['Imported contract PDF']
                    r['_pdf_available'] = True
                except Exception as exc:
                    pdf_errors.append(f'{r["NAME"]}: could not copy imported PDF ({exc})')
            r['_delivery_certificate'] = 'Not requested'
            r['_delivery_certificate_method'] = 'Imported PDF source'
            r['_delivery_certificate_name'] = ''
        else:
            # V6/V7 fast direct-link packet for ShedSuite. This part only runs
            # when the user leaves Download PDFs enabled.
            if download_requested:
                ordered_documents = build_ordered_document_list(srcrow, [])
                if pdf_mode != 'all':
                    ordered_documents = [
                        (name, url)
                        for name, url in ordered_documents
                        if is_invoice(name) or is_contract(name)
                    ]
                urls = [url for _name, url in ordered_documents]
                if urls:
                    errs = download_and_combine(urls, target)
                    pdf_errors.extend([f'{r["NAME"]}: {x}' for x in errs])
                r['_pdf_manifest'] = [name for name, _url in ordered_documents]
                r['_pdf_available'] = target.exists()
                r['_delivery_certificate'] = 'Queued'
                r['_delivery_certificate_method'] = ''
                r['_delivery_certificate_name'] = ''
                delivery_order_ids.add(order_id)
            else:
                r['_delivery_certificate'] = 'Not requested'
                r['_delivery_certificate_method'] = ''
                r['_delivery_certificate_name'] = ''

        # Extract any immediately available PDF fields. For ShedSuite rows the
        # Delivery Certificate worker may enrich these again later (coordinates).
        if not target.exists():
            r['_pdf_available'] = False
            continue

        try:
            fields = extract_pdf_fields(target)
        except Exception as exc:
            fields = {}
            pdf_errors.append(f'{r["NAME"]}: PDF field extraction failed ({exc})')

        r['_pdf_salesman'] = fields.get('salesman', '')
        r['_pdf_coordinates'] = fields.get('coordinates', '')
        r['_pdf_latitude'] = fields.get('latitude', '')
        r['_pdf_longitude'] = fields.get('longitude', '')
        r['_pdf_invoice'] = fields.get('invoice', '')

        # A RentaBarn-style contract parser may already have a salesperson. The
        # generic packet extractor only overwrites it when it actually finds one.
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

    if download_requested and delivery_order_ids:
        _initialize_certificate_status(jp, rows, active=True)
        # Let the worker select all queued ShedSuite/CSV rows from the saved job.
        # Retry requests still pass a specific order id.
        _start_certificate_worker(job)
    else:
        _initialize_certificate_status(jp, rows, active=False)

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

    input_summary = {}
    input_summary_path = jp / 'input_files.json'
    if input_summary_path.exists():
        try:
            input_summary = json.loads(input_summary_path.read_text(encoding='utf-8'))
        except Exception:
            input_summary = {}

    page_html = render_template(
        'review.html',
        job=job,
        rows=rows,
        refs=refs,
        warnings=warnings,
        headers=RTO_HEADERS,
        sources=sources,
        input_summary=input_summary,
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

    # V7.7: cards can be reordered in the browser. Merge every submitted row
    # by its permanent _row_uid instead of by list position, while still using
    # the freshest server copy so background certificate/coordinate updates win.
    with _job_lock(job):
        try:
            latest_server_rows = json.loads((jp / 'rows.json').read_text(encoding='utf-8'))
            _ensure_row_uids(latest_server_rows)
            old_uids = {str(x.get('_row_uid', '') or '').strip() for x in old_rows}
            latest_uids = {str(x.get('_row_uid', '') or '').strip() for x in latest_server_rows}
            if len(latest_server_rows) == len(old_rows) and latest_uids == old_uids:
                old_rows = latest_server_rows
        except Exception:
            pass

    metadata = sorted({
        k
        for old in old_rows
        for k in old.keys()
        if k.startswith('_')
    })

    old_by_uid = {
        str(old.get('_row_uid', '') or '').strip(): old
        for old in old_rows
        if str(old.get('_row_uid', '') or '').strip()
    }
    incoming_uids = [str(new.get('_row_uid', '') or '').strip() for new in incoming]
    uid_merge = (
        len(old_by_uid) == len(old_rows)
        and all(incoming_uids)
        and len(set(incoming_uids)) == len(incoming_uids)
        and set(incoming_uids) == set(old_by_uid)
    )

    if uid_merge:
        row_pairs = [(old_by_uid[uid], new) for uid, new in zip(incoming_uids, incoming)]
    else:
        # Compatibility fallback for jobs created before stable row ids existed.
        row_pairs = list(zip(old_rows, incoming))

    rows = []

    for old, new in row_pairs:
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

        server_owned_metadata = {
            '_delivery_certificate', '_delivery_certificate_method',
            '_delivery_certificate_name', '_pdf_manifest', '_pdf_available',
            '_pdf_pages',
            # V7.2: these can arrive from the background certificate worker
            # after the review page was opened, so stale browser JSON must not
            # be allowed to wipe them out.
            '_pdf_coordinates', '_pdf_latitude', '_pdf_longitude',
        }
        for k in metadata:
            # These fields may change in the background after the browser loaded
            # the review page. Never let stale browser JSON overwrite them.
            if k in server_owned_metadata:
                r[k] = old.get(k, '')
            else:
                r[k] = new.get(
                    k,
                    old.get(k, ''),
                )

        # V7.2: if coordinates arrived while this review page was open, preserve
        # them in DIRECTIONS even when the browser submitted an older copy of
        # the row. This never removes or replaces user-entered directions.
        live_coords = str(old.get('_pdf_coordinates', '') or '').strip()
        if live_coords:
            directions = str(r.get('DIRECTIONS', '') or '')
            if live_coords not in directions:
                r['DIRECTIONS'] = live_coords + (('      ' + directions) if directions else '')

        # Keep the same phone normalization on manually edited values as on
        # the original CSV transform. A leading US country code (1) is ignored.
        for phone_field in ['CELL', 'PHONE5', 'REFPHONE1', 'REFPHONE2']:
            r[phone_field] = phone(r.get(phone_field, ''))
        if '_secondary_phone' in r:
            r['_secondary_phone'] = phone(r.get('_secondary_phone', ''))
        _prepare_phone_and_customer_comment(r)

        # Save UI helper metadata as well.
        for k in [
            '_store_title',
            '_dealer_name',
            '_set_contract',
            '_warnings',
            '_used_building',
            '_used_detected',
            '_used_detection_reason',
            '_used_base_contract',
            '_used_contract_suffix',
            '_used_suffix_manual',
            '_used_contract_error',
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


@app.post('/api/reorder/<job>')
def api_reorder_contracts(job):
    """Persist the contract-card order without changing any contract data."""
    jp = job_path(job)
    if not (jp / 'rows.json').exists():
        abort(404)

    payload = request.get_json(silent=True) or {}
    requested = payload.get('order', [])
    if not isinstance(requested, list):
        return jsonify({'ok': False, 'error': 'Invalid order payload.'}), 400
    requested = [str(x or '').strip() for x in requested]
    if not requested or any(not x for x in requested) or len(set(requested)) != len(requested):
        return jsonify({'ok': False, 'error': 'Invalid or duplicate row ids.'}), 400

    with _job_lock(job):
        rows = json.loads((jp / 'rows.json').read_text(encoding='utf-8'))
        _ensure_row_uids(rows)
        by_uid = {str(r.get('_row_uid', '') or '').strip(): r for r in rows}
        if len(by_uid) != len(rows) or len(requested) != len(rows) or set(requested) != set(by_uid):
            return jsonify({'ok': False, 'error': 'Contract list changed. Refresh and try sorting again.'}), 409
        ordered = [by_uid[uid] for uid in requested]
        save_job_rows(jp, ordered)

    return jsonify({'ok': True, 'count': len(requested)})


@app.post('/api/discard-row/<job>/<row_uid>')
def api_discard_contract_uid(job, row_uid):
    """Stable-id version of discard used by the sortable V7.7 review UI."""
    _jp, rows, _refs, _warnings = load_job(job)
    index = next(
        (i for i, r in enumerate(rows) if str(r.get('_row_uid', '') or '').strip() == str(row_uid or '').strip()),
        -1,
    )
    if index < 0:
        return jsonify({'ok': False, 'error': 'Contract no longer exists.'}), 404
    return api_discard_contract(job, index)


@app.post('/api/discard/<job>/<int:index>')
def api_discard_contract(job, index):
    """Permanently remove one uploaded contract from the current job.

    This immediately rebuilds the RTO CSV/XML/ZIP so a discarded ShedSuite
    row cannot accidentally make it into a later RTO Pro import.
    """
    jp, rows, _refs, _warnings = load_job(job)
    if index < 0 or index >= len(rows):
        return jsonify({'ok': False, 'error': 'Contract no longer exists.'}), 404

    removed = rows.pop(index)
    order_id = str(removed.get('_source_order_id', '') or '').strip()
    display_name = str(removed.get('NAME', '') or '').strip() or f'Row {index + 1}'

    # Keep source-row details aligned with the remaining contracts.
    source_path = jp / 'source_rows.json'
    if source_path.exists():
        try:
            source_rows = json.loads(source_path.read_text(encoding='utf-8'))
            if isinstance(source_rows, list):
                removed_source = False
                kept = []
                for src in source_rows:
                    src_order = str((src or {}).get('Customer Order Id', '') or '').strip() if isinstance(src, dict) else ''
                    if not removed_source and order_id and src_order == order_id:
                        removed_source = True
                        continue
                    kept.append(src)
                source_path.write_text(json.dumps(kept, indent=2), encoding='utf-8')
        except Exception:
            # Source-row cleanup is helpful for the review dialog, but should
            # never block the actual contract discard.
            pass

    # Remove this row's combined packet so it cannot remain inside Full ZIP.
    pdf_base = safe_filename(removed.get('_pdf_name') or removed.get('NAME') or 'contract')
    pdf_path = jp / 'Combined_Files' / f'{pdf_base}.pdf'
    try:
        if pdf_path.exists():
            pdf_path.unlink()
    except Exception:
        pass

    # Anything created from a prior launch/import is now stale and should be
    # regenerated only from the reduced row set.
    for stale_name in [
        'PDF_Import_Report.csv',
        'RTOProImport_LAUNCHED.csv',
        'Run_RTO_Import.cmd',
        'RTO_Launch_Command.txt',
    ]:
        stale = jp / stale_name
        try:
            if stale.exists():
                stale.unlink()
        except Exception:
            pass

    with _job_lock(job):
        st = _read_certificate_status(jp)
        if order_id:
            st.setdefault('items', {})[order_id] = {
                'status': 'discarded',
                'name': display_name,
                'detail': 'Discarded from this import.',
            }
            _write_certificate_status(jp, st)
    save_job_rows(jp, rows)
    return jsonify({
        'ok': True,
        'removed': display_name,
        'order_id': order_id,
        'remaining': len(rows),
    })


@app.get('/api/delivery-status/<job>')
def api_delivery_status(job):
    jp = job_path(job)
    if not (jp / 'rows.json').exists():
        abort(404)
    with _job_lock(job):
        status = _read_certificate_status(jp)
    items = status.get('items', {}) if isinstance(status.get('items', {}), dict) else {}
    counts = {'queued': 0, 'working': 0, 'ready': 0, 'missing': 0, 'skipped': 0, 'discarded': 0, 'not_requested': 0}
    for item in items.values():
        key = str((item or {}).get('status', '')).lower()
        if key in counts:
            counts[key] += 1
    status['counts'] = counts
    status['total'] = counts['queued'] + counts['working'] + counts['ready'] + counts['missing'] + counts['skipped']
    status['done'] = counts['ready'] + counts['missing'] + counts['skipped']
    return jsonify(status)


@app.post('/api/delivery-retry/<job>/<order_id>')
def api_delivery_retry(job, order_id):
    jp = job_path(job)
    if not (jp / 'rows.json').exists():
        abort(404)
    with CERT_ACTIVE_LOCK:
        if job in CERT_ACTIVE:
            return jsonify({'ok': False, 'error': 'A Delivery Certificate worker is still running. Retry when it finishes.'}), 409
    with _job_lock(job):
        rows = json.loads((jp / 'rows.json').read_text(encoding='utf-8'))
        match = next((r for r in rows if str(r.get('_source_order_id', '') or '').strip() == order_id), None)
        if not match:
            return jsonify({'ok': False, 'error': 'Contract no longer exists.'}), 404
        st = _read_certificate_status(jp)
        st.setdefault('items', {})[order_id] = {
            'status': 'queued',
            'name': str(match.get('NAME', '') or order_id),
            'detail': 'Queued for retry…',
        }
        st['active'] = True
        st['complete'] = False
        _write_certificate_status(jp, st)
    if not _start_certificate_worker(job, {order_id}):
        return jsonify({'ok': False, 'error': 'A Delivery Certificate worker is already running. Retry when it finishes.'}), 409
    return jsonify({'ok': True})


@app.post('/api/delivery-skip/<job>/<order_id>')
def api_delivery_skip(job, order_id):
    jp = job_path(job)
    if not (jp / 'rows.json').exists():
        abort(404)
    with _job_lock(job):
        st = _read_certificate_status(jp)
        item = st.setdefault('items', {}).setdefault(order_id, {})
        item['status'] = 'skipped'
        item['detail'] = 'Skipped by user.'
        _write_certificate_status(jp, st)
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
            f"{result['accounts_matched']} account number(s) matched; "
            f"{result.get('customer_comments_synced', 0)} customer comment(s) synced. "
            "Last-used model numbers were refreshed from Firebird.",
            'ok',
        )
        if result.get('customer_comment_errors'):
            flash(
                'Customer comment sync warnings: ' + ' | '.join(result['customer_comment_errors'][:6]),
                'error',
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
    jp, rows, refs, _ = load_job(job)

    # Recheck used-building suffixes against the newest RTO data just before
    # import. If U was taken meanwhile, V6.3.1 advances to A/B/C... automatically.
    suffix_changes = ensure_used_contracts_available(jp, rows, refs)
    if suffix_changes:
        save_job_rows(jp, rows)
        flash('Used-building contract suffix refreshed: ' + ' | '.join(suffix_changes[:6]), 'ok')

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
