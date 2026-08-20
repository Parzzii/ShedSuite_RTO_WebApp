from __future__ import annotations

import asyncio
import base64
import os
import re
import tempfile
from pathlib import Path

import fitz
from openpyxl import load_workbook
from playwright.async_api import async_playwright, expect

from pdf_tools import bytes_to_pdf, safe_filename

BASE_URL = 'https://app.shedsuite.com/rto/order/'
HOME_URL = 'https://app.shedsuite.com/'


def _safe_email(email: str) -> str:
    return re.sub(r'[^a-zA-Z0-9._-]', '_', (email or '').strip().lower())


def _find_legacy_workbook(base: Path) -> Path | None:
    """Find the old workbook that already contains the Logininfo table.

    V5 intentionally does NOT ask for a ShedSuite password. It reuses the same
    credentials source as Contract_Import.xlsm / Contract_Import.py.
    """
    explicit = os.getenv('LEGACY_XLSM_PATH', '').strip()
    candidates: list[Path] = []
    if explicit:
        candidates.append(Path(explicit))

    fixed_roots = [base, base.parent]
    home = Path(os.environ.get('USERPROFILE') or Path.home())
    fixed_roots += [home / 'Downloads', home / 'Desktop', home / 'Documents']

    for root in fixed_roots:
        candidates += [
            root / 'Contract_Import.xlsm',
            root / 'Shedsuite_Import V1.1.7.6' / 'Contract_Import.xlsm',
            root / 'Shedsuite_Import V1.1.7.6' / 'Shedsuite_Import V1.1.7.6' / 'Contract_Import.xlsm',
        ]

    for p in candidates:
        try:
            if p.exists() and p.is_file():
                return p.resolve()
        except Exception:
            pass

    # Last-resort shallow search of the normal user folders. This avoids a
    # full drive scan while still finding an extracted old program folder.
    for root in [home / 'Downloads', home / 'Desktop', home / 'Documents']:
        if not root.exists():
            continue
        try:
            for p in root.glob('*/Contract_Import.xlsm'):
                if p.is_file():
                    return p.resolve()
            for p in root.glob('*/*/Contract_Import.xlsm'):
                if p.is_file():
                    return p.resolve()
        except Exception:
            pass
    return None


def _legacy_logins(base: Path) -> tuple[list[tuple[str, str]], Path | None, str]:
    """Read the exact Excel Logininfo table. Passwords remain local/in-memory."""
    book = _find_legacy_workbook(base)
    if not book:
        return [], None, ''

    wb = load_workbook(book, data_only=True, read_only=False, keep_vba=True)
    try:
        selected = str(wb['Import Tools']['D4'].value or '').strip() if 'Import Tools' in wb.sheetnames else ''
        if 'RTOPro Data' not in wb.sheetnames:
            return [], book, selected
        ws = wb['RTOPro Data']
        table = ws.tables.get('Logininfo')
        result: list[tuple[str, str]] = []
        if table:
            rows = list(ws[table.ref])
            for row in rows[1:]:
                email = str(row[0].value or '').strip()
                password = str(row[1].value or '')
                if email and password:
                    result.append((email, password))
    finally:
        wb.close()

    if selected:
        result.sort(key=lambda x: 0 if x[0].lower() == selected.lower() else 1)
    return result, book, selected


def _preferred_login(company: str, dealer: str) -> list[str]:
    """Same company/login selection idea used by the old Excel macro."""
    text = f'{company} {dealer}'.lower()
    if 'smart shed' in text or 'magnolia' in text:
        return ['office@magnoliarto.com']
    if 'lonestar' in text or re.search(r'\blsr\b', text):
        return ['lsrbarns@gmail.com']
    if 'westwood' in text or re.search(r'\bwwp\b', text):
        return ['office@westwoodpayments.com']
    if '4 seasons' in text or 'four seasons' in text:
        return ['conrad.s+wv@evergreenrto.com']
    if 'phoenix' in text and ('dbm' in text or 'dutch' in text):
        return ['conrad.s+dbm@evergreenrto.com']
    if 'yoder storage' in text or 'yss' in text:
        return ['conrad.s+ysb@evergreenrto.com', 'joy.f+davis@evergreenrto.com']
    if 'alpine' in text:
        return ['tori.h@evergreenrto.com', 'rentabarn2@gmail.com']
    if 'genesis' in text:
        return ['conrad.s@evergreenrto.com']
    return []


def _unwrap_url(url: str) -> str:
    """Legacy ShedSuite viewer URLs sometimes contain a second real https URL."""
    if not url:
        return ''
    hits = [m.start() for m in re.finditer('https', str(url))]
    if len(hits) >= 2:
        return str(url)[hits[1]:]
    if hits:
        return str(url)[hits[0]:]
    return str(url)


async def _url_to_pdf_bytes(context, url: str) -> bytes | None:
    url = _unwrap_url(url)
    if not url or url.startswith(('blob:', 'data:')):
        return None
    try:
        response = await context.request.get(url, timeout=60000)
        if not response.ok:
            return None
        return bytes_to_pdf(await response.body())
    except Exception:
        return None


async def _page_blob_bytes(page) -> bytes | None:
    try:
        data_url = await page.evaluate('''async () => {
            const response = await fetch(window.location.href);
            const blob = await response.blob();
            return await new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = () => resolve(reader.result);
                reader.onerror = reject;
                reader.readAsDataURL(blob);
            });
        }''')
        if not data_url or ',' not in data_url:
            return None
        return bytes_to_pdf(base64.b64decode(data_url.split(',', 1)[1]))
    except Exception:
        return None


async def _login_context(browser, email: str, password: str, base: Path, legacy_book: Path | None):
    """Open ShedSuite without prompting; auto-login from old Excel if needed."""
    local_auth = base / 'auth_state' / f'shedsuite_auth_{_safe_email(email)}.json'
    local_auth.parent.mkdir(parents=True, exist_ok=True)
    legacy_auth = legacy_book.parent / 'auth_state' / local_auth.name if legacy_book else None
    auth = local_auth if local_auth.exists() else (legacy_auth if legacy_auth and legacy_auth.exists() else None)

    kwargs = {
        'viewport': {'width': 1100, 'height': 850},
        'accept_downloads': True,
        'user_agent': (
            'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
            'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        ),
    }
    if auth:
        kwargs['storage_state'] = str(auth)

    context = await browser.new_context(**kwargs)
    context.set_default_navigation_timeout(90000)
    context.set_default_timeout(15000)
    page = await context.new_page()
    await page.goto(HOME_URL, wait_until='domcontentloaded', timeout=90000)
    await page.wait_for_timeout(2500)

    login_button = page.get_by_role('button', name='Log in')
    try:
        needs_login = await login_button.is_visible()
    except Exception:
        needs_login = False

    if needs_login:
        # No prompt: this is exactly why we read the old workbook's Logininfo.
        await page.get_by_label('Email').fill(email)
        await page.get_by_label('Password').fill(password)
        await login_button.click()
        await expect(page).to_have_url(re.compile(r'https://app\.shedsuite\.com/rto/.*'), timeout=300000)
        await context.storage_state(path=str(local_auth))

    return context, page


async def _capture_delivery_certificate(page, context, order_id: str) -> tuple[bytes | None, str, str]:
    """Capture only the Delivery Certificate using the robust V4 capture paths."""
    order_url = BASE_URL + str(order_id)
    responses = []

    def remember(response):
        responses.append(response)

    try:
        await page.goto(order_url, wait_until='domcontentloaded', timeout=90000)
        await page.wait_for_selector('span[class^="fileName"]', timeout=90000)
        spans = page.locator('span[class^="fileName"]')
        target = None
        found_name = ''
        for i in range(await spans.count()):
            loc = spans.nth(i)
            name = ((await loc.text_content()) or '').strip()
            low = name.lower()
            if 'delivery' in low and ('certificate' in low or 'cert' in low):
                target = loc
                found_name = name
                break
        if target is None:
            return None, 'Delivery Certificate not found', ''

        # Some ShedSuite builds expose an authenticated anchor directly.
        try:
            href = await target.evaluate("el => (el.closest('a') && el.closest('a').href) || ''")
            if href:
                data = await _url_to_pdf_bytes(context, href)
                if data:
                    return data, 'authenticated link', found_name
        except Exception:
            pass

        page.on('response', remember)
        popup_task = asyncio.create_task(page.wait_for_event('popup'))
        download_task = asyncio.create_task(page.wait_for_event('download'))
        before_url = page.url

        try:
            await target.click(timeout=15000)
        except Exception as exc:
            for task in (popup_task, download_task):
                if not task.done():
                    task.cancel()
            return None, f'click failed: {exc}', found_name

        done, pending = await asyncio.wait(
            [popup_task, download_task], timeout=18, return_when=asyncio.FIRST_COMPLETED
        )
        for task in pending:
            task.cancel()

        if download_task in done:
            try:
                download = await download_task
                with tempfile.NamedTemporaryFile(delete=False, suffix='.bin') as tf:
                    temp_path = Path(tf.name)
                try:
                    await download.save_as(str(temp_path))
                    return bytes_to_pdf(temp_path.read_bytes()), 'browser download', found_name
                finally:
                    temp_path.unlink(missing_ok=True)
            except Exception:
                pass

        if popup_task in done:
            try:
                popup = await popup_task
                try:
                    await popup.wait_for_load_state('domcontentloaded', timeout=30000)
                except Exception:
                    pass
                data = None
                if popup.url.startswith(('blob:', 'data:')):
                    data = await _page_blob_bytes(popup)
                else:
                    data = await _url_to_pdf_bytes(context, popup.url)
                    if not data:
                        data = await _page_blob_bytes(popup)
                try:
                    await popup.close()
                except Exception:
                    pass
                if data:
                    return data, 'popup/viewer', found_name
            except Exception:
                pass

        # A PDF viewer may fetch the bytes in the background without a useful URL.
        await page.wait_for_timeout(2500)
        for response in reversed(responses):
            try:
                ctype = str((response.headers or {}).get('content-type', '')).lower()
                url = str(response.url or '').lower()
                if 'pdf' in ctype or 'application/octet-stream' in ctype or 'delivery' in url or 'certificate' in url:
                    raw = await response.body()
                    if raw:
                        return bytes_to_pdf(raw), 'network response', found_name
            except Exception:
                continue

        # Final fallback: the document replaced the order page in the same tab.
        if page.url != before_url:
            try:
                data = await (_page_blob_bytes(page) if page.url.startswith(('blob:', 'data:')) else _url_to_pdf_bytes(context, page.url))
                if not data:
                    data = await _page_blob_bytes(page)
                if data:
                    return data, 'same-tab viewer', found_name
            except Exception:
                pass

        return None, 'found and clicked, but no document bytes were captured', found_name
    except Exception as exc:
        return None, str(exc), ''
    finally:
        try:
            page.remove_listener('response', remember)
        except Exception:
            pass


def _append_pdf(target: Path, cert_bytes: bytes) -> None:
    cert = fitz.open(stream=cert_bytes, filetype='pdf')
    try:
        if target.exists():
            current = fitz.open(target)
            try:
                current.insert_pdf(cert)
                temp = target.with_suffix('.deliverytmp.pdf')
                current.save(temp, garbage=4, deflate=True)
            finally:
                current.close()
            temp.replace(target)
        else:
            cert.save(target, garbage=4, deflate=True)
    finally:
        cert.close()


async def _run(rows: list[dict], source_rows: list[dict], pdf_dir: Path, base: Path) -> list[str]:
    logins, legacy_book, _selected = _legacy_logins(base)
    if not logins:
        return [
            'Delivery Certificate: V6 could not find Contract_Import.xlsm / Logininfo. '
            'Put the old Contract_Import.xlsm beside V6 or set LEGACY_XLSM_PATH. '
            'V6 intentionally never asks for a ShedSuite password.'
        ]

    src_by_order = {str(x.get('Customer Order Id', '')).strip(): x for x in source_rows}
    login_map = {email.lower(): (email, password) for email, password in logins}
    warnings: list[str] = []

    async with async_playwright() as playwright:
        # The old Excel automation explicitly required visible Chromium; headless
        # mode was unreliable with ShedSuite.
        browser = await playwright.chromium.launch(headless=False)
        contexts: dict[str, tuple] = {}
        try:
            for row in rows:
                oid = str(row.get('_source_order_id', '')).strip()
                if not oid:
                    continue
                src = src_by_order.get(oid, {})

                # V6 keeps V3's company/login selection authoritative.  The
                # transformed V3 row already contains the exact ShedSuite login
                # selected by COMPANY_RULES as `_login`.  When it is present we
                # use ONLY that account, so one company's order is never searched
                # through unrelated company logins.
                v3_login = str(row.get('_login') or '').strip()
                ordered: list[tuple[str, str]] = []
                if v3_login:
                    item = login_map.get(v3_login.lower())
                    if item:
                        ordered.append(item)
                    else:
                        warnings.append(
                            f"{row.get('NAME') or oid}: Delivery Certificate: "
                            f"V3 selected ShedSuite login {v3_login}, but that login was not found in Logininfo."
                        )
                        row['_delivery_certificate'] = 'Missing'
                        row['_delivery_certificate_method'] = ''
                        continue
                else:
                    # Unknown/manual companies do not have a V3 `_login`. In that
                    # case use the narrow V5 preference map first; only if it has
                    # no match do we fall back to the Logininfo list.
                    preferred = _preferred_login(src.get('Company Name', ''), src.get('Dealer', ''))
                    for email in preferred:
                        item = login_map.get(email.lower())
                        if item and item not in ordered:
                            ordered.append(item)
                    if not ordered:
                        ordered.extend(logins)

                cert_data = None
                method = ''
                cert_name = ''
                last_error = 'Delivery Certificate not found'

                # Wrong-company logins simply fail the order lookup, so try the
                # preferred Excel login first and then the remaining Logininfo rows.
                for email, password in ordered:
                    try:
                        if email.lower() not in contexts:
                            contexts[email.lower()] = await _login_context(
                                browser, email, password, base, legacy_book
                            )
                        context, page = contexts[email.lower()]
                        cert_data, method, cert_name = await _capture_delivery_certificate(page, context, oid)
                        if cert_data:
                            break
                        last_error = method
                    except Exception as exc:
                        last_error = str(exc)

                if cert_data:
                    target = pdf_dir / f"{safe_filename(row.get('_pdf_name') or row.get('NAME') or oid)}.pdf"
                    _append_pdf(target, cert_data)
                    row['_delivery_certificate'] = 'Downloaded'
                    row['_delivery_certificate_method'] = method
                    row['_delivery_certificate_name'] = cert_name or 'Delivery Certificate'
                else:
                    row['_delivery_certificate'] = 'Missing'
                    row['_delivery_certificate_method'] = ''
                    warnings.append(f"{row.get('NAME') or oid}: Delivery Certificate: {last_error}")

            # Keep the fresh cookies for later runs, just as the Excel macro did.
            for email_key, (context, _page) in contexts.items():
                try:
                    await context.storage_state(path=str(base / 'auth_state' / f'shedsuite_auth_{_safe_email(email_key)}.json'))
                except Exception:
                    pass
        finally:
            for context, _page in contexts.values():
                try:
                    await context.close()
                except Exception:
                    pass
            await browser.close()

    return warnings


def append_delivery_certificates(rows: list[dict], source_rows: list[dict], pdf_dir: Path, base: Path) -> list[str]:
    """Synchronous Flask wrapper. ONLY Delivery Certificate uses Chromium."""
    return asyncio.run(_run(rows, source_rows, pdf_dir, base))
