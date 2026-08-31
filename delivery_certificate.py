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


def _candidate_legacy_workbooks(base: Path) -> list[Path]:
    """Return every nearby Contract_Import.xlsm in deterministic priority order.

    V7.20.1 stopped at the first workbook it found.  Users often have several
    historical ShedSuite app folders, so the newly-added login could live in a
    different copy.  V7.20.2 merges Logininfo across all nearby workbooks.
    """
    explicit = os.getenv('LEGACY_XLSM_PATH', '').strip()
    home = Path(os.environ.get('USERPROFILE') or Path.home())
    candidates: list[Path] = []
    if explicit:
        candidates.append(Path(explicit))

    fixed_roots = [base, base.parent, home / 'Downloads', home / 'Desktop', home / 'Documents']
    for root in fixed_roots:
        candidates += [
            root / 'Contract_Import.xlsm',
            root / 'Shedsuite_Import V1.1.7.6' / 'Contract_Import.xlsm',
            root / 'Shedsuite_Import V1.1.7.6' / 'Shedsuite_Import V1.1.7.6' / 'Contract_Import.xlsm',
        ]

    # Bounded search; gather all results rather than returning the first one.
    search_roots = [base.parent, base.parent.parent, home / 'Downloads', home / 'Desktop', home / 'Documents']
    seen_roots = set()
    for root in search_roots:
        try:
            root = root.resolve()
        except Exception:
            continue
        if root in seen_roots or not root.exists():
            continue
        seen_roots.add(root)
        try:
            root_depth = len(root.parts)
            for current, dirs, files in os.walk(root):
                current_path = Path(current)
                depth = len(current_path.parts) - root_depth
                if depth >= 4:
                    dirs[:] = []
                if 'Contract_Import.xlsm' in files:
                    candidates.append(current_path / 'Contract_Import.xlsm')
        except Exception:
            pass

    result: list[Path] = []
    seen = set()
    for candidate in candidates:
        try:
            resolved = candidate.resolve()
            key = str(resolved).casefold()
            if key in seen or not resolved.exists() or not resolved.is_file():
                continue
            seen.add(key)
            result.append(resolved)
        except Exception:
            continue
    return result


def _find_legacy_workbook(base: Path) -> Path | None:
    books = _candidate_legacy_workbooks(base)
    return books[0] if books else None


def _legacy_logins(base: Path) -> tuple[list[tuple[str, str]], Path | None, str]:
    """Merge Logininfo credentials from every nearby legacy workbook.

    First occurrence wins for duplicate email addresses, so an explicit
    LEGACY_XLSM_PATH / workbook beside the current app has highest priority.
    Passwords remain local and are held only in memory.
    """
    books = _candidate_legacy_workbooks(base)
    if not books:
        return [], None, ''

    merged: list[tuple[str, str]] = []
    seen_emails: set[str] = set()
    selected = ''
    primary_book: Path | None = None

    for book in books:
        try:
            wb = load_workbook(book, data_only=True, read_only=False, keep_vba=True)
        except Exception:
            continue
        try:
            if primary_book is None:
                primary_book = book
            if not selected and 'Import Tools' in wb.sheetnames:
                selected = str(wb['Import Tools']['D4'].value or '').strip()
            if 'RTOPro Data' not in wb.sheetnames:
                continue
            ws = wb['RTOPro Data']
            table = ws.tables.get('Logininfo')
            if not table:
                continue
            rows = list(ws[table.ref])
            for row in rows[1:]:
                email = str(row[0].value or '').strip()
                password = str(row[1].value or '')
                key = email.casefold()
                if email and password and key not in seen_emails:
                    seen_emails.add(key)
                    merged.append((email, password))
        finally:
            try:
                wb.close()
            except Exception:
                pass

    if selected:
        merged.sort(key=lambda x: 0 if x[0].lower() == selected.lower() else 1)
    return merged, primary_book, selected

def _preferred_login(company: str, dealer: str) -> list[str]:
    """Same company/login selection idea used by the old Excel macro."""
    text = f'{company} {dealer}'.lower()
    if 'crestwood storage barns' in text or 'crestwood' in text:
        return ['info@whiteriverrto.com']
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

    login_button = page.get_by_role('button', name=re.compile(r'log\s*in', re.I))
    email_box = page.get_by_label(re.compile(r'email', re.I))
    password_box = page.get_by_label(re.compile(r'password', re.I))
    try:
        needs_login = await login_button.is_visible()
    except Exception:
        needs_login = False
    if not needs_login:
        try:
            needs_login = await email_box.is_visible() and await password_box.is_visible()
        except Exception:
            needs_login = False

    if needs_login:
        # No prompt: this is exactly why we read the old workbook's Logininfo.
        try:
            await email_box.fill(email)
        except Exception:
            await page.locator('input[type="email"]').first.fill(email)
        try:
            await password_box.fill(password)
        except Exception:
            await page.locator('input[type="password"]').first.fill(password)
        try:
            await login_button.click()
        except Exception:
            await page.locator('button[type="submit"]').first.click()
        try:
            await expect(page).to_have_url(re.compile(r'https://app\.shedsuite\.com/rto/.*'), timeout=300000)
        except Exception:
            # ShedSuite occasionally lands on another authenticated app route.
            # A successful disappearance of the login fields is sufficient.
            await page.wait_for_timeout(2500)
        await context.storage_state(path=str(local_auth))

    return context, page


async def _capture_delivery_certificate(page, context, order_id: str) -> tuple[bytes | None, str, str]:
    """Capture only the Delivery Certificate from a ShedSuite order.

    V7.6.2 deliberately waits for ShedSuite's React attachment list.  Earlier
    V7.6.x builds inspected the DOM immediately after ``domcontentloaded`` and
    could therefore report a false Missing result while the Files list was
    still rendering.
    """
    order_url = BASE_URL + str(order_id)
    responses = []

    def remember(response):
        responses.append(response)

    async def login_screen_visible() -> bool:
        try:
            if '/login' in str(page.url).lower():
                return True
        except Exception:
            pass
        try:
            email = page.locator('input[type="email"]')
            password = page.locator('input[type="password"]')
            return (await email.count() > 0 and await password.count() > 0
                    and await email.first.is_visible() and await password.first.is_visible())
        except Exception:
            return False

    async def locate_delivery_target():
        candidate_selectors = [
            'span[class^="fileName"]',
            'span[class*="fileName"]',
            '[class*="fileName"]',
        ]
        for selector in candidate_selectors:
            try:
                locs = page.locator(selector)
                count = await locs.count()
                for i in range(count):
                    loc = locs.nth(i)
                    name = ((await loc.text_content()) or '').strip()
                    low = re.sub(r'\s+', ' ', name.lower())
                    if 'delivery' in low and ('certificate' in low or re.search(r'\bcert\b', low)):
                        return loc, name
            except Exception:
                continue

        # Fall back to text rather than generated CSS classes.  Exact=False is
        # intentional because ShedSuite sometimes adds a date/file extension.
        try:
            text_matches = page.get_by_text(
                re.compile(r'delivery\s*(certificate|cert(?:ificate)?)', re.I)
            )
            count = await text_matches.count()
            for i in range(count):
                loc = text_matches.nth(i)
                try:
                    if not await loc.is_visible():
                        continue
                except Exception:
                    pass
                name = ((await loc.text_content()) or '').strip()
                if name:
                    return loc, name
        except Exception:
            pass
        return None, ''

    async def wait_for_delivery_target(timeout_ms: int):
        """Poll because ShedSuite's Files section is populated after page load."""
        loop = asyncio.get_running_loop()
        deadline = loop.time() + (timeout_ms / 1000.0)
        files_seen_at = None
        last_body = ''
        while loop.time() < deadline:
            if await login_screen_visible():
                return None, '', 'ShedSuite session is not authenticated (login page was shown).'

            target, name = await locate_delivery_target()
            if target is not None:
                return target, name, ''

            # Once the attachment list itself is clearly present, allow a few
            # extra seconds for all rows to render before deciding the certificate
            # truly is absent.
            try:
                file_count = await page.locator('[class*="fileName"]').count()
            except Exception:
                file_count = 0
            if file_count:
                if files_seen_at is None:
                    files_seen_at = loop.time()
                elif loop.time() - files_seen_at >= 8:
                    try:
                        last_body = (await page.locator('body').inner_text()).lower()
                    except Exception:
                        last_body = ''
                    if 'delivery certificate' not in last_body and 'delivery cert' not in last_body:
                        return None, '', 'Delivery Certificate was not listed on the ShedSuite order page.'

            await page.wait_for_timeout(1000)

        try:
            last_body = (await page.locator('body').inner_text()).lower()
        except Exception:
            last_body = ''
        if 'delivery certificate' in last_body or 'delivery cert' in last_body:
            return None, '', 'Delivery Certificate text was present but its clickable element could not be located.'
        return None, '', 'Timed out waiting for ShedSuite to load the Delivery Certificate/file list.'

    try:
        # Try twice. A soft reload fixes occasional ShedSuite orders whose Files
        # request stalls even though the order page itself has loaded.
        target = None
        found_name = ''
        find_error = ''
        for attempt in range(2):
            if attempt == 0:
                await page.goto(order_url, wait_until='domcontentloaded', timeout=90000)
            else:
                try:
                    await page.reload(wait_until='domcontentloaded', timeout=90000)
                except Exception:
                    await page.goto(order_url, wait_until='domcontentloaded', timeout=90000)

            # Keep the historical 90-second tolerance on the first attempt. The
            # retry is shorter because it is only a recovery path.
            target, found_name, find_error = await wait_for_delivery_target(90000 if attempt == 0 else 30000)
            if target is not None:
                break
            # Authentication problems will not be repaired by refreshing.
            if 'not authenticated' in find_error.lower():
                return None, find_error, ''

        if target is None:
            return None, find_error or 'Delivery Certificate not found on the ShedSuite order page.', ''

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
            # Prefer the nearest clickable ancestor. Clicking a nested text span
            # does not always trigger the document viewer in newer ShedSuite UI.
            try:
                clickable = target.locator('xpath=ancestor-or-self::a[1] | ancestor-or-self::button[1] | ancestor-or-self::*[@role="button"][1]').first
                if await clickable.count() and await clickable.is_visible():
                    await clickable.click(timeout=15000)
                else:
                    await target.click(timeout=15000)
            except Exception:
                await target.click(timeout=15000)
        except Exception as exc:
            for task in (popup_task, download_task):
                if not task.done():
                    task.cancel()
            return None, f'click failed: {exc}', found_name

        done, pending = await asyncio.wait(
            [popup_task, download_task], timeout=20, return_when=asyncio.FIRST_COMPLETED
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
                    raw = temp_path.read_bytes()
                    converted = bytes_to_pdf(raw)
                    if converted:
                        return converted, 'browser download', found_name
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
        await page.wait_for_timeout(3000)
        for response in reversed(responses):
            try:
                ctype = str((response.headers or {}).get('content-type', '')).lower()
                url = str(response.url or '').lower()
                if 'pdf' in ctype or 'application/octet-stream' in ctype or 'delivery' in url or 'certificate' in url:
                    raw = await response.body()
                    if raw:
                        converted = bytes_to_pdf(raw)
                        if converted:
                            return converted, 'network response', found_name
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

        return None, 'Delivery Certificate was found and clicked, but no PDF bytes were captured.', found_name
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


async def _run(
    rows: list[dict],
    source_rows: list[dict],
    pdf_dir: Path,
    base: Path,
    progress=None,
    should_process=None,
    headless: bool = True,
    prefetch_cache_dir: Path | None = None,
    packet_ready_flag: Path | None = None,
    concurrency: int = 3,
) -> list[str]:
    """Download Delivery Certificates with session reuse and bounded concurrency.

    V7.15 can start this before the direct ShedSuite packet PDFs are finished.
    Certificates are fetched into a small local cache first; once the base-packet
    marker appears they are appended to the combined PDFs.  This lets Chromium do
    network/login work while the normal /process request is still assembling the
    faster direct-link documents.

    One authenticated context/page is reused per ShedSuite login.  A per-login
    asyncio lock keeps that page serial, while different company logins can run
    concurrently (default: 3 accounts at once).
    """
    logins, legacy_book, _selected = _legacy_logins(base)
    if prefetch_cache_dir:
        prefetch_cache_dir.mkdir(parents=True, exist_ok=True)

    async def wait_for_packet_ready(timeout_seconds: float = 600.0) -> bool:
        if packet_ready_flag is None:
            return True
        loop = asyncio.get_running_loop()
        deadline = loop.time() + timeout_seconds
        while loop.time() < deadline:
            if packet_ready_flag.exists():
                return True
            await asyncio.sleep(0.25)
        return packet_ready_flag.exists()

    if not logins:
        msg = (
            'Delivery Certificate: V7 could not find Contract_Import.xlsm / Logininfo. '
            'Put the old Contract_Import.xlsm beside V7 or set LEGACY_XLSM_PATH. '
            'V7 intentionally never asks for a ShedSuite password.'
        )
        for row in rows:
            oid = str(row.get('_source_order_id', '')).strip()
            if oid:
                row['_delivery_certificate'] = 'Missing'
                row['_delivery_certificate_method'] = ''
                if progress:
                    progress(oid, 'missing', 'Logininfo credentials were not found.', row)
        # In prefetch mode, do not let the app merge/rebuild while /process is
        # still constructing the base packets.
        await wait_for_packet_ready()
        return [msg]

    src_by_order = {str(x.get('Customer Order Id', '')).strip(): x for x in source_rows}
    login_map = {email.lower(): (email, password) for email, password in logins}
    warnings: list[str] = []

    try:
        concurrency = max(1, min(5, int(concurrency or 3)))
    except Exception:
        concurrency = 3

    async with async_playwright() as playwright:
        launch_kwargs = {'headless': headless}
        if not headless:
            launch_kwargs['args'] = [
                '--window-position=-32000,-32000',
                '--window-size=1200,900',
                '--disable-background-timer-throttling',
                '--disable-renderer-backgrounding',
                '--disable-backgrounding-occluded-windows',
                '--disable-features=CalculateNativeWinOcclusion',
            ]
        browser = await playwright.chromium.launch(**launch_kwargs)
        contexts: dict[str, tuple] = {}
        login_locks: dict[str, asyncio.Lock] = {}
        semaphore = asyncio.Semaphore(concurrency)
        pending_prefetch: dict[str, tuple[dict, Path, str, str]] = {}

        def ordered_logins_for(row: dict) -> list[tuple[str, str]]:
            oid = str(row.get('_source_order_id', '')).strip()
            src = src_by_order.get(oid, {})
            v3_login = str(row.get('_login') or '').strip()
            ordered: list[tuple[str, str]] = []
            if v3_login:
                item = login_map.get(v3_login.lower())
                if item:
                    ordered.append(item)
                return ordered

            preferred = _preferred_login(
                src.get('Company Name', '') or row.get('_source_company', ''),
                src.get('Dealer', '') or row.get('_source_dealer', ''),
            )
            for email in preferred:
                item = login_map.get(email.lower())
                if item and item not in ordered:
                    ordered.append(item)
            if not ordered:
                ordered.extend(logins)
            return ordered

        async def get_context(email: str, password: str):
            key = email.lower()
            if key not in contexts:
                contexts[key] = await _login_context(browser, email, password, base, legacy_book)
            return contexts[key]

        async def append_prefetched(row: dict, cache_path: Path, method: str, cert_name: str) -> bool:
            oid = str(row.get('_source_order_id', '')).strip()
            if should_process and not should_process(oid):
                try:
                    cache_path.unlink(missing_ok=True)
                except Exception:
                    pass
                return False
            try:
                cert_data = cache_path.read_bytes()
                target = pdf_dir / f"{safe_filename(row.get('_pdf_name') or row.get('NAME') or oid)}.pdf"
                _append_pdf(target, cert_data)
                try:
                    cache_path.unlink(missing_ok=True)
                except Exception:
                    pass
                row['_delivery_certificate'] = 'Downloaded'
                row['_delivery_certificate_method'] = method
                row['_delivery_certificate_name'] = cert_name or 'Delivery Certificate'
                if progress:
                    progress(oid, 'ready', f'Downloaded via {method}.', row)
                return True
            except Exception as exc:
                row['_delivery_certificate'] = 'Missing'
                row['_delivery_certificate_method'] = ''
                warnings.append(f"{row.get('NAME') or oid}: Delivery Certificate append failed: {exc}")
                if progress:
                    progress(oid, 'missing', f'Certificate was prefetched but could not be appended: {exc}', row)
                return False

        async def process_row(row: dict) -> None:
            oid = str(row.get('_source_order_id', '')).strip()
            if not oid:
                return
            if should_process and not should_process(oid):
                return

            ordered = ordered_logins_for(row)
            v3_login = str(row.get('_login') or '').strip()
            if v3_login and not ordered:
                warnings.append(
                    f"{row.get('NAME') or oid}: Delivery Certificate: "
                    f"V3 selected ShedSuite login {v3_login}, but that login was not found in any nearby Contract_Import.xlsm Logininfo table."
                )
                row['_delivery_certificate'] = 'Missing'
                row['_delivery_certificate_method'] = ''
                if progress:
                    progress(oid, 'missing', f'Credential {v3_login} was not found in any nearby Contract_Import.xlsm Logininfo table.', row)
                return

            if progress:
                progress(oid, 'working', 'Prefetching Delivery Certificate from ShedSuite…', row)

            cert_data = None
            method = ''
            cert_name = ''
            last_error = 'Delivery Certificate not found'

            # Different company logins can proceed at the same time. Rows that
            # share one login intentionally serialize on the same authenticated page.
            async with semaphore:
                for email, password in ordered:
                    if should_process and not should_process(oid):
                        return
                    key = email.lower()
                    lock = login_locks.setdefault(key, asyncio.Lock())
                    try:
                        async with lock:
                            context, page = await get_context(email, password)
                            cert_data, method, cert_name = await _capture_delivery_certificate(page, context, oid)
                        if cert_data:
                            break
                        last_error = method
                    except Exception as exc:
                        last_error = str(exc)

            if not cert_data:
                row['_delivery_certificate'] = 'Missing'
                row['_delivery_certificate_method'] = ''
                warnings.append(f"{row.get('NAME') or oid}: Delivery Certificate: {last_error}")
                if progress:
                    progress(oid, 'missing', last_error, row)
                return

            if prefetch_cache_dir is None:
                target = pdf_dir / f"{safe_filename(row.get('_pdf_name') or row.get('NAME') or oid)}.pdf"
                _append_pdf(target, cert_data)
                row['_delivery_certificate'] = 'Downloaded'
                row['_delivery_certificate_method'] = method
                row['_delivery_certificate_name'] = cert_name or 'Delivery Certificate'
                if progress:
                    progress(oid, 'ready', f'Downloaded via {method}.', row)
                return

            cache_path = prefetch_cache_dir / f'{safe_filename(oid)}.pdf'
            temp = cache_path.with_suffix('.tmp')
            temp.write_bytes(cert_data)
            temp.replace(cache_path)
            row['_delivery_certificate'] = 'Prefetched'
            row['_delivery_certificate_method'] = method
            row['_delivery_certificate_name'] = cert_name or 'Delivery Certificate'

            # If the direct packet has already finished, append now. Otherwise
            # keep fetching later contracts rather than blocking this login on a
            # file-system wait; the second phase appends all cached certificates.
            if packet_ready_flag is None or packet_ready_flag.exists():
                await append_prefetched(row, cache_path, method, cert_name)
            else:
                pending_prefetch[oid] = (row, cache_path, method, cert_name)
                if progress:
                    progress(oid, 'prefetched', 'Certificate prefetched; waiting for the base contract packet…', row)

        try:
            tasks = [asyncio.create_task(process_row(row)) for row in rows]
            if tasks:
                await asyncio.gather(*tasks)

            if pending_prefetch:
                ready = await wait_for_packet_ready()
                if not ready:
                    for oid, (row, cache_path, _method, _cert_name) in list(pending_prefetch.items()):
                        if should_process and not should_process(oid):
                            continue
                        row['_delivery_certificate'] = 'Missing'
                        row['_delivery_certificate_method'] = ''
                        msg = 'Delivery Certificate was prefetched, but the base packet never became ready.'
                        warnings.append(f"{row.get('NAME') or oid}: {msg}")
                        if progress:
                            progress(oid, 'missing', msg, row)
                else:
                    for oid, (row, cache_path, method, cert_name) in list(pending_prefetch.items()):
                        await append_prefetched(row, cache_path, method, cert_name)

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

def append_delivery_certificates(
    rows: list[dict],
    source_rows: list[dict],
    pdf_dir: Path,
    base: Path,
    progress=None,
    should_process=None,
    headless: bool = True,
    prefetch_cache_dir: Path | None = None,
    packet_ready_flag: Path | None = None,
    concurrency: int = 3,
) -> list[str]:
    """Synchronous wrapper used by the V7 background worker."""
    return asyncio.run(_run(
        rows, source_rows, pdf_dir, base,
        progress=progress, should_process=should_process, headless=headless,
        prefetch_cache_dir=prefetch_cache_dir, packet_ready_flag=packet_ready_flag,
        concurrency=concurrency,
    ))
