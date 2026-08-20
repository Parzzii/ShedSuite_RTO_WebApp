from __future__ import annotations

import io
import re
from pathlib import Path

import fitz
import requests
from PIL import Image, ImageOps

LETTER_WIDTH = 612
LETTER_HEIGHT = 792


def safe_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', '_', name).strip().strip('.')
    return name[:180] or 'contract'


def _insert_image_page(output_pdf, content: bytes):
    img = Image.open(io.BytesIO(content))
    img = ImageOps.exif_transpose(img)
    if img.mode not in ('RGB', 'L'):
        img = img.convert('RGB')
    bio = io.BytesIO()
    img.save(bio, format='JPEG', quality=90)
    w, h = img.size
    page = output_pdf.new_page(width=LETTER_WIDTH, height=LETTER_HEIGHT)
    scale = min(LETTER_WIDTH / w, LETTER_HEIGHT / h)
    nw, nh = w * scale, h * scale
    x, y = (LETTER_WIDTH - nw) / 2, (LETTER_HEIGHT - nh) / 2
    page.insert_image(fitz.Rect(x, y, x + nw, y + nh), stream=bio.getvalue())


def bytes_to_pdf(data: bytes) -> bytes:
    """Normalize authenticated ShedSuite document bytes to a valid PDF.

    Added for V6 Delivery Certificate capture only.  The existing V3 direct
    download/combine behavior below is intentionally unchanged.
    """
    if not data:
        raise ValueError('Empty document bytes')
    if data[:5] == b'%PDF-':
        doc = fitz.open(stream=data, filetype='pdf')
        try:
            return doc.tobytes()
        finally:
            doc.close()

    out = fitz.open()
    try:
        _insert_image_page(out, data)
        return out.tobytes()
    except Exception as e:
        raise ValueError(f'Unsupported document type: {e}') from e
    finally:
        out.close()


def download_and_combine(urls: list[str], target: Path, timeout: int = 45) -> list[str]:
    target.parent.mkdir(parents=True, exist_ok=True)
    out = fitz.open()
    errors: list[str] = []
    seen: set[str] = set()
    for url in urls:
        url = (url or '').strip()
        if not url or url in seen:
            continue
        seen.add(url)
        try:
            resp = requests.get(url, timeout=timeout)
            resp.raise_for_status()
            content = resp.content
            ctype = (resp.headers.get('Content-Type') or '').lower()
            if 'pdf' in ctype or content[:4] == b'%PDF':
                d = fitz.open(stream=content, filetype='pdf')
                out.insert_pdf(d)
                d.close()
            elif ctype.startswith('image/'):
                _insert_image_page(out, content)
            else:
                try:
                    d = fitz.open(stream=content)
                    out.insert_pdf(d)
                    d.close()
                except Exception:
                    _insert_image_page(out, content)
        except Exception as e:
            errors.append(f'{url}: {e}')
    if len(out) > 0:
        out.save(target)
    out.close()
    return errors


def _next_value(lines: list[str], index: int) -> str:
    for j in range(index + 1, min(len(lines), index + 5)):
        value = lines[j].strip()
        if value:
            return value
    return ''


def _clean_person(value: str) -> str:
    value = re.sub(r'\s+', ' ', (value or '').strip())
    # Guard against labels/fields being mistaken for a name.
    bad = {
        'description', 'quantity', 'total', 'work order', 'work order#', 'building',
        'invoice', 'invoice #', 'order type', 'date', 'salesman', 'agent'
    }
    if value.casefold() in bad:
        return ''
    if re.match(r'^(?:work order|building|invoice|order type|description|quantity|total|date)\b', value, re.I):
        return ''
    if len(value) > 80:
        return ''
    return value


def extract_pdf_fields(path: Path) -> dict[str, str]:
    """Extract the same fields used by the workbook PDF Data sheet.

    The original Python reads the invoice's Agent field.  ShedSuite invoice layouts
    vary, so this accepts "Agent:", "Agent", "Salesman:", and same-line values.
    """
    fields = {
        'salesman': '',
        'coordinates': '',
        'latitude': '',
        'longitude': '',
        'invoice': '',
    }
    if not path.exists():
        return fields

    doc = fitz.open(path)
    try:
        pages: list[tuple[list[str], str]] = []
        for page in doc:
            lines = [re.sub(r'\s+', ' ', x).strip() for x in page.get_text().splitlines()]
            lines = [x for x in lines if x]
            pages.append((lines, '\n'.join(lines)))

        # Coordinates can occur on a delivery/happy sheet near the end of the PDF.
        for lines, text in pages:
            if not fields['coordinates']:
                m = re.search(
                    r'lat\s*:\s*([-+]?\d+(?:\.\d+)?)\s*,?\s*lon\s*:\s*([-+]?\d+(?:\.\d+)?)',
                    text,
                    re.I,
                )
                if m:
                    fields['latitude'], fields['longitude'] = m.group(1), m.group(2)
                    fields['coordinates'] = f'Lat: {m.group(1)} Lon: {m.group(2)}'

        # Prefer an Agent field on an invoice-like page, which mirrors extract_agent()
        # from the original Contract_Import.py.
        candidate_pages = sorted(
            pages,
            key=lambda p: (0 if re.search(r'Invoice\s*#|Invoice From:', p[1], re.I) else 1),
        )
        for lines, text in candidate_pages:
            if fields['salesman']:
                break
            for i, line in enumerate(lines):
                m = re.match(r'^(?:Agent|Salesman)\s*:\s*(.+)$', line, re.I)
                if m:
                    value = _clean_person(m.group(1))
                    if value:
                        fields['salesman'] = value
                        break
                if re.match(r'^(?:Agent|Salesman)\s*:?\s*$', line, re.I):
                    value = _clean_person(_next_value(lines, i))
                    if value:
                        fields['salesman'] = value
                        break

        # Invoice number: first explicit invoice label/value wins.
        for lines, text in candidate_pages:
            if fields['invoice']:
                break
            for i, line in enumerate(lines):
                m = re.match(r'^Invoice\s*#\s*:?\s*(\S.+)$', line, re.I)
                if m:
                    value = m.group(1).strip()
                    if value:
                        fields['invoice'] = value
                        break
                if re.match(r'^INVOICE\s*#\s*:?\s*$', line, re.I):
                    value = _next_value(lines, i)
                    if value:
                        fields['invoice'] = value
                        break
    finally:
        doc.close()
    return fields
