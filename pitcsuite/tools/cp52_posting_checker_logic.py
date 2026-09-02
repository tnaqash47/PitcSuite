from __future__ import annotations

import os
import re
import threading
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import END, BOTH, LEFT, RIGHT, X, Y, StringVar, Tk, filedialog, messagebox
from tkinter import ttk

from openpyxl import load_workbook
from pypdf import PdfReader, PdfWriter
from reportlab.lib.colors import Color, black
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
import pdfplumber


ACCOUNT_DIGITS = re.compile(r"(?<![0-9])[0-9]{13,15}(?![0-9])")
ACCOUNT_DIGITS_BYTES = re.compile(rb"(?<![0-9])[0-9]{13,15}(?![0-9])")
ACCOUNT_LINE = re.compile(r"(?m)^\s*[0-9]{13,15}\s*$")
AMOUNT_TOKEN = re.compile(r"(?<![0-9])-?[0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{1,2})?(?![0-9])|-?(?<![0-9])[0-9]+(?:\.[0-9]{1,2})?(?![0-9])")
CODE_TOKEN = re.compile(
    r"(?<![A-Z0-9])(?:[0-9]{1,4}\s*[CBF]|[BF]\s*[0-9]{2,})(?![A-Z0-9])",
    re.IGNORECASE,
)
HIGHLIGHT = Color(1.0, 0.88, 0.15, alpha=0.40)
AMOUNT_LAST_DIGIT_TOLERANCE = 9
CONTINUATION_LINES = 6
BOUNDARY_CONTEXT_RECORDS = 3


@dataclass
class SourcePage:
    number: int
    text: str
    source_index: int
    source_path: Path
    compact_output: bool = False
    visual_rows: list[list[str]] | None = None
    text_offset: int = 0


@dataclass
class MatchResult:
    page: SourcePage | None
    account_hits: list[tuple[int, int]]
    code_hits: list[tuple[int, int]]
    amount_hits: list[tuple[int, int]]
    amount_found: bool
    amount_exact: bool
    code: str = ""
    amount_values: tuple[int, ...] = ()
    code_values: tuple[str, ...] = ()
    visual_account_rows: tuple[int, ...] = ()


def compact_digits(value: object) -> str:
    digits = re.sub(r"[^0-9]", "", "" if value is None else str(value))
    # Excel often drops a leading zero when a fixed-width AC number is entered as a number.
    return digits.zfill(14) if len(digits) == 13 else digits


def dashed_account(value: object) -> str:
    digits = compact_digits(value)
    if len(digits) < 9:
        return ""
    return f"{digits[:2]}-{digits[2:-7]}-{digits[-7:]}"


def numeric_value(value: object) -> float | None:
    if value is None or str(value).strip() == "":
        return None
    try:
        return round(float(str(value).replace(",", "").strip()), 2)
    except ValueError:
        return None


def amount_key(value: object) -> int | None:
    """Return a signed whole amount, ignoring commas and decimal digits."""
    if value is None or str(value).strip() == "":
        return None
    try:
        return int(Decimal(str(value).strip().replace(",", "")))
    except (InvalidOperation, ValueError):
        return None


def amount_matches(token: str, wanted: object) -> bool:
    source = amount_key(token)
    target = amount_key(wanted)
    return source is not None and target is not None and abs(source - target) <= AMOUNT_LAST_DIGIT_TOLERANCE


def code_letter(value: str) -> str:
    compact = re.sub(r"\s+", "", value).upper()
    if compact.endswith(("C", "B", "F")) and compact[:-1].isdigit():
        return compact[-1]
    if compact.startswith(("B", "F")) and len(compact) >= 3:
        return compact[0]
    return ""


def code_matches(code: str, code_filter: str) -> bool:
    selected = code_filter.upper().replace("CODE-", "").replace("CODES", "ALL").strip()
    return selected == "ALL" or code == selected


def line_bounds(text: str) -> list[tuple[int, int, str]]:
    """Return absolute offsets and contents for each displayed text line."""
    result = []
    start = 0
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    for line in normalized.split("\n"):
        result.append((start, start + len(line), line))
        start += len(line) + 1
    return result


def line_number_at(bounds: list[tuple[int, int, str]], offset: int) -> int:
    return next(
        (number for number, (start, end, _) in enumerate(bounds) if start <= offset <= end),
        max(len(bounds) - 1, 0),
    )


def _amount_tokens_for_line(line: str, line_start: int) -> list[tuple[int, int, str]]:
    return [(line_start + m.start(), line_start + m.end(), m.group()) for m in AMOUNT_TOKEN.finditer(line)]


def _whole_digits(value: str) -> str:
    return re.sub(r"[^0-9]", "", value.split(".", 1)[0])


def record_total_spans(text: str) -> list[tuple[int, int]]:
    """Find the TOTAL AMOUNT value on an adjustment-ID row."""
    for line_start, _, line in line_bounds(text):
        if not re.match(r"^\s*\d{7,9}\b", line):
            continue
        tokens = _amount_tokens_for_line(line, line_start)
        useful = [item for item in tokens if len(_whole_digits(item[2])) < 7]
        useful = [item for item in useful if item[2].replace(",", "") not in {"0", "00", "01"}]
        if useful:
            return [(useful[-1][0], useful[-1][1])]
    return []


def record_amount_candidates(text: str) -> list[tuple[int, int, str]]:
    """Return amount tokens while excluding account/code/date metadata."""
    candidates = []
    for line_start, _, line in line_bounds(text):
        stripped = line.strip()
        if not stripped or re.fullmatch(r"\d+", stripped) or re.match(r"^\d{1,2}/\d{1,2}/\d{4}", stripped):
            continue
        if stripped.upper().startswith(("MTR INF", "ERRORS")):
            continue
        code_ranges = [(m.start(), m.end()) for m in CODE_TOKEN.finditer(line)]
        for start, end, token in _amount_tokens_for_line(line, line_start):
            local_start, local_end = start - line_start, end - line_start
            if any(local_start < code_end and local_end > code_start for code_start, code_end in code_ranges):
                continue
            if len(re.sub(r"[^0-9]", "", token)) >= 10:
                continue
            candidates.append((start, end, token))
    return candidates


def choose_amount_spans(text: str, wanted: object | None) -> tuple[list[tuple[int, int]], bool, tuple[int, ...]]:
    total_spans = record_total_spans(text)
    if amount_key(wanted) is None:
        values = tuple(amount_key(text[start:end]) for start, end in total_spans)
        return total_spans, False, tuple(value for value in values if value is not None)
    wanted_key = amount_key(wanted)
    preferred = [(start, end, text[start:end]) for start, end in total_spans]
    pools = [preferred, record_amount_candidates(text)] if preferred else [record_amount_candidates(text)]
    scored = []
    for pool in pools:
        for start, end, token in pool:
            value = amount_key(token)
            if value is None:
                continue
            difference = abs(value - wanted_key)
            if difference <= AMOUNT_LAST_DIGIT_TOLERANCE:
                scored.append((difference, start, end, value))
        if scored:
            break
    if not scored:
        return [], False, ()
    closest = min(item[0] for item in scored)
    chosen = [(start, end) for difference, start, end, _ in scored if difference == closest]
    values = tuple(value for difference, _, _, value in scored if difference == closest)
    return chosen, closest == 0, values


def amount_spans(text: str, wanted: object) -> tuple[list[tuple[int, int]], bool]:
    spans, exact, _ = choose_amount_spans(text, wanted)
    return spans, exact


def _find_code(text: str) -> tuple[str, list[tuple[int, int]]]:
    # Search only the lines directly below the account; this avoids headers
    # such as F-TAX being interpreted as a posting code.
    for start, _, line in line_bounds(text)[:10]:
        for match in CODE_TOKEN.finditer(line):
            code = code_letter(match.group())
            if code:
                normalized = re.sub(r"\s+", "", match.group()).upper()
                return normalized, [(start + match.start(), start + match.end())]
    return "", []


def pdf_word_rows(words: list[dict]) -> list[list[dict]]:
    """Group pdfplumber words into visual rows from top to bottom."""
    rows: list[list[dict]] = []
    for word in sorted(words, key=lambda item: (float(item["top"]), float(item["x0"]))):
        height = float(word["bottom"]) - float(word["top"])
        tolerance = max(3.0, height * 0.6)
        if rows and abs(float(word["top"]) - float(rows[-1][0]["top"])) <= tolerance:
            rows[-1].append(word)
        else:
            rows.append([word])
    return rows


def split_text_pages(path: Path, account_keys: set[str] | None = None) -> list[SourcePage]:
    raw = path.read_bytes()
    if account_keys is not None:
        found_keys = {compact_digits(match.group().decode("ascii")) for match in ACCOUNT_DIGITS_BYTES.finditer(raw)}
        if not found_keys.intersection(account_keys):
            return []
    # Keep matching and PDF rendering on the same character offsets. Windows
    # posting exports are commonly CRLF; retaining the CR while line_bounds()
    # uses LF would shift every later highlight span on the page.
    text = raw.decode("utf-8", errors="replace").replace("\r\n", "\n").replace("\r", "\n")
    chunks = text.split("\f")
    if len(chunks) > 1:
        return [SourcePage(i + 1, chunk, i, path) for i, chunk in enumerate(chunks) if chunk.strip()]
    starts = [m.start() for m in re.finditer(r"WAPDA\s+ELECTRICITY", text, re.IGNORECASE)]
    if not starts:
        return [SourcePage(1, text, 0, path)]
    pages = []
    for i, start in enumerate(starts):
        end = starts[i + 1] if i + 1 < len(starts) else len(text)
        pages.append(SourcePage(i + 1, text[start:end], i, path))
    return pages


def pdf_pages(path: Path, page_indices: set[int] | None = None) -> list[SourcePage]:
    reader = PdfReader(str(path))
    result = []
    selected = range(len(reader.pages)) if page_indices is None else sorted(page_indices)
    if not selected:
        return result
    with pdfplumber.open(str(path)) as pdf:
        for i in selected:
            page = reader.pages[i]
            words = pdf.pages[i].extract_words()
            rows = [[word["text"] for word in row] for row in pdf_word_rows(words)]
            result.append(SourcePage(i + 1, page.extract_text() or "", i, path, False, rows))
    return result


def _first_continuation_lines(next_text: str) -> str:
    """Get the six posting lines that continue a page-ending account."""
    lines = next_text.replace("\r\n", "\n").replace("\r", "\n").splitlines(keepends=True)
    ref_starts = [i for i, line in enumerate(lines) if re.match(r"^\s*(?:\d+\s+)?\d{2}-\d", line)]
    if ref_starts:
        start = ref_starts[0]
        if start > 0 and re.fullmatch(r"\s*\d+\s*", lines[start - 1]):
            start -= 1
    else:
        code_starts = [i for i, line in enumerate(lines) if CODE_TOKEN.search(line)]
        if not code_starts:
            return ""
        start = max(code_starts[0] - 1, 0)
    return "".join(lines[start : start + CONTINUATION_LINES])


def _continuation_page(pages: list[SourcePage], page_index: int, account_end: int) -> SourcePage | None:
    page = pages[page_index]
    if page.text[account_end:].strip() or page_index + 1 >= len(pages):
        return None
    next_page = pages[page_index + 1]
    if next_page.source_path != page.source_path or next_page.source_index != page.source_index + 1:
        return None
    continuation = _first_continuation_lines(next_page.text)
    if not continuation.strip():
        return None
    page_text = page.text.replace("\r\n", "\n").replace("\r", "\n")
    lines = page_text.splitlines(keepends=True)
    bounds = line_bounds(page_text)
    target_line = line_number_at(bounds, max(account_end - 1, 0))
    record_starts = [
        index for index, line in enumerate(lines)
        if re.fullmatch(r"\s*\d{13,15}\s*\n?", line)
    ]
    target_position = next(
        (position for position, line_number in enumerate(record_starts) if line_number == target_line),
        None,
    )
    if target_position is None:
        start_line = max(0, target_line - 24)
    else:
        # Keep the tail of the old page (the target plus three preceding
        # records) and append the six real posting lines from the next page.
        start_line = record_starts[max(0, target_position - BOUNDARY_CONTEXT_RECORDS)]
    crop_offset = sum(len(line) for line in lines[:start_line])
    cropped_text = "".join(lines[start_line : target_line + 1]).rstrip("\n")
    combined_text = cropped_text + "\n" + continuation.lstrip("\r\n")
    return SourcePage(page.number, combined_text, page.source_index, page.source_path, True, None, crop_offset)


def _account_hits(page: SourcePage, dashed: str, digits: str) -> list[tuple[int, int]]:
    # Posting-list TXT records normally print the real AC as a compact
    # 13/14-digit value, then repeat a dashed reference number on the next
    # line.  The dashed reference is not a second occurrence of the account;
    # including it makes the renderer highlight the wrong record block.
    compact_hits = [
        (match.start(), match.end())
        for match in re.finditer(re.escape(digits), page.text, re.IGNORECASE)
    ]
    if compact_hits:
        return sorted(set(compact_hits))
    # Keep support for older/source files that contain only the dashed form.
    return sorted(
        {
            (match.start(), match.end())
            for match in re.finditer(re.escape(dashed), page.text, re.IGNORECASE)
        }
    )


def build_account_index(pages: list[SourcePage]) -> dict[str, list[int]]:
    """Index pages once so every workbook row does not rescan the full archive."""
    index: dict[str, list[int]] = {}
    for page_number, page in enumerate(pages):
        keys = {compact_digits(match.group()) for match in ACCOUNT_DIGITS.finditer(page.text)}
        for row in page.visual_rows or []:
            for word in row:
                value = word.replace(" ", "")
                if re.fullmatch(r"\d{13,15}", value):
                    keys.add(compact_digits(value))
        for key in keys:
            index.setdefault(key, []).append(page_number)
    return index


def _page_account_keys(page: SourcePage) -> set[str]:
    keys = {compact_digits(match.group()) for match in ACCOUNT_DIGITS.finditer(page.text)}
    for row in page.visual_rows or []:
        for word in row:
            value = word.replace(" ", "")
            if re.fullmatch(r"\d{13,15}", value):
                keys.add(compact_digits(value))
    return keys


def _pdf_candidate_indices(path: Path, account_keys: set[str]) -> set[int]:
    """Use fast PDF text extraction to locate pages before visual parsing."""
    reader = PdfReader(str(path))
    page_texts = [page.extract_text() or "" for page in reader.pages]
    indices = set()
    for index, text in enumerate(page_texts):
        if _page_account_keys(SourcePage(index + 1, text, index, path)).intersection(account_keys):
            indices.add(index)
            last_line = next((line.strip() for line in reversed(text.splitlines()) if line.strip()), "")
            if compact_digits(last_line) in account_keys and re.fullmatch(r"\d{13,15}", last_line) and index + 1 < len(page_texts):
                indices.add(index + 1)
    return indices


def _load_relevant_pages(source_files: list[Path], account_keys: set[str], log) -> list[SourcePage]:
    """Load only pages that contain a workbook AC or its page continuation."""
    pages = []
    loaded_count = 0
    for file_number, source_path in enumerate(source_files, start=1):
        log(f"Loading {file_number}/{len(source_files)}: {source_path.name}")
        if source_path.suffix.lower() == ".pdf":
            file_pages = pdf_pages(source_path, _pdf_candidate_indices(source_path, account_keys))
        else:
            file_pages = split_text_pages(source_path, account_keys)
        keep_indices = set()
        for index, page in enumerate(file_pages):
            if _page_account_keys(page).intersection(account_keys):
                keep_indices.add(index)
                # A page-ending AC may have its code/amount on the next page.
                last_line = next((line.strip() for line in reversed(page.text.splitlines()) if line.strip()), "")
                if compact_digits(last_line) in account_keys and re.fullmatch(r"\d{13,15}", last_line) and index + 1 < len(file_pages):
                    keep_indices.add(index + 1)
        selected = [page for index, page in enumerate(file_pages) if index in keep_indices]
        pages.extend(selected)
        loaded_count += len(selected)
        log(f"  retained {len(selected)} relevant page(s) from {source_path.name}")
    log(f"Retained {loaded_count} relevant source page(s); building account index...")
    return pages


def _visual_total_values(rows: list[list[str]], account_row: int) -> tuple[int, ...]:
    """Read the total amount from the visual adjustment-ID row in a PDF."""
    for row in rows[account_row + 1 : account_row + 10]:
        words = [word.strip() for word in row]
        if not any(re.fullmatch(r"\d{7,9}", word) for word in words):
            continue
        values = []
        for word in words:
            value = amount_key(word)
            if value is None or len(_whole_digits(word)) >= 7 or word.replace(",", "") in {"0", "00", "01"}:
                continue
            values.append(value)
        if values:
            return (values[-1],)
    return ()


def _find_visual_pdf_matches(
    page: SourcePage, raw_account: object, amount: object, code_filter: str
) -> tuple[list[MatchResult], bool]:
    """Match PDF records using visual rows, which preserve spaced codes."""
    dashed = dashed_account(raw_account)
    digits = compact_digits(raw_account)
    # CP-52 PDF rows usually print the AC as a compact 14-digit value. The
    # dashed form also appears on the reference row, so do not use it here or
    # that reference would look like a second posting occurrence.
    targets = {digits}
    account_rows = {
        row_number
        for row_number, row in enumerate(page.visual_rows or [])
        if any(word.replace(" ", "") in targets for word in row)
    }
    matches = []
    has_posting_occurrence = False
    rows = page.visual_rows or []
    for account_row in sorted(account_rows):
        code = ""
        for row in rows[account_row : account_row + 10]:
            row_code, _ = _find_code(" ".join(row))
            if row_code:
                code = row_code
                break
        if not code:
            continue
        has_posting_occurrence = True
        if not code_matches(code_letter(code), code_filter):
            continue
        total_values = _visual_total_values(rows, account_row)
        wanted_key = amount_key(amount)
        if wanted_key is None:
            amount_found, amount_exact, values = bool(total_values), False, total_values
        else:
            candidates = [(abs(value - wanted_key), value) for value in total_values if abs(value - wanted_key) <= AMOUNT_LAST_DIGIT_TOLERANCE]
            if not candidates:
                continue
            closest = min(difference for difference, _ in candidates)
            values = tuple(value for difference, value in candidates if difference == closest)
            amount_found, amount_exact = True, closest == 0
        matches.append(
            MatchResult(
                page, [], [], [], amount_found, amount_exact, code, values, (code,), (account_row,)
            )
        )
    return matches, has_posting_occurrence


def find_matches(
    pages: list[SourcePage], raw_account: object, amount: object, code_filter: str = "ALL",
    candidate_pages: list[int] | None = None,
) -> list[MatchResult]:
    """Return every matching account occurrence across every source page."""
    dashed = dashed_account(raw_account)
    digits = compact_digits(raw_account)
    if not dashed:
        return []
    wanted_amount = amount if amount_key(amount) is not None else None
    if candidate_pages is None:
        candidate_pages = build_account_index(pages).get(digits, [])
    matches = []
    for page_index in candidate_pages:
        page = pages[page_index]
        if page.visual_rows is not None:
            visual_matches, has_posting_occurrence = _find_visual_pdf_matches(page, raw_account, amount, code_filter)
            if has_posting_occurrence:
                matches.extend(visual_matches)
                continue
        hits = _account_hits(page, dashed, digits)
        if not hits:
            continue
        bounds = line_bounds(page.text)
        for hit_index, (account_start, account_end) in enumerate(hits):
            account_line = line_number_at(bounds, account_start)
            account_line_start = bounds[account_line][0]
            next_account_start = None
            if hit_index + 1 < len(hits):
                next_line = line_number_at(bounds, hits[hit_index + 1][0])
                next_account_start = bounds[next_line][0]
            block_end = next_account_start if next_account_start is not None else len(page.text)
            output_page = page
            if block_end == len(page.text):
                continued = _continuation_page(pages, page_index, account_end)
                if continued is not None:
                    output_page = continued
                    account_start -= output_page.text_offset
                    account_end -= output_page.text_offset
                    account_line_start -= output_page.text_offset
                    block_end = len(output_page.text)
            block_text = output_page.text[account_line_start:block_end]
            code, relative_code_hits = _find_code(block_text)
            # PDF exports often repeat every AC number in a compact index at
            # the top of the page. Those index entries have no code row and
            # are not posting occurrences.
            if not code:
                continue
            if not code_matches(code_letter(code), code_filter):
                continue
            amount_hits, amount_exact, values = choose_amount_spans(block_text, wanted_amount)
            # With an amount in Excel, only an occurrence matching that amount
            # (including the permitted +/- 1..9 final-digit range) is output.
            if wanted_amount is not None and not amount_hits:
                continue
            block_offset = account_line_start
            matches.append(
                MatchResult(
                    output_page,
                    [(account_start, account_end)],
                    [(block_offset + start, block_offset + end) for start, end in relative_code_hits],
                    [(block_offset + start, block_offset + end) for start, end in amount_hits],
                    wanted_amount is None or bool(amount_hits),
                    amount_exact,
                    code,
                    values,
                    (code,) if code else (),
                )
            )
    return matches


def find_match(pages: list[SourcePage], raw_account: object, amount: object) -> MatchResult:
    """Backward-compatible first-match wrapper for callers outside the UI."""
    matches = find_matches(pages, raw_account, amount, "ALL")
    return matches[0] if matches else MatchResult(None, [], [], [], False, False)


def add_sr_overlay(page, sr: object, output: Path) -> None:
    if sr is None or str(sr).strip() == "":
        return
    writer = PdfWriter()
    with open(output, "rb") as f:
        source = PdfReader(f)
        page_obj = source.pages[0]
        width = float(page_obj.mediabox.width)
        height = float(page_obj.mediabox.height)
        overlay_path = output.with_suffix(".overlay.pdf")
        c = canvas.Canvas(str(overlay_path), pagesize=(width, height))
        label = str(sr).strip()
        c.setFont("Helvetica-Bold", 16)
        text_width = c.stringWidth(label, "Helvetica-Bold", 16)
        x = width - 24 - text_width
        y = height - 28
        c.setFillColor(black)
        c.drawString(x, y, label)
        c.setLineWidth(1.2)
        c.line(x, y - 2, x + text_width, y - 2)
        c.save()
        overlay = PdfReader(str(overlay_path)).pages[0]
        page_obj.merge_page(overlay)
        writer.add_page(page_obj)
        with open(output, "wb") as out:
            writer.write(out)
    overlay_path.unlink(missing_ok=True)


def render_text_page(page: SourcePage, result: MatchResult, sr: object, output: Path) -> None:
    # Posting-list TXT output is intended for landscape A4 printing.
    page_width, page_height = landscape(A4)
    margin_x, margin_y = (6, 8) if page.compact_output else (12, 18)
    lines = page.text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    max_len = max((len(line) for line in lines), default=1)
    width_font_size = (page_width - 2 * margin_x) / (max_len * 0.60)
    # A boundary page can arrive from an older/serialized SourcePage without
    # its marker. A long page is unambiguously the same fit-to-A4 case.
    fit_to_a4 = page.compact_output or len(lines) > 35
    if fit_to_a4:
        # Boundary output retains the full previous page plus the next
        # heading/record. Condense only this combined page so no prior AC
        # rows are lost and the continuation amount still fits.
        margin_x, margin_y = 6, 8
        width_font_size = (page_width - 2 * margin_x) / (max_len * 0.60)
        available_height = page_height - margin_y - 6
        leading = min(12.0, available_height / max(len(lines), 1))
        # Use more of the A4 canvas while preserving the fixed-width column
        # alignment of the posting export.
        font_size = min(11.0, width_font_size, max(3.0, leading * 0.88))
    else:
        font_size = min(8.5, max(4.5, width_font_size))
        leading = min(8.2, max(5.1, (page_height - 2 * margin_y) / max(len(lines), 1)))
    c = canvas.Canvas(str(output), pagesize=(page_width, page_height))
    c.setFont("Courier", font_size)
    for line_no, line in enumerate(lines):
        y = page_height - margin_y - (line_no + 1) * leading
        if y < 10 and not fit_to_a4:
            break
        spans = []
        line_start = sum(len(x) + 1 for x in lines[:line_no])
        line_end = line_start + len(line)
        for start, end in result.account_hits + result.code_hits + result.amount_hits:
            if start < line_end and end > line_start:
                local_start, local_end = max(start, line_start) - line_start, min(end, line_end) - line_start
                spans.append((local_start, local_end))
        for start, end in spans:
            x = margin_x + start * font_size * 0.60
            w = max((end - start) * font_size * 0.60, 3)
            c.setFillColor(HIGHLIGHT)
            c.rect(x - 1, y - 1, w + 2, leading, stroke=0, fill=1)
        c.setFillColor(black)
        c.drawString(margin_x, y, line)
    if sr is not None and str(sr).strip():
        label = str(sr).strip()
        c.setFont("Helvetica-Bold", 16)
        width = c.stringWidth(label, "Helvetica-Bold", 16)
        x, y = page_width - 24 - width, page_height - 28
        c.drawString(x, y, label)
        c.setLineWidth(1.2)
        c.line(x, y - 2, x + width, y - 2)
    c.save()


def render_pdf_page(source_path: Path, page_number: int, result: MatchResult, account: object, sr: object, output: Path) -> None:
    reader = PdfReader(str(source_path))
    original = reader.pages[page_number - 1]
    width, height = float(original.mediabox.width), float(original.mediabox.height)
    with pdfplumber.open(str(source_path)) as pdf:
        words = pdf.pages[page_number - 1].extract_words()
    targets = {dashed_account(account), compact_digits(account)}
    wanted = []
    rows = pdf_word_rows(words)
    # Match the requested account directly against extracted words; this avoids
    # highlighting unrelated 14-digit values on the page.
    account_row_numbers = set(result.visual_account_rows)
    if not account_row_numbers:
        for row_number, row in enumerate(rows):
            for word in row:
                normalized = word["text"].replace(" ", "")
                if normalized in targets:
                    account_row_numbers.add(row_number)
    # Highlight the code row and the total amount in each matching account
    # block. PDF extraction does not preserve character offsets reliably, so
    # row/word matching is used here instead of the TXT offset spans.
    # Keep only account rows that lead into a real code row. This filters the
    # compact AC-number index that some PDF exports place above the report.
    valid_account_rows = list(account_row_numbers) if result.visual_account_rows else []
    if not valid_account_rows:
        for account_row in sorted(account_row_numbers):
            for row_number in range(account_row, min(account_row + 10, len(rows))):
                row_code, _ = _find_code(" ".join(word["text"] for word in rows[row_number]))
                if row_code:
                    valid_account_rows.append(account_row)
                    break
    account_rows_sorted = sorted(set(valid_account_rows))
    for position, account_row in enumerate(account_rows_sorted):
        for word in rows[account_row]:
            if word["text"].replace(" ", "") in targets:
                wanted.append((word["x0"] - 1, height - word["bottom"] - 1, word["x1"] - word["x0"] + 2, word["bottom"] - word["top"] + 2))
        next_account_row = account_rows_sorted[position + 1] if position + 1 < len(account_rows_sorted) else len(rows)
        for row_number in range(account_row, min(account_row + 10, next_account_row)):
            row = rows[row_number]
            row_text = " ".join(word["text"] for word in row)
            row_code, _ = _find_code(row_text)
            row_category = code_letter(row_code)
            if row_code and row_code in result.code_values:
                for word_index, word in enumerate(row):
                    normalized_word = word["text"].replace(" ", "").upper()
                    is_code_word = normalized_word == row_code or code_letter(normalized_word) == row_category
                    is_split_code_letter = normalized_word == row_category and word_index > 0 and row[word_index - 1]["text"].strip().isdigit()
                    if is_code_word or is_split_code_letter:
                        wanted.append((word["x0"] - 1, height - word["bottom"] - 1, word["x1"] - word["x0"] + 2, word["bottom"] - word["top"] + 2))
                        if is_split_code_letter:
                            previous = row[word_index - 1]
                            wanted.append((previous["x0"] - 1, height - previous["bottom"] - 1, previous["x1"] - previous["x0"] + 2, previous["bottom"] - previous["top"] + 2))
            for word in row:
                if amount_key(word["text"]) in result.amount_values:
                    wanted.append((word["x0"] - 1, height - word["bottom"] - 1, word["x1"] - word["x0"] + 2, word["bottom"] - word["top"] + 2))
    overlay_path = output.with_suffix(".overlay.pdf")
    c = canvas.Canvas(str(overlay_path), pagesize=(width, height))
    c.setFillColor(HIGHLIGHT)
    for x, y, w, h in wanted:
        c.rect(x, y, w, h, stroke=0, fill=1)
    if sr is not None and str(sr).strip():
        label = str(sr).strip()
        c.setFillColor(black)
        c.setFont("Helvetica-Bold", 16)
        tw = c.stringWidth(label, "Helvetica-Bold", 16)
        x, y = width - 24 - tw, height - 28
        c.drawString(x, y, label)
        c.setLineWidth(1.2)
        c.line(x, y - 2, x + tw, y - 2)
    c.save()
    overlay = PdfReader(str(overlay_path)).pages[0]
    original.merge_page(overlay)
    writer = PdfWriter()
    writer.add_page(original)
    with open(output, "wb") as f:
        writer.write(f)
    overlay_path.unlink(missing_ok=True)


def _is_within(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False


def _merge_results(results: list[MatchResult]) -> MatchResult:
    first = results[0]
    return MatchResult(
        first.page,
        sorted(set(hit for result in results for hit in result.account_hits)),
        sorted(set(hit for result in results for hit in result.code_hits)),
        sorted(set(hit for result in results for hit in result.amount_hits)),
        any(result.amount_found for result in results),
        all(result.amount_exact for result in results),
        ", ".join(sorted({result.code for result in results if result.code})),
        tuple(sorted({value for result in results for value in result.amount_values})),
        tuple(sorted({value for result in results for value in result.code_values})),
        tuple(sorted({row for result in results for row in result.visual_account_rows})),
    )


def _codes_text(results: list[MatchResult]) -> str:
    codes = sorted({result.code for result in results if result.code})
    return ", ".join(codes) if codes else "unknown"


def _normalise_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


def _prepare_output_columns(ws) -> tuple[int, int, int]:
    """Ensure separate Debit Code, Found Amount, and Remarks columns."""
    code_col = None
    amount_col = None
    remarks_col = None
    for column in range(1, ws.max_column + 1):
        header = _normalise_header(ws.cell(1, column).value)
        if header in {"debitcode", "postingcode", "code"}:
            code_col = column
        elif header in {"foundamount", "amountfound", "extractedamount"}:
            amount_col = column
        elif header in {"remarks", "remark", "status"}:
            remarks_col = column

    # The original checker used column D for remarks without a header, and the
    # previous revision used column E for remarks. Move either legacy column to
    # the new final Remarks column before writing the new output fields.
    legacy_remarks_col = None
    if code_col is None and remarks_col == 4:
        legacy_remarks_col = 4
    elif code_col is None and remarks_col is None and any(
        ws.cell(row, 4).value not in (None, "") for row in range(2, ws.max_row + 1)
    ):
        legacy_remarks_col = 4

    code_col = code_col or 4
    amount_col = amount_col or (code_col + 1)
    if remarks_col is None or remarks_col in {code_col, amount_col}:
        legacy_remarks_col = legacy_remarks_col or remarks_col
        remarks_col = max(code_col, amount_col) + 1

    if legacy_remarks_col is not None and legacy_remarks_col != remarks_col:
        for row in range(1, ws.max_row + 1):
            ws.cell(row, remarks_col).value = ws.cell(row, legacy_remarks_col).value
            ws.cell(row, legacy_remarks_col).value = None

    ws.cell(1, code_col).value = "Debit Code"
    ws.cell(1, amount_col).value = "Found Amount"
    ws.cell(1, remarks_col).value = "Remarks"
    return code_col, amount_col, remarks_col


def _found_amount_value(results: list[MatchResult]) -> object:
    values = sorted({value for result in results for value in result.amount_values})
    if not values:
        return ""
    if len(values) == 1:
        return values[0]
    return ", ".join(f"{value:,}" for value in values)


def process_workbook(
    source_dir: Path, workbook_path: Path, output_dir: Path, log, code_filter: str = "ALL"
) -> Path:
    output_dir = output_dir.resolve()
    log(f"Scanning posting folder recursively: {source_dir}")
    source_files = sorted(
        [
            p for p in source_dir.rglob("*")
            if p.is_file()
            and p.suffix.lower() in {".txt", ".pdf"}
            and not _is_within(p, output_dir)
        ],
        key=lambda p: str(p.relative_to(source_dir)).lower(),
    )
    if not source_files:
        raise ValueError("The selected folder does not contain any .txt or .pdf files.")
    log(f"Found {len(source_files)} posting file(s). Loading source pages...")
    wb = load_workbook(workbook_path)
    ws = wb.active
    account_keys = {
        compact_digits(ws.cell(row, 1).value)
        for row in range(2, ws.max_row + 1)
        if ws.cell(row, 1).value not in (None, "")
    }
    pages = _load_relevant_pages(source_files, account_keys, log)
    account_index = build_account_index(pages)
    log(f"Account index ready. Checking workbook rows for {code_filter}...")
    output_dir.mkdir(parents=True, exist_ok=True)
    code_col, amount_col, remarks_col = _prepare_output_columns(ws)
    for row in range(2, ws.max_row + 1):
        account = ws.cell(row, 1).value
        if account is None or str(account).strip() == "":
            continue
        amount = ws.cell(row, 2).value
        sr = ws.cell(row, 3).value
        code_cell = ws.cell(row, code_col)
        remarks_cell = ws.cell(row, remarks_col)
        label = dashed_account(account) or str(account)
        try:
            results = find_matches(pages, account, amount, code_filter, account_index.get(compact_digits(account), []))
            if not results:
                code_cell.value = ""
                ws.cell(row, amount_col).value = ""
                remarks_cell.value = "Not found"
                log(f"Row {row}: Not found — {label} ({code_filter})")
                continue
            # One PDF is produced for each source page containing a match. If an
            # account occurs more than once on that page, all matching occurrences
            # are highlighted together in that page's output.
            groups = {}
            for result in results:
                key = (str(result.page.source_path.resolve()), result.page.source_index, result.page.compact_output)
                groups.setdefault(key, []).append(result)
            requested_name = str(sr).strip() if sr not in (None, "") else compact_digits(account)
            safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", requested_name)
            made = []
            for group_number, grouped in enumerate(groups.values(), start=1):
                result = _merge_results(grouped)
                suffix = f"_{group_number}" if len(groups) > 1 else ""
                out_pdf = output_dir / f"{safe_name}{suffix}.pdf"
                if out_pdf.exists():
                    duplicate = 2
                    base = out_pdf.with_suffix("")
                    while Path(f"{base}_{duplicate}.pdf").exists():
                        duplicate += 1
                    out_pdf = Path(f"{base}_{duplicate}.pdf")
                if result.page.compact_output or result.page.source_path.suffix.lower() != ".pdf":
                    render_text_page(result.page, result, sr, out_pdf)
                else:
                    render_pdf_page(result.page.source_path, result.page.number, result, account, sr, out_pdf)
                made.append((result, out_pdf))
            code_cell.value = _codes_text(results)
            ws.cell(row, amount_col).value = _found_amount_value(results)
            remarks_cell.value = "Found and extracted only"
            log(f"Row {row}: {remarks_cell.value} — {label} ({code_cell.value})")
        except Exception as exc:
            code_cell.value = ""
            ws.cell(row, amount_col).value = ""
            remarks_cell.value = f"Error: {exc}"
            log(f"Row {row}: {remarks_cell.value} — {label}")
    # Update the workbook selected by the user; only PDFs are written to the output folder.
    wb.save(workbook_path)
    return workbook_path


class App:
    def __init__(self, root: Tk):
        self.root = root
        root.title("Posting Checker")
        root.geometry("680x500")
        root.minsize(680, 420)
        self.source_dir = StringVar()
        self.workbook = StringVar()
        self.output = StringVar(value=str(Path.cwd() / "output"))
        self.code_filter = StringVar(value="All Codes")
        self.status = StringVar(value="Ready")
        self._build()

    def _build(self):
        frame = ttk.Frame(self.root, padding=18)
        frame.pack(fill=BOTH, expand=True)
        ttk.Label(frame, text="Posting Checker", font=("Segoe UI", 20, "bold")).pack(anchor="w")
        ttk.Label(frame, text="Developed by Tahir Naqash", wraplength=700).pack(anchor="w", pady=(4, 18))
        self._file_row(frame, "Posting folder (.txt/.pdf)", self.source_dir, self.pick_source_dir, None)
        self._file_row(frame, "Excel input (.xlsx)", self.workbook, self.pick_workbook, ("Excel files", "*.xlsx"))
        self._file_row(frame, "Output folder", self.output, self.pick_output, None)
        filter_row = ttk.Frame(frame); filter_row.pack(fill=X, pady=5)
        ttk.Label(filter_row, text="Find Posting of", width=25).pack(side=LEFT)
        ttk.Combobox(
            filter_row,
            textvariable=self.code_filter,
            values=("Code-C", "Code-B", "Code-F", "All Codes"),
            state="readonly",
            width=18,
        ).pack(side=LEFT)
        buttons = ttk.Frame(frame); buttons.pack(fill=X, pady=(12, 8))
        self.run_button = ttk.Button(buttons, text="Check posting", command=self.start)
        self.run_button.pack(side=LEFT)
        ttk.Button(buttons, text="Download Template", command=self.download_template, width=20).pack(side=LEFT, padx=8)
        ttk.Button(buttons, text="Open Folder", command=self.open_output, width=18).pack(side=LEFT)
        ttk.Label(frame, textvariable=self.status).pack(anchor="w", pady=(0, 5))
        self.log = ttk.Treeview(frame, columns=("message",), show="headings", height=14)
        self.log.heading("message", text="Activity")
        self.log.column("message", width=690)
        self.log.pack(fill=BOTH, expand=True)

    def _file_row(self, parent, label, var, command, filetypes):
        row = ttk.Frame(parent); row.pack(fill=X, pady=5)
        ttk.Label(row, text=label, width=25).pack(side=LEFT)
        ttk.Entry(row, textvariable=var).pack(side=LEFT, fill=X, expand=True)
        ttk.Button(row, text="Browse", command=lambda: command(filetypes)).pack(side=LEFT, padx=(8, 0))

    def pick_source_dir(self, _):
        value = filedialog.askdirectory()
        if value: self.source_dir.set(value)

    def pick_workbook(self, filetypes):
        value = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if value: self.workbook.set(value)

    def pick_output(self, _):
        value = filedialog.askdirectory()
        if value: self.output.set(value)

    def download_template(self):
        path = filedialog.asksaveasfilename(
            title="Save Posting Checker Template",
            defaultextension=".xlsx",
            initialfile="CP52_Posting_Checker_Template.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return
        from openpyxl import Workbook

        wb = Workbook()
        ws = wb.active
        ws.title = "Data"
        headers = ["AC No", "Amount", "Para Sr No.", "Debit Code", "Found Amount", "Remarks"]
        for column, header in enumerate(headers, 1):
            ws.cell(1, column).value = header
            ws.cell(1, column).font = ws.cell(1, column).font.copy(bold=True)
            ws.column_dimensions[ws.cell(1, column).column_letter].width = max(len(header) + 4, 16)
        ws.freeze_panes = "A2"
        wb.save(path)
        messagebox.showinfo("Template Saved", f"Template saved to:\n{path}")

    def open_output(self):
        path = Path(self.output.get())
        path.mkdir(parents=True, exist_ok=True)
        os.startfile(path)

    def start(self):
        if not self.source_dir.get() or not self.workbook.get():
            messagebox.showerror("Missing input", "Select both a posting folder and an Excel workbook.")
            return
        source_dir, workbook, output = Path(self.source_dir.get()), Path(self.workbook.get()), Path(self.output.get())
        if not source_dir.exists() or not source_dir.is_dir() or not workbook.exists():
            messagebox.showerror("Invalid input", "One or more selected files do not exist.")
            return
        self.run_button.configure(state="disabled")
        self.status.set("Processing...")
        threading.Thread(target=self._worker, args=(source_dir, workbook, output), daemon=True).start()

    def _worker(self, source_dir, workbook, output):
        try:
            selected_filter = {"Code-C": "C", "Code-B": "B", "Code-F": "F", "All Codes": "ALL"}.get(self.code_filter.get(), "ALL")
            result = process_workbook(source_dir, workbook, output, lambda msg: self.root.after(0, lambda: self._log(msg)), selected_filter)
            self.root.after(0, lambda: self._done(result))
        except Exception as exc:
            self.root.after(0, lambda: self._error(exc))

    def _log(self, message):
        self.log.insert("", END, values=(message,))
        self.log.yview_moveto(1)
        self.status.set(f"{message}")

    def _done(self, result):
        self.run_button.configure(state="normal")
        self.status.set(f"Completed: {result}")
        messagebox.showinfo("Completed", f"Excel updated:\n{result}\n\nPDFs saved in:\n{result.parent}")

    def _error(self, exc):
        self.run_button.configure(state="normal")
        self.status.set("Error")
        messagebox.showerror("Processing error", str(exc))


if __name__ == "__main__":
    root = Tk()
    App(root)
    root.mainloop()
