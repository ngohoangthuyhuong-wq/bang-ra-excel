"""Đọc bảng từ PDF scan và ảnh, xuất Excel. Chạy hoàn toàn trên máy."""

from __future__ import annotations

import io
import re
import threading
from html import escape
from pathlib import Path

import cv2
import numpy as np
import xlsxwriter


def _patch_opencv() -> None:
    """OpenCV 5 bỏ niBlackThreshold mà img2table cần để tách nét chữ."""
    ximg = getattr(cv2, "ximgproc", None)
    if ximg is None or hasattr(ximg, "niBlackThreshold"):
        return
    if not hasattr(ximg, "BINARIZATION_SAUVOLA"):
        ximg.BINARIZATION_SAUVOLA = 2

    def ni_black_threshold(src, max_value, _threshold_type, block_size, k, binarizationMethod=2, r=128.0):
        gray = src.astype(np.float32)
        block = max(3, int(block_size) | 1)
        mean = cv2.boxFilter(gray, ddepth=cv2.CV_32F, ksize=(block, block), borderType=cv2.BORDER_REPLICATE)
        square = cv2.boxFilter(gray * gray, ddepth=cv2.CV_32F, ksize=(block, block), borderType=cv2.BORDER_REPLICATE)
        std = np.sqrt(np.maximum(square - mean * mean, 0))
        radius = float(r) if r else 128.0
        if binarizationMethod == getattr(ximg, "BINARIZATION_SAUVOLA", 2):
            level = mean * (1.0 + float(k) * (std / radius - 1.0))
        else:
            level = mean + float(k) * std
        return np.where(gray <= level, max_value, 0).astype(np.uint8)

    ximg.niBlackThreshold = ni_black_threshold


_patch_opencv()

from img2table.document import Image as DocImage
from img2table.document import PDF
from img2table.ocr import EasyOCR
from img2table.ocr._types import OCRData
from PIL import Image as PILImage
from PIL import ImageOps

MAX_PAGES = 60
MIN_CONFIDENCE = 18
PREVIEW_ROWS = 100

_ocr: "RememberOCR | None" = None
_ocr_lock = threading.Lock()
_ocr_error: str | None = None


class RememberOCR(EasyOCR):
    """EasyOCR tiếng Việt + Anh. Giữ lại từng chữ để không bỏ phần ngoài bảng."""

    def __init__(self) -> None:
        import torch

        gpu = bool(torch.cuda.is_available())
        super().__init__(lang=["vi", "en"], kw={"gpu": gpu})
        self.using_gpu = gpu
        self.last_pages: list[list[dict]] | None = None

    def of(self, document):
        self.last_pages = []
        records: dict[int, list[dict]] = {}
        for page, image in enumerate(document.images):
            raw = self.reader.readtext(
                image,
                detail=1,
                paragraph=False,
                mag_ratio=1.2,
                min_size=5,
                text_threshold=0.45,
                low_text=0.28,
                link_threshold=0.3,
                canvas_size=3200,
                contrast_ths=0.08,
                adjust_contrast=0.6,
                decoder="greedy",
            )
            page_words: list[dict] = []
            for idx, word in enumerate(raw):
                box, text, conf = word[0], word[1], word[2]
                item = {
                    "id": f"word_{page + 1}_{idx + 1}",
                    "parent": f"word_{page + 1}_{idx + 1}",
                    "value": text,
                    "confidence": int(round(100 * float(conf))),
                    "x1": int(round(min(p[0] for p in box))),
                    "y1": int(round(min(p[1] for p in box))),
                    "x2": int(round(max(p[0] for p in box))),
                    "y2": int(round(max(p[1] for p in box))),
                }
                page_words.append(item)
                records.setdefault(page, []).append(item)
            self.last_pages.append(page_words)
        return OCRData(records=records) if records else None


def ocr_status() -> dict:
    if _ocr is not None:
        device = "GPU" if _ocr.using_gpu else "CPU"
        return {"ready": True, "message": f"Bộ đọc tiếng Việt đã sẵn ({device})."}
    if _ocr_error:
        return {"ready": False, "message": _ocr_error}
    return {
        "ready": False,
        "message": "Đang tải bộ đọc tiếng Việt. Lần đầu mất vài phút, các lần sau mở ngay.",
    }


def get_ocr() -> RememberOCR:
    global _ocr, _ocr_error
    if _ocr is not None:
        return _ocr
    with _ocr_lock:
        if _ocr is not None:
            return _ocr
        try:
            _ocr = RememberOCR()
            _ocr_error = None
        except Exception as exc:
            _ocr_error = f"Không mở được bộ đọc chữ: {exc}"
            raise
        return _ocr


def _open_pdf(data: bytes):
    try:
        import pymupdf

        return pymupdf.open(stream=data, filetype="pdf")
    except ImportError:
        import fitz

        return fitz.open(stream=data, filetype="pdf")


def _looks_native(text: str) -> bool:
    compact = re.sub(r"\s+", "", text or "")
    if len(compact) < 40:
        return False
    letters = sum(ch.isalpha() for ch in compact)
    return letters / len(compact) > 0.35


def _prepare_image(raw: bytes) -> bytes:
    im = PILImage.open(io.BytesIO(raw))
    im = ImageOps.exif_transpose(im)
    if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
        rgba = im.convert("RGBA")
        base = PILImage.new("RGB", rgba.size, (255, 255, 255))
        base.paste(rgba, mask=rgba.split()[-1])
        im = base
    else:
        im = im.convert("RGB")

    span = im.convert("L").getextrema()
    if span[1] - span[0] < 80:
        im = ImageOps.autocontrast(im, cutoff=1)

    long_edge = max(im.size)
    if long_edge < 2400:
        scale = min(3.0, 2400 / long_edge)
        im = im.resize(
            (max(1, int(im.width * scale)), max(1, int(im.height * scale))),
            PILImage.Resampling.LANCZOS,
        )
    long_edge = max(im.size)
    if long_edge > 4200:
        scale = 4200 / long_edge
        im = im.resize(
            (max(1, int(im.width * scale)), max(1, int(im.height * scale))),
            PILImage.Resampling.LANCZOS,
        )

    buf = io.BytesIO()
    im.save(buf, format="PNG")
    return buf.getvalue()


def _page_png(page) -> bytes:
    pix = page.get_pixmap(dpi=320, alpha=False)
    return pix.tobytes("png")


def _safe_sheet_name(raw: str, used: set[str]) -> str:
    cleaned = re.sub(r"[\[\]\:\*\?\/\\]", " ", raw)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" '")
    cleaned = cleaned or "Bang"
    base = cleaned[:31]
    name = base
    n = 2
    while name.lower() in used:
        suffix = f" {n}"
        name = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(name.lower())
    return name


def _fix_digits(value: str) -> str:
    """OCR hay đọc số 0 thành chữ o. Chỉ sửa token toàn số."""

    def repl(match: re.Match) -> str:
        token = match.group(0)
        if not any(ch.isdigit() for ch in token):
            return token
        return token.translate(str.maketrans({"o": "0", "O": "0"}))

    return re.sub(r"[0-9oO]{2,}", repl, value)


def _clean_table(table) -> None:
    for row in table.content.values():
        for cell in row:
            if cell.value:
                cell.value = _fix_digits(str(cell.value).strip())


def _matrix(table) -> list[list[str]]:
    rows: list[list[str]] = []
    for row in table.content.values():
        rows.append(["" if cell.value is None else str(cell.value).strip() for cell in row])
    while rows and not any(rows[-1]):
        rows.pop()
    return rows


def _html_table(rows: list[list[str]], limit: int = PREVIEW_ROWS) -> str:
    shown = rows[:limit]
    parts = ["<table>"]
    for row in shown:
        parts.append("<tr>")
        for cell in row:
            parts.append(f"<td>{escape(cell) if cell else ''}</td>")
        parts.append("</tr>")
    parts.append("</table>")
    return "".join(parts)


def _median(values: list[float]) -> float:
    if not values:
        return 12.0
    ordered = sorted(values)
    return float(ordered[len(ordered) // 2])


def _cluster_rows(words: list[dict]) -> list[list[dict]]:
    usable = [w for w in words if w.get("value") and w.get("confidence", 0) >= MIN_CONFIDENCE]
    usable.sort(key=lambda w: ((w["y1"] + w["y2"]) / 2, w["x1"]))
    if not usable:
        return []
    height = _median([w["y2"] - w["y1"] for w in usable]) or 12
    rows: list[dict] = []
    for word in usable:
        cy = (word["y1"] + word["y2"]) / 2
        if not rows or abs(cy - rows[-1]["cy"]) > height * 0.65:
            rows.append({"cy": cy, "words": [word]})
        else:
            rows[-1]["words"].append(word)
            rows[-1]["cy"] = sum((w["y1"] + w["y2"]) / 2 for w in rows[-1]["words"]) / len(
                rows[-1]["words"]
            )
    return [sorted(row["words"], key=lambda w: w["x1"]) for row in rows]


def _words_to_lines(words: list[dict]) -> list[str]:
    lines = []
    for row in _cluster_rows(words):
        text = " ".join(w["value"].strip() for w in row if w["value"].strip())
        if text:
            lines.append(text)
    return lines


def _words_to_grid(words: list[dict]) -> list[list[str]]:
    clustered = _cluster_rows(words)
    if len(clustered) < 2:
        return []
    widths = [w["x2"] - w["x1"] for row in clustered for w in row]
    gap = max(18.0, _median(widths) * 0.45)
    edges: list[float] = []
    for row in clustered:
        for word in row:
            if not edges or word["x1"] - edges[-1] > gap:
                edges.append(float(word["x1"]))
    if len(edges) < 2:
        return []
    grid: list[list[str]] = []
    for row in clustered:
        cells = [""] * len(edges)
        for word in row:
            idx = min(range(len(edges)), key=lambda i: abs(word["x1"] - edges[i]))
            cells[idx] = (cells[idx] + " " + word["value"]).strip()
        if any(cells):
            grid.append(cells)
    return grid


def _inside(word: dict, boxes) -> bool:
    cx = (word["x1"] + word["x2"]) / 2
    cy = (word["y1"] + word["y2"]) / 2
    for box in boxes:
        if box.x1 - 6 <= cx <= box.x2 + 6 and box.y1 - 6 <= cy <= box.y2 + 6:
            return True
    return False


def _leftover_native(lines: list[str], tables) -> list[str]:
    blob = " ".join(
        (cell.value or "")
        for table in tables
        for row in table.content.values()
        for cell in row
    ).lower()
    kept = []
    for line in lines:
        tokens = [tok for tok in re.split(r"\s+", line) if len(tok) > 1]
        if not tokens:
            continue
        hit = sum(1 for tok in tokens if tok.lower() in blob)
        if hit / len(tokens) < 0.55:
            kept.append(line)
    return kept


def _extract_kwargs(ocr: RememberOCR, borderless: bool) -> dict:
    return {
        "ocr": ocr,
        "implicit_rows": True,
        "implicit_columns": True,
        "borderless_tables": borderless,
        "min_confidence": MIN_CONFIDENCE,
    }


def _take_tables(result):
    if isinstance(result, dict):
        tables = []
        for group in result.values():
            tables.extend(group or [])
        return tables
    return list(result or [])


class _Book:
    def __init__(self) -> None:
        self.buffer = io.BytesIO()
        self.book = xlsxwriter.Workbook(self.buffer, {"in_memory": True})
        self.used: set[str] = set()
        self.cell_fmt = self.book.add_format(
            {
                "font_name": "Calibri",
                "font_size": 12,
                "border": 1,
                "border_color": "#c5d5cb",
                "valign": "vcenter",
                "text_wrap": True,
            }
        )
        self.head_fmt = self.book.add_format(
            {
                "font_name": "Calibri",
                "font_size": 12,
                "bold": True,
                "font_color": "#f7fbf8",
                "bg_color": "#143d2c",
                "border": 1,
                "border_color": "#143d2c",
                "valign": "vcenter",
                "text_wrap": True,
            }
        )
        self.text_fmt = self.book.add_format(
            {
                "font_name": "Calibri",
                "font_size": 12,
                "valign": "vcenter",
                "text_wrap": True,
            }
        )
        self.sheets: list[dict] = []

    def add_table(self, name: str, table) -> None:
        _clean_table(table)
        rows = _matrix(table)
        if not rows:
            return
        sheet_name = _safe_sheet_name(name, self.used)
        ws = self.book.add_worksheet(sheet_name)
        table._to_worksheet(ws, self.cell_fmt)
        ws.freeze_panes(1, 0)
        self.sheets.append(
            {
                "name": sheet_name,
                "html": _html_table(rows),
                "rows": len(rows),
                "cols": max(len(r) for r in rows),
                "preview_limited": len(rows) > PREVIEW_ROWS,
                "kind": "table",
            }
        )

    def add_grid(self, name: str, rows: list[list[str]]) -> None:
        rows = [[_fix_digits(cell) for cell in row] for row in rows if any(row)]
        if not rows:
            return
        width = max(len(r) for r in rows)
        sheet_name = _safe_sheet_name(name, self.used)
        ws = self.book.add_worksheet(sheet_name)
        for r, row in enumerate(rows):
            fmt = self.head_fmt if r == 0 else self.cell_fmt
            for c in range(width):
                value = row[c] if c < len(row) else ""
                ws.write(r, c, value, fmt)
        ws.autofit()
        ws.freeze_panes(1, 0)
        self.sheets.append(
            {
                "name": sheet_name,
                "html": _html_table(rows),
                "rows": len(rows),
                "cols": width,
                "preview_limited": len(rows) > PREVIEW_ROWS,
                "kind": "table",
            }
        )

    def add_lines(self, name: str, lines: list[tuple[str, str]]) -> None:
        if not lines:
            return
        sheet_name = _safe_sheet_name(name, self.used)
        ws = self.book.add_worksheet(sheet_name)
        ws.write(0, 0, "Nguồn", self.head_fmt)
        ws.write(0, 1, "Chữ đọc được", self.head_fmt)
        for i, (source, line) in enumerate(lines, start=1):
            ws.write(i, 0, source, self.text_fmt)
            ws.write(i, 1, line, self.text_fmt)
        ws.set_column(0, 0, 22)
        ws.set_column(1, 1, 88)
        ws.freeze_panes(1, 0)
        preview = [["Nguồn", "Chữ đọc được"], *[[a, b] for a, b in lines]]
        self.sheets.append(
            {
                "name": sheet_name,
                "html": _html_table(preview),
                "rows": len(lines),
                "cols": 2,
                "preview_limited": len(preview) > PREVIEW_ROWS,
                "kind": "text",
            }
        )

    def finish(self) -> bytes:
        self.book.close()
        return self.buffer.getvalue()


def _label(path_name: str, multi: bool) -> str:
    if not multi:
        return ""
    stem = Path(path_name).stem
    stem = re.sub(r"\s+", " ", stem).strip()
    return stem[:14]


def read_document(
    name: str,
    data: bytes,
    *,
    borderless: bool,
    keep_outside: bool,
    multi: bool,
    book: _Book,
    on_page,
) -> tuple[int, int, str]:
    """Trả về (số trang đã đọc, số trang gốc, ghi chú cắt trang)."""
    ocr = get_ocr()
    suffix = Path(name).suffix.lower()
    prefix = _label(name, multi)
    outside: list[tuple[str, str]] = []
    note = ""
    pages_read = 0

    def tagged(page_no: int, index: int, title: str | None) -> str:
        base = (title or "").strip()
        if len(base) < 2:
            base = f"Trang {page_no} bang {index}"
        else:
            base = f"T{page_no} {base}"
        return f"{prefix} {base}".strip() if prefix else base

    if suffix == ".pdf":
        doc = _open_pdf(data)
        try:
            if getattr(doc, "is_encrypted", False) and getattr(doc, "needs_pass", False):
                raise ValueError("PDF đang khóa mật khẩu.")
            total = doc.page_count
            limit = min(total, MAX_PAGES)
            if total > MAX_PAGES:
                note = f"{name}: file có {total} trang, đã đọc {MAX_PAGES} trang đầu."
            for index in range(limit):
                page = doc[index]
                page_no = index + 1
                on_page(f"Đang đọc {name}, trang {page_no}/{limit}")
                native = page.get_text("text") or ""
                ocr.last_pages = None
                if _looks_native(native):
                    result = PDF(
                        src=data,
                        pages=[index],
                        detect_rotation=False,
                        pdf_text_extraction=True,
                    ).extract_tables(**_extract_kwargs(ocr, borderless))
                    tables = _take_tables(result)
                    # Trang chữ thật mà không ra bảng: vẫn OCR ảnh để không bỏ bảng scan kèm chữ.
                    if not tables:
                        png = _prepare_image(_page_png(page))
                        ocr.last_pages = None
                        tables = DocImage(src=png, detect_rotation=True).extract_tables(
                            **_extract_kwargs(ocr, True)
                        )
                else:
                    png = _prepare_image(_page_png(page))
                    tables = DocImage(src=png, detect_rotation=True).extract_tables(
                        **_extract_kwargs(ocr, borderless)
                    )
                for n, table in enumerate(tables, start=1):
                    book.add_table(tagged(page_no, n, table.title), table)
                source = f"{prefix + ' ' if prefix else ''}trang {page_no}".strip()
                words = (ocr.last_pages or [[]])[0] if ocr.last_pages else []
                if words:
                    loose = [w for w in words if not _inside(w, [t.bbox for t in tables])]
                    if not tables:
                        grid = _words_to_grid(words)
                        if grid:
                            book.add_grid(tagged(page_no, 1, None), grid)
                            loose = []
                    if keep_outside:
                        for line in _words_to_lines(loose):
                            outside.append((source, _fix_digits(line)))
                elif keep_outside:
                    lines = [ln.strip() for ln in native.splitlines() if ln.strip()]
                    for line in _leftover_native(lines, tables):
                        outside.append((source, line))
                pages_read += 1
        finally:
            doc.close()
    else:
        on_page(f"Đang đọc ảnh {name}")
        png = _prepare_image(data)
        ocr.last_pages = None
        tables = DocImage(src=png, detect_rotation=True).extract_tables(
            **_extract_kwargs(ocr, borderless)
        )
        for n, table in enumerate(tables, start=1):
            book.add_table(tagged(1, n, table.title), table)
        source = prefix or Path(name).stem[:22]
        words = (ocr.last_pages or [[]])[0] if ocr.last_pages else []
        loose = [w for w in words if not _inside(w, [t.bbox for t in tables])]
        if not tables:
            grid = _words_to_grid(words)
            if grid:
                book.add_grid(tagged(1, 1, None), grid)
                loose = []
        if keep_outside:
            for line in _words_to_lines(loose):
                outside.append((source, _fix_digits(line)))
        pages_read = 1
        note = ""

    if outside:
        title = f"{prefix} chu doc duoc".strip() if prefix else "Chu doc duoc"
        book.add_lines(title, outside)
    return pages_read, pages_read, note


def convert_files(files: list[tuple[str, bytes]], *, borderless: bool, keep_outside: bool, on_page) -> dict:
    if not files:
        raise ValueError("Chưa có file.")
    book = _Book()
    notes = []
    pages = 0
    multi = len(files) > 1
    for name, data in files:
        _pages, _total, note = read_document(
            name,
            data,
            borderless=borderless,
            keep_outside=keep_outside,
            multi=multi,
            book=book,
            on_page=on_page,
        )
        pages += _pages
        if note:
            notes.append(note)
    if not book.sheets:
        raise ValueError(
            "Không thấy chữ. Ảnh cần sáng, thẳng, và chữ đủ lớn. Có thể bật “Bảng không có đường kẻ”."
        )
    tables = sum(1 for sheet in book.sheets if sheet["kind"] == "table")
    return {
        "xlsx": book.finish(),
        "sheets": book.sheets,
        "tables": tables,
        "pages": pages,
        "note": " ".join(notes),
    }
