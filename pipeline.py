"""Core pipeline: PDF/image -> text or vision -> Gemini -> structured data -> Excel.

Two extraction paths:
- PDF with a text layer  -> text is sent to the LLM (fast, cheap, accurate).
- Scanned PDF / image    -> rendered pages are sent to Gemini vision directly.

Extraction modes: "auto" (text if a layer exists, else vision),
"text" (text layer only), "vision" (always send page images).

Warnings are structured: {"message", "severity" (critical|check|info), "field"}.
"""

import base64
import json
import logging
import os
import re
import time
from datetime import datetime
from pathlib import Path
from typing import NamedTuple

import pymupdf as fitz
from dotenv import load_dotenv
from google import genai
from google.genai import types
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field

load_dotenv()
log = logging.getLogger("doclift")

MODEL = "gemini-3.8-flash"
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp"}
MIME_BY_EXT = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
}
MAX_VISION_PAGES = 5
VALID_TAX_RATES = {0, 5, 8, 10}
LOG_DIR = Path(__file__).parent / "logs"


# ---------------------------------------------------------------- schema

class InvoiceItem(BaseModel):
    ten_hang: str = ""
    don_vi_tinh: str = ""
    so_luong: float | None = None
    don_gia: float | None = None
    thanh_tien: float | None = None
    thue_suat: float | None = None  # per-line VAT %, when printed per line


class TaxLine(BaseModel):
    """One row of the VAT breakdown table on multi-rate invoices."""
    thue_suat: float | None = None
    tien_hang: float | None = None
    tien_thue: float | None = None


class Invoice(BaseModel):
    ky_hieu: str = ""
    so_hoa_don: str = ""
    ngay_lap: str = ""
    ten_don_vi_ban: str = ""
    ma_so_thue_ban: str = ""
    dia_chi_ban: str = ""
    ten_nguoi_mua: str = ""
    ten_don_vi_mua: str = ""
    ma_so_thue_mua: str = ""
    dia_chi_mua: str = ""
    hinh_thuc_thanh_toan: str = ""
    items: list[InvoiceItem] = Field(default_factory=list)
    bang_ke_thue: list[TaxLine] = Field(default_factory=list)
    tong_tien_hang: float | None = None
    thue_suat: float | None = None  # single-rate invoices; null when multi-rate
    tien_thue: float | None = None
    tong_thanh_toan: float | None = None
    so_tien_bang_chu: str = ""
    # fields the model itself is unsure about (blur, stamp over text, ...)
    truong_can_kiem_tra: list[str] = Field(default_factory=list)


class Warning(BaseModel):
    message: str
    severity: str = "check"  # critical | check | info
    field: str = ""


class ExtractionResult(NamedTuple):
    invoice: Invoice
    warnings: list[Warning]
    mode: str
    duration_ms: int


REQUIRED_FIELDS = ("so_hoa_don", "tong_thanh_toan")


# ---------------------------------------------------------------- input handling

def extract_text_from_pdf(path: str | Path) -> str:
    doc = fitz.open(path)
    try:
        return "\n".join(page.get_text() for page in doc).strip()
    finally:
        doc.close()


def pdf_to_page_images(path: str | Path, dpi: int = 200) -> list[bytes]:
    """Render PDF pages to PNG bytes, capped to bound token usage."""
    doc = fitz.open(path)
    try:
        return [p.get_pixmap(dpi=dpi).tobytes("png")
                for p in doc[:MAX_VISION_PAGES]]
    finally:
        doc.close()


def _vision_parts(path: Path) -> list:
    if path.suffix.lower() == ".pdf":
        return [types.Part.from_bytes(data=b, mime_type="image/png")
                for b in pdf_to_page_images(path)]
    return [types.Part.from_bytes(
        data=path.read_bytes(), mime_type=MIME_BY_EXT[path.suffix.lower()])]


def build_contents(path: str | Path, mode: str = "auto") -> tuple[str, list]:
    """Return (effective_mode, contents_parts). mode: auto|text|vision."""
    path = Path(path)
    ext = path.suffix.lower()
    if ext != ".pdf" and ext not in IMAGE_EXTS:
        raise ValueError(f"Unsupported file type: {ext}")

    if mode == "vision":
        return "vision", [EXTRACTION_PROMPT, *_vision_parts(path)]

    if ext == ".pdf":
        text = extract_text_from_pdf(path)
        if mode == "text" and len(text) <= 100:
            raise RuntimeError(
                "PDF không có lớp text (file scan). Đổi sang chế độ Chính xác "
                "(vision) hoặc Tự động.")
        if len(text) > 100:
            return "text", [EXTRACTION_PROMPT, f"---\n{text}\n---"]
        return "vision", [EXTRACTION_PROMPT, *_vision_parts(path)]

    # image input: 'text' mode is meaningless here, vision always applies
    return "vision", [EXTRACTION_PROMPT, *_vision_parts(path)]


# ---------------------------------------------------------------- LLM extraction

EXTRACTION_PROMPT = """Bạn là hệ thống trích xuất dữ liệu hóa đơn Việt Nam (hóa đơn GTGT/VAT).

Đọc nội dung hóa đơn (văn bản hoặc hình ảnh) được cung cấp và trả về JSON đúng schema yêu cầu.

Quy tắc chung:
- Chỉ trích xuất thông tin CÓ TRONG tài liệu. Không suy đoán, không bịa.
- Trường không tìm thấy: để chuỗi rỗng "" (text) hoặc null (số).
- ngay_lap: định dạng DD/MM/YYYY.
- Tiền/số: bỏ dấu phân cách hàng nghìn và đơn vị tiền tệ.
- ma_so_thue: chỉ giữ chữ số và dấu gạch (ví dụ "0101234567" hoặc "0101234567-001").

Trường hợp khó:
- Tài liệu nghiêng/mờ/đóng dấu đè chữ: cố đọc tốt nhất, đừng đoán số không rõ.
- Nhiều thuế suất trên một hóa đơn: điền từng dòng vào bang_ke_thue
  (thue_suat, tien_hang, tien_thue), và để invoice-level thue_suat = null.
- Dòng hàng dài bị xuống dòng: gộp về đúng một item, tên hàng gộp đủ 2 dòng.
- Dòng hàng có thuế suất riêng từng dòng: điền thue_suat vào từng item.

Độ tin cậy:
- Với mỗi trường đọc không chắc (chữ mờ, dấu đè, số nghi ngờ): thêm tên trường
  vào truong_can_kiem_tra (ví dụ "so_hoa_don", "tong_thanh_toan", "ma_so_thue_ban",
  "items"). Chỉ liệt kê trường thật sự nghi ngờ.
"""


LLM_TIMEOUT_S = 90

# LLM provider: "gemini" (default) or "mock" — mock returns a canned invoice
# so backend features can be developed/tested while the API is unavailable.
PROVIDER = os.environ.get("LLM_PROVIDER", "gemini")


def _mock_invoice() -> Invoice:
    return Invoice(
        ky_hieu="1C25TAA", so_hoa_don="MOCK-001", ngay_lap="11/11/2026",
        ten_don_vi_ban="MOCK PROVIDER (không phải dữ liệu thật)",
        ma_so_thue_ban="0101248141",
        ten_don_vi_mua="CÔNG TY CỔ PHẦN XYZ", ma_so_thue_mua="0100109106",
        hinh_thuc_thanh_toan="Chuyển khoản",
        items=[InvoiceItem(ten_hang="Dữ liệu giả từ mock provider",
                           don_vi_tinh="Cái", so_luong=1, don_gia=1000000,
                           thanh_tien=1000000)],
        tong_tien_hang=1000000, thue_suat=10, tien_thue=100000,
        tong_thanh_toan=1100000,
        so_tien_bang_chu="Một triệu một trăm nghìn đồng",
    )


def get_client() -> genai.Client:
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "Missing GEMINI_API_KEY. Copy .env.example to .env and set your key.")
    return genai.Client(
        api_key=api_key,
        http_options=types.HttpOptions(timeout=LLM_TIMEOUT_S))


def _call_llm(client, contents) -> Invoice | None:
    """One LLM call; returns parsed Invoice, or None to signal a retryable failure."""
    try:
        resp = client.models.generate_content(
            model=MODEL,
            contents=contents,
            config=types.GenerateContentConfig(
                response_mime_type="application/json",
                response_schema=Invoice,
                temperature=0,
            ),
        )
    except Exception as e:
        # 503 high demand, timeouts, transient network — let the caller retry once
        log.warning("LLM call failed: %s", e)
        return None
    try:
        return resp.parsed or Invoice.model_validate_json(resp.text)
    except Exception as e:
        log.warning("LLM returned unusable payload: %s", e)
        return None


def _missing_required(inv: Invoice) -> list[str]:
    return [f for f in REQUIRED_FIELDS if not getattr(inv, f)]


def extract_invoice(path: str | Path, mode: str = "auto") -> ExtractionResult:
    """Full pipeline for one file. mode: auto | text | vision."""
    effective_mode, contents = build_contents(path, mode)

    t0 = time.monotonic()
    if PROVIDER == "mock":
        invoice = _mock_invoice()
    else:
        client = get_client()
        invoice = _call_llm(client, contents)

        if invoice is not None and _missing_required(invoice):
            log.info("missing required fields %s — retrying",
                     _missing_required(invoice))
            invoice = None

        if invoice is None:
            # one retry, nudged: remind the model about required fields
            retry_contents = [*contents,
                              "Nhắc lại: bắt buộc có so_hoa_don và "
                              "tong_thanh_toan trong JSON trả về."]
            invoice = _call_llm(client, retry_contents)

    duration_ms = int((time.monotonic() - t0) * 1000)

    if invoice is None:
        raise RuntimeError(
            "AI không xử lý được sau 2 lần thử — có thể dịch vụ đang quá "
            "tải, hoặc chứng từ quá khó đọc. Thử lại sau ít phút hoặc đổi "
            "chế độ đọc.")

    warnings = validate_invoice(invoice)
    return ExtractionResult(invoice, warnings, effective_mode, duration_ms)


# ---------------------------------------------------------------- validation

def check_mst(mst: str) -> bool | None:
    """Vietnam tax code checksum. Returns None if format doesn't apply."""
    digits = re.sub(r"\D", "", mst)
    if not digits:
        return None
    if len(digits) not in (10, 13):
        return False
    weights = [31, 29, 23, 19, 17, 13, 7, 5, 3]
    s = sum(int(d) * w for d, w in zip(digits[:9], weights))
    check = (10 - (s % 11)) % 10
    return check == int(digits[9])


def _check_date(ngay: str) -> bool | None:
    if not ngay:
        return None
    try:
        datetime.strptime(ngay.strip(), "%d/%m/%Y")
        return True
    except ValueError:
        return False


def validate_invoice(inv: Invoice) -> list[Warning]:
    """Soft validation -> structured warnings with severity."""
    w: list[Warning] = []
    TOL = 1000  # VND

    def add(msg, severity="check", field=""):
        w.append(Warning(message=msg, severity=severity, field=field))

    # --- required fields present ---
    for field, label in (("so_hoa_don", "Số hóa đơn"),
                         ("ma_so_thue_ban", "MST bên bán"),
                         ("tong_thanh_toan", "Tổng thanh toán")):
        if getattr(inv, field) in (None, ""):
            add(f"Thiếu {label}", "critical" if field != "ma_so_thue_ban"
                else "check", field)

    # --- formats ---
    ok = _check_date(inv.ngay_lap)
    if ok is False:
        add(f"Ngày lập '{inv.ngay_lap}' không đúng định dạng DD/MM/YYYY",
            "check", "ngay_lap")
    elif not inv.ngay_lap:
        add("Thiếu ngày lập", "check", "ngay_lap")

    for label, mst in (("bên bán", inv.ma_so_thue_ban),
                       ("bên mua", inv.ma_so_thue_mua)):
        ok = check_mst(mst)
        if ok is False:
            add(f"MST {label} ({mst}) sai checksum — kiểm tra lại ký tự đọc được",
                "check",
                "ma_so_thue_ban" if label == "bên bán" else "ma_so_thue_mua")

    # --- line-item arithmetic: thanh_tien ≈ so_luong * don_gia ---
    for i, it in enumerate(inv.items, start=1):
        if (it.so_luong is not None and it.don_gia is not None
                and it.thanh_tien is not None):
            expected = it.so_luong * it.don_gia
            if abs(expected - it.thanh_tien) > TOL:
                add(f"Dòng {i}: {it.so_luong:g} × {it.don_gia:,.0f} = "
                    f"{expected:,.0f} nhưng thành tiền ghi {it.thanh_tien:,.0f}",
                    "check", "items")

    # --- totals ---
    if inv.items and inv.tong_tien_hang is not None:
        item_sum = sum(i.thanh_tien or 0 for i in inv.items)
        if abs(item_sum - inv.tong_tien_hang) > TOL:
            add(f"Tổng các dòng hàng ({item_sum:,.0f}) != cộng tiền hàng "
                f"({inv.tong_tien_hang:,.0f})", "critical", "tong_tien_hang")

    if inv.tong_tien_hang is not None and inv.tien_thue is not None \
            and inv.tong_thanh_toan is not None:
        expected = inv.tong_tien_hang + inv.tien_thue
        if abs(expected - inv.tong_thanh_toan) > TOL:
            add(f"Cộng tiền hàng + thuế ({expected:,.0f}) != tổng thanh toán "
                f"({inv.tong_thanh_toan:,.0f})", "critical", "tong_thanh_toan")

    # --- multi-tax breakdown consistency ---
    if inv.bang_ke_thue:
        s_hang = sum(t.tien_hang or 0 for t in inv.bang_ke_thue)
        s_thue = sum(t.tien_thue or 0 for t in inv.bang_ke_thue)
        if inv.tong_tien_hang is not None and abs(s_hang - inv.tong_tien_hang) > TOL:
            add(f"Bảng kê thuế: tổng tiền hàng ({s_hang:,.0f}) != cộng tiền hàng "
                f"({inv.tong_tien_hang:,.0f})", "check", "tong_tien_hang")
        if inv.tien_thue is not None and abs(s_thue - inv.tien_thue) > TOL:
            add(f"Bảng kê thuế: tổng thuế ({s_thue:,.0f}) != tiền thuế "
                f"({inv.tien_thue:,.0f})", "check", "tien_thue")
        rates = {t.thue_suat for t in inv.bang_ke_thue if t.thue_suat is not None}
        weird = rates - VALID_TAX_RATES
        if weird:
            add(f"Thuế suất lạ trong bảng kê: {sorted(weird)}% "
                f"(thường gặp: 0, 5, 8, 10)", "check", "tien_thue")
    elif inv.thue_suat is not None and inv.thue_suat not in VALID_TAX_RATES:
        add(f"Thuế suất {inv.thue_suat:g}% bất thường "
            f"(thường gặp: 0, 5, 8, 10)", "check", "thue_suat")

    # --- anomalies ---
    if inv.tong_thanh_toan is not None:
        if inv.tong_thanh_toan > 10_000_000_000:
            add("Tổng thanh toán trên 10 tỷ — kiểm tra lại số chữ số",
                "info", "tong_thanh_toan")
        elif inv.tong_thanh_toan <= 0:
            add("Tổng thanh toán <= 0 — có thể đọc sai số",
                "check", "tong_thanh_toan")

    # --- model-reported low confidence ---
    FIELD_LABELS = {
        "so_hoa_don": "Số hóa đơn", "ky_hieu": "Ký hiệu",
        "ngay_lap": "Ngày lập", "ma_so_thue_ban": "MST bên bán",
        "ma_so_thue_mua": "MST bên mua", "items": "Bảng hàng hóa",
        "tong_tien_hang": "Cộng tiền hàng", "tien_thue": "Tiền thuế",
        "tong_thanh_toan": "Tổng thanh toán", "thue_suat": "Thuế suất",
    }
    for f in inv.truong_can_kiem_tra:
        add(f"AI không chắc về trường '{FIELD_LABELS.get(f, f)}' "
            f"— nên đối chiếu với chứng từ gốc", "check", f)

    return w


# ---------------------------------------------------------------- Excel output

_THIN = Side(style="thin", color="999999")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
HEADER_FILL = PatternFill("solid", fgColor="DCE6F1")
CRIT_FILL = PatternFill("solid", fgColor="FFEDEA")
CHECK_FILL = PatternFill("solid", fgColor="FFF9E6")
NUMBER_FMT = '#,##0" ₫"'

# field name -> row number on the info sheet (1-based)
KV_LAYOUT = [
    ("ky_hieu", "Ký hiệu"),
    ("so_hoa_don", "Số hóa đơn"),
    ("ngay_lap", "Ngày lập"),
    ("ten_don_vi_ban", "Đơn vị bán"),
    ("ma_so_thue_ban", "MST bên bán"),
    ("dia_chi_ban", "Địa chỉ bên bán"),
    ("ten_nguoi_mua", "Người mua"),
    ("ten_don_vi_mua", "Đơn vị mua"),
    ("ma_so_thue_mua", "MST bên mua"),
    ("dia_chi_mua", "Địa chỉ bên mua"),
    ("hinh_thuc_thanh_toan", "Hình thức thanh toán"),
    ("tong_tien_hang", "Cộng tiền hàng"),
    ("thue_suat", "Thuế suất (%)"),
    ("tien_thue", "Tiền thuế GTGT"),
    ("tong_thanh_toan", "Tổng thanh toán"),
    ("so_tien_bang_chu", "Số tiền bằng chữ"),
]


def _flagged_fills(warnings: list[Warning]) -> dict[str, PatternFill]:
    """field -> fill for the worst severity touching that field."""
    fills: dict[str, PatternFill] = {}
    for w in warnings:
        if not w.field:
            continue
        if w.severity == "critical":
            fills[w.field] = CRIT_FILL
        elif w.field not in fills:
            fills[w.field] = CHECK_FILL
    return fills


def _kv_sheet(ws, inv: Invoice, fills):
    ws.title = "Thông tin hóa đơn"
    for r, (key, label) in enumerate(KV_LAYOUT, start=1):
        ws.cell(r, 1, label).font = Font(bold=True)
        v = getattr(inv, key)
        c = ws.cell(r, 2, v if v is not None else "")
        if isinstance(v, (int, float)):
            c.number_format = NUMBER_FMT
        if key in fills:
            c.fill = fills[key]
            c.font = Font(color="831C1E" if fills[key] is CRIT_FILL
                          else "92400E")
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 60


def _items_sheet(ws, inv: Invoice):
    ws.title = "Hàng hóa - Dịch vụ"
    headers = ["STT", "Tên hàng hóa, dịch vụ", "ĐVT",
               "Số lượng", "Đơn giá", "Thành tiền", "Thuế %"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(1, c, h)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center")

    for r, item in enumerate(inv.items, start=2):
        row = [r - 1, item.ten_hang, item.don_vi_tinh, item.so_luong,
               item.don_gia, item.thanh_tien, item.thue_suat]
        for c, v in enumerate(row, start=1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c in (4, 5, 6) and isinstance(v, (int, float)):
                cell.number_format = NUMBER_FMT

    for c, w in enumerate([6, 46, 10, 11, 15, 18, 9], start=1):
        ws.column_dimensions[get_column_letter(c)].width = w

    if inv.bang_ke_thue:
        r0 = len(inv.items) + 4
        ws.cell(r0, 2, "Bảng kê thuế GTGT").font = Font(bold=True)
        for i, t in enumerate(inv.bang_ke_thue, start=1):
            ws.cell(r0 + i, 2, f"Thuế {t.thue_suat:g}%"
                if t.thue_suat is not None else "Thuế ?")
            ws.cell(r0 + i, 5, t.tien_hang).number_format = NUMBER_FMT
            ws.cell(r0 + i, 6, t.tien_thue).number_format = NUMBER_FMT


def _warnings_sheet(ws, warnings: list[Warning], meta: dict):
    ws.title = "Cảnh báo & Ghi chú"
    ws.cell(1, 1, "Cảnh báo & ghi chú trích xuất").font = Font(bold=True)

    SEV_LABEL = {"critical": "Quan trọng", "check": "Nên kiểm tra",
                 "info": "Thông tin"}
    if warnings:
        for c, h in enumerate(["Mức độ", "Trường", "Nội dung"], start=1):
            cell = ws.cell(3, c, h)
            cell.font = Font(bold=True)
            cell.border = BORDER
        for r, w in enumerate(warnings, start=4):
            ws.cell(r, 1, SEV_LABEL.get(w.severity, w.severity))
            ws.cell(r, 2, w.field)
            ws.cell(r, 3, w.message)
            if w.severity == "critical":
                for c in range(1, 4):
                    ws.cell(r, c).fill = CRIT_FILL
    else:
        ws.cell(3, 1, "Không có cảnh báo — số liệu nhất quán.")

    r = 5 + len(warnings)
    ws.cell(r, 1, "Thông tin trích xuất").font = Font(bold=True)
    for i, (k, v) in enumerate(meta.items(), start=1):
        ws.cell(r + i, 1, k)
        ws.cell(r + i, 2, v)

    for c, w in enumerate([16, 20, 70], start=1):
        ws.column_dimensions[get_column_letter(c)].width = w


def write_excel(inv: Invoice, warnings: list[Warning], out: str | Path,
                meta: dict | None = None):
    """Template layout: info sheet + items sheet + warnings sheet."""
    wb = Workbook()
    fills = _flagged_fills(warnings)
    _kv_sheet(wb.active, inv, fills)
    _items_sheet(wb.create_sheet(), inv)
    _warnings_sheet(wb.create_sheet(), warnings, meta or {})
    wb.save(out)


def write_excel_flat(inv: Invoice, warnings: list[Warning], out: str | Path,
                     meta: dict | None = None):
    """Flat layout: one row per line item, invoice fields repeated —
    ready for accounting-software import."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Du lieu phang"
    headers = ["So hoa don", "Ky hieu", "Ngay lap", "MST ban", "MST mua",
               "Ten hang", "DVT", "So luong", "Don gia", "Thanh tien",
               "Thue %", "Tong tien hang", "Tien thue", "Tong thanh toan"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(1, c, h)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.border = BORDER

    for r, it in enumerate(inv.items or [InvoiceItem()], start=2):
        row = [inv.so_hoa_don, inv.ky_hieu, inv.ngay_lap,
               inv.ma_so_thue_ban, inv.ma_so_thue_mua,
               it.ten_hang, it.don_vi_tinh, it.so_luong, it.don_gia,
               it.thanh_tien, it.thue_suat,
               inv.tong_tien_hang, inv.tien_thue, inv.tong_thanh_toan]
        for c, v in enumerate(row, start=1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c in (8, 9, 10, 12, 13, 14) and isinstance(v, (int, float)):
                cell.number_format = NUMBER_FMT

    widths = [11, 10, 11, 13, 13, 42, 9, 10, 14, 15, 8, 15, 13, 16]
    for c, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c)].width = w
    _warnings_sheet(wb.create_sheet(), warnings, meta or {})
    wb.save(out)


# ---------------------------------------------------------------- logging

def log_extraction(record: dict):
    """Append one JSON line per extraction — survives without a DB."""
    LOG_DIR.mkdir(exist_ok=True)
    record = {"ts": datetime.now().isoformat(timespec="seconds"), **record}
    with (LOG_DIR / "extractions.jsonl").open("a", encoding="utf-8") as f:
        f.write(json.dumps(record, ensure_ascii=False) + "\n")


def read_extraction_stats() -> dict:
    path = LOG_DIR / "extractions.jsonl"
    if not path.exists():
        return {"total": 0, "ok": 0, "failed": 0,
                "success_rate": None, "by_mode": {}, "avg_ms": None}
    rows = [json.loads(l) for l in path.read_text("utf-8").splitlines() if l.strip()]
    ok = [r for r in rows if r.get("ok")]
    by_mode: dict[str, int] = {}
    for r in rows:
        by_mode[r.get("mode") or "?"] = by_mode.get(r.get("mode") or "?", 0) + 1
    durations = [r["duration_ms"] for r in ok if r.get("duration_ms")]
    return {
        "total": len(rows), "ok": len(ok), "failed": len(rows) - len(ok),
        "success_rate": round(len(ok) / len(rows) * 100, 1) if rows else None,
        "by_mode": by_mode,
        "avg_ms": int(sum(durations) / len(durations)) if durations else None,
        "last": rows[-1] if rows else None,
    }


# ---------------------------------------------------------------- convenience

def process(path: str | Path, out: str | Path, mode: str = "auto",
            layout: str = "template") -> dict:
    """Run the full pipeline; returns a summary dict (used by CLI)."""
    result = extract_invoice(path, mode)
    meta = {"file": str(path), "mode": result.mode,
            "extracted_at": datetime.now().isoformat(timespec="seconds")}
    writer = write_excel_flat if layout == "flat" else write_excel
    writer(result.invoice, result.warnings, out, meta)
    return {
        "invoice": json.loads(result.invoice.model_dump_json()),
        "warnings": [w.model_dump() for w in result.warnings],
        "mode": result.mode, "duration_ms": result.duration_ms,
        "excel": str(out),
    }
