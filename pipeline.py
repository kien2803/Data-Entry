"""Core pipeline: PDF/image -> text or vision -> Gemini -> structured data -> Excel.

Two extraction paths:
- PDF with a text layer  -> text is sent to the LLM (fast, cheap, accurate).
- Scanned PDF / image    -> rendered pages are sent to Gemini vision directly,
  so no separate OCR engine is needed for the MVP.
"""

import json
import os
import re
from pathlib import Path

import pymupdf as fitz
from dotenv import load_dotenv
from google import genai
from google.genai import types
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field

load_dotenv()

MODEL = "gemini-3.8-flash"
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp"}
MIME_BY_EXT = {
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".webp": "image/webp",
}


# ---------------------------------------------------------------- schema

class InvoiceItem(BaseModel):
    ten_hang: str = ""
    don_vi_tinh: str = ""
    so_luong: float | None = None
    don_gia: float | None = None
    thanh_tien: float | None = None


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
    tong_tien_hang: float | None = None
    thue_suat: float | None = None  # percent, e.g. 10 or 8
    tien_thue: float | None = None
    tong_thanh_toan: float | None = None
    so_tien_bang_chu: str = ""


# ---------------------------------------------------------------- input handling

def extract_text_from_pdf(path: str | Path) -> str:
    doc = fitz.open(path)
    try:
        return "\n".join(page.get_text() for page in doc).strip()
    finally:
        doc.close()


MAX_VISION_PAGES = 5


def pdf_to_page_images(path: str | Path, dpi: int = 200) -> list[bytes]:
    """Render PDF pages to PNG bytes (for scanned PDFs -> Gemini vision).
    Capped so a large scan cannot blow up token usage."""
    doc = fitz.open(path)
    try:
        return [
            page.get_pixmap(dpi=dpi).tobytes("png")
            for page in doc[:MAX_VISION_PAGES]
        ]
    finally:
        doc.close()


def build_contents(path: str | Path) -> tuple[str, list]:
    """Return (mode, contents_parts) for the LLM request.

    mode 'text'   -> contents = [prompt, extracted_text]
    mode 'vision' -> contents = [prompt, image_part, ...]
    """
    path = Path(path)
    ext = path.suffix.lower()
    prompt = EXTRACTION_PROMPT

    if ext == ".pdf":
        text = extract_text_from_pdf(path)
        if len(text) > 100:  # has a real text layer
            return "text", [prompt, f"---\n{text}\n---"]
        # scanned PDF -> send page images to vision
        parts = [types.Part.from_bytes(data=b, mime_type="image/png")
                 for b in pdf_to_page_images(path)]
        return "vision", [prompt, *parts]

    if ext in IMAGE_EXTS:
        return "vision", [
            prompt,
            types.Part.from_bytes(
                data=path.read_bytes(), mime_type=MIME_BY_EXT[ext]),
        ]

    raise ValueError(f"Unsupported file type: {ext}")


# ---------------------------------------------------------------- LLM extraction

EXTRACTION_PROMPT = """Bạn là hệ thống trích xuất dữ liệu hóa đơn Việt Nam (hóa đơn GTGT/VAT).

Đọc nội dung hóa đơn (văn bản hoặc hình ảnh) được cung cấp và trả về JSON đúng schema yêu cầu.

Quy tắc:
- Chỉ trích xuất thông tin CÓ TRONG tài liệu. Không suy đoán, không bịa.
- Trường không tìm thấy: để chuỗi rỗng "" (text) hoặc null (số).
- ngay_lap: định dạng DD/MM/YYYY.
- so_luong, don_gia, thanh_tien, tong_*: dạng số, bỏ dấu phân cách hàng nghìn và đơn vị tiền tệ.
- thue_suat: phần trăm thuế GTGT dạng số (ví dụ 10, 8, 5). Nếu "KCT" hoặc không chịu thuế: để null.
- items: mỗi dòng hàng hóa/dịch vụ là một phần tử, giữ nguyên thứ tự trên hóa đơn.
- ma_so_thue: chỉ giữ chữ số và dấu gạch (ví dụ "0101234567" hoặc "0101234567-001").
"""


def get_client() -> genai.Client:
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "Missing GEMINI_API_KEY. Copy .env.example to .env and set your key.")
    return genai.Client(api_key=api_key)


def extract_invoice(path: str | Path) -> tuple[Invoice, list[str]]:
    """Full pipeline for one file -> (Invoice, warnings)."""
    mode, contents = build_contents(path)
    client = get_client()
    resp = client.models.generate_content(
        model=MODEL,
        contents=contents,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=Invoice,
            temperature=0,
        ),
    )
    invoice = resp.parsed or Invoice.model_validate_json(resp.text)
    warnings = validate_invoice(invoice)
    warnings.append(f"mode={mode}")
    return invoice, warnings


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


def validate_invoice(inv: Invoice) -> list[str]:
    """Soft validation -> list of human-readable warnings."""
    w: list[str] = []
    TOL = 1000  # VND

    if inv.tong_tien_hang is not None and inv.tien_thue is not None \
            and inv.tong_thanh_toan is not None:
        expected = inv.tong_tien_hang + inv.tien_thue
        if abs(expected - inv.tong_thanh_toan) > TOL:
            w.append(
                f"Tổng tiền hàng + thuế ({expected:,.0f}) != tổng thanh toán "
                f"({inv.tong_thanh_toan:,.0f})")

    if inv.items and inv.tong_tien_hang is not None:
        item_sum = sum(i.thanh_tien or 0 for i in inv.items)
        if abs(item_sum - inv.tong_tien_hang) > TOL:
            w.append(
                f"Tổng các dòng hàng ({item_sum:,.0f}) != tổng tiền hàng "
                f"({inv.tong_tien_hang:,.0f})")

    for label, mst in (("bên bán", inv.ma_so_thue_ban),
                       ("bên mua", inv.ma_so_thue_mua)):
        ok = check_mst(mst)
        if ok is False:
            w.append(f"MST {label} ({mst}) không đúng checksum")

    for field_name, value in (("số hóa đơn", inv.so_hoa_don),
                              ("ngày lập", inv.ngay_lap)):
        if not value:
            w.append(f"Thiếu {field_name}")

    return w


# ---------------------------------------------------------------- Excel output

_THIN = Side(style="thin", color="999999")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
HEADER_FILL = PatternFill("solid", fgColor="DCE6F1")
NUMBER_FMT = "#,##0"


def _kv_sheet(ws, inv: Invoice):
    ws.title = "Thông tin hóa đơn"
    rows = [
        ("Ký hiệu", inv.ky_hieu),
        ("Số hóa đơn", inv.so_hoa_don),
        ("Ngày lập", inv.ngay_lap),
        ("Đơn vị bán", inv.ten_don_vi_ban),
        ("MST bên bán", inv.ma_so_thue_ban),
        ("Địa chỉ bên bán", inv.dia_chi_ban),
        ("Người mua", inv.ten_nguoi_mua),
        ("Đơn vị mua", inv.ten_don_vi_mua),
        ("MST bên mua", inv.ma_so_thue_mua),
        ("Địa chỉ bên mua", inv.dia_chi_mua),
        ("Hình thức thanh toán", inv.hinh_thuc_thanh_toan),
        ("Tổng tiền hàng", inv.tong_tien_hang),
        ("Thuế suất (%)", inv.thue_suat),
        ("Tiền thuế GTGT", inv.tien_thue),
        ("Tổng thanh toán", inv.tong_thanh_toan),
        ("Số tiền bằng chữ", inv.so_tien_bang_chu),
    ]
    for r, (k, v) in enumerate(rows, start=1):
        ws.cell(r, 1, k).font = Font(bold=True)
        c = ws.cell(r, 2, v)
        if isinstance(v, (int, float)):
            c.number_format = NUMBER_FMT
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 60


def _items_sheet(ws, inv: Invoice):
    ws.title = "Hàng hóa - Dịch vụ"
    headers = ["STT", "Tên hàng hóa, dịch vụ", "ĐVT",
               "Số lượng", "Đơn giá", "Thành tiền"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(1, c, h)
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center")

    for r, item in enumerate(inv.items, start=2):
        row = [r - 1, item.ten_hang, item.don_vi_tinh,
               item.so_luong, item.don_gia, item.thanh_tien]
        for c, v in enumerate(row, start=1):
            cell = ws.cell(r, c, v)
            cell.border = BORDER
            if c >= 4 and isinstance(v, (int, float)):
                cell.number_format = NUMBER_FMT

    for c, w in enumerate([6, 50, 10, 12, 15, 18], start=1):
        ws.column_dimensions[get_column_letter(c)].width = w


def write_excel(inv: Invoice, warnings: list[str], out: str | Path):
    wb = Workbook()
    _kv_sheet(wb.active, inv)
    _items_sheet(wb.create_sheet(), inv)

    soft_warnings = [w for w in warnings if not w.startswith("mode=")]
    if soft_warnings:
        ws = wb["Thông tin hóa đơn"]
        r = ws.max_row + 2
        ws.cell(r, 1, "CẢNH BÁO - CẦN KIỂM TRA LẠI").font = Font(
            bold=True, color="CC0000")
        for i, w in enumerate(soft_warnings, start=1):
            ws.cell(r + i, 1, f"- {w}")

    wb.save(out)


# ---------------------------------------------------------------- convenience

def process(path: str | Path, out: str | Path) -> dict:
    """Run the full pipeline; returns a summary dict (used by CLI + API)."""
    invoice, warnings = extract_invoice(path)
    write_excel(invoice, warnings, out)
    return {
        "invoice": json.loads(invoice.model_dump_json()),
        "warnings": warnings,
        "excel": str(out),
    }
