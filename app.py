"""FastAPI app: serves the web UI + extraction endpoints.

POST /extract       -> xlsx download        (form: file, mode, layout)
POST /extract/json  -> invoice JSON         (form: file, mode)
POST /export/xlsx   -> xlsx from edited JSON (body: {invoice, layout})
GET  /sample        -> bundled sample invoice (no API key needed)
GET  /sample/xlsx   -> xlsx of the sample invoice
GET  /sample/file   -> the sample PDF itself
GET  /sample/preview.png -> rendered page 1 of the sample PDF
GET  /stats         -> extraction log stats

Run: uvicorn app:app --reload
"""

import base64
import tempfile
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

import pymupdf
from fastapi import FastAPI, File, Form, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from pipeline import (IMAGE_EXTS, Invoice, InvoiceItem, TaxLine,
                      extract_invoice, log_extraction,
                      read_extraction_stats, validate_invoice,
                      write_excel, write_excel_flat)

app = FastAPI(title="DocLift — trích xuất chứng từ")

ALLOWED = IMAGE_EXTS | {".pdf"}
MODES = {"auto", "text", "vision"}
LAYOUTS = {"template", "flat"}
MAX_SIZE_MB = 15
XLSX_MIME = ("application/vnd.openxmlformats"
             "-officedocument.spreadsheetml.sheet")
WEB_DIR = Path(__file__).parent / "web"
SAMPLE_PDF = Path(__file__).parent / "samples" / "hoa_don_mau.pdf"

# Data matching samples/hoa_don_mau.pdf — lets the demo work without an API key.
SAMPLE_INVOICE = Invoice(
    ky_hieu="1C25TAA", so_hoa_don="0001234", ngay_lap="05/11/2026",
    ten_don_vi_ban="CÔNG TY TNHH THƯƠNG MẠI VÀ DỊCH VỤ ABC",
    ma_so_thue_ban="0101248141",
    dia_chi_ban="Số 12 Đường Láng, Quận Đống Đa, Hà Nội",
    ten_nguoi_mua="Nguyễn Văn An",
    ten_don_vi_mua="CÔNG TY CỔ PHẦN XYZ",
    ma_so_thue_mua="0100109106",
    dia_chi_mua="456 Lê Lợi, Quận 1, TP. Hồ Chí Minh",
    hinh_thuc_thanh_toan="Chuyển khoản",
    items=[
        InvoiceItem(ten_hang="Máy in laser HP M404dn", don_vi_tinh="Chiếc",
                    so_luong=2, don_gia=4500000, thanh_tien=9000000),
        InvoiceItem(ten_hang="Giấy A4 Double A 80gsm", don_vi_tinh="Thùng",
                    so_luong=5, don_gia=250000, thanh_tien=1250000),
        InvoiceItem(ten_hang="Dịch vụ bảo trì thiết bị", don_vi_tinh="Lần",
                    so_luong=1, don_gia=1500000, thanh_tien=1500000),
    ],
    tong_tien_hang=11750000, thue_suat=10, tien_thue=1175000,
    tong_thanh_toan=12925000,
    so_tien_bang_chu="Mười hai triệu chín trăm hai mươi lăm nghìn đồng",
)


def _write_and_stream(invoice: Invoice, warnings, meta: dict,
                      layout: str, name: str):
    with tempfile.TemporaryDirectory() as tmp:
        out = Path(tmp) / name
        writer = write_excel_flat if layout == "flat" else write_excel
        writer(invoice, warnings, out, meta)
        xlsx_bytes = out.read_bytes()
    return StreamingResponse(
        iter([xlsx_bytes]), media_type=XLSX_MIME,
        headers={"Content-Disposition": f'attachment; filename="{name}"'})


def _read_upload(file: UploadFile) -> tuple[str, bytes]:
    ext = Path(file.filename or "").suffix.lower()
    if ext not in ALLOWED:
        raise HTTPException(
            400, f"Chỉ hỗ trợ: {', '.join(sorted(ALLOWED))}")
    data = file.file.read()
    if len(data) > MAX_SIZE_MB * 1024 * 1024:
        raise HTTPException(400, f"File vượt quá {MAX_SIZE_MB} MB")
    return ext, data


def _check_file_opens(ext: str, data: bytes):
    """Reject corrupt/truncated files with a friendly 400 before the LLM call."""
    try:
        # open() alone doesn't decode image data — render a probe pixmap
        doc = pymupdf.open(stream=data, filetype=ext.lstrip("."))
        doc[0].get_pixmap(dpi=10)
        doc.close()
    except Exception:
        raise HTTPException(
            400, "File không đọc được — file hỏng hoặc sai định dạng.")


def _run_extraction(ext: str, data: bytes, mode: str):
    _check_file_opens(ext, data)
    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / f"input{ext}"
        src.write_bytes(data)
        try:
            return extract_invoice(src, mode)
        except RuntimeError as e:
            raise HTTPException(500, str(e))


def _meta(file_name: str, result) -> dict:
    return {"file": file_name, "mode": result.mode,
            "duration_ms": result.duration_ms,
            "extracted_at": datetime.now().isoformat(timespec="seconds")}


@app.post("/extract")
def extract(file: UploadFile = File(...), mode: str = Form("auto"),
            layout: str = Form("template")):
    if mode not in MODES:
        raise HTTPException(400, f"mode phải là: {sorted(MODES)}")
    if layout not in LAYOUTS:
        raise HTTPException(400, f"layout phải là: {sorted(LAYOUTS)}")
    ext, data = _read_upload(file)
    try:
        result = _run_extraction(ext, data, mode)
    except HTTPException as e:
        log_extraction({"file": file.filename, "mode": mode, "ok": False,
                        "error": str(e.detail)})
        raise
    log_extraction({"file": file.filename, "mode": result.mode, "ok": True,
                    "warnings": len(result.warnings),
                    "duration_ms": result.duration_ms})
    soft = [w for w in result.warnings if w.severity != "info"]
    resp = _write_and_stream(result.invoice, result.warnings,
                             _meta(file.filename or "?", result),
                             layout, "ket_qua.xlsx")
    if soft:
        resp.headers["X-Extraction-Warnings"] = quote(
            "; ".join(w.message for w in soft))
    return resp


@app.post("/extract/json")
def extract_json(file: UploadFile = File(...), mode: str = Form("auto")):
    if mode not in MODES:
        raise HTTPException(400, f"mode phải là: {sorted(MODES)}")
    ext, data = _read_upload(file)
    try:
        result = _run_extraction(ext, data, mode)
    except HTTPException as e:
        log_extraction({"file": file.filename, "mode": mode, "ok": False,
                        "error": str(e.detail)})
        raise
    log_extraction({"file": file.filename, "mode": result.mode, "ok": True,
                    "warnings": len(result.warnings),
                    "duration_ms": result.duration_ms})
    return {"invoice": result.invoice.model_dump(),
            "warnings": [w.model_dump() for w in result.warnings],
            "mode": result.mode,
            "duration_ms": result.duration_ms,
            "file_name": file.filename,
            "preview_b64": _pdf_preview_b64(data) if ext == ".pdf" else None}


class ExportRequest(BaseModel):
    invoice: Invoice
    layout: str = "template"


@app.post("/export/xlsx")
def export_xlsx(req: ExportRequest):
    """Build xlsx from (possibly user-edited) invoice JSON.
    Re-validates server-side so edited errors still get flagged."""
    if req.layout not in LAYOUTS:
        raise HTTPException(400, f"layout phải là: {sorted(LAYOUTS)}")
    warnings = validate_invoice(req.invoice)
    meta = {"source": "user-edited",
            "extracted_at": datetime.now().isoformat(timespec="seconds")}
    return _write_and_stream(req.invoice, warnings, meta,
                             req.layout, "ket_qua.xlsx")


def _pdf_preview_b64(data: bytes) -> str | None:
    """First page of an uploaded PDF as a PNG data payload for the UI."""
    try:
        doc = pymupdf.open(stream=data, filetype="pdf")
        png = doc[0].get_pixmap(dpi=100).tobytes("png")
        doc.close()
        return base64.b64encode(png).decode()
    except Exception:
        return None


@app.get("/sample")
def sample():
    return {"invoice": SAMPLE_INVOICE.model_dump(),
            "warnings": [w.model_dump()
                         for w in validate_invoice(SAMPLE_INVOICE)],
            "mode": "sample"}


@app.get("/sample/xlsx")
def sample_xlsx(layout: str = Query("template")):
    if layout not in LAYOUTS:
        raise HTTPException(400, f"layout phải là: {sorted(LAYOUTS)}")
    return _write_and_stream(
        SAMPLE_INVOICE, validate_invoice(SAMPLE_INVOICE),
        {"source": "sample", "file": "hoa_don_mau.pdf"},
        layout, "mau.xlsx")


@app.get("/sample/file")
def sample_file():
    if not SAMPLE_PDF.exists():
        raise HTTPException(404, "Chưa có file mẫu — chạy sample_invoice.py")
    return FileResponse(SAMPLE_PDF, media_type="application/pdf")


@app.get("/sample/preview.png")
def sample_preview():
    if not SAMPLE_PDF.exists():
        raise HTTPException(404, "Chưa có file mẫu — chạy sample_invoice.py")
    doc = pymupdf.open(SAMPLE_PDF)
    png = doc[0].get_pixmap(dpi=110).tobytes("png")
    doc.close()
    return Response(png, media_type="image/png")


@app.get("/stats")
def stats():
    return read_extraction_stats()


@app.get("/health")
def health():
    return {"status": "ok"}


app.mount("/", StaticFiles(directory=WEB_DIR, html=True), name="web")
