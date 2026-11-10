"""FastAPI endpoint: POST /extract (multipart file) -> xlsx download.

Run: uvicorn app:app --reload
"""

import tempfile
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import StreamingResponse

from pipeline import IMAGE_EXTS, extract_invoice, write_excel

app = FastAPI(title="Invoice -> Excel extractor")

ALLOWED = IMAGE_EXTS | {".pdf"}
MAX_SIZE_MB = 15


@app.post("/extract")
async def extract(file: UploadFile = File(...)):
    ext = Path(file.filename or "").suffix.lower()
    if ext not in ALLOWED:
        raise HTTPException(400, f"Chỉ hỗ trợ: {', '.join(sorted(ALLOWED))}")

    data = await file.read()
    if len(data) > MAX_SIZE_MB * 1024 * 1024:
        raise HTTPException(400, f"File vượt quá {MAX_SIZE_MB} MB")

    with tempfile.TemporaryDirectory() as tmp:
        src = Path(tmp) / f"input{ext}"
        src.write_bytes(data)
        try:
            invoice, warnings = extract_invoice(src)
        except RuntimeError as e:
            raise HTTPException(500, str(e))
        out = Path(tmp) / "ket_qua.xlsx"
        write_excel(invoice, warnings, out)
        xlsx_bytes = out.read_bytes()

    soft = [w for w in warnings if not w.startswith("mode=")]
    headers = {"X-Extraction-Warnings": "; ".join(soft)} if soft else {}
    return StreamingResponse(
        iter([xlsx_bytes]),
        media_type="application/vnd.openxmlformats"
                   "-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": 'attachment; filename="ket_qua.xlsx"',
            **headers,
        },
    )


@app.get("/health")
def health():
    return {"status": "ok"}
