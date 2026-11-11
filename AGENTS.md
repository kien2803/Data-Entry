# Data-Entry — SaaS trích xuất hóa đơn VN -> Excel

PoC pipeline: PDF/ảnh -> text layer hoặc vision -> Gemini -> JSON (Pydantic) -> xlsx.

## Commands

- Venv: `source venv/bin/activate` (hoặc `./venv/bin/python ...`)
- Tạo hóa đơn mẫu: `./venv/bin/python sample_invoice.py` -> `samples/hoa_don_mau.pdf`
- Chạy CLI: `./venv/bin/python cli.py samples/hoa_don_mau.pdf --json`
- Chạy API: `./venv/bin/uvicorn app:app --reload` -> POST `/extract` (multipart file) trả về xlsx
- Cần `GEMINI_API_KEY` trong `.env` (copy từ `.env.example`, lấy key free tại aistudio.google.com)

## Files

- `pipeline.py` — lõi: `build_contents` (text vs vision), `extract_invoice`, `validate_invoice`, `write_excel`
- `cli.py` — chạy 1 file từ terminal
- `app.py` — FastAPI: `/extract` (xlsx), `/extract/json` (JSON + preview_b64), `/export/xlsx` (build từ JSON đã sửa), `/sample*` (demo không cần key), `/stats`, mount `web/` tại `/`. Endpoints dùng `def` (threadpool) — KHÔNG dùng async def cho LLM calls đồng bộ (chặn event loop, kể cả SIGTERM).
- `web/` — UI tĩnh (index.html + app.css + app.js), style "achromatic-ledger" (taste-skill): ground #f2f4f7, Geist + Geist Mono cho số liệu, accent #1d5fbf chỉ trên tổng tiền
- `sample_invoice.py` — sinh PDF test có text layer tiếng Việt

## Notes

- PDF có text layer (>100 ký tự) -> gửi text; scan/ảnh -> render PNG gửi Gemini vision. Không dùng OCR riêng ở MVP.
- Checksum MST VN: weights [31,29,23,19,17,13,7,5,3], check = (10 - sum%11) % 10.
- Model: gemini-3.8-flash, structured output bằng Pydantic schema `Invoice`. Timeout 90s, retry 1 lần. (gemini-2.5-flash đã ngừng hỗ trợ user mới)
- Modes: auto|text|vision (param trên endpoints + segmented UI). Layout Excel: template|flat.
- Warnings là objects {message, severity(critical|check|info), field} — không phải strings.
- Log trích xuất: `logs/extractions.jsonl` (gitignored), xem qua `/stats`.
- E2E verified 11/11/2026: cả 2 mode (text + vision) trích đúng 100% trên hóa đơn mẫu.
