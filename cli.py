"""CLI: python cli.py input.pdf -o output.xlsx [--mode auto|text|vision] [--layout template|flat]"""

import argparse
import sys
from datetime import datetime
from pathlib import Path

from pipeline import extract_invoice, write_excel, write_excel_flat

SEV_ICON = {"critical": "!!", "check": "! ", "info": "- "}


def main():
    p = argparse.ArgumentParser(
        description="Trích xuất hóa đơn PDF/ảnh -> Excel")
    p.add_argument("input", help="Đường dẫn file PDF hoặc ảnh hóa đơn")
    p.add_argument("-o", "--output", help="File Excel đầu ra "
                   "(mặc định: <input>.xlsx)")
    p.add_argument("--mode", default="auto",
                   choices=["auto", "text", "vision"])
    p.add_argument("--layout", default="template",
                   choices=["template", "flat"])
    p.add_argument("--json", action="store_true",
                   help="In JSON kết quả ra stdout")
    args = p.parse_args()

    src = Path(args.input)
    if not src.exists():
        sys.exit(f"File không tồn tại: {src}")
    out = Path(args.output) if args.output else src.with_suffix(".xlsx")

    try:
        result = extract_invoice(src, args.mode)
    except RuntimeError as e:
        sys.exit(str(e))

    meta = {"file": str(src), "mode": result.mode,
            "extracted_at": datetime.now().isoformat(timespec="seconds")}
    writer = write_excel_flat if args.layout == "flat" else write_excel
    writer(result.invoice, result.warnings, out, meta)

    print(f"Đã xuất: {out}  (mode={result.mode}, {result.duration_ms}ms)")
    for w in result.warnings:
        print(f"  [{SEV_ICON.get(w.severity, '?')}] {w.message}")

    if args.json:
        print(result.invoice.model_dump_json(indent=2))


if __name__ == "__main__":
    main()
