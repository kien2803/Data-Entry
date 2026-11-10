"""CLI: python cli.py input.pdf -o output.xlsx"""

import argparse
import json
import sys
from pathlib import Path

from pipeline import extract_invoice, write_excel


def main():
    p = argparse.ArgumentParser(
        description="Trích xuất hóa đơn PDF/ảnh -> Excel")
    p.add_argument("input", help="Đường dẫn file PDF hoặc ảnh hóa đơn")
    p.add_argument("-o", "--output", help="File Excel đầu ra "
                   "(mặc định: <input>.xlsx)")
    p.add_argument("--json", action="store_true",
                   help="In JSON kết quả ra stdout")
    args = p.parse_args()

    src = Path(args.input)
    if not src.exists():
        sys.exit(f"File không tồn tại: {src}")
    out = Path(args.output) if args.output else src.with_suffix(".xlsx")

    try:
        invoice, warnings = extract_invoice(src)
    except RuntimeError as e:
        sys.exit(str(e))

    write_excel(invoice, warnings, out)
    print(f"Đã xuất: {out}")

    soft = [w for w in warnings if not w.startswith("mode=")]
    for w in soft:
        print(f"  [!] {w}")

    if args.json:
        print(invoice.model_dump_json(indent=2))


if __name__ == "__main__":
    main()
