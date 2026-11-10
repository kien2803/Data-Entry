"""Generate a sample Vietnamese VAT invoice PDF (with real text layer)
to test the pipeline without needing real documents.

Run: python sample_invoice.py
"""

from pathlib import Path

import pymupdf as fitz

FONT_CANDIDATES = [
    "/usr/share/fonts/google-noto-vf/NotoSans[wght].ttf",
    "/usr/share/fonts/open-sans/OpenSans-Regular.ttf",
    "/usr/share/fonts/liberation-sans-fonts/LiberationSans-Regular.ttf",
]

LINES = [
    ("HOÁ ĐƠN GIÁ TRỊ GIA TĂNG", 16),
    ("(VAT INVOICE)", 10),
    ("Ký hiệu: 1C25TAA", 10),
    ("Số: 0001234", 10),
    ("Ngày 05 tháng 11 năm 2026", 10),
    ("", 10),
    ("Đơn vị bán hàng: CÔNG TY TNHH THƯƠNG MẠI VÀ DỊCH VỤ ABC", 10),
    ("Mã số thuế: 0101248141", 10),
    ("Địa chỉ: Số 12 Đường Láng, Quận Đống Đa, Hà Nội", 10),
    ("", 10),
    ("Họ tên người mua: Nguyễn Văn An", 10),
    ("Tên đơn vị mua: CÔNG TY CỔ PHẦN XYZ", 10),
    ("Mã số thuế: 0100109106", 10),
    ("Địa chỉ: 456 Lê Lợi, Quận 1, TP. Hồ Chí Minh", 10),
    ("Hình thức thanh toán: Chuyển khoản", 10),
    ("", 10),
    ("STT | Tên hàng hóa, dịch vụ | ĐVT | Số lượng | Đơn giá | Thành tiền", 9),
    ("1 | Máy in laser HP M404dn | Chiếc | 2 | 4500000 | 9000000", 9),
    ("2 | Giấy A4 Double A 80gsm | Thùng | 5 | 250000 | 1250000", 9),
    ("3 | Dịch vụ bảo trì thiết bị | Lần | 1 | 1500000 | 1500000", 9),
    ("", 10),
    ("Cộng tiền hàng: 11.750.000", 10),
    ("Thuế suất GTGT: 10%   Tiền thuế GTGT: 1.175.000", 10),
    ("Tổng tiền thanh toán: 12.925.000", 10),
    ("Số tiền viết bằng chữ: Mười hai triệu chín trăm hai mươi lăm nghìn đồng", 10),
]


def main():
    font_path = next((p for p in FONT_CANDIDATES if Path(p).exists()), None)
    if not font_path:
        raise SystemExit("Không tìm thấy font hỗ trợ tiếng Việt")

    out_dir = Path("samples")
    out_dir.mkdir(exist_ok=True)
    out = out_dir / "hoa_don_mau.pdf"

    doc = fitz.open()
    page = doc.new_page()
    page.insert_font(fontname="noto", fontfile=font_path)

    y = 50
    for text, size in LINES:
        page.insert_text((50, y), text, fontname="noto", fontsize=size)
        y += size + 8

    doc.save(out)
    doc.close()
    print(f"Đã tạo: {out}")


if __name__ == "__main__":
    main()
