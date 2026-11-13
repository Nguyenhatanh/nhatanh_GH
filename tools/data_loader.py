"""
Tool 1: Tiếp nhận và kiểm tra dữ liệu
- Đọc Excel 4 sheet
- Kiểm tra mã ưu đãi hợp lệ
- Tách đơn hợp lệ / không hợp lệ
- Ghi nhật ký lỗi
"""

import pandas as pd
from typing import Any


def validate_data(file_path: str) -> dict[str, Any]:
    """Đọc file Excel, kiểm tra dữ liệu, trả về kết quả phân loại."""

    # Đọc 4 sheet
    giao_dich = pd.read_excel(file_path, sheet_name="Giao_Dich")
    doi_tra = pd.read_excel(file_path, sheet_name="Doi_Tra")
    san_pham = pd.read_excel(file_path, sheet_name="SanPham_GiaVon")
    chi_phi = pd.read_excel(file_path, sheet_name="ChiPhi_ChuongTrinh")
    return classify_orders(giao_dich, doi_tra, san_pham, chi_phi)


def classify_orders(giao_dich: pd.DataFrame, doi_tra: pd.DataFrame,
                    san_pham: pd.DataFrame, chi_phi: pd.DataFrame) -> dict[str, Any]:
    """Phân loại đơn hợp lệ / không hợp lệ từ các DataFrame đã đọc (đã làm sạch)."""

    # Danh sách mã ưu đãi hợp lệ
    valid_promos = set(chi_phi["promo_code"].tolist())

    # Phân loại đơn hàng
    nhat_ky = []  # Nhật ký lỗi
    don_hop_le = []
    don_khong_hop_le = []

    for _, row in giao_dich.iterrows():
        promo = row["promo_code"]

        # Kiểm tra mã trống (NaN)
        if pd.isna(promo) or str(promo).strip() == "":
            don_khong_hop_le.append({**row.to_dict(), "ly_do": "Mã ưu đãi trống"})
            nhat_ky.append(f"[LỖI] {row['order_id']}: Mã ưu đãi trống")
            continue

        promo_str = str(promo).strip()

        # Kiểm tra mã NONE
        if promo_str.upper() == "NONE":
            don_khong_hop_le.append({**row.to_dict(), "ly_do": "Mã ưu đãi = NONE"})
            nhat_ky.append(f"[LỖI] {row['order_id']}: Mã ưu đãi = NONE")
            continue

        # Tách mã (có thể nhiều mã phân cách bằng dấu phẩy)
        codes = [c.strip() for c in promo_str.split(",")]

        # Kiểm tra tất cả mã có hợp lệ không
        invalid_codes = [c for c in codes if c not in valid_promos]
        if invalid_codes:
            don_khong_hop_le.append({
                **row.to_dict(),
                "ly_do": f"Mã không hợp lệ: {', '.join(invalid_codes)}"
            })
            nhat_ky.append(
                f"[LỖI] {row['order_id']}: Mã không hợp lệ {invalid_codes}"
            )
            continue

        don_hop_le.append({**row.to_dict(), "promo_codes": codes})

    df_hop_le = pd.DataFrame(don_hop_le)
    df_khong_hop_le = pd.DataFrame(don_khong_hop_le)

    # Tổng kết
    summary = {
        "tong_don": len(giao_dich),
        "don_hop_le": len(df_hop_le),
        "don_khong_hop_le": len(df_khong_hop_le),
        "ma_trong": len([d for d in don_khong_hop_le if "trống" in d["ly_do"]]),
        "ma_none": len([d for d in don_khong_hop_le if "NONE" in d["ly_do"]]),
        "ma_khong_hop_le": len([d for d in don_khong_hop_le if "không hợp lệ" in d["ly_do"]]),
    }

    return {
        "don_hop_le": df_hop_le,
        "don_khong_hop_le": df_khong_hop_le,
        "doi_tra": doi_tra,
        "san_pham": san_pham,
        "chi_phi": chi_phi,
        "nhat_ky": nhat_ky,
        "summary": summary,
    }
