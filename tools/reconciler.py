"""
Tool 4: Đối soát và xuất kết quả
- Cộng tổng doanh thu phân bổ + đơn chưa phân bổ = tổng gốc
- Nếu khớp → xuất Excel 4 sheet
- Nếu không → báo lỗi
"""

import pandas as pd
from typing import Any
from pathlib import Path


def reconcile_and_export(data: dict[str, Any], output_path: str | None,
                         export: bool = True,
                         extra_sheets: dict[str, pd.DataFrame] | None = None) -> dict[str, Any]:
    """Đối soát tổng và (tuỳ chọn) xuất file kết quả.

    export=False: chỉ đối soát, không ghi file (dùng ở chặng xử lý trước khi xác nhận).
    extra_sheets: các sheet bổ sung ghi thêm vào file (vd: báo cáo làm sạch).
    """

    df_alloc = data["phan_bo"]
    df_invalid = data["don_khong_hop_le"]
    ket_qua = data["ket_qua_chuong_trinh"]
    nhat_ky = data["nhat_ky"]

    # Tính tổng đối soát
    # Phần 1: Doanh thu đã phân bổ (tổng gross_revenue của đơn hợp lệ)
    tong_phan_bo = df_alloc["gross_revenue"].sum()

    # Phần 2: Doanh thu chưa phân bổ (đơn không hợp lệ)
    if len(df_invalid) > 0 and "gross_revenue" in df_invalid.columns:
        tong_chua_phan_bo = df_invalid["gross_revenue"].sum()
    else:
        tong_chua_phan_bo = 0

    tong_doi_soat = tong_phan_bo + tong_chua_phan_bo

    # Tổng gốc: ưu tiên tổng doanh thu của dữ liệu đầu vào (sau làm sạch) nếu có
    if data.get("tong_goc_dau_vao") is not None:
        tong_goc = data["tong_goc_dau_vao"]
    else:
        tong_goc = data["don_hop_le"]["gross_revenue"].sum() + (
            df_invalid["gross_revenue"].sum() if len(df_invalid) > 0 else 0
        )

    sai_lech = abs(tong_doi_soat - tong_goc)
    doi_soat_ok = sai_lech < 1  # Cho phép sai lệch < 1đ do làm tròn

    doi_soat_info = {
        "tong_phan_bo": tong_phan_bo,
        "tong_chua_phan_bo": tong_chua_phan_bo,
        "tong_doi_soat": tong_doi_soat,
        "tong_goc": tong_goc,
        "sai_lech": sai_lech,
        "ket_qua": "KHỚP" if doi_soat_ok else "KHÔNG KHỚP",
    }

    if not doi_soat_ok:
        return {
            **data,
            "doi_soat": doi_soat_info,
            "xuat_file": False,
            "loi": f"Đối soát KHÔNG KHỚP: sai lệch {sai_lech:,.0f}đ",
        }

    if not export:
        return {**data, "doi_soat": doi_soat_info, "xuat_file": False}

    # Xuất Excel 4 sheet
    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)

    # Chuẩn bị sheet kết quả chương trình
    cols_ket_qua = [
        "promo_code", "program_name", "so_don", "doanh_thu_gop",
        "tong_doi_tra", "doanh_thu_thuan", "tong_gia_von", "hoan_gia_von",
        "gia_von_thuan", "chi_phi_co_dinh", "phan_dong_gop", "bien_dong_gop_pct",
    ]
    df_ket_qua = ket_qua[cols_ket_qua].copy()

    # Sheet đối soát
    df_doi_soat = pd.DataFrame([doi_soat_info])

    # Sheet nhật ký
    df_nhat_ky = pd.DataFrame({"nhat_ky": nhat_ky})

    # Sheet chưa phân bổ
    df_chua_phan_bo = df_invalid.copy() if len(df_invalid) > 0 else pd.DataFrame()

    with pd.ExcelWriter(str(output_file), engine="openpyxl") as writer:
        df_ket_qua.to_excel(writer, sheet_name="KetQua_ChuongTrinh", index=False)
        df_chua_phan_bo.to_excel(writer, sheet_name="Chua_PhanBo", index=False)
        df_doi_soat.to_excel(writer, sheet_name="DoiSoat", index=False)
        df_nhat_ky.to_excel(writer, sheet_name="NhatKy", index=False)
        for name, df in (extra_sheets or {}).items():
            df.to_excel(writer, sheet_name=name[:31], index=False)

    return {
        **data,
        "doi_soat": doi_soat_info,
        "xuat_file": True,
        "output_path": str(output_file),
    }
