"""
Tool 3: Tính kết quả từng chương trình
- Doanh thu thuần = Doanh thu gộp - Đổi trả
- Giá vốn thuần = Giá vốn - Hoàn giá vốn hàng trả
- Chi phí cố định theo bảng (A:3tr, B:5tr, C:4tr, D:2tr)
- Phần đóng góp = Thuần - Giá vốn thuần - Chi phí
- Biên đóng góp (%)
"""

import pandas as pd
from typing import Any


def calculate_results(data: dict[str, Any]) -> dict[str, Any]:
    """Tính kết quả kinh doanh từng chương trình."""

    df_alloc = data["phan_bo"]
    chi_phi = data["chi_phi"]
    san_pham = data["san_pham"]
    doi_tra = data["doi_tra"]

    # Dict giá vốn
    gia_von_map = dict(zip(san_pham["product_id"], san_pham["unit_cost"]))

    # Dict chi phí cố định
    chi_phi_map = dict(zip(chi_phi["promo_code"], chi_phi["fixed_cost"]))

    # Xác định đơn nào bị đổi trả
    don_doi_tra = set(doi_tra["order_id"].tolist())

    # Tính giá vốn hoàn lại cho hàng trả
    # Chỉ hoàn giá vốn nếu đơn đó nằm trong danh sách đổi trả
    df_alloc = df_alloc.copy()
    df_alloc["is_returned"] = df_alloc["order_id"].isin(don_doi_tra)
    df_alloc["hoan_gia_von"] = df_alloc.apply(
        lambda r: r["total_cost"] if r["is_returned"] else 0, axis=1
    )

    # Tổng hợp theo chương trình
    by_program = df_alloc.groupby("promo_code").agg(
        doanh_thu_gop=("gross_revenue", "sum"),
        tong_gia_von=("total_cost", "sum"),
        tong_doi_tra=("return_amount", "sum"),
        hoan_gia_von=("hoan_gia_von", "sum"),
        so_don=("order_id", "nunique"),
        so_dong_phan_bo=("order_id", "count"),
    ).reset_index()

    # Tính các chỉ số
    by_program["doanh_thu_thuan"] = by_program["doanh_thu_gop"] - by_program["tong_doi_tra"]
    by_program["gia_von_thuan"] = by_program["tong_gia_von"] - by_program["hoan_gia_von"]
    by_program["chi_phi_co_dinh"] = by_program["promo_code"].map(chi_phi_map)
    by_program["phan_dong_gop"] = (
        by_program["doanh_thu_thuan"]
        - by_program["gia_von_thuan"]
        - by_program["chi_phi_co_dinh"]
    )
    by_program["bien_dong_gop_pct"] = (
        by_program["phan_dong_gop"] / by_program["doanh_thu_thuan"] * 100
    ).round(2)

    # Merge tên chương trình
    by_program = by_program.merge(
        chi_phi[["promo_code", "program_name"]], on="promo_code", how="left"
    )

    return {
        **data,
        "ket_qua_chuong_trinh": by_program,
    }
