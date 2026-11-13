"""
Tool 2: Xử lý đơn nhiều ưu đãi và đổi trả
- Đơn 1 mã → gán trọn cho chương trình đó
- Đơn nhiều mã → chia đều 1/N
- Trừ đổi trả theo mã đơn hàng
"""

import pandas as pd
from typing import Any


def allocate_revenue(data: dict[str, Any]) -> dict[str, Any]:
    """Phân bổ doanh thu theo chương trình và trừ đổi trả."""

    df = data["don_hop_le"]
    doi_tra = data["doi_tra"]
    san_pham = data["san_pham"]

    # Tạo dict giá vốn theo product_id
    gia_von_map = dict(zip(san_pham["product_id"], san_pham["unit_cost"]))

    # Tạo dict đổi trả theo order_id
    doi_tra_map = dict(zip(doi_tra["order_id"], doi_tra["return_amount"]))

    # Phân bổ doanh thu
    records = []
    for _, row in df.iterrows():
        codes = row["promo_codes"]
        n = len(codes)
        order_id = row["order_id"]
        gross = row["gross_revenue"]
        product_id = row["product_id"]
        quantity = row["quantity"]
        unit_cost = gia_von_map.get(product_id, 0)

        # Số tiền đổi trả (nếu có)
        return_amt = doi_tra_map.get(order_id, 0)

        for code in codes:
            records.append({
                "order_id": order_id,
                "product_id": product_id,
                "promo_code": code,
                "quantity": quantity,
                "gross_revenue": gross / n,          # Chia đều doanh thu gộp
                "unit_cost": unit_cost,
                "total_cost": unit_cost * quantity / n,  # Chia đều giá vốn
                "return_amount": return_amt / n,     # Chia đều đổi trả
                "so_chuong_trinh": n,
            })

    df_allocated = pd.DataFrame(records)

    # Tổng kết phân bổ
    summary_alloc = df_allocated.groupby("promo_code").agg(
        so_dong=("order_id", "count"),
        doanh_thu_gop=("gross_revenue", "sum"),
        gia_von=("total_cost", "sum"),
        doi_tra=("return_amount", "sum"),
    ).reset_index()

    return {
        **data,
        "phan_bo": df_allocated,
        "summary_phan_bo": summary_alloc,
    }
