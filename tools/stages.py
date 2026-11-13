"""
Chặng 1 & 2 của quy trình:
- Chặng 1: Kiểm tra cấu trúc file đầu vào (sheet, cột, kiểu dữ liệu)
- Chặng 2: Làm sạch dữ liệu (chuẩn hoá, xoá trùng, xử lý thiếu, giá trị sai...)
"""

from __future__ import annotations

from typing import Any

import pandas as pd

# Cấu trúc bắt buộc: sheet -> {cột: kiểu} ("text" | "number")
SCHEMA: dict[str, dict[str, str]] = {
    "Giao_Dich": {
        "order_id": "text",
        "customer_id": "text",
        "product_id": "text",
        "quantity": "number",
        "gross_revenue": "number",
        "promo_code": "text",
    },
    "Doi_Tra": {
        "order_id": "text",
        "return_amount": "number",
    },
    "SanPham_GiaVon": {
        "product_id": "text",
        "product_name": "text",
        "unit_cost": "number",
    },
    "ChiPhi_ChuongTrinh": {
        "promo_code": "text",
        "program_name": "text",
        "fixed_cost": "number",
    },
}

SHEET_LABEL = {
    "Giao_Dich": "Giao dịch",
    "Doi_Tra": "Đổi trả",
    "SanPham_GiaVon": "Sản phẩm & giá vốn",
    "ChiPhi_ChuongTrinh": "Chi phí chương trình",
}

# Các tuỳ chọn làm sạch mặc định (người dùng bật/tắt được ở giao diện)
DEFAULT_OPTIONS: dict[str, bool] = {
    "chuan_hoa": True,              # Cắt khoảng trắng, viết hoa mã
    "xoa_trung_hoan_toan": True,    # Xoá dòng trùng 100%
    "xoa_trung_ma_don": True,       # Trùng order_id nhưng khác nội dung -> giữ dòng đầu
    "xoa_thieu_bat_buoc": True,     # Thiếu order_id/product_id/quantity/gross_revenue
    "dien_khach_hang_trong": True,  # customer_id trống -> "KH_KHONG_RO"
    "xoa_gia_tri_sai": True,        # quantity <= 0, doanh thu âm, không phải số
    "xoa_san_pham_la": True,        # product_id không có trong bảng giá vốn
    "gop_doi_tra_trung": True,      # Đổi trả trùng order_id -> cộng dồn
    "xoa_doi_tra_mo_coi": True,     # Đổi trả không có đơn gốc hợp lệ
    "xoa_doi_tra_vuot": False,      # Tiền trả > doanh thu đơn
}

OPTION_LABEL = {
    "chuan_hoa": "Chuẩn hoá: cắt khoảng trắng, viết hoa mã SP / mã KM",
    "xoa_trung_hoan_toan": "Xoá dòng giao dịch trùng lặp hoàn toàn",
    "xoa_trung_ma_don": "Mã đơn bị lặp (khác nội dung) → chỉ giữ dòng đầu tiên",
    "xoa_thieu_bat_buoc": "Xoá dòng thiếu mã đơn / mã SP / số lượng / doanh thu",
    "dien_khach_hang_trong": "Mã khách hàng trống → điền \"KH_KHONG_RO\"",
    "xoa_gia_tri_sai": "Xoá dòng có số lượng ≤ 0 hoặc doanh thu âm / không phải số",
    "xoa_san_pham_la": "Xoá dòng có mã SP không nằm trong bảng giá vốn",
    "gop_doi_tra_trung": "Đổi trả trùng mã đơn → cộng dồn thành 1 dòng",
    "xoa_doi_tra_mo_coi": "Xoá đổi trả không khớp với đơn hàng hợp lệ nào",
    "xoa_doi_tra_vuot": "Xoá đổi trả có số tiền trả lớn hơn doanh thu của đơn",
}


def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def _is_blank(s: pd.Series) -> pd.Series:
    return s.isna() | (s.astype(str).str.strip() == "") | (s.astype(str).str.lower() == "nan")


# ════════════════════════════════════════════════════════════
# CHẶNG 1: KIỂM TRA CẤU TRÚC FILE
# ════════════════════════════════════════════════════════════

def check_input_file(file_path: str) -> dict[str, Any]:
    """Kiểm tra file có đúng mẫu không. Trả về ok, errors, warnings, sheets, raw."""
    errors: list[str] = []
    warnings: list[str] = []
    sheets_info: list[dict[str, Any]] = []
    raw: dict[str, pd.DataFrame] = {}

    try:
        book = pd.read_excel(file_path, sheet_name=None)
    except Exception as e:  # noqa: BLE001
        return {
            "ok": False,
            "errors": [f"Không đọc được file Excel ({type(e).__name__}). "
                       "File có thể bị hỏng, có mật khẩu hoặc không phải định dạng .xlsx."],
            "warnings": [], "sheets": [], "raw": {},
        }

    # So khớp tên sheet không phân biệt hoa thường / khoảng trắng
    lookup = {str(n).strip().lower(): n for n in book}

    for sheet, cols in SCHEMA.items():
        real = lookup.get(sheet.lower())
        if real is None:
            errors.append(f"Thiếu sheet \"{sheet}\" ({SHEET_LABEL[sheet]}).")
            sheets_info.append({"sheet": sheet, "label": SHEET_LABEL[sheet], "status": "missing"})
            continue

        df = _norm_cols(book[real])
        missing = [c for c in cols if c not in df.columns]
        info = {
            "sheet": sheet, "label": SHEET_LABEL[sheet], "rows": int(len(df)),
            "cols": list(df.columns), "missing_cols": missing, "status": "ok",
        }

        if missing:
            errors.append(f"Sheet \"{sheet}\" thiếu cột: {', '.join(missing)}.")
            info["status"] = "error"
        elif len(df) == 0 and sheet != "Doi_Tra":
            errors.append(f"Sheet \"{sheet}\" không có dòng dữ liệu nào.")
            info["status"] = "error"
        else:
            # Cột số phải đa số là số
            for c, kind in cols.items():
                if kind != "number" or len(df) == 0:
                    continue
                nonblank = df[c][~_is_blank(df[c])]
                bad = pd.to_numeric(nonblank, errors="coerce").isna().sum()
                if len(nonblank) and bad / len(nonblank) > 0.5:
                    errors.append(f"Sheet \"{sheet}\", cột \"{c}\" phải là số "
                                  f"nhưng {bad}/{len(nonblank)} ô không phải số.")
                    info["status"] = "error"
                elif bad:
                    warnings.append(f"Sheet \"{sheet}\", cột \"{c}\": {bad} ô không phải số "
                                    "(sẽ xử lý ở bước làm sạch).")
            nulls = {c: int(_is_blank(df[c]).sum()) for c in cols if c in df.columns}
            info["nulls"] = {k: v for k, v in nulls.items() if v}

        raw[sheet] = df
        sheets_info.append(info)

    extra = [n for n in book if str(n).strip().lower() not in {s.lower() for s in SCHEMA}]
    if extra:
        warnings.append(f"Bỏ qua các sheet không dùng tới: {', '.join(map(str, extra))}.")

    return {
        "ok": not errors,
        "errors": errors,
        "warnings": warnings,
        "sheets": sheets_info,
        "raw": raw if not errors else {},
    }


# ════════════════════════════════════════════════════════════
# CHẶNG 2: LÀM SẠCH DỮ LIỆU
# ════════════════════════════════════════════════════════════

def clean_data(raw: dict[str, pd.DataFrame], options: dict[str, bool] | None = None) -> dict[str, Any]:
    """Làm sạch dữ liệu theo các tuỳ chọn. Không thay đổi `raw`."""
    opt = {**DEFAULT_OPTIONS, **(options or {})}
    gd = raw["Giao_Dich"].copy()
    dt = raw["Doi_Tra"].copy()
    sp = raw["SanPham_GiaVon"].copy()
    cp = raw["ChiPhi_ChuongTrinh"].copy()

    steps: list[dict[str, Any]] = []
    removed_frames: list[pd.DataFrame] = []
    log: list[str] = []

    def record(key: str, sheet: str, affected: pd.DataFrame, action: str, removed: bool = True):
        n = int(len(affected))
        steps.append({
            "key": key, "label": OPTION_LABEL.get(key, key), "sheet": sheet,
            "count": n, "action": action, "enabled": opt.get(key, True),
            "sample": affected.head(5).astype(str).to_dict(orient="records"),
        })
        if n and removed:
            r = affected.copy()
            r.insert(0, "ly_do_loai", OPTION_LABEL.get(key, key))
            r.insert(0, "sheet", sheet)
            removed_frames.append(r)
        if n:
            log.append(f"[{sheet}] {OPTION_LABEL.get(key, key)}: {n} dòng")

    rows_before = {"Giao_Dich": len(gd), "Doi_Tra": len(dt),
                   "SanPham_GiaVon": len(sp), "ChiPhi_ChuongTrinh": len(cp)}

    # ── 0. Chuẩn hoá ──
    text_cols = {"Giao_Dich": gd, "Doi_Tra": dt, "SanPham_GiaVon": sp, "ChiPhi_ChuongTrinh": cp}
    changed = 0
    if opt["chuan_hoa"]:
        for name, df in text_cols.items():
            for c, kind in SCHEMA[name].items():
                if kind == "text" and c in df.columns:
                    before = df[c].copy()
                    mask = ~before.isna()
                    s = before[mask].astype(str).str.strip()
                    if c in ("order_id", "customer_id", "product_id", "promo_code"):
                        s = s.str.upper().str.replace(r"\s*,\s*", ",", regex=True)
                    df.loc[mask, c] = s
                    changed += int((before[mask].astype(str) != s).sum())
    steps.append({"key": "chuan_hoa", "label": OPTION_LABEL["chuan_hoa"], "sheet": "Tất cả",
                  "count": changed, "action": "sửa", "enabled": opt["chuan_hoa"], "sample": []})
    if changed:
        log.append(f"[Tất cả] Chuẩn hoá {changed} ô")

    # Ép kiểu số (ô không phải số -> NaN)
    for name, df in text_cols.items():
        for c, kind in SCHEMA[name].items():
            if kind == "number":
                df[c] = pd.to_numeric(df[c], errors="coerce")

    # Bảng danh mục: xoá trùng khoá (giữ dòng đầu)
    for name, key in (("SanPham_GiaVon", "product_id"), ("ChiPhi_ChuongTrinh", "promo_code")):
        df = text_cols[name]
        dup = df[df.duplicated(subset=[key], keep="first")]
        if len(dup):
            text_cols[name] = df.drop(dup.index)
            r = dup.copy(); r.insert(0, "ly_do_loai", f"Trùng {key} trong danh mục"); r.insert(0, "sheet", name)
            removed_frames.append(r)
            log.append(f"[{name}] Xoá {len(dup)} dòng trùng {key}")
    sp, cp = text_cols["SanPham_GiaVon"], text_cols["ChiPhi_ChuongTrinh"]

    # ── 1. Trùng hoàn toàn ──
    dup_full = gd[gd.duplicated(keep="first")]
    record("xoa_trung_hoan_toan", "Giao_Dich", dup_full, "xoá", removed=opt["xoa_trung_hoan_toan"])
    if opt["xoa_trung_hoan_toan"]:
        gd = gd.drop(dup_full.index)

    # ── 2. Thiếu trường bắt buộc ──
    req = ["order_id", "product_id", "quantity", "gross_revenue"]
    miss_mask = pd.Series(False, index=gd.index)
    for c in req:
        miss_mask |= _is_blank(gd[c])
    record("xoa_thieu_bat_buoc", "Giao_Dich", gd[miss_mask], "xoá", removed=opt["xoa_thieu_bat_buoc"])
    if opt["xoa_thieu_bat_buoc"]:
        gd = gd[~miss_mask]

    # ── 3. Trùng mã đơn (khác nội dung) ──
    dup_id = gd[gd["order_id"].notna() & gd.duplicated(subset=["order_id"], keep="first")]
    record("xoa_trung_ma_don", "Giao_Dich", dup_id, "xoá", removed=opt["xoa_trung_ma_don"])
    if opt["xoa_trung_ma_don"]:
        gd = gd.drop(dup_id.index)

    # ── 4. Giá trị sai ──
    bad_mask = (gd["quantity"] <= 0) | (gd["gross_revenue"] < 0)
    if not opt["xoa_thieu_bat_buoc"]:
        bad_mask |= gd["quantity"].isna() | gd["gross_revenue"].isna()
    record("xoa_gia_tri_sai", "Giao_Dich", gd[bad_mask], "xoá", removed=opt["xoa_gia_tri_sai"])
    if opt["xoa_gia_tri_sai"]:
        gd = gd[~bad_mask]

    # ── 5. Sản phẩm lạ ──
    known_sp = set(sp["product_id"].dropna())
    la_mask = gd["product_id"].notna() & ~gd["product_id"].isin(known_sp)
    record("xoa_san_pham_la", "Giao_Dich", gd[la_mask], "xoá", removed=opt["xoa_san_pham_la"])
    if opt["xoa_san_pham_la"]:
        gd = gd[~la_mask]

    # ── 6. Khách hàng trống ──
    kh_mask = _is_blank(gd["customer_id"])
    record("dien_khach_hang_trong", "Giao_Dich", gd[kh_mask], "điền", removed=False)
    if opt["dien_khach_hang_trong"]:
        gd.loc[kh_mask, "customer_id"] = "KH_KHONG_RO"

    # ── 7. Đổi trả ──
    dt_blank = _is_blank(dt["order_id"]) | dt["return_amount"].isna() | (dt["return_amount"] <= 0)
    if dt_blank.any():
        r = dt[dt_blank].copy(); r.insert(0, "ly_do_loai", "Đổi trả thiếu mã đơn / số tiền không hợp lệ")
        r.insert(0, "sheet", "Doi_Tra"); removed_frames.append(r)
        log.append(f"[Doi_Tra] Xoá {int(dt_blank.sum())} dòng thiếu mã đơn / số tiền ≤ 0")
        dt = dt[~dt_blank]

    dup_dt = dt[dt.duplicated(subset=["order_id"], keep=False)]
    record("gop_doi_tra_trung", "Doi_Tra", dup_dt, "gộp", removed=False)
    if opt["gop_doi_tra_trung"] and len(dup_dt):
        dt = dt.groupby("order_id", as_index=False, sort=False)["return_amount"].sum()

    valid_orders = gd.groupby("order_id")["gross_revenue"].sum()
    mo_coi = dt[~dt["order_id"].isin(valid_orders.index)]
    record("xoa_doi_tra_mo_coi", "Doi_Tra", mo_coi, "xoá", removed=opt["xoa_doi_tra_mo_coi"])
    if opt["xoa_doi_tra_mo_coi"]:
        dt = dt[dt["order_id"].isin(valid_orders.index)]

    gross_of = dt["order_id"].map(valid_orders)
    vuot = dt[gross_of.notna() & (dt["return_amount"] > gross_of)]
    record("xoa_doi_tra_vuot", "Doi_Tra", vuot, "xoá", removed=opt["xoa_doi_tra_vuot"])
    if opt["xoa_doi_tra_vuot"]:
        dt = dt.drop(vuot.index)

    # ── Thống kê mã khuyến mãi (không xoá: đơn không mã -> "chưa phân bổ") ──
    valid_promos = set(cp["promo_code"].dropna())
    promo = gd["promo_code"]
    p_blank = _is_blank(promo)
    p_none = ~p_blank & (promo.astype(str).str.upper() == "NONE")

    def _unknown(v):
        return any(c.strip() not in valid_promos for c in str(v).split(","))

    p_unknown = ~p_blank & ~p_none & promo.apply(_unknown)
    promo_stats = {
        "trong": int(p_blank.sum()),
        "none": int(p_none.sum()),
        "ma_la": int(p_unknown.sum()),
        "nhieu_ma": int((~p_blank & promo.astype(str).str.contains(",")).sum()),
    }

    gd = gd.reset_index(drop=True)
    dt = dt.reset_index(drop=True)
    gd["quantity"] = gd["quantity"].astype("int64", errors="ignore")

    removed = pd.concat(removed_frames, ignore_index=True) if removed_frames else pd.DataFrame()

    rows_after = {"Giao_Dich": len(gd), "Doi_Tra": len(dt),
                  "SanPham_GiaVon": len(sp), "ChiPhi_ChuongTrinh": len(cp)}

    errors = []
    if len(gd) == 0:
        errors.append("Sau khi làm sạch không còn giao dịch nào. Hãy kiểm tra lại file hoặc tắt bớt quy tắc.")

    return {
        "ok": not errors,
        "errors": errors,
        "options": opt,
        "steps": steps,
        "promo_stats": promo_stats,
        "rows_before": rows_before,
        "rows_after": rows_after,
        "tong_doanh_thu": float(gd["gross_revenue"].sum()),
        "log": log,
        "removed": removed,
        "preview": gd.head(15).astype(str).to_dict(orient="records"),
        "clean": {"Giao_Dich": gd, "Doi_Tra": dt, "SanPham_GiaVon": sp, "ChiPhi_ChuongTrinh": cp},
    }
