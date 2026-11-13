"""
Web server cho Agent Đánh Giá Chương Trình Bán Hàng — Công ty Gia Hòa.

Quy trình 5 chặng, mỗi chặng phải qua thì mới sang chặng sau:
  1. Nhập dữ liệu   — kiểm tra file đúng mẫu; sai → báo lỗi, bắt nhập lại
  2. Làm sạch       — xoá trùng, xử lý thiếu, giá trị sai...; chưa ưng → chỉnh & làm lại
  3. Xử lý          — phân loại mã KM, phân bổ doanh thu, tính kết quả, đối soát
  4. Xác nhận       — người dùng duyệt kết quả; chưa đúng → quay lại chặng trước
  5. Xuất báo cáo   — file Excel + dashboard
"""

import math
import os
import sys
import threading
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from flask import Flask, jsonify, request, send_from_directory
from werkzeug.utils import secure_filename

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from tools.data_loader import classify_orders  # noqa: E402
from tools.program_calculator import calculate_results  # noqa: E402
from tools.reconciler import reconcile_and_export  # noqa: E402
from tools.revenue_allocator import allocate_revenue  # noqa: E402
from tools.stages import DEFAULT_OPTIONS, OPTION_LABEL, check_input_file, clean_data  # noqa: E402

app = Flask(__name__, static_folder=None)
app.json.sort_keys = False  # giữ đúng thứ tự cột / quy tắc

DATA_FOLDER = BASE_DIR / "data"
UPLOAD_FOLDER = DATA_FOLDER / "uploads"
OUTPUT_FOLDER = BASE_DIR / "output"
OUTPUT_FILE = OUTPUT_FOLDER / "KetQua_GiaHoa.xlsx"
DEFAULT_FILE = DATA_FOLDER / "Du_Lieu_Gia_Hoa.xlsx"
for d in (DATA_FOLDER, UPLOAD_FOLDER, OUTPUT_FOLDER):
    d.mkdir(parents=True, exist_ok=True)

lock = threading.Lock()
S: dict = {}


def reset_session():
    S.clear()
    S.update({
        "stage": 1,          # chặng hiện tại (1..5)
        "file": None,
        "check": None,       # kết quả chặng 1
        "raw": None,
        "clean": None,       # kết quả chặng 2
        "process": None,     # kết quả chặng 3 (state pipeline)
        "final": None,       # kết quả chặng 5
    })


reset_session()


# ── Tiện ích ──

def sanitize(obj):
    """numpy/pandas → kiểu Python, NaN → None để JSON hợp lệ."""
    if isinstance(obj, dict):
        return {str(k): sanitize(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [sanitize(i) for i in obj]
    if isinstance(obj, np.integer):
        return int(obj)
    if isinstance(obj, (np.floating, float)):
        f = float(obj)
        return None if math.isnan(f) or math.isinf(f) else f
    if isinstance(obj, np.bool_):
        return bool(obj)
    if isinstance(obj, np.ndarray):
        return sanitize(obj.tolist())
    if obj is pd.NaT:
        return None
    return obj


def ok(**kw):
    return jsonify(sanitize({"ok": True, "stage": S["stage"], **kw}))


def fail(msg, code=400, **kw):
    errors = msg if isinstance(msg, list) else [msg]
    return jsonify(sanitize({"ok": False, "stage": S["stage"], "errors": errors, **kw})), code


def recommend(margin: float) -> tuple[str, str]:
    if margin >= 38:
        return "TIẾP TỤC", "Biên cao, hiệu quả tốt"
    if margin >= 37.5:
        return "ĐIỀU CHỈNH", "Biên trung bình, cần tối ưu chi phí"
    if margin >= 0:
        return "ĐIỀU CHỈNH", "Biên thấp, chi phí/đơn cao"
    return "DỪNG", "Chương trình lỗ"


def build_programs(kq: pd.DataFrame) -> list[dict]:
    out = []
    for _, r in kq.iterrows():
        m = float(r["bien_dong_gop_pct"]) if pd.notna(r["bien_dong_gop_pct"]) else 0.0
        rec, why = recommend(m)
        out.append({
            "promo_code": r["promo_code"],
            "program_name": r["program_name"],
            "so_don": int(r["so_don"]),
            "doanh_thu_gop": float(r["doanh_thu_gop"]),
            "tong_doi_tra": float(r["tong_doi_tra"]),
            "doanh_thu_thuan": float(r["doanh_thu_thuan"]),
            "gia_von_thuan": float(r["gia_von_thuan"]),
            "chi_phi_co_dinh": float(r["chi_phi_co_dinh"]),
            "phan_dong_gop": float(r["phan_dong_gop"]),
            "bien_dong_gop_pct": m,
            "khuyen_nghi": rec,
            "ly_do": why,
        })
    out.sort(key=lambda p: p["bien_dong_gop_pct"], reverse=True)
    return out


def clean_public(c: dict) -> dict:
    return {k: v for k, v in c.items() if k not in ("clean", "removed")} | {
        "removed_count": int(len(c["removed"])),
    }


# ── Routes ──

@app.route("/")
def index():
    return send_from_directory(str(BASE_DIR), "dashboard.html")


@app.route("/api/reset", methods=["POST"])
def api_reset():
    with lock:
        reset_session()
        return ok()


@app.route("/api/back", methods=["POST"])
def api_back():
    """Quay lại chặng trước; xoá kết quả các chặng sau."""
    to = int((request.get_json(silent=True) or {}).get("to", 1))
    with lock:
        if to < 1 or to > S["stage"]:
            return fail("Không thể chuyển tới chặng chưa mở")
        if to <= 1:
            reset_session()
        if to <= 2:
            S["process"] = None
        if to <= 3:
            S["process"] = None
        S["final"] = None
        S["stage"] = to
        return ok()


# ═══ CHẶNG 1: NHẬP & KIỂM TRA FILE ═══

def _do_check(path: Path, display_name: str):
    res = check_input_file(str(path))
    reset_session()
    S["file"] = display_name
    S["check"] = {k: v for k, v in res.items() if k != "raw"}
    if res["ok"]:
        S["raw"] = res["raw"]
        S["stage"] = 2
        return ok(file=display_name, **S["check"])
    return fail(res["errors"], 422, file=display_name, warnings=res["warnings"], sheets=res["sheets"])


@app.route("/api/s1/upload", methods=["POST"])
def s1_upload():
    f = request.files.get("file")
    if not f or not f.filename:
        return fail("Chưa chọn file")
    name = f.filename
    if not name.lower().endswith((".xlsx", ".xlsm", ".xls")):
        with lock:
            reset_session()
        return fail(f"\"{name}\" không phải file Excel. Chỉ nhận .xlsx / .xls.", 422, file=name)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = UPLOAD_FOLDER / f"{stamp}_{secure_filename(name) or 'upload.xlsx'}"
    f.save(path)
    with lock:
        return _do_check(path, name)


@app.route("/api/s1/default", methods=["POST"])
def s1_default():
    if not DEFAULT_FILE.exists():
        return fail("Chưa có file dữ liệu mẫu trong thư mục data/", 404)
    with lock:
        return _do_check(DEFAULT_FILE, DEFAULT_FILE.name)


# ═══ CHẶNG 2: LÀM SẠCH ═══

@app.route("/api/s2/options")
def s2_options():
    return ok(defaults=DEFAULT_OPTIONS, labels=OPTION_LABEL)


@app.route("/api/s2/clean", methods=["POST"])
def s2_clean():
    opts = (request.get_json(silent=True) or {}).get("options") or {}
    opts = {k: bool(v) for k, v in opts.items() if k in DEFAULT_OPTIONS}
    with lock:
        if S["stage"] < 2 or S["raw"] is None:
            return fail("Cần nhập file hợp lệ ở chặng 1 trước")
        c = clean_data(S["raw"], opts)
        S["clean"] = c
        S["process"] = None
        S["final"] = None
        S["stage"] = 2
        return ok(**clean_public(c))


@app.route("/api/s2/accept", methods=["POST"])
def s2_accept():
    with lock:
        if not S["clean"]:
            return fail("Chưa chạy làm sạch")
        if not S["clean"]["ok"]:
            return fail(S["clean"]["errors"])
        S["stage"] = 3
        return ok()


# ═══ CHẶNG 3: XỬ LÝ ═══

@app.route("/api/s3/process", methods=["POST"])
def s3_process():
    with lock:
        if S["stage"] < 3 or not S["clean"]:
            return fail("Cần hoàn tất chặng làm sạch trước")
        c = S["clean"]["clean"]
        steps = []
        try:
            st = classify_orders(c["Giao_Dich"], c["Doi_Tra"], c["SanPham_GiaVon"], c["ChiPhi_ChuongTrinh"])
            st["tong_goc_dau_vao"] = float(c["Giao_Dich"]["gross_revenue"].sum())
            s = st["summary"]
            if s["don_hop_le"] == 0:
                return fail("Không có đơn nào mang mã khuyến mãi hợp lệ để phân tích.")
            steps.append({"name": "Phân loại mã khuyến mãi", "detail": s})

            st = allocate_revenue(st)
            steps.append({"name": "Phân bổ doanh thu",
                          "detail": st["summary_phan_bo"].to_dict(orient="records")})

            st = calculate_results(st)
            steps.append({"name": "Tính kết quả chương trình",
                          "detail": {"so_chuong_trinh": int(len(st["ket_qua_chuong_trinh"]))}})

            st = reconcile_and_export(st, None, export=False)
            steps.append({"name": "Đối soát", "detail": st["doi_soat"]})
        except Exception as e:  # noqa: BLE001
            return fail(f"Lỗi ở bước \"{['Phân loại', 'Phân bổ', 'Tính kết quả', 'Đối soát'][len(steps)]}\": {e}",
                        500, steps=steps)

        S["process"] = st
        S["stage"] = 4
        return ok(steps=steps,
                  programs=build_programs(st["ket_qua_chuong_trinh"]),
                  summary=st["summary"],
                  doi_soat=st["doi_soat"],
                  chua_phan_bo=int(len(st["don_khong_hop_le"])))


# ═══ CHẶNG 4: XÁC NHẬN → CHẶNG 5: XUẤT ═══

@app.route("/api/s4/confirm", methods=["POST"])
def s4_confirm():
    with lock:
        if S["stage"] < 4 or not S["process"]:
            return fail("Cần chạy xử lý trước khi xác nhận")
        st = dict(S["process"])
        if st["doi_soat"]["ket_qua"] != "KHỚP":
            return fail("Đối soát KHÔNG KHỚP — không thể xác nhận. Hãy quay lại kiểm tra dữ liệu.")

        cl = S["clean"]
        lam_sach = pd.DataFrame([
            {"quy_tac": s["label"], "sheet": s["sheet"], "so_dong": s["count"],
             "hanh_dong": s["action"], "bat": "Có" if s["enabled"] else "Không"}
            for s in cl["steps"]
        ])
        extra = {"LamSach_BaoCao": lam_sach}
        if len(cl["removed"]):
            extra["DuLieu_BiLoai"] = cl["removed"]
        extra["DuLieu_DaLamSach"] = cl["clean"]["Giao_Dich"]

        st["nhat_ky"] = list(cl["log"]) + list(st["nhat_ky"])
        try:
            out = reconcile_and_export(st, str(OUTPUT_FILE), export=True, extra_sheets=extra)
        except PermissionError:
            return fail("Không ghi được file kết quả — hãy đóng file KetQua_GiaHoa.xlsx nếu đang mở trong Excel.", 500)

        S["final"] = {
            "file": S["file"],
            "generated_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "programs": build_programs(out["ket_qua_chuong_trinh"]),
            "summary": out["summary"],
            "doi_soat": out["doi_soat"],
            "clean": {"rows_before": cl["rows_before"], "rows_after": cl["rows_after"],
                      "removed_count": int(len(cl["removed"]))},
            "output_file": OUTPUT_FILE.name,
        }
        S["stage"] = 5
        return ok(**S["final"])


@app.route("/api/download")
def download():
    if not OUTPUT_FILE.exists():
        return jsonify({"error": "Chưa có file kết quả"}), 404
    return send_from_directory(str(OUTPUT_FOLDER), OUTPUT_FILE.name, as_attachment=True)


if __name__ == "__main__":
    print("=" * 50)
    print("  GIA HOA - Agent Danh Gia Chuong Trinh")
    print("  http://localhost:5000")
    print("=" * 50)
    app.run(host="0.0.0.0", port=5000, debug=True, threaded=True)
