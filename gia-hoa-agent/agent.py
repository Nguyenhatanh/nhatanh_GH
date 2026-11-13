"""
Agent đánh giá chương trình bán hàng - Công ty Gia Hòa
Sử dụng Google Gemini API để điều phối 4 bước xử lý dữ liệu
và đưa ra nhận định kinh doanh.
"""

import json
import os
import sys
from pathlib import Path
from google import genai
from google.genai import types

from tools.data_loader import validate_data
from tools.revenue_allocator import allocate_revenue
from tools.program_calculator import calculate_results
from tools.reconciler import reconcile_and_export

# ============================================================
# Cấu hình
# ============================================================
DATA_FILE = "data/Du_Lieu_Gia_Hoa.xlsx"
OUTPUT_FILE = "output/KetQua_GiaHoa.xlsx"

# Đọc API key từ biến môi trường
API_KEY = os.environ.get("GEMINI_API_KEY")
if not API_KEY:
    print("[LỖI] Chưa đặt biến môi trường GEMINI_API_KEY")
    print("Chạy: export GEMINI_API_KEY='your-key-here'")
    sys.exit(1)

# ============================================================
# Định nghĩa tool cho Gemini (function calling)
# ============================================================

validate_data_decl = types.FunctionDeclaration(
    name="validate_data",
    description=(
        "Bước 1: Đọc file Excel 4 sheet, kiểm tra mã ưu đãi hợp lệ. "
        "Tách đơn hợp lệ và không hợp lệ (mã trống, mã NONE). "
        "Ghi nhật ký lỗi phát hiện."
    ),
    parameters=types.Schema(
        type=types.Type.OBJECT,
        properties={
            "file_path": types.Schema(
                type=types.Type.STRING,
                description="Đường dẫn file Excel đầu vào",
            ),
        },
        required=["file_path"],
    ),
)

allocate_revenue_decl = types.FunctionDeclaration(
    name="allocate_revenue",
    description=(
        "Bước 2: Phân bổ doanh thu theo chương trình. "
        "Đơn 1 mã → gán trọn. Đơn nhiều mã → chia đều 1/N. "
        "Trừ đổi trả theo mã đơn hàng."
    ),
    parameters=types.Schema(
        type=types.Type.OBJECT,
        properties={},
    ),
)

calculate_results_decl = types.FunctionDeclaration(
    name="calculate_results",
    description=(
        "Bước 3: Tính kết quả kinh doanh từng chương trình. "
        "Doanh thu thuần, giá vốn thuần, chi phí cố định, "
        "phần đóng góp và biên đóng góp (%)."
    ),
    parameters=types.Schema(
        type=types.Type.OBJECT,
        properties={},
    ),
)

reconcile_export_decl = types.FunctionDeclaration(
    name="reconcile_and_export",
    description=(
        "Bước 4: Đối soát tổng doanh thu (phân bổ + chưa phân bổ = tổng gốc). "
        "Nếu khớp → xuất Excel 4 sheet. Nếu không → báo lỗi."
    ),
    parameters=types.Schema(
        type=types.Type.OBJECT,
        properties={
            "output_path": types.Schema(
                type=types.Type.STRING,
                description="Đường dẫn file Excel kết quả đầu ra",
            ),
        },
        required=["output_path"],
    ),
)

# ============================================================
# Trạng thái pipeline (lưu giữ data giữa các bước)
# ============================================================
pipeline_state: dict = {}


def execute_tool(function_call) -> str:
    """Thực thi tool và trả kết quả dạng text cho Gemini."""
    global pipeline_state
    name = function_call.name
    args = dict(function_call.args) if function_call.args else {}

    try:
        if name == "validate_data":
            file_path = args.get("file_path", DATA_FILE)
            pipeline_state = validate_data(file_path)
            s = pipeline_state["summary"]
            return json.dumps({
                "status": "OK",
                "tong_don": s["tong_don"],
                "don_hop_le": s["don_hop_le"],
                "don_khong_hop_le": s["don_khong_hop_le"],
                "ma_trong": s["ma_trong"],
                "ma_none": s["ma_none"],
                "so_dong_nhat_ky": len(pipeline_state["nhat_ky"]),
            }, ensure_ascii=False)

        elif name == "allocate_revenue":
            if not pipeline_state:
                return '{"error": "Chưa chạy validate_data"}'
            pipeline_state = allocate_revenue(pipeline_state)
            sa = pipeline_state["summary_phan_bo"]
            return sa.to_json(orient="records", force_ascii=False)

        elif name == "calculate_results":
            if "phan_bo" not in pipeline_state:
                return '{"error": "Chưa chạy allocate_revenue"}'
            pipeline_state = calculate_results(pipeline_state)
            kq = pipeline_state["ket_qua_chuong_trinh"]
            # Chỉ gửi các cột quan trọng
            cols = [
                "promo_code", "program_name", "so_don",
                "doanh_thu_thuan", "gia_von_thuan", "chi_phi_co_dinh",
                "phan_dong_gop", "bien_dong_gop_pct",
            ]
            return kq[cols].to_json(orient="records", force_ascii=False)

        elif name == "reconcile_and_export":
            if "ket_qua_chuong_trinh" not in pipeline_state:
                return '{"error": "Chưa chạy calculate_results"}'
            output = args.get("output_path", OUTPUT_FILE)
            pipeline_state = reconcile_and_export(pipeline_state, output)
            ds = pipeline_state["doi_soat"]
            return json.dumps({
                "doi_soat": ds["ket_qua"],
                "tong_phan_bo": ds["tong_phan_bo"],
                "tong_chua_phan_bo": ds["tong_chua_phan_bo"],
                "tong_doi_soat": ds["tong_doi_soat"],
                "sai_lech": ds["sai_lech"],
                "xuat_file": pipeline_state["xuat_file"],
                "output_path": pipeline_state.get("output_path", ""),
            }, ensure_ascii=False)

        else:
            return f'{{"error": "Tool không tồn tại: {name}"}}'

    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


# ============================================================
# Agent chính
# ============================================================

SYSTEM_PROMPT = """\
Bạn là chuyên viên phân tích kinh doanh của Công ty Gia Hòa.

Nhiệm vụ: Đánh giá hiệu quả 4 chương trình bán hàng (PROMO_A → D) để Trưởng phòng
Kinh doanh quyết định chương trình nào nên tiếp tục, điều chỉnh, hoặc dừng.

Bạn có 4 công cụ phải gọi TUẦN TỰ:
1. validate_data       → Kiểm tra dữ liệu đầu vào
2. allocate_revenue    → Phân bổ doanh thu & trừ đổi trả
3. calculate_results   → Tính kết quả từng chương trình
4. reconcile_and_export → Đối soát & xuất Excel

Sau khi hoàn tất 4 bước, hãy:
- Tóm tắt kết quả từng chương trình (doanh thu thuần, phần đóng góp, biên %)
- Xếp hạng chương trình từ tốt nhất đến kém nhất
- Đưa ra khuyến nghị cụ thể: TIẾP TỤC / ĐIỀU CHỈNH / DỪNG cho từng chương trình
- Giải thích lý do dựa trên số liệu

Luôn trả lời bằng tiếng Việt. Trình bày rõ ràng, có số liệu cụ thể.
"""


def run_agent():
    """Chạy agent với Gemini function calling."""
    print("=" * 60)
    print("  AGENT ĐÁNH GIÁ CHƯƠNG TRÌNH BÁN HÀNG - GIA HÒA")
    print("=" * 60)

    # Khởi tạo Gemini client
    client = genai.Client(api_key=API_KEY)

    tool = types.Tool(function_declarations=[
        validate_data_decl,
        allocate_revenue_decl,
        calculate_results_decl,
        reconcile_export_decl,
    ])

    config = types.GenerateContentConfig(
        system_instruction=SYSTEM_PROMPT,
        tools=[tool],
        temperature=0.1,  # Giảm tính ngẫu nhiên cho phân tích số liệu
    )

    # Tin nhắn khởi đầu
    user_msg = (
        f"Hãy phân tích file dữ liệu '{DATA_FILE}' và đánh giá hiệu quả "
        f"4 chương trình bán hàng của Công ty Gia Hòa. "
        f"Chạy tuần tự 4 bước xử lý, sau đó đưa ra khuyến nghị."
    )

    print(f"\n[USER] {user_msg}\n")

    contents = [types.Content(role="user", parts=[types.Part.from_text(text=user_msg)])]

    # Vòng lặp function calling
    max_turns = 10
    for turn in range(max_turns):
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents=contents,
            config=config,
        )

        # Xử lý response
        has_function_call = False

        for candidate in response.candidates:
            for part in candidate.content.parts:
                # Nếu là function call
                if part.function_call:
                    has_function_call = True
                    fc = part.function_call
                    print(f"[TOOL] Gọi: {fc.name}({dict(fc.args) if fc.args else ''})")

                    # Thực thi tool
                    result = execute_tool(fc)
                    print(f"[TOOL] Kết quả: {result[:200]}...")

                    # Thêm response của model vào history
                    contents.append(candidate.content)

                    # Thêm kết quả tool vào history
                    contents.append(
                        types.Content(
                            role="user",
                            parts=[types.Part.from_function_response(
                                name=fc.name,
                                response=json.loads(result),
                            )],
                        )
                    )

                # Nếu là text response (kết luận cuối)
                elif part.text:
                    if not has_function_call:
                        print(f"\n{'=' * 60}")
                        print("  BÁO CÁO PHÂN TÍCH")
                        print(f"{'=' * 60}")
                        print(part.text)
                        return part.text

        if not has_function_call:
            # Nếu không có function call và không có text → kết thúc
            break

    print("\n[CẢNH BÁO] Agent kết thúc sau", max_turns, "lượt")
    return None


if __name__ == "__main__":
    run_agent()
