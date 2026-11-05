import io
import json
import sqlite3
import threading
import time

# Thư viện xuất Excel và giao diện
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st

# 1. Cấu hình giao diện trang web chuẩn Sản phẩm App Chuyên Nghiệp
st.set_page_config(
    page_title="DataClean Pro - Hệ Thống Xử Lý & Kiểm Tra Dữ Liệu Thông Minh",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS giao diện Dark Mode hiện đại dạng Card
st.markdown(
    """
    <style>
    .stApp {
        background-color: #121212;
        color: #e0e0e0;
    }
    [data-testid="stSidebar"] {
        background-color: #1a1a1a;
        border-right: 1px solid #2d2d2d;
    }
    div.block-container {
        background-color: #1e1e1e;
        padding: 2.5rem 3rem;
        border-radius: 16px;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.4);
        margin-top: 1.5rem;
        margin-bottom: 2rem;
        border: 1px solid #2d2d2d;
    }
    div.block-container p, div.block-container span, div.block-container label, div.block-container div {
        color: #e0e0e0 !important;
    }
    h1, h2, h3, h4 {
        color: #ffffff !important;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 12px;
        background-color: #252525;
        padding: 8px 12px;
        border-radius: 12px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 10px 24px;
        font-weight: 600;
        color: #b0b0b0;
        border: none;
        background-color: #2a2a2a;
        transition: all 0.3s ease;
    }
    .stTabs [aria-selected="true"] {
        background-color: #3b82f6 !important;
        color: #ffffff !important;
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3);
    }
    .stTabs [aria-selected="true"] p {
        color: #ffffff !important;
    }
    div[data-testid="stMetric"] {
        background-color: #252525;
        padding: 18px 22px;
        border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        border-left: 5px solid #3b82f6;
        border-top: 1px solid #333333;
        border-right: 1px solid #333333;
        border-bottom: 1px solid #333333;
    }
    div[data-testid="stMetric"] label {
        color: #9ca3af !important;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #ffffff !important;
    }
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        padding: 10px 24px;
        background-color: #3b82f6;
        color: #ffffff;
        border: none;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #2563eb;
    }
    
    /* Nút bấm Menu Sidebar giả lập đường dẫn Tab */
    div[data-testid="stSidebarNav"] {display: none;}
    </style>
""",
    unsafe_allow_html=True,
)

DB_NAME = "dataclean_app_integrated.db"


# 2. Khởi tạo Cơ sở dữ liệu SQLite
def init_db():
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS import_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            total_rows INTEGER,
            processed_rows INTEGER DEFAULT 0,
            valid_rows INTEGER DEFAULT 0,
            invalid_rows INTEGER DEFAULT 0,
            status TEXT DEFAULT 'ĐANG CHẠY',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS valid_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER,
            row_number INTEGER,
            data_json TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS invalid_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER,
            row_number INTEGER,
            raw_data TEXT,
            error_reason TEXT
        )
    """)

    conn.commit()
    conn.close()


init_db()


# 3. Các hàm xuất Excel chuyên nghiệp
def export_history_to_excel(df_history):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Lich_Su_Phien"

    ws.merge_cells("A1:I1")
    title_cell = ws["A1"]
    title_cell.value = "DANH SÁCH LỊCH SỬ CÁC PHIÊN NHẬP DỮ LIỆU - DATACLEAN PRO"
    title_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(
        start_color="1E3C72", end_color="1E3C72", fill_type="solid"
    )
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    headers = [
        "Mã phiên",
        "Tên tệp dữ liệu",
        "Tổng số dòng",
        "Số dòng đã xử lý",
        "Số dòng hợp lệ",
        "Số dòng bị lỗi",
        "Tỉ lệ sai",
        "Trạng thái",
        "Thời gian tạo",
    ]
    ws.append([])
    ws.append(headers)

    header_fill = PatternFill(
        start_color="2A5298", end_color="2A5298", fill_type="solid"
    )
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )

    ws.row_dimensions[3].height = 25
    for col_num in range(1, 10):
        cell = ws.cell(row=3, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    for row_idx, row in enumerate(df_history.itertuples(), start=4):
        session_id_str = f"#{row.id}"
        total = row.total_rows if row.total_rows > 0 else 1
        error_rate = f"{(row.invalid_rows / total) * 100:.2f}%"

        row_data = [
            session_id_str,
            row.filename,
            row.total_rows,
            row.processed_rows,
            row.valid_rows,
            row.invalid_rows,
            error_rate,
            row.status,
            str(row.created_at),
        ]
        ws.append(row_data)
        ws.row_dimensions[row_idx].height = 22

        for col_idx in range(1, 10):
            c = ws.cell(row=row_idx, column=col_idx)
            c.border = thin_border
            if col_idx in [1, 3, 4, 5, 6, 7, 8, 9]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    col_widths = [12, 32, 16, 18, 16, 16, 14, 16, 22]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def export_marked_file_to_excel(session_id, original_df):
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    err_records = pd.read_sql_query(
        f"SELECT row_number, error_reason FROM invalid_data WHERE session_id = {session_id}",
        conn,
    )
    conn.close()

    error_dict = {}
    if not err_records.empty:
        for _, r in err_records.iterrows():
            error_dict[int(r["row_number"])] = r["error_reason"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Kiem_Tra_Loi_File_Goc"

    export_df = original_df.copy()
    export_df.insert(
        len(export_df.columns), "Ghi chú kiểm tra lỗi", "Hợp lệ"
    )

    headers = list(export_df.columns)
    ws.append(headers)

    header_fill = PatternFill(
        start_color="1E3C72", end_color="1E3C72", fill_type="solid"
    )
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    ws.row_dimensions[1].height = 25

    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    red_fill = PatternFill(
        start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"
    )
    red_font = Font(name="Calibri", size=10, color="9C0006", bold=True)

    for i, row in export_df.iterrows():
        row_num_db = i + 1
        row_values = list(row)

        if row_num_db in error_dict:
            row_values[-1] = f"LỖI: {error_dict[row_num_db]}"
        else:
            row_values[-1] = "Hợp lệ"

        ws.append(row_values)
        current_row_idx = i + 2
        ws.row_dimensions[current_row_idx].height = 20

        is_error = row_num_db in error_dict
        for col_idx in range(1, len(headers) + 1):
            c = ws.cell(row=current_row_idx, column=col_idx)
            c.border = thin_border
            if is_error:
                c.fill = red_fill
                if col_idx == len(headers):
                    c.font = red_font
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# 4. Hàm đọc file & chuẩn hóa tên cột
def load_uploaded_file(uploaded_file):
    file_name = uploaded_file.name.lower()
    try:
        if file_name.endswith(".csv"):
            try:
                df = pd.read_csv(uploaded_file, encoding="utf-8-sig")
            except:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding="latin1")
        elif file_name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(uploaded_file)
        elif file_name.endswith(".json"):
            df = pd.read_json(uploaded_file)
        else:
            return None

        new_cols = []
        for i, col in enumerate(df.columns):
            if (
                "Unnamed" in str(col)
                or str(col).strip() == ""
                or pd.isna(col)
            ):
                new_cols.append(f"Cột {i}")
            else:
                new_cols.append(str(col).strip())
        df.columns = new_cols
        return df
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")
        return None


# 5. Kiểm tra tính hợp lệ dữ liệu
def validate_row(row, required_cols):
    errors = []
    for col in required_cols:
        val = row.get(col)
        if pd.isna(val) or str(val).strip() == "":
            errors.append(f"Cột '{col}' bị trống")
    return errors


# 6. Xử lý nền độc lập & Resume
def process_data_background(session_id, df, required_cols, start_row=0):
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    cursor = conn.cursor()

    batch_valid = []
    batch_invalid = []

    try:
        for index in range(start_row, len(df)):
            row = df.iloc[index]
            row_num = index + 1
            errors = validate_row(row, required_cols)

            row_dict = row.to_dict()
            clean_dict = {
                str(k): (None if pd.isna(v) else str(v))
                for k, v in row_dict.items()
            }

            if not errors:
                batch_valid.append(
                    (session_id, row_num, json.dumps(clean_dict, ensure_ascii=False))
                )
            else:
                batch_invalid.append(
                    (
                        session_id,
                        row_num,
                        json.dumps(clean_dict, ensure_ascii=False),
                        "; ".join(errors),
                    )
                )

            if len(batch_valid) + len(batch_invalid) >= 50 or index == len(df) - 1:
                if batch_valid:
                    cursor.executemany(
                        "INSERT INTO valid_data (session_id, row_number, data_json) VALUES (?, ?, ?)",
                        batch_valid,
                    )
                if batch_invalid:
                    cursor.executemany(
                        "INSERT INTO invalid_data (session_id, row_number, raw_data, error_reason) VALUES (?, ?, ?, ?)",
                        batch_invalid,
                    )

                cursor.execute(
                    """
                    UPDATE import_sessions 
                    SET processed_rows = ?, 
                        valid_rows = valid_rows + ?, 
                        invalid_rows = invalid_rows + ?
                    WHERE id = ?
                """,
                    (row_num, len(batch_valid), len(batch_invalid), session_id),
                )

                conn.commit()
                batch_valid.clear()
                batch_invalid.clear()

            time.sleep(0.0005)

        cursor.execute(
            "UPDATE import_sessions SET status = 'HOÀN THÀNH' WHERE id = ?",
            (session_id,),
        )
        conn.commit()
    except Exception as e:
        print(f"Lỗi tiến trình: {e}")
    finally:
        conn.close()


# 7. Quản lý trạng thái chuyển Tab thông minh từ Menu Sidebar
if "selected_nav" not in st.session_state:
    st.session_state["selected_nav"] = "📥 Tải File & Cấu Hình"

with st.sidebar:
    st.title("⚡ DataClean Pro")
    st.caption("Hệ Thống Kiểm Tra Dữ Liệu Tự Động")
    st.markdown("---")

    st.markdown("### **Menu Chức Năng**")

    # Các nút Menu có khả năng click để đổi trang/tab trực tiếp
    if st.button("📥 Tải File & Cấu Hình", use_container_width=True):
        st.session_state["selected_nav"] = "📥 Tải File & Cấu Hình"
        st.rerun()

    if st.button("🔄 Quản Lý Phiên & Tiến Độ", use_container_width=True):
        st.session_state["selected_nav"] = "🔄 Quản Lý Phiên & Tiếp Tục (Resume)"
        st.rerun()

    if st.button("📜 Lịch Sử & Xuất File", use_container_width=True):
        st.session_state["selected_nav"] = "📜 Lịch Sử & Xuất File Báo Cáo"
        st.rerun()

    st.markdown("---")
    with st.expander("📖 **Giới Thiệu & Hướng Dẫn**", expanded=False):
        st.write(
            "Phần mềm hỗ trợ kiểm tra dữ liệu lớn (5000+ dòng), tải nhiều file cùng lúc, "
            "tự động quét lỗi khuyết thiếu, chạy ngầm mượt mà và xuất báo cáo Excel chuyên nghiệp."
        )

# Tiêu đề chính của ứng dụng
st.title("⚡ DataClean Pro - Hệ Thống Kiểm Tra & Chuẩn Hóa Dữ Liệu")
st.caption(
    "Giải pháp xử lý tệp lớn thông minh, chạy ngầm độc lập, tự động phát hiện lỗi và xuất báo cáo Excel."
)
st.markdown("---")

# Danh sách Tab chính của ứng dụng
tab_options = [
    "📥 Tải File & Cấu Hình",
    "🔄 Quản Lý Phiên & Tiếp Tục (Resume)",
    "📜 Lịch Sử & Xuất File Báo Cáo",
]

# Tự động chọn Tab theo Lựa chọn của Menu Sidebar
default_tab_idx = 0
if st.session_state["selected_nav"] in tab_options:
    default_tab_idx = tab_options.index(st.session_state["selected_nav"])

# Khởi tạo các Tab trên giao diện
tab1, tab2, tab3 = st.tabs(tab_options)

# --- TAB 1: TẢI FILE & CẤU HÌNH ---
with tab1:
    st.subheader(
        "1. Tải lên tệp dữ liệu (Hỗ trợ tải nhiều file CSV, Excel, JSON cùng lúc)"
    )
    uploaded_files = st.file_uploader(
        "Kéo thả hoặc chọn các tệp dữ liệu của bạn:",
        type=["csv", "xlsx", "xls", "json"],
        accept_multiple_files=True,
    )

    if uploaded_files:
        st.success(
            f"📄 Đã tải lên thành công tổng cộng **{len(uploaded_files)}** tệp dữ liệu!"
        )

        if "df_cache" not in st.session_state:
            st.session_state["df_cache"] = {}

        for idx, uploaded_file in enumerate(uploaded_files):
            df = load_uploaded_file(uploaded_file)
            if df is not None:
                with st.expander(
                    f"📁 Tệp {idx+1}: {uploaded_file.name} (Tổng dòng: {len(df):,}, Số cột: {len(df.columns)})",
                    expanded=True,
                ):
                    st.dataframe(df.head(), use_container_width=True)

                st.markdown(
                    f"**2. Quy tắc kiểm tra cột bắt buộc cho: `{uploaded_file.name}`**"
                )
                all_cols = df.columns.tolist()

                # Tự động Select All toàn bộ các cột
                required_cols = st.multiselect(
                    f"Chọn các cột BẮT BUỘC phải có dữ liệu (File {idx+1}):",
                    options=all_cols,
                    default=all_cols,  # Select All mặc định
                    key=f"req_cols_{idx}_{uploaded_file.name}",
                )

                st.session_state["current_raw_df"] = df

                if st.button(
                    f"🚀 Bắt Đầu Xử Lý Tệp: {uploaded_file.name}",
                    key=f"btn_process_{idx}_{uploaded_file.name}",
                    type="primary",
                ):
                    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
                    cursor = conn.cursor()
                    cursor.execute(
                        "INSERT INTO import_sessions (filename, total_rows, status) VALUES (?, ?, 'ĐANG CHẠY')",
                        (uploaded_file.name, len(df)),
                    )
                    session_id = cursor.lastrowid
                    conn.commit()
                    conn.close()

                    st.session_state["df_cache"][session_id] = (df, required_cols)

                    thread = threading.Thread(
                        target=process_data_background,
                        args=(session_id, df, required_cols, 0),
                    )
                    thread.daemon = True
                    thread.start()

                    st.session_state["latest_session_id"] = session_id

                    # Tự động chuyển hướng giao diện sang Tab "Quản Lý Phiên & Tiếp Tục"
                    st.session_state["selected_nav"] = (
                        "🔄 Quản Lý Phiên & Tiếp Tục (Resume)"
                    )
                    st.rerun()

# --- TAB 2: QUẢN LÝ PHIÊN & RESUME ---
with tab2:
    st.subheader(
        "🔄 Theo Dõi Tiến Độ Thời Gian Thực & Khôi Phục Phiên Gián Đoạn"
    )
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    sessions_df = pd.read_sql_query(
        "SELECT * FROM import_sessions ORDER BY id DESC", conn
    )
    conn.close()

    if not sessions_df.empty:
        display_df = sessions_df.copy()
        display_df["Tỉ lệ sai"] = display_df.apply(
            lambda r: f"{(r['invalid_rows'] / r['total_rows']) * 100:.2f}%"
            if r["total_rows"] > 0
            else "0.00%",
            axis=1,
        )
        display_df["id"] = display_df["id"].apply(lambda x: f"#{x}")
        display_df.columns = [
            "Mã phiên",
            "Tên tệp dữ liệu",
            "Tổng số dòng",
            "Số dòng đã xử lý",
            "Số dòng hợp lệ",
            "Số dòng bị lỗi",
            "Trạng thái",
            "Thời gian tạo",
            "Tỉ lệ sai",
        ]
        display_df = display_df[[
            "Mã phiên",
            "Tên tệp dữ liệu",
            "Tổng số dòng",
            "Số dòng đã xử lý",
            "Số dòng hợp lệ",
            "Số dòng bị lỗi",
            "Tỉ lệ sai",
            "Trạng thái",
            "Thời gian tạo",
        ]]
        st.dataframe(display_df, use_container_width=True)

        default_index = 0
        if (
            "latest_session_id" in st.session_state
            and st.session_state["latest_session_id"] in sessions_df["id"].tolist()
        ):
            default_index = sessions_df["id"].tolist().index(
                st.session_state["latest_session_id"]
            )

        selected_id = st.selectbox(
            "Chọn ID Phiên cần theo dõi chi tiết:",
            sessions_df["id"].tolist(),
            index=default_index,
        )

        if selected_id:
            conn = sqlite3.connect(DB_NAME, check_same_thread=False)
            info = pd.read_sql_query(
                f"SELECT * FROM import_sessions WHERE id = {selected_id}", conn
            ).iloc[0]
            conn.close()

            total = info["total_rows"]
            processed = info["processed_rows"]
            status = info["status"]
            valid = info["valid_rows"]
            invalid = info["invalid_rows"]
            pct = (processed / total) if total > 0 else 0
            error_rate_val = (invalid / total * 100) if total > 0 else 0

            st.markdown("---")
            st.markdown(f"#### 📊 Chi tiết tiến độ Phiên #{selected_id}")
            st.progress(pct)

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric(label="📁 Tổng số dòng", value=f"{total:,}")
            m2.metric(
                label="⚙️ Đã xử lý",
                value=f"{processed:,}",
                delta=f"{(pct*100):.1f}%",
            )
            m3.metric(label="✅ Hợp lệ", value=f"{valid:,}")
            m4.metric(
                label="⚠️ Dòng lỗi",
                value=f"{invalid:,}",
                delta_color="inverse",
            )
            m5.metric(label="📉 Tỉ lệ sai", value=f"{error_rate_val:.2f}%")

            st.markdown("<br>", unsafe_allow_html=True)
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("🔄 Cập Nhật Trạng Thái"):
                    st.rerun()
            with col_btn2:
                if status != "HOÀN THÀNH" and processed < total:
                    if st.button(
                        "▶️ Tiếp Tục Xử Lý (Resume từ điểm dừng)", type="primary"
                    ):
                        if (
                            "df_cache" in st.session_state
                            and selected_id in st.session_state["df_cache"]
                        ):
                            df_cached, req_cached = st.session_state["df_cache"][
                                selected_id
                            ]
                            t_resume = threading.Thread(
                                target=process_data_background,
                                args=(
                                    selected_id,
                                    df_cached,
                                    req_cached,
                                    processed,
                                ),
                            )
                            t_resume.daemon = True
                            t_resume.start()
                            st.success(
                                f"Đã khôi phục thành công Phiên #{selected_id} từ dòng số {processed + 1}!"
                            )
                            st.rerun()
                        else:
                            st.error(
                                "⚠️ Vui lòng tải lại tệp ở Tab 1 để bộ nhớ tạm ghi nhận dữ liệu phục hồi."
                            )
    else:
        st.info("Chưa có phiên nhập dữ liệu nào trong hệ thống.")

# --- TAB 3: LỊCH SỬ & XUẤT EXCEL BÁO CÁO ---
with tab3:
    st.subheader("📜 Bảng Lịch Sử & Xuất Báo Cáo Xử Lý")
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    history_df = pd.read_sql_query(
        "SELECT * FROM import_sessions ORDER BY id DESC", conn
    )

    if not history_df.empty:
        view_history = history_df.copy()
        view_history["Tỉ lệ sai"] = view_history.apply(
            lambda r: f"{(r['invalid_rows'] / r['total_rows']) * 100:.2f}%"
            if r["total_rows"] > 0
            else "0.00%",
            axis=1,
        )
        view_history["id"] = view_history["id"].apply(lambda x: f"#{x}")
        view_history.columns = [
            "Mã phiên",
            "Tên tệp dữ liệu",
            "Tổng số dòng",
            "Số dòng đã xử lý",
            "Số dòng hợp lệ",
            "Số dòng bị lỗi",
            "Trạng thái",
            "Thời gian tạo",
            "Tỉ lệ sai",
        ]
        view_history = view_history[[
            "Mã phiên",
            "Tên tệp dữ liệu",
            "Tổng số dòng",
            "Số dòng đã xử lý",
            "Số dòng hợp lệ",
            "Số dòng bị lỗi",
            "Tỉ lệ sai",
            "Trạng thái",
            "Thời gian tạo",
        ]]
        st.dataframe(view_history, use_container_width=True)

        st.markdown("---")
        st.subheader("📥 Xuất Báo Cáo Định Dạng Excel Chuyên Nghiệp")

        col_ex1, col_ex2 = st.columns(2)

        with col_ex1:
            st.markdown("#### 1. Báo cáo lịch sử phiên")
            st.write(
                "Tải danh sách phiên làm việc tổng hợp kèm theo tỉ lệ sai."
            )
            excel_buffer = export_history_to_excel(history_df)
            st.download_button(
                label="📊 Tải Báo Cáo Lịch Sử Phiên (Excel)",
                data=excel_buffer,
                file_name="DataClean_Lich_Su_Phien.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

        with col_ex2:
            st.markdown("#### 2. File gốc đánh dấu ô lỗi")
            st.write(
                "Tải tệp dữ liệu gốc có thêm cột ghi chú và tô đỏ các dòng lỗi."
            )
            target_session_id = st.selectbox(
                "Chọn Mã Phiên muốn xuất file đánh dấu lỗi:",
                history_df["id"].tolist(),
                key="select_session_for_marking",
            )

            if target_session_id:
                if (
                    "current_raw_df" in st.session_state
                    and st.session_state["current_raw_df"] is not None
                ):
                    marked_buffer = export_marked_file_to_excel(
                        target_session_id, st.session_state["current_raw_df"]
                    )
                    st.download_button(
                        label=f"📝 Tải File Gốc Đánh Dấu Lỗi (Phiên #{target_session_id})",
                        data=marked_buffer,
                        file_name=f"DataClean_File_Goc_Dau_Loi_Phien_{target_session_id}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                    )
                else:
                    st.info(
                        "💡 Vui lòng tải file tại Tab 1 để lưu cache xuất file đánh dấu lỗi."
                    )
    else:
        st.info("Chưa có dữ liệu lịch sử phiên nhập nào.")

    st.markdown("---")
    st.subheader("🔍 Tra Cứu Dòng Lỗi Chi Tiết")
    view_id = st.number_input(
        "Nhập ID Phiên muốn xem chi tiết dòng lỗi:",
        min_value=1,
        step=1,
        value=1,
    )

    if st.button("🔍 Trích Xuất Dữ Liệu Lỗi"):
        err_df = pd.read_sql_query(
            f"SELECT row_number AS 'Dòng số', error_reason AS 'Lý do lỗi' FROM invalid_data WHERE session_id = {view_id}",
            conn,
        )

        if not err_df.empty:
            st.error(
                f"⚠️ Phát hiện tổng cộng **{len(err_df)}** dòng lỗi ở Phiên #{view_id}:"
            )
            st.dataframe(err_df, use_container_width=True)

            csv_bytes = err_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                label="📥 Tải Báo Cáo Lỗi Chi Tiết (CSV)",
                data=csv_bytes,
                file_name=f"DataClean_Bao_Cao_Loi_Phien_{view_id}.csv",
                mime="text/csv",
            )
        else:
            st.success(
                f"🎉 Phiên #{view_id} hoàn toàn chính xác, không có dòng lỗi nào!"
            )
    conn.close()
