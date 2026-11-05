import hashlib
import io
import json
import os
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

UPLOAD_DIR = "uploaded_temp"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Custom CSS giao diện Dark Mode + Khung cuộn tùy chỉnh cho danh sách tệp
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
        padding: 2rem 2.5rem;
        border-radius: 16px;
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.4);
        margin-top: 1rem;
        margin-bottom: 2rem;
        border: 1px solid #2d2d2d;
    }
    div.block-container p, div.block-container span, div.block-container label, div.block-container div {
        color: #e0e0e0 !important;
    }
    h1, h2, h3, h4 {
        color: #ffffff !important;
    }
    
    [data-testid="stTabs"] {
        display: none !important;
    }

    .scrollable-file-box {
        max-height: 420px;
        overflow-y: auto;
        padding-right: 15px;
        margin-bottom: 1.5rem;
        border: 1px solid #2d2d2d;
        border-radius: 12px;
        padding: 15px;
        background-color: #161616;
    }
    .scrollable-file-box::-webkit-scrollbar {
        width: 8px;
    }
    .scrollable-file-box::-webkit-scrollbar-track {
        background: #1a1a1a;
        border-radius: 4px;
    }
    .scrollable-file-box::-webkit-scrollbar-thumb {
        background: #3b82f6;
        border-radius: 4px;
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
    </style>
""",
    unsafe_allow_html=True,
)

DB_NAME = "dataclean_app_integrated.db"


def init_db():
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS import_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            file_hash TEXT,
            total_rows INTEGER,
            processed_rows INTEGER DEFAULT 0,
            valid_rows INTEGER DEFAULT 0,
            invalid_rows INTEGER DEFAULT 0,
            status TEXT DEFAULT 'ĐANG CHẠY',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    cursor.execute("PRAGMA table_info(import_sessions)")
    columns = [col[1] for col in cursor.fetchall()]
    if "file_hash" not in columns:
        cursor.execute("ALTER TABLE import_sessions ADD COLUMN file_hash TEXT")
        conn.commit()

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


def get_file_content_hash(file_bytes):
    hasher = hashlib.md5()
    hasher.update(file_bytes)
    return hasher.hexdigest()


def reconstruct_df_from_db(session_id):
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    valid_df = pd.read_sql_query(
        f"SELECT row_number, data_json FROM valid_data WHERE session_id = {session_id}",
        conn,
    )
    invalid_df = pd.read_sql_query(
        f"SELECT row_number, raw_data FROM invalid_data WHERE session_id = {session_id}",
        conn,
    )
    conn.close()

    rows = []
    for _, r in valid_df.iterrows():
        try:
            data = json.loads(r["data_json"])
            data["_row_num"] = int(r["row_number"])
            rows.append(data)
        except:
            pass

    for _, r in invalid_df.iterrows():
        try:
            data = json.loads(r["raw_data"])
            data["_row_num"] = int(r["row_number"])
            rows.append(data)
        except:
            pass

    if not rows:
        return None

    df = pd.DataFrame(rows)
    if "_row_num" in df.columns:
        df = df.sort_values("_row_num").drop(columns=["_row_num"])
    return df


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


def load_uploaded_file(uploaded_file, header_row=0):
    file_name = uploaded_file.name.lower()
    try:
        if file_name.endswith(".csv"):
            try:
                df = pd.read_csv(uploaded_file, encoding="utf-8-sig", header=header_row)
            except:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding="latin1", header=header_row)
        elif file_name.endswith((".xlsx", ".xls")):
            df = pd.read_excel(uploaded_file, header=header_row)
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
        df = df.dropna(how="all").reset_index(drop=True)
        return df
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")
        return None


def validate_row(row, required_cols):
    errors = []
    for col in required_cols:
        if col in row:
            val = row.get(col)
            if pd.isna(val) or str(val).strip() == "":
                errors.append(f"Cột '{col}' bị trống")
    return errors


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

        cursor.execute(
            "UPDATE import_sessions SET status = 'HOÀN THÀNH', processed_rows = total_rows WHERE id = ?",
            (session_id,),
        )
        conn.commit()
    except Exception as e:
        print(f"Lỗi tiến trình nền: {e}")
    finally:
        conn.close()


if "selected_nav" not in st.session_state:
    st.session_state["selected_nav"] = "📥 Tải File & Cấu Hình"

with st.sidebar:
    st.title("⚡ DataClean Pro")
    st.caption("Hệ Thống Kiểm Tra Dữ Liệu Tự Động")
    st.markdown("---")

    st.markdown("### **Menu Chức Năng**")

    if st.button("📥 Tải File & Cấu Hình", use_container_width=True, key="nav_btn_1"):
        st.session_state["selected_nav"] = "📥 Tải File & Cấu Hình"
        st.rerun()

    if st.button("🔄 Quản Lý Phiên & Tiến Độ", use_container_width=True, key="nav_btn_2"):
        st.session_state["selected_nav"] = "🔄 Quản Lý Phiên & Tiếp Tục (Resume)"
        st.rerun()

    if st.button("📜 Lịch Sử & Xuất File", use_container_width=True, key="nav_btn_3"):
        st.session_state["selected_nav"] = "📜 Lịch Sử & Xuất File Báo Cáo"
        st.rerun()

    st.markdown("---")
    with st.expander("📖 **Giới Thiệu & Hướng Dẫn**", expanded=False):
        st.write(
            "Phần mềm tự động nhận diện nội dung tệp (File Content Hash). "
            "Nếu tệp đã từng được quét trước đó, hệ thống sẽ tự động thông báo và bỏ qua quét lại để tiết kiệm thời gian."
        )

st.title("⚡ DataClean Pro - Hệ Thống Kiểm Tra & Chuẩn Hóa Dữ Liệu")
st.caption(
    "Giải pháp xử lý tệp lớn thông minh, nhận diện nội dung trùng lặp, chạy ngầm độc lập và xuất báo cáo Excel."
)
st.markdown("---")

current_view = st.session_state["selected_nav"]

# --- VIEW 1: TẢI FILE & CẤU HÌNH ---
if current_view == "📥 Tải File & Cấu Hình":
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

        st.markdown("### ⚙️ Cấu Hình Dòng Tiêu Đề (Header Row)")
        header_row_idx = st.number_input(
            "Chọn số thứ tự dòng chứa tiêu đề (Dòng bắt đầu tính từ 0):",
            min_value=0, max_value=10, value=0, step=1,
            help="Nếu dòng đầu tiên của file Excel là tiêu đề phụ hoặc tên biên bản, hãy tăng số này lên (ví dụ: 1 hoặc 2) để lấy đúng tên cột."
        )

        st.markdown("### ⚙️ Danh Sách Tệp Tải Lên & Chọn Tệp Xử Lý")

        file_info_list = []
        file_dfs = {}
        file_hashes = {}

        conn_check = sqlite3.connect(DB_NAME, check_same_thread=False)

        for idx, uploaded_file in enumerate(uploaded_files):
            file_bytes = uploaded_file.getvalue()
            f_hash = get_file_content_hash(file_bytes)
            file_hashes[uploaded_file.name] = f_hash

            try:
                existing_session = pd.read_sql_query(
                    f"SELECT id, filename, status as created_status FROM import_sessions WHERE file_hash = '{f_hash}'",
                    conn_check,
                )
            except:
                existing_session = pd.DataFrame()

            is_already_scanned = not existing_session.empty
            scanned_session_id = existing_session.iloc[0]["id"] if is_already_scanned else None

            df = load_uploaded_file(uploaded_file, header_row=int(header_row_idx))
            if df is not None:
                file_dfs[uploaded_file.name] = (df, uploaded_file)
                
                status_text = f"Đã từng quét (Phiên #{scanned_session_id})" if is_already_scanned else "Mới / Chưa quét"
                file_info_list.append({
                    "idx_ref": idx,
                    "Tên tệp": uploaded_file.name,
                    "Tổng số dòng": f"{len(df):,}",
                    "Trạng thái": status_text
                })

        conn_check.close()

        if file_info_list:
            files_df = pd.DataFrame(file_info_list)
            
            default_selection = list(range(len(files_df)))
            
            upload_event = st.dataframe(
                files_df.drop(columns=["idx_ref"]),
                use_container_width=True,
                selection_mode="multi-row",
                on_select="rerun",
                key="upload_files_table_selection"
            )

            selected_files_to_process = []
            selected_rows = []
            
            if upload_event and "selection" in upload_event and "rows" in upload_event.selection:
                selected_rows = upload_event.selection["rows"]
            
            if not selected_rows and not (upload_event and "selection" in upload_event):
                selected_rows = default_selection

            for r_idx in selected_rows:
                if r_idx < len(files_df):
                    selected_files_to_process.append(files_df.iloc[r_idx]["Tên tệp"])

            if selected_files_to_process:
                st.info(f"📦 Bạn đang chọn **{len(selected_files_to_process)}** tệp để xử lý.")

            st.markdown("---")
            if st.button(
                "🚀 Bắt Đầu Xử Lý Tất Cả Các Tệp Đã Chọn", type="primary", key="main_process_btn"
            ):
                if not selected_files_to_process:
                    st.error("⚠️ Vui lòng tích chọn ít nhất một tệp ở bảng trên để xử lý!")
                else:
                    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
                    cursor = conn.cursor()

                    last_created_id = None
                    for fname in selected_files_to_process:
                        df, up_file = file_dfs[fname]
                        all_cols = df.columns.tolist()
                        f_hash = file_hashes[fname]

                        cursor.execute(
                            "INSERT INTO import_sessions (filename, file_hash, total_rows, status) VALUES (?, ?, ?, 'ĐANG CHẠY')",
                            (fname, f_hash, len(df)),
                        )
                        session_id = cursor.lastrowid
                        conn.commit()

                        file_save_path = os.path.join(UPLOAD_DIR, f"session_{session_id}_{fname}")
                        up_file.seek(0)
                        with open(file_save_path, "wb") as f:
                            f.write(up_file.read())

                        st.session_state["df_cache"][session_id] = (df, all_cols, file_save_path)

                        thread = threading.Thread(
                            target=process_data_background,
                            args=(session_id, df, all_cols, 0),
                        )
                        thread.daemon = True
                        thread.start()
                        last_created_id = session_id

                    conn.close()

                    if last_created_id:
                        st.session_state["latest_session_id"] = last_created_id

                    st.session_state["selected_nav"] = (
                        "📜 Lịch Sử & Xuất File Báo Cáo"
                    )
                    st.rerun()

# --- VIEW 2: QUẢN LÝ PHIÊN & RESUME ---
elif current_view == "🔄 Quản Lý Phiên & Tiếp Tục (Resume)":
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
        
        display_df = display_df[[
            "id",
            "filename",
            "total_rows",
            "processed_rows",
            "valid_rows",
            "invalid_rows",
            "Tỉ lệ sai",
            "status",
            "created_at"
        ]]
        display_df.columns = [
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
            key="selectbox_session_tracking",
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
            filename_val = info["filename"]
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
            col_btn1, col_btn2, col_btn3 = st.columns(3)
            with col_btn1:
                if st.button("🔄 Cập Nhật Trạng Thái", key="btn_refresh_status"):
                    st.rerun()
            with col_btn2:
                # Nút cứu hộ cưỡng chế hoàn thành phiên nếu bị kẹt trạng thái ĐANG CHẠY
                if status != "HOÀN THÀNH":
                    if st.button("⚡ Đánh Dấu Hoàn Thành Ngay", key="btn_force_complete"):
                        conn_f = sqlite3.connect(DB_NAME, check_same_thread=False)
                        cursor_f = conn_f.cursor()
                        cursor_f.execute("UPDATE import_sessions SET status = 'HOÀN THÀNH' WHERE id = ?", (selected_id,))
                        conn_f.commit()
                        conn_f.close()
                        st.success(f"Đã cập nhật Phiên #{selected_id} thành HOÀN THÀNH thành công!")
                        st.rerun()
            with col_btn3:
                if status != "HOÀN THÀNH" and processed < total:
                    possible_path = os.path.join(UPLOAD_DIR, f"session_{selected_id}_{filename_val}")
                    file_exists_on_disk = os.path.exists(possible_path) or ("df_cache" in st.session_state and selected_id in st.session_state["df_cache"])

                    if file_exists_on_disk:
                        if st.button(
                            "▶️ Tiếp Tục Xử Lý", type="primary", key="btn_resume_process"
                        ):
                            df_cached = None
                            req_cached = None
                            
                            if "df_cache" in st.session_state and selected_id in st.session_state["df_cache"]:
                                cached_data = st.session_state["df_cache"][selected_id]
                                df_cached = cached_data[0]
                                req_cached = cached_data[1]
                            else:
                                if os.path.exists(possible_path):
                                    try:
                                        if filename_val.endswith(".csv"):
                                            df_cached = pd.read_csv(possible_path, encoding="utf-8-sig")
                                        elif filename_val.endswith((".xlsx", ".xls")):
                                            df_cached = pd.read_excel(possible_path)
                                        elif filename_val.endswith(".json"):
                                            df_cached = pd.read_json(possible_path)
                                        req_cached = df_cached.columns.tolist()
                                    except Exception as e:
                                        st.error(f"Không thể đọc lại tệp từ đĩa: {e}")

                            if df_cached is not None:
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
        st.info("Chưa có phiên nhập dữ liệu nào trong hệ thống.")

# --- VIEW 3: LỊCH SỬ & XUẤT FILE BÁO CÁO ---
elif current_view == "📜 Lịch Sử & Xuất File Báo Cáo":
    st.subheader("📜 Bảng Lịch Sử & Xuất Báo Cáo Xử Lý")
    
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    history_df = pd.read_sql_query(
        "SELECT * FROM import_sessions ORDER BY id DESC", conn
    )

    if not history_df.empty:
        search_query = st.text_input(
            "🔍 Tìm kiếm nhanh tên tệp dữ liệu hoặc mã phiên:",
            value="",
            placeholder="Nhập tên file hoặc mã phiên (ví dụ: 7A1, data_test)...",
            key="global_search_input"
        )

        filtered_history_df = history_df.copy()
        if search_query.strip():
            q = search_query.strip().lower()
            filtered_history_df = filtered_history_df[
                filtered_history_df["filename"].str.lower().str.contains(q, na=False) |
                filtered_history_df["id"].astype(str).str.contains(q, na=False)
            ]

        view_history = filtered_history_df.copy()
        view_history["Tỉ lệ sai"] = view_history.apply(
            lambda r: f"{(r['invalid_rows'] / r['total_rows']) * 100:.2f}%"
            if r["total_rows"] > 0
            else "0.00%",
            axis=1,
        )
        view_history["raw_id"] = view_history["id"]
        view_history["id"] = view_history["id"].apply(lambda x: f"#{x}")
        
        view_history = view_history[[
            "raw_id",
            "id",
            "filename",
            "total_rows",
            "processed_rows",
            "valid_rows",
            "invalid_rows",
            "Tỉ lệ sai",
            "status",
            "created_at"
        ]]
        view_history.columns = [
            "raw_id",
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
        
        event = st.dataframe(
            view_history.drop(columns=["raw_id"]),
            use_container_width=True,
            selection_mode="multi-row",
            on_select="rerun",
            key="history_multi_select_table"
        )

        if "selected_target_session_id" not in st.session_state:
            st.session_state["selected_target_session_id"] = history_df["id"].iloc[0] if not history_df.empty else 1

        if event and "selection" in event and "rows" in event.selection:
            selected_rows = event.selection["rows"]
            if selected_rows:
                row_idx = selected_rows[-1]
                clicked_id = int(view_history.iloc[row_idx]["raw_id"])
                if st.session_state["selected_target_session_id"] != clicked_id:
                    st.session_state["selected_target_session_id"] = clicked_id

        available_ids = filtered_history_df["id"].tolist() if not filtered_history_df.empty else history_df["id"].tolist()
        
        if st.session_state["selected_target_session_id"] not in available_ids and available_ids:
            st.session_state["selected_target_session_id"] = available_ids[0]

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
                key="dl_btn_history_excel"
            )

        with col_ex2:
            st.markdown("#### 2. File gốc đánh dấu ô lỗi")
            st.write(
                "Tải tệp dữ liệu gốc có thêm cột ghi chú và tô đỏ các dòng lỗi dựa trên phiên bạn đang chọn."
            )
            
            target_session_id = st.selectbox(
                "Chọn Mã Phiên muốn xuất file đánh dấu lỗi:",
                available_ids,
                index=(
                    available_ids.index(
                        st.session_state["selected_target_session_id"]
                    )
                    if st.session_state["selected_target_session_id"]
                    in available_ids
                    else 0
                ),
                key="select_session_for_marking_tab3",
                format_func=lambda x: (
                    f"Phiên #{x} - {history_df[history_df['id'] == x]['filename'].values[0] if not history_df[history_df['id'] == x].empty else ''}"
                ),
            )

            if target_session_id != st.session_state["selected_target_session_id"]:
                st.session_state["selected_target_session_id"] = target_session_id
                st.rerun()

            if target_session_id:
                target_filename = history_df[history_df['id'] == target_session_id]['filename'].values[0]
                export_df_source = None
                
                if "df_cache" in st.session_state and target_session_id in st.session_state["df_cache"]:
                    export_df_source = st.session_state["df_cache"][target_session_id][0]
                else:
                    possible_path = os.path.join(UPLOAD_DIR, f"session_{target_session_id}_{target_filename}")
                    if os.path.exists(possible_path):
                        try:
                            if target_filename.endswith(".csv"):
                                export_df_source = pd.read_csv(possible_path, encoding="utf-8-sig")
                            elif target_filename.endswith((".xlsx", ".xls")):
                                export_df_source = pd.read_excel(possible_path)
                            elif target_filename.endswith(".json"):
                                export_df_source = pd.read_json(possible_path)
                        except:
                            pass
                    
                    if export_df_source is None:
                        export_df_source = reconstruct_df_from_db(target_session_id)

                if export_df_source is not None:
                    marked_buffer = export_marked_file_to_excel(
                        target_session_id, export_df_source
                    )
                    st.download_button(
                        label=f"📝 Tải File Gốc Đánh Dấu Lỗi (Phiên #{target_session_id})",
                        data=marked_buffer,
                        file_name=f"DataClean_File_Goc_Dau_Loi_Phiên_{target_session_id}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        key="dl_btn_marked_excel"
                    )
                else:
                    st.info(
                        "💡 Không tìm thấy dữ liệu gốc để xuất file đánh dấu lỗi."
                    )
    else:
        st.info("Chưa có phiên nhập dữ liệu nào trong hệ thống.")

    st.markdown("---")
    st.subheader("🔍 Tra Cứu Dòng Lỗi Chi Tiết")
    
    lookup_default_id = st.session_state.get(
        "selected_target_session_id",
        available_ids[0] if 'available_ids' in locals() and available_ids else 1,
    )

    view_id = st.selectbox(
        "Chọn Mã Phiên muốn xem chi tiết dòng lỗi:",
        history_df["id"].tolist() if not history_df.empty else [1],
        index=(
            history_df["id"].tolist().index(lookup_default_id)
            if not history_df.empty
            and lookup_default_id in history_df["id"].tolist()
            else 0
        ),
        key="selectbox_lookup_error_id",
        format_func=lambda x: (
            f"Phiên #{x} - {history_df[history_df['id'] == x]['filename'].values[0] if not history_df[history_df['id'] == x].empty else ''}"
        ),
    )

    if st.button("🔍 Trích Xuất Dữ Liệu Lỗi", key="btn_lookup_error_details"):
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
                file_name=f"DataClean_Bao_Cao_Loi_Phiên_{view_id}.csv",
                mime="text/csv",
                key="dl_btn_error_csv"
            )
        else:
            st.success(
                f"🎉 Phiên #{view_id} hoàn toàn chính xác, không có dòng lỗi nào!"
            )
    conn.close()
