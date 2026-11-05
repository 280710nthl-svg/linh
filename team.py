import hashlib
import io
import json
import os
import sqlite3
import threading

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st


# =========================
# CẤU HÌNH ỨNG DỤNG
# =========================
st.set_page_config(
    page_title="DataClean Pro - Hệ Thống Xử Lý & Kiểm Tra Dữ Liệu Thông Minh",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

UPLOAD_DIR = "uploaded_temp"
os.makedirs(UPLOAD_DIR, exist_ok=True)

DB_NAME = "dataclean_app_integrated.db"


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
    div.block-container p,
    div.block-container span,
    div.block-container label,
    div.block-container div {
        color: #e0e0e0 !important;
    }
    h1, h2, h3, h4 {
        color: #ffffff !important;
    }
    [data-testid="stTabs"] {
        display: none !important;
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
    .stButton > button {
        border-radius: 8px;
        font-weight: 600;
        padding: 10px 24px;
        background-color: #3b82f6;
        color: #ffffff;
        border: none;
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        background-color: #2563eb;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================
# DATABASE
# =========================
def get_connection():
    return sqlite3.connect(DB_NAME, check_same_thread=False, timeout=30)


def init_db():
    conn = get_connection()
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

    # Ngăn việc resume ghi trùng cùng một dòng.
    cursor.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_valid_session_row
        ON valid_data(session_id, row_number)
    """)
    cursor.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS idx_invalid_session_row
        ON invalid_data(session_id, row_number)
    """)

    conn.commit()
    conn.close()


init_db()


# =========================
# HÀM XỬ LÝ FILE / DỮ LIỆU
# =========================
def get_file_content_hash(file_bytes):
    return hashlib.md5(file_bytes).hexdigest()


def load_file_from_path(path, filename):
    """Đọc file từ đường dẫn, luôn dùng header ở dòng đầu tiên."""
    filename_lower = filename.lower()

    if filename_lower.endswith(".csv"):
        try:
            df = pd.read_csv(path, encoding="utf-8-sig", header=0)
        except UnicodeDecodeError:
            df = pd.read_csv(path, encoding="latin1", header=0)
    elif filename_lower.endswith((".xlsx", ".xls")):
        df = pd.read_excel(path, header=0)
    elif filename_lower.endswith(".json"):
        df = pd.read_json(path)
    else:
        raise ValueError("Định dạng file không được hỗ trợ.")

    return normalize_dataframe(df)


def normalize_dataframe(df):
    new_cols = []
    for i, col in enumerate(df.columns):
        if "Unnamed" in str(col) or str(col).strip() == "" or pd.isna(col):
            new_cols.append(f"Cột {i}")
        else:
            new_cols.append(str(col).strip())

    df.columns = new_cols
    return df.dropna(how="all").reset_index(drop=True)


def load_uploaded_file(uploaded_file):
    try:
        uploaded_file.seek(0)
        filename = uploaded_file.name.lower()

        if filename.endswith(".csv"):
            try:
                df = pd.read_csv(uploaded_file, encoding="utf-8-sig", header=0)
            except UnicodeDecodeError:
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding="latin1", header=0)
        elif filename.endswith((".xlsx", ".xls")):
            df = pd.read_excel(uploaded_file, header=0)
        elif filename.endswith(".json"):
            df = pd.read_json(uploaded_file)
        else:
            return None

        return normalize_dataframe(df)
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
    """
    start_row là vị trí bắt đầu theo chỉ số 0.
    row_number trong database được tính từ 1.
    """
    conn = get_connection()
    cursor = conn.cursor()

    batch_valid = []
    batch_invalid = []

    try:
        start_row = max(0, int(start_row))

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
                    (
                        session_id,
                        row_num,
                        json.dumps(clean_dict, ensure_ascii=False),
                    )
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

            is_last_row = index == len(df) - 1
            batch_size_reached = len(batch_valid) + len(batch_invalid) >= 50

            if batch_size_reached or is_last_row:
                # INSERT OR IGNORE phòng trường hợp phiên bị resume lặp.
                if batch_valid:
                    cursor.executemany(
                        """
                        INSERT OR IGNORE INTO valid_data
                        (session_id, row_number, data_json)
                        VALUES (?, ?, ?)
                        """,
                        batch_valid,
                    )

                if batch_invalid:
                    cursor.executemany(
                        """
                        INSERT OR IGNORE INTO invalid_data
                        (session_id, row_number, raw_data, error_reason)
                        VALUES (?, ?, ?, ?)
                        """,
                        batch_invalid,
                    )

                cursor.execute(
                    """
                    UPDATE import_sessions
                    SET processed_rows = MAX(processed_rows, ?),
                        valid_rows = (
                            SELECT COUNT(*) FROM valid_data
                            WHERE session_id = ?
                        ),
                        invalid_rows = (
                            SELECT COUNT(*) FROM invalid_data
                            WHERE session_id = ?
                        )
                    WHERE id = ?
                    """,
                    (row_num, session_id, session_id, session_id),
                )

                conn.commit()
                batch_valid.clear()
                batch_invalid.clear()

        cursor.execute(
            """
            UPDATE import_sessions
            SET status = 'HOÀN THÀNH',
                processed_rows = total_rows
            WHERE id = ?
            """,
            (session_id,),
        )
        conn.commit()

    except Exception as e:
        print(f"Lỗi tiến trình nền phiên #{session_id}: {e}")
        try:
            cursor.execute(
                "UPDATE import_sessions SET status = 'TẠM DỪNG' WHERE id = ?",
                (session_id,),
            )
            conn.commit()
        except Exception:
            pass
    finally:
        conn.close()


def reconstruct_df_from_db(session_id):
    conn = get_connection()

    valid_df = pd.read_sql_query(
        "SELECT row_number, data_json FROM valid_data WHERE session_id = ?",
        conn,
        params=(session_id,),
    )
    invalid_df = pd.read_sql_query(
        "SELECT row_number, raw_data FROM invalid_data WHERE session_id = ?",
        conn,
        params=(session_id,),
    )
    conn.close()

    rows = []

    for _, record in valid_df.iterrows():
        try:
            data = json.loads(record["data_json"])
            data["_row_num"] = int(record["row_number"])
            rows.append(data)
        except Exception:
            pass

    for _, record in invalid_df.iterrows():
        try:
            data = json.loads(record["raw_data"])
            data["_row_num"] = int(record["row_number"])
            rows.append(data)
        except Exception:
            pass

    if not rows:
        return None

    df = pd.DataFrame(rows)
    if "_row_num" in df.columns:
        df = df.sort_values("_row_num").drop(columns=["_row_num"])
    return df.reset_index(drop=True)


# =========================
# XUẤT EXCEL
# =========================
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
        total = row.total_rows if row.total_rows > 0 else 1
        error_rate = f"{(row.invalid_rows / total) * 100:.2f}%"

        ws.append(
            [
                f"#{row.id}",
                row.filename,
                row.total_rows,
                row.processed_rows,
                row.valid_rows,
                row.invalid_rows,
                error_rate,
                row.status,
                str(row.created_at),
            ]
        )
        ws.row_dimensions[row_idx].height = 22

        for col_idx in range(1, 10):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            cell.alignment = Alignment(
                horizontal="left" if col_idx == 2 else "center",
                vertical="center",
            )

    col_widths = [12, 32, 16, 18, 16, 16, 14, 16, 22]
    for i, width in enumerate(col_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def export_marked_file_to_excel(session_id, original_df):
    conn = get_connection()
    err_records = pd.read_sql_query(
        "SELECT row_number, error_reason FROM invalid_data WHERE session_id = ?",
        conn,
        params=(session_id,),
    )
    conn.close()

    error_dict = {
        int(record["row_number"]): record["error_reason"]
        for _, record in err_records.iterrows()
    }

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Kiem_Tra_Loi_File_Goc"

    export_df = original_df.copy()
    export_df["Ghi chú kiểm tra lỗi"] = "Hợp lệ"

    headers = list(export_df.columns)
    ws.append(headers)

    header_fill = PatternFill(
        start_color="1E3C72", end_color="1E3C72", fill_type="solid"
    )
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
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

        ws.append(row_values)
        excel_row = i + 2
        ws.row_dimensions[excel_row].height = 20

        is_error = row_num_db in error_dict
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=excel_row, column=col_idx)
            cell.border = thin_border
            if is_error:
                cell.fill = red_fill
                if col_idx == len(headers):
                    cell.font = red_font
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# =========================
# SIDEBAR
# =========================
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
            "Hệ thống nhận diện tệp trùng lặp bằng mã hash. "
            "Nếu tệp đã từng được quét, hệ thống sẽ thông báo và bỏ qua việc quét lại."
        )


st.title("⚡ DataClean Pro - Hệ Thống Kiểm Tra & Chuẩn Hóa Dữ Liệu")
st.caption(
    "Giải pháp xử lý tệp lớn thông minh, nhận diện nội dung trùng lặp, "
    "chạy ngầm và xuất báo cáo Excel."
)
st.markdown("---")

current_view = st.session_state["selected_nav"]


# =========================
# VIEW 1: TẢI FILE
# =========================
if current_view == "📥 Tải File & Cấu Hình":
    st.subheader("1. Tải lên tệp dữ liệu")
    uploaded_files = st.file_uploader(
        "Kéo thả hoặc chọn các tệp dữ liệu của bạn:",
        type=["csv", "xlsx", "xls", "json"],
        accept_multiple_files=True,
    )

    if uploaded_files:
        st.success(f"📄 Đã tải lên thành công {len(uploaded_files)} tệp dữ liệu!")

        if "df_cache" not in st.session_state:
            st.session_state["df_cache"] = {}

        st.caption("Dòng đầu tiên của tệp được dùng làm tiêu đề cột.")
        st.markdown("### 📂 Danh Sách Tệp Tải Lên & Chọn Tệp Xử Lý")

        file_info_list = []
        file_dfs = {}
        file_hashes = {}

        conn_check = get_connection()

        for idx, uploaded_file in enumerate(uploaded_files):
            file_bytes = uploaded_file.getvalue()
            file_hash = get_file_content_hash(file_bytes)
            file_hashes[uploaded_file.name] = file_hash

            existing_session = pd.read_sql_query(
                "SELECT id, filename, status FROM import_sessions WHERE file_hash = ?",
                conn_check,
                params=(file_hash,),
            )

            is_already_scanned = not existing_session.empty
            scanned_session_id = (
                int(existing_session.iloc[0]["id"]) if is_already_scanned else None
            )

            df = load_uploaded_file(uploaded_file)
            if df is not None:
                file_dfs[uploaded_file.name] = (df, uploaded_file)

                status_text = (
                    f"Đã từng quét (Phiên #{scanned_session_id})"
                    if is_already_scanned
                    else "Mới / Chưa quét"
                )
                file_info_list.append(
                    {
                        "idx_ref": idx,
                        "Tên tệp": uploaded_file.name,
                        "Tổng số dòng": f"{len(df):,}",
                        "Trạng thái": status_text,
                    }
                )

        conn_check.close()

        if file_info_list:
            files_df = pd.DataFrame(file_info_list)
            default_selection = list(range(len(files_df)))

            upload_event = st.dataframe(
                files_df.drop(columns=["idx_ref"]),
                use_container_width=True,
                selection_mode="multi-row",
                on_select="rerun",
                key="upload_files_table_selection",
            )

            selected_rows = []
            if upload_event and "selection" in upload_event:
                selected_rows = upload_event.selection.get("rows", [])

            # Mặc định chọn tất cả nếu người dùng chưa thao tác chọn dòng.
            if not selected_rows:
                selected_rows = default_selection

            selected_files_to_process = [
                files_df.iloc[row_index]["Tên tệp"]
                for row_index in selected_rows
                if row_index < len(files_df)
            ]

            if selected_files_to_process:
                st.info(f"📦 Đang chọn {len(selected_files_to_process)} tệp để xử lý.")

            st.markdown("---")
            if st.button(
                "🚀 Bắt Đầu Xử Lý Các Tệp Đã Chọn",
                type="primary",
                key="main_process_btn",
            ):
                if not selected_files_to_process:
                    st.error("Vui lòng chọn ít nhất một tệp.")
                else:
                    conn = get_connection()
                    cursor = conn.cursor()
                    last_created_id = None

                    for filename in selected_files_to_process:
                        df, uploaded_file = file_dfs[filename]
                        file_hash = file_hashes[filename]

                        # Nếu nội dung đã có trong session thì không quét lại.
                        existing = cursor.execute(
                            "SELECT id FROM import_sessions WHERE file_hash = ?",
                            (file_hash,),
                        ).fetchone()

                        if existing:
                            st.warning(
                                f"Đã bỏ qua {filename}: nội dung đã được quét ở phiên #{existing[0]}."
                            )
                            continue

                        cursor.execute(
                            """
                            INSERT INTO import_sessions
                            (filename, file_hash, total_rows, status)
                            VALUES (?, ?, ?, 'ĐANG CHẠY')
                            """,
                            (filename, file_hash, len(df)),
                        )
                        session_id = cursor.lastrowid
                        conn.commit()

                        file_path = os.path.join(
                            UPLOAD_DIR, f"session_{session_id}_{filename}"
                        )
                        uploaded_file.seek(0)
                        with open(file_path, "wb") as output_file:
                            output_file.write(uploaded_file.read())

                        st.session_state["df_cache"][session_id] = (
                            df.copy(),
                            df.columns.tolist(),
                            file_path,
                        )

                        thread = threading.Thread(
                            target=process_data_background,
                            args=(session_id, df.copy(), df.columns.tolist(), 0),
                            daemon=True,
                        )
                        thread.start()
                        last_created_id = session_id

                    conn.close()

                    if last_created_id is not None:
                        st.session_state["latest_session_id"] = last_created_id
                        st.session_state["selected_nav"] = (
                            "📜 Lịch Sử & Xuất File Báo Cáo"
                        )
                        st.rerun()


# =========================
# VIEW 2: QUẢN LÝ / RESUME
# =========================
elif current_view == "🔄 Quản Lý Phiên & Tiếp Tục (Resume)":
    st.subheader("🔄 Theo Dõi Tiến Độ & Khôi Phục Phiên")

    conn = get_connection()
    sessions_df = pd.read_sql_query(
        "SELECT * FROM import_sessions ORDER BY id DESC",
        conn,
    )
    conn.close()

    if not sessions_df.empty:
        display_df = sessions_df.copy()
        display_df["Tỉ lệ sai"] = display_df.apply(
            lambda row: (
                f"{row['invalid_rows'] / row['total_rows'] * 100:.2f}%"
                if row["total_rows"] > 0
                else "0.00%"
            ),
            axis=1,
        )
        display_df["id"] = display_df["id"].apply(lambda value: f"#{value}")

        display_df = display_df[
            [
                "id",
                "filename",
                "total_rows",
                "processed_rows",
                "valid_rows",
                "invalid_rows",
                "Tỉ lệ sai",
                "status",
                "created_at",
            ]
        ]
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

        session_ids = sessions_df["id"].astype(int).tolist()
        default_index = 0

        latest_id = st.session_state.get("latest_session_id")
        if latest_id in session_ids:
            default_index = session_ids.index(latest_id)

        selected_id = st.selectbox(
            "Chọn phiên cần theo dõi:",
            session_ids,
            index=default_index,
            key="selectbox_session_tracking",
            format_func=lambda value: f"Phiên #{value}",
        )

        conn = get_connection()
        info = pd.read_sql_query(
            "SELECT * FROM import_sessions WHERE id = ?",
            conn,
            params=(selected_id,),
        ).iloc[0]
        conn.close()

        total = int(info["total_rows"])
        processed = int(info["processed_rows"])
        status = info["status"]
        valid = int(info["valid_rows"])
        invalid = int(info["invalid_rows"])
        filename_val = info["filename"]

        progress = processed / total if total > 0 else 0
        error_rate = invalid / total * 100 if total > 0 else 0

        st.markdown("---")
        st.markdown(f"#### 📊 Chi tiết tiến độ Phiên #{selected_id}")
        st.progress(min(progress, 1.0))

        col1, col2, col3, col4, col5 = st.columns(5)
        col1.metric("📁 Tổng số dòng", f"{total:,}")
        col2.metric("⚙️ Đã xử lý", f"{processed:,}", delta=f"{progress * 100:.1f}%")
        col3.metric("✅ Hợp lệ", f"{valid:,}")
        col4.metric("⚠️ Dòng lỗi", f"{invalid:,}", delta_color="inverse")
        col5.metric("📉 Tỉ lệ sai", f"{error_rate:.2f}%")

        btn1, btn2, btn3 = st.columns(3)

        with btn1:
            if st.button("🔄 Cập Nhật Trạng Thái", key="btn_refresh_status"):
                st.rerun()

        with btn2:
            if status != "HOÀN THÀNH":
                if st.button("✅ Đánh Dấu Hoàn Thành", key="btn_force_complete"):
                    conn = get_connection()
                    conn.execute(
                        "UPDATE import_sessions SET status = 'HOÀN THÀNH' WHERE id = ?",
                        (selected_id,),
                    )
                    conn.commit()
                    conn.close()
                    st.success(f"Đã đánh dấu phiên #{selected_id} hoàn thành.")
                    st.rerun()

        with btn3:
            if status != "HOÀN THÀNH" and processed < total:
                file_path = os.path.join(
                    UPLOAD_DIR, f"session_{selected_id}_{filename_val}"
                )
                cache_exists = (
                    "df_cache" in st.session_state
                    and selected_id in st.session_state["df_cache"]
                )

                if os.path.exists(file_path) or cache_exists:
                    if st.button(
                        "▶️ Tiếp Tục Xử Lý",
                        type="primary",
                        key="btn_resume_process",
                    ):
                        df_cached = None
                        required_cols = None

                        if cache_exists:
                            cached_data = st.session_state["df_cache"][selected_id]
                            df_cached = cached_data[0].copy()
                            required_cols = cached_data[1]
                        elif os.path.exists(file_path):
                            try:
                                df_cached = load_file_from_path(
                                    file_path, filename_val
                                )
                                if df_cached is not None:
                                    required_cols = df_cached.columns.tolist()
                            except Exception as e:
                                st.error(f"Không thể đọc lại file: {e}")

                        if df_cached is not None:
                            # processed_rows là số dòng đã lưu; dùng trực tiếp
                            # làm chỉ số bắt đầu theo quy ước index từ 0.
                            resume_from = processed

                            thread = threading.Thread(
                                target=process_data_background,
                                args=(
                                    selected_id,
                                    df_cached,
                                    required_cols,
                                    resume_from,
                                ),
                                daemon=True,
                            )
                            thread.start()

                            st.success(
                                f"Đã tiếp tục phiên #{selected_id} từ dòng "
                                f"{resume_from + 1}."
                            )
                            st.rerun()
                else:
                    st.warning(
                        "Không tìm thấy file gốc trên máy chủ nên không thể Resume."
                    )


# =========================
# VIEW 3: LỊCH SỬ / XUẤT FILE
# =========================
elif current_view == "📜 Lịch Sử & Xuất File Báo Cáo":
    st.subheader("📜 Bảng Lịch Sử & Xuất Báo Cáo")

    conn = get_connection()
    history_df = pd.read_sql_query(
        "SELECT * FROM import_sessions ORDER BY id DESC",
        conn,
    )

    available_ids = []

    if not history_df.empty:
        search_query = st.text_input(
            "🔍 Tìm kiếm tên tệp hoặc mã phiên:",
            placeholder="Ví dụ: 7 hoặc data_test",
            key="global_search_input",
        )

        filtered_history_df = history_df.copy()
        if search_query.strip():
            query = search_query.strip().lower()
            filtered_history_df = filtered_history_df[
                filtered_history_df["filename"].str.lower().str.contains(
                    query, na=False
                )
                | filtered_history_df["id"].astype(str).str.contains(
                    query, na=False
                )
            ]

        view_history = filtered_history_df.copy()
        view_history["Tỉ lệ sai"] = view_history.apply(
            lambda row: (
                f"{row['invalid_rows'] / row['total_rows'] * 100:.2f}%"
                if row["total_rows"] > 0
                else "0.00%"
            ),
            axis=1,
        )
        view_history["raw_id"] = view_history["id"].astype(int)
        view_history["id"] = view_history["id"].apply(lambda value: f"#{value}")

        view_history = view_history[
            [
                "raw_id",
                "id",
                "filename",
                "total_rows",
                "processed_rows",
                "valid_rows",
                "invalid_rows",
                "Tỉ lệ sai",
                "status",
                "created_at",
            ]
        ]
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
            selection_mode="single-row",
            on_select="rerun",
            key="history_multi_select_table",
        )

        if "selected_target_session_id" not in st.session_state:
            st.session_state["selected_target_session_id"] = int(
                history_df["id"].iloc[0]
            )

        selected_rows = []
        if event and "selection" in event:
            selected_rows = event.selection.get("rows", [])

        if selected_rows:
            row_index = selected_rows[-1]
            if row_index < len(view_history):
                clicked_id = int(view_history.iloc[row_index]["raw_id"])
                st.session_state["selected_target_session_id"] = clicked_id

        available_ids = (
            filtered_history_df["id"].astype(int).tolist()
            if not filtered_history_df.empty
            else history_df["id"].astype(int).tolist()
        )

        if available_ids:
            current_target = st.session_state.get("selected_target_session_id")
            if current_target not in available_ids:
                st.session_state["selected_target_session_id"] = available_ids[0]

        st.markdown("---")
        st.subheader("📥 Xuất Báo Cáo Excel")

        col_export1, col_export2 = st.columns(2)

        with col_export1:
            st.markdown("#### 1. Báo cáo lịch sử phiên")
            st.write("Tải danh sách phiên làm việc tổng hợp.")

            excel_buffer = export_history_to_excel(history_df)
            st.download_button(
                label="📊 Tải Báo Cáo Lịch Sử Phiên",
                data=excel_buffer,
                file_name="DataClean_Lich_Su_Phien.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="dl_btn_history_excel",
            )

        with col_export2:
            st.markdown("#### 2. File gốc đánh dấu lỗi")
            st.write("Chọn phiên để xuất dữ liệu gốc, kèm ghi chú và tô đỏ dòng lỗi.")

            selected_target = st.session_state.get(
                "selected_target_session_id",
                available_ids[0] if available_ids else None,
            )

            if available_ids:
                target_index = (
                    available_ids.index(selected_target)
                    if selected_target in available_ids
                    else 0
                )

                target_session_id = st.selectbox(
                    "Chọn phiên cần xuất:",
                    available_ids,
                    index=target_index,
                    key="select_session_for_marking_tab3",
                    format_func=lambda value: (
                        f"Phiên #{value} - "
                        f"{history_df.loc[history_df['id'] == value, 'filename'].iloc[0]}"
                    ),
                )

                st.session_state["selected_target_session_id"] = target_session_id

                target_filename = history_df.loc[
                    history_df["id"] == target_session_id, "filename"
                ].iloc[0]

                export_df_source = None

                if (
                    "df_cache" in st.session_state
                    and target_session_id in st.session_state["df_cache"]
                ):
                    export_df_source = st.session_state["df_cache"][
                        target_session_id
                    ][0]
                else:
                    file_path = os.path.join(
                        UPLOAD_DIR,
                        f"session_{target_session_id}_{target_filename}",
                    )

                    if os.path.exists(file_path):
                        try:
                            export_df_source = load_file_from_path(
                                file_path, target_filename
                            )
                        except Exception:
                            export_df_source = None

                    if export_df_source is None:
                        export_df_source = reconstruct_df_from_db(target_session_id)

                if export_df_source is not None:
                    marked_buffer = export_marked_file_to_excel(
                        target_session_id,
                        export_df_source,
                    )
                    st.download_button(
                        label=f"📝 Tải File Đánh Dấu Lỗi (Phiên #{target_session_id})",
                        data=marked_buffer,
                        file_name=f"DataClean_File_Danh_Dau_Loi_Phiên_{target_session_id}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        key="dl_btn_marked_excel",
                    )
                else:
                    st.info("Không tìm thấy dữ liệu gốc để xuất.")

    else:
        st.info("Chưa có phiên nhập dữ liệu nào trong hệ thống.")

    st.markdown("---")
    st.subheader("🔍 Tra Cứu Dòng Lỗi Chi Tiết")

    if not history_df.empty:
        all_session_ids = history_df["id"].astype(int).tolist()
        default_lookup_id = st.session_state.get(
            "selected_target_session_id",
            all_session_ids[0],
        )
        lookup_index = (
            all_session_ids.index(default_lookup_id)
            if default_lookup_id in all_session_ids
            else 0
        )

        lookup_session_id = st.selectbox(
            "Chọn phiên muốn xem chi tiết dòng lỗi:",
            all_session_ids,
            index=lookup_index,
            key="selectbox_lookup_error_id",
            format_func=lambda value: (
                f"Phiên #{value} - "
                f"{history_df.loc[history_df['id'] == value, 'filename'].iloc[0]}"
            ),
        )

        if st.button("🔍 Trích Xuất Dữ Liệu Lỗi", key="btn_lookup_error_details"):
            err_df = pd.read_sql_query(
                """
                SELECT row_number AS 'Dòng số',
                       error_reason AS 'Lý do lỗi'
                FROM invalid_data
                WHERE session_id = ?
                ORDER BY row_number
                """,
                conn,
                params=(lookup_session_id,),
            )

            if not err_df.empty:
                st.error(
                    f"Phát hiện {len(err_df)} dòng lỗi ở phiên "
                    f"#{lookup_session_id}:"
                )
                st.dataframe(err_df, use_container_width=True)

                csv_bytes = err_df.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    label="📥 Tải Báo Cáo Lỗi Chi Tiết (CSV)",
                    data=csv_bytes,
                    file_name=f"DataClean_Bao_Cao_Loi_Phien_{lookup_session_id}.csv",
                    mime="text/csv",
                    key="dl_btn_error_csv",
                )
            else:
                st.success(
                    f"Phiên #{lookup_session_id} không có dòng lỗi."
                )

    conn.close()
