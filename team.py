import hashlib
import io
import json
import os
import sqlite3
import threading
from pathlib import Path

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


# =========================
# CẤU HÌNH
# =========================
st.set_page_config(
    page_title="DataClean Pro - Kiểm tra dữ liệu",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

DB_NAME = "dataclean_app_integrated.db"
UPLOAD_DIR = Path("uploaded_temp")
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

# Khóa dùng để tránh khởi chạy trùng một phiên trong cùng tiến trình ứng dụng.
_active_sessions = set()
_active_sessions_lock = threading.Lock()


# =========================
# GIAO DIỆN
# =========================
st.markdown(
    """
    <style>
    .stApp { background-color: #121212; color: #e0e0e0; }
    [data-testid="stSidebar"] {
        background-color: #1a1a1a;
        border-right: 1px solid #2d2d2d;
    }
    div.block-container {
        background-color: #1e1e1e;
        padding: 2rem 2.5rem;
        border-radius: 16px;
        margin-top: 1rem;
        margin-bottom: 2rem;
        border: 1px solid #2d2d2d;
    }
    h1, h2, h3, h4 { color: #ffffff !important; }
    div[data-testid="stMetric"] {
        background-color: #252525;
        padding: 16px;
        border-radius: 12px;
        border-left: 5px solid #3b82f6;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================
# CƠ SỞ DỮ LIỆU
# =========================
def db_connect():
    return sqlite3.connect(DB_NAME, timeout=30)


def init_db():
    with db_connect() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS import_sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                file_hash TEXT,
                file_path TEXT,
                required_cols TEXT,
                total_rows INTEGER NOT NULL DEFAULT 0,
                processed_rows INTEGER NOT NULL DEFAULT 0,
                valid_rows INTEGER NOT NULL DEFAULT 0,
                invalid_rows INTEGER NOT NULL DEFAULT 0,
                status TEXT NOT NULL DEFAULT 'ĐANG CHẠY',
                error_message TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # Bổ sung cột cho cơ sở dữ liệu đã tạo bởi phiên bản cũ.
        columns = {
            row[1]
            for row in conn.execute("PRAGMA table_info(import_sessions)").fetchall()
        }
        migrations = {
            "file_hash": "TEXT",
            "file_path": "TEXT",
            "required_cols": "TEXT",
            "error_message": "TEXT",
        }
        for column, definition in migrations.items():
            if column not in columns:
                conn.execute(
                    f"ALTER TABLE import_sessions ADD COLUMN {column} {definition}"
                )

        conn.execute("""
            CREATE TABLE IF NOT EXISTS valid_data (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id INTEGER NOT NULL,
                row_number INTEGER NOT NULL,
                data_json TEXT NOT NULL
            )
        """)

        conn.execute("""
            CREATE TABLE IF NOT EXISTS invalid_data (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id INTEGER NOT NULL,
                row_number INTEGER NOT NULL,
                raw_data TEXT NOT NULL,
                error_reason TEXT NOT NULL
            )
        """)

        conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_valid_session
            ON valid_data(session_id)
        """)
        conn.execute("""
            CREATE INDEX IF NOT EXISTS idx_invalid_session
            ON invalid_data(session_id)
        """)


init_db()


# =========================
# ĐỌC VÀ KIỂM TRA FILE
# =========================
def get_file_hash(file_bytes):
    # SHA-256 phù hợp hơn MD5 cho nhận diện nội dung.
    return hashlib.sha256(file_bytes).hexdigest()


def load_uploaded_file(file_obj, header_row=0):
    filename = file_obj.name.lower()
    file_obj.seek(0)

    try:
        if filename.endswith(".csv"):
            try:
                df = pd.read_csv(
                    file_obj, encoding="utf-8-sig", header=header_row
                )
            except UnicodeDecodeError:
                file_obj.seek(0)
                df = pd.read_csv(
                    file_obj, encoding="latin1", header=header_row
                )
        elif filename.endswith((".xlsx", ".xls")):
            df = pd.read_excel(file_obj, header=header_row)
        elif filename.endswith(".json"):
            df = pd.read_json(file_obj)
        else:
            raise ValueError("Định dạng file chưa được hỗ trợ.")

        columns = []
        for i, column in enumerate(df.columns):
            if pd.isna(column) or str(column).strip() == "" or "Unnamed" in str(column):
                columns.append(f"Cột {i + 1}")
            else:
                columns.append(str(column).strip())

        df.columns = columns
        df = df.dropna(how="all").reset_index(drop=True)
        return df

    except Exception as exc:
        raise ValueError(f"Không đọc được file: {exc}") from exc


def validate_row(row, required_cols):
    errors = []
    for col in required_cols:
        value = row.get(col)
        if pd.isna(value) or str(value).strip() == "":
            errors.append(f"Cột '{col}' bị trống")
    return errors


def json_safe_dict(row):
    result = {}
    for key, value in row.items():
        if pd.isna(value):
            result[str(key)] = None
        elif hasattr(value, "item"):
            result[str(key)] = value.item()
        else:
            result[str(key)] = value
    return result


# =========================
# TIẾN TRÌNH NỀN
# =========================
def process_data_background(session_id, df, required_cols, start_row=0):
    with _active_sessions_lock:
        if session_id in _active_sessions:
            return
        _active_sessions.add(session_id)

    try:
        start_row = max(0, min(int(start_row), len(df)))

        with db_connect() as conn:
            conn.execute(
                "UPDATE import_sessions SET status = 'ĐANG CHẠY', error_message = NULL WHERE id = ?",
                (session_id,),
            )

            for index in range(start_row, len(df)):
                row = df.iloc[index]
                row_number = index + 1
                errors = validate_row(row, required_cols)
                row_json = json.dumps(
                    json_safe_dict(row.to_dict()),
                    ensure_ascii=False,
                    default=str,
                )

                if errors:
                    conn.execute(
                        """
                        INSERT INTO invalid_data
                            (session_id, row_number, raw_data, error_reason)
                        VALUES (?, ?, ?, ?)
                        """,
                        (session_id, row_number, row_json, "; ".join(errors)),
                    )
                    conn.execute(
                        """
                        UPDATE import_sessions
                        SET processed_rows = ?, invalid_rows = invalid_rows + 1
                        WHERE id = ?
                        """,
                        (row_number, session_id),
                    )
                else:
                    conn.execute(
                        """
                        INSERT INTO valid_data
                            (session_id, row_number, data_json)
                        VALUES (?, ?, ?)
                        """,
                        (session_id, row_number, row_json),
                    )
                    conn.execute(
                        """
                        UPDATE import_sessions
                        SET processed_rows = ?, valid_rows = valid_rows + 1
                        WHERE id = ?
                        """,
                        (row_number, session_id),
                    )

                # Commit định kỳ để tiến độ được lưu dù ứng dụng dừng bất ngờ.
                if row_number % 50 == 0:
                    conn.commit()

            conn.execute(
                """
                UPDATE import_sessions
                SET status = 'HOÀN THÀNH',
                    processed_rows = total_rows,
                    error_message = NULL
                WHERE id = ?
                """,
                (session_id,),
            )
            conn.commit()

    except Exception as exc:
        try:
            with db_connect() as conn:
                conn.execute(
                    """
                    UPDATE import_sessions
                    SET status = 'LỖI', error_message = ?
                    WHERE id = ?
                    """,
                    (str(exc), session_id),
                )
        except Exception:
            pass
    finally:
        with _active_sessions_lock:
            _active_sessions.discard(session_id)


def start_processing(session_id, df, required_cols, start_row=0):
    thread = threading.Thread(
        target=process_data_background,
        args=(session_id, df.copy(), list(required_cols), start_row),
        daemon=True,
    )
    thread.start()


# =========================
# XUẤT EXCEL
# =========================
def export_history_to_excel(history_df):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Lịch sử phiên"

    headers = [
        "Mã phiên",
        "Tên tệp",
        "Tổng số dòng",
        "Đã xử lý",
        "Hợp lệ",
        "Bị lỗi",
        "Tỉ lệ lỗi",
        "Trạng thái",
        "Thời gian tạo",
    ]

    ws.append(["DANH SÁCH LỊCH SỬ PHIÊN - DATACLEAN PRO"])
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    ws["A1"].fill = PatternFill("solid", fgColor="1E3C72")
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 30

    ws.append(headers)

    header_fill = PatternFill("solid", fgColor="2A5298")
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )

    for cell in ws[2]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = thin_border

    for row in history_df.itertuples(index=False):
        total = int(row.total_rows or 0)
        invalid = int(row.invalid_rows or 0)
        error_rate = f"{invalid / total * 100:.2f}%" if total else "0.00%"

        ws.append([
            f"#{row.id}",
            row.filename,
            total,
            int(row.processed_rows or 0),
            int(row.valid_rows or 0),
            invalid,
            error_rate,
            row.status,
            str(row.created_at),
        ])

    for row in ws.iter_rows(min_row=3):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    widths = [12, 32, 16, 16, 14, 14, 14, 18, 22]
    for index, width in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(index)].width = width

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def export_marked_file_to_excel(session_id, original_df):
    with db_connect() as conn:
        errors_df = pd.read_sql_query(
            """
            SELECT row_number, error_reason
            FROM invalid_data
            WHERE session_id = ?
            """,
            conn,
            params=(session_id,),
        )

    errors = {
        int(row.row_number): row.error_reason
        for row in errors_df.itertuples(index=False)
    }

    export_df = original_df.copy()
    export_df["Ghi chú kiểm tra lỗi"] = [
        f"LỖI: {errors[row_number]}" if row_number in errors else "Hợp lệ"
        for row_number in range(1, len(export_df) + 1)
    ]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Kết quả kiểm tra"

    ws.append(list(export_df.columns))
    header_fill = PatternFill("solid", fgColor="1E3C72")
    red_fill = PatternFill("solid", fgColor="FFC7CE")
    red_font = Font(color="9C0006", bold=True)
    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = thin_border

    for row_number, values in enumerate(
        export_df.itertuples(index=False, name=None), start=1
    ):
        safe_values = [
            None if pd.isna(value) else value
            for value in values
        ]
        ws.append(safe_values)
        excel_row = row_number + 1
        is_error = row_number in errors

        for cell in ws[excel_row]:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if is_error:
                cell.fill = red_fill
        if is_error:
            ws.cell(
                row=excel_row, column=len(export_df.columns)
            ).font = red_font

    for col_idx, column in enumerate(export_df.columns, start=1):
        max_length = len(str(column))
        for cell in ws.iter_cols(
            min_col=col_idx, max_col=col_idx, min_row=2
        ):
            for item in cell:
                if item.value is not None:
                    max_length = max(max_length, len(str(item.value)))
        ws.column_dimensions[
            openpyxl.utils.get_column_letter(col_idx)
        ].width = min(max_length + 2, 45)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


def reconstruct_original_df(session_id):
    with db_connect() as conn:
        session = conn.execute(
            "SELECT total_rows FROM import_sessions WHERE id = ?",
            (session_id,),
        ).fetchone()

        valid = pd.read_sql_query(
            """
            SELECT row_number, data_json
            FROM valid_data WHERE session_id = ?
            """,
            conn,
            params=(session_id,),
        )
        invalid = pd.read_sql_query(
            """
            SELECT row_number, raw_data AS data_json
            FROM invalid_data WHERE session_id = ?
            """,
            conn,
            params=(session_id,),
        )

    records = []
    for frame in (valid, invalid):
        for item in frame.itertuples(index=False):
            try:
                data = json.loads(item.data_json)
                records.append((int(item.row_number), data))
            except (ValueError, TypeError):
                continue

    if not records:
        return pd.DataFrame()

    records.sort(key=lambda item: item[0])
    return pd.DataFrame([item[1] for item in records]).reset_index(drop=True)


# =========================
# ĐIỀU HƯỚNG
# =========================
if "selected_nav" not in st.session_state:
    st.session_state["selected_nav"] = "Tải file & cấu hình"

with st.sidebar:
    st.title("⚡ DataClean Pro")
    st.caption("Hệ thống kiểm tra dữ liệu tự động")
    st.markdown("---")

    if st.button("📥 Tải file & cấu hình", use_container_width=True):
        st.session_state["selected_nav"] = "Tải file & cấu hình"

    if st.button("🔄 Quản lý phiên", use_container_width=True):
        st.session_state["selected_nav"] = "Quản lý phiên"

    if st.button("📜 Lịch sử & xuất file", use_container_width=True):
        st.session_state["selected_nav"] = "Lịch sử & xuất file"

    st.markdown("---")
    st.caption(
        "Hỗ trợ CSV, XLSX, XLS và JSON. "
        "Tệp tải lên được lưu cục bộ trong thư mục uploaded_temp."
    )

st.title("⚡ DataClean Pro")
st.caption("Tải, kiểm tra, theo dõi tiến độ và xuất báo cáo dữ liệu.")
st.markdown("---")


# =========================
# TRANG 1: TẢI FILE
# =========================
if st.session_state["selected_nav"] == "Tải file & cấu hình":
    st.subheader("1. Tải tệp dữ liệu")

    header_row = st.number_input(
        "Dòng tiêu đề (đếm từ 0)",
        min_value=0,
        max_value=50,
        value=0,
        help="Ví dụ nhập 1 nếu dòng thứ hai trong file chứa tên cột.",
    )

    uploaded_files = st.file_uploader(
        "Chọn một hoặc nhiều tệp",
        type=["csv", "xlsx", "xls", "json"],
        accept_multiple_files=True,
    )

    parsed_files = {}

    if uploaded_files:
        st.success(f"Đã chọn {len(uploaded_files)} tệp.")

        for uploaded_file in uploaded_files:
            try:
                raw_bytes = uploaded_file.getvalue()
                file_hash = get_file_hash(raw_bytes)
                df = load_uploaded_file(uploaded_file, int(header_row))

                with db_connect() as conn:
                    existing = conn.execute(
                        """
                        SELECT id, status
                        FROM import_sessions
                        WHERE file_hash = ?
                        ORDER BY id DESC LIMIT 1
                        """,
                        (file_hash,),
                    ).fetchone()

                parsed_files[uploaded_file.name] = {
                    "df": df,
                    "bytes": raw_bytes,
                    "hash": file_hash,
                }

                status = (
                    f"Đã quét trước: phiên #{existing[0]} ({existing[1]})"
                    if existing
                    else "Chưa quét"
                )
                st.write(
                    f"**{uploaded_file.name}** — "
                    f"{len(df):,} dòng, {len(df.columns)} cột — {status}"
                )
                st.dataframe(df.head(5), use_container_width=True)

            except Exception as exc:
                st.error(f"{uploaded_file.name}: {exc}")

        if parsed_files:
            st.markdown("### 2. Chọn cột bắt buộc")
            all_columns = sorted({
                str(column)
                for item in parsed_files.values()
                for column in item["df"].columns
            })

            required_cols = st.multiselect(
                "Các cột không được để trống",
                options=all_columns,
                default=[],
            )

            selected_names = st.multiselect(
                "Chọn tệp cần xử lý",
                options=list(parsed_files.keys()),
                default=list(parsed_files.keys()),
            )

            if st.button(
                "🚀 Bắt đầu xử lý",
                type="primary",
                disabled=not selected_names,
            ):
                new_session_ids = []

                for filename in selected_names:
                    item = parsed_files[filename]
                    df = item["df"]
                    file_path = UPLOAD_DIR / (
                        f"{item['hash']}_{Path(filename).name}"
                    )
                    file_path.write_bytes(item["bytes"])

                    with db_connect() as conn:
                        cursor = conn.execute(
                            """
                            INSERT INTO import_sessions
                                (filename, file_hash, file_path, required_cols,
                                 total_rows, status)
                            VALUES (?, ?, ?, ?, ?, 'ĐANG CHẠY')
                            """,
                            (
                                filename,
                                item["hash"],
                                str(file_path),
                                json.dumps(required_cols, ensure_ascii=False),
                                len(df),
                            ),
                        )
                        session_id = cursor.lastrowid

                    start_processing(session_id, df, required_cols)
                    new_session_ids.append(session_id)

                st.success(
                    "Đã tạo phiên: "
                    + ", ".join(f"#{session_id}" for session_id in new_session_ids)
                )
                st.session_state["selected_nav"] = "Quản lý phiên"
                st.rerun()


# =========================
# TRANG 2: QUẢN LÝ VÀ RESUME
# =========================
elif st.session_state["selected_nav"] == "Quản lý phiên":
    st.subheader("Theo dõi tiến độ và tiếp tục phiên")

    with db_connect() as conn:
        sessions_df = pd.read_sql_query(
            "SELECT * FROM import_sessions ORDER BY id DESC",
            conn,
        )

    if sessions_df.empty:
        st.info("Chưa có phiên xử lý nào.")
    else:
        for row in sessions_df.itertuples(index=False):
            total = int(row.total_rows or 0)
            processed = int(row.processed_rows or 0)
            progress = min(processed / total, 1.0) if total else 1.0

            with st.container(border=True):
                st.markdown(f"### Phiên #{row.id}: {row.filename}")
                st.write(
                    f"Trạng thái: **{row.status}** · "
                    f"Tiến độ: **{processed:,}/{total:,}** · "
                    f"Hợp lệ: **{row.valid_rows:,}** · "
                    f"Lỗi: **{row.invalid_rows:,}**"
                )
                st.progress(progress)

                if row.error_message:
                    st.error(row.error_message)

                resumable = row.status in ("LỖI", "TẠM DỪNG", "ĐANG CHẠY")
                if resumable and st.button(
                    f"▶️ Tiếp tục phiên #{row.id}",
                    key=f"resume_{row.id}",
                ):
                    if not row.file_path or not Path(row.file_path).exists():
                        st.error(
                            "Không tìm thấy tệp gốc. "
                            "Hãy tải lại tệp ở trang Tải file."
                        )
                    else:
                        try:
                            df = load_uploaded_file(
                                type(
                                    "UploadedFileProxy",
                                    (),
                                    {
                                        "name": row.filename,
                                        "seek": lambda self, *_: None,
                                    },
                                )(),
                                0,
                            )
                        except Exception:
                            df = None

                        # Đọc lại trực tiếp từ file đã lưu, dùng cấu hình header 0.
                        try:
                            if row.filename.lower().endswith(".csv"):
                                try:
                                    df = pd.read_csv(
                                        row.file_path, encoding="utf-8-sig"
                                    )
                                except UnicodeDecodeError:
                                    df = pd.read_csv(
                                        row.file_path, encoding="latin1"
                                    )
                            elif row.filename.lower().endswith((".xlsx", ".xls")):
                                df = pd.read_excel(row.file_path)
                            elif row.filename.lower().endswith(".json"):
                                df = pd.read_json(row.file_path)

                            if df is None:
                                raise ValueError("Không đọc được tệp đã lưu.")

                            df = df.dropna(how="all").reset_index(drop=True)
                            required = json.loads(row.required_cols or "[]")
                            start_row = int(row.processed_rows or 0)

                            if start_row >= len(df):
                                with db_connect() as conn:
                                    conn.execute(
                                        """
                                        UPDATE import_sessions
                                        SET status = 'HOÀN THÀNH'
                                        WHERE id = ?
                                        """,
                                        (row.id,),
                                    )
                                st.success("Phiên đã xử lý hết dữ liệu.")
                            else:
                                start_processing(
                                    row.id, df, required, start_row
                                )
                                st.success(f"Đã tiếp tục phiên #{row.id}.")
                            st.rerun()

                        except Exception as exc:
                            st.error(f"Không thể khôi phục phiên: {exc}")

        if st.button("🔄 Làm mới tiến độ"):
            st.rerun()


# =========================
# TRANG 3: LỊCH SỬ VÀ XUẤT
# =========================
else:
    st.subheader("Lịch sử phiên và xuất báo cáo Excel")

    with db_connect() as conn:
        history_df = pd.read_sql_query(
            "SELECT * FROM import_sessions ORDER BY id DESC",
            conn,
        )

    if history_df.empty:
        st.info("Chưa có lịch sử xử lý.")
    else:
        display_df = history_df.copy()
        display_df["Tỉ lệ lỗi"] = display_df.apply(
            lambda row: (
                f"{row['invalid_rows'] / row['total_rows'] * 100:.2f}%"
                if row["total_rows"] > 0
                else "0.00%"
            ),
            axis=1,
        )

        st.dataframe(
            display_df[
                [
                    "id",
                    "filename",
                    "total_rows",
                    "processed_rows",
                    "valid_rows",
                    "invalid_rows",
                    "Tỉ lệ lỗi",
                    "status",
                    "created_at",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )

        st.download_button(
            "⬇️ Tải lịch sử Excel",
            data=export_history_to_excel(history_df),
            file_name="lich_su_dataclean.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        completed_ids = history_df["id"].astype(int).tolist()
        selected_session = st.selectbox(
            "Chọn phiên để xuất file đã đánh dấu lỗi",
            completed_ids,
            format_func=lambda session_id: (
                f"Phiên #{session_id} — "
                f"{history_df.loc[history_df['id'] == session_id, 'filename'].iloc[0]}"
            ),
        )

        if st.button("📊 Tạo file Excel có đánh dấu lỗi"):
            session = history_df.loc[
                history_df["id"] == selected_session
            ].iloc[0]

            original_df = None
            file_path = session.get("file_path")

            if file_path and isinstance(file_path, str) and Path(file_path).exists():
                filename = str(session["filename"]).lower()
                try:
                    if filename.endswith(".csv"):
                        try:
                            original_df = pd.read_csv(
                                file_path, encoding="utf-8-sig"
                            )
                        except UnicodeDecodeError:
                            original_df = pd.read_csv(
                                file_path, encoding="latin1"
                            )
                    elif filename.endswith((".xlsx", ".xls")):
                        original_df = pd.read_excel(file_path)
                    elif filename.endswith(".json"):
                        original_df = pd.read_json(file_path)
                except Exception as exc:
                    st.error(f"Không đọc được tệp gốc: {exc}")

            if original_df is None:
                original_df = reconstruct_original_df(selected_session)

            if original_df.empty:
                st.error("Không có dữ liệu để xuất.")
            else:
                output = export_marked_file_to_excel(
                    selected_session, original_df
                )
                st.download_button(
                    "⬇️ Tải file kết quả",
                    data=output,
                    file_name=f"ket_qua_phien_{selected_session}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"download_marked_{selected_session}",
                )
