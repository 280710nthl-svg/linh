import sqlite3
import threading
import time
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="DA11 - Nhập Dữ Liệu Nâng Cao", layout="wide"
)

DB_NAME = "da11_advanced.db"


# ---------------------------------------------------------
# 1. Khởi tạo Cơ sở dữ liệu SQLite
# ---------------------------------------------------------
def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    # Bảng quản lý phiên nhập nâng cao (hỗ trợ khôi phục tiến trình)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS import_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            total_rows INTEGER,
            processed_rows INTEGER DEFAULT 0,
            valid_rows INTEGER DEFAULT 0,
            invalid_rows INTEGER DEFAULT 0,
            status TEXT DEFAULT 'RUNNING', -- RUNNING, PAUSED, COMPLETED
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Bảng dữ liệu hợp lệ
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS valid_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER,
            name TEXT,
            email TEXT,
            age INTEGER
        )
    """)

    # Bảng lưu vết dòng lỗi
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


# ---------------------------------------------------------
# 2. Logic Kiểm tra dữ liệu (Validate)
# ---------------------------------------------------------
def validate_row(row):
    errors = []
    name = str(row.get("name", "")).strip()
    email = str(row.get("email", "")).strip()

    if not name or pd.isna(row.get("name")):
        errors.append("Tên bị trống")

    if not email or "@" not in email or pd.isna(row.get("email")):
        errors.append("Email sai định dạng")

    try:
        age = int(row.get("age"))
        if age <= 0 or age > 120:
            errors.append("Tuổi ngoài khoảng 1-120")
    except (ValueError, TypeError):
        errors.append("Tuổi không phải số")

    return errors


# ---------------------------------------------------------
# 3. Hàm Xử Lý Nền (Background Processing Worker)
# ---------------------------------------------------------
def process_data_background(session_id, df, start_row=0):
    """Chạy ngầm trong một Thread riêng biệt, có khả năng resume từ start_row."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    batch_valid = []
    batch_invalid = []

    for index in range(start_row, len(df)):
        row = df.iloc[index]
        row_num = index + 1
        errors = validate_row(row)

        if not errors:
            batch_valid.append(
                (session_id, str(row["name"]), str(row["email"]), int(row["age"]))
            )
        else:
            batch_invalid.append(
                (session_id, row_num, str(row.to_dict()), "; ".join(errors))
            )

        # Ghi vào CSDL theo từng đợt (Batch size = 50 dòng) để tối ưu hiệu năng
        if len(batch_valid) + len(batch_invalid) >= 50 or index == len(df) - 1:
            if batch_valid:
                cursor.executemany(
                    "INSERT INTO valid_data (session_id, name, email, age) VALUES (?, ?, ?, ?)",
                    batch_valid,
                )
            if batch_invalid:
                cursor.executemany(
                    "INSERT INTO invalid_data (session_id, row_number, raw_data, error_reason) VALUES (?, ?, ?, ?)",
                    batch_invalid,
                )

            # Cập nhật tiến độ của phiên vào DB
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

        # Giả lập thời gian xử lý nhẹ để quan sát tiến độ thời gian thực
        time.sleep(0.001)

    # Đánh dấu phiên đã hoàn tất
    cursor.execute(
        "UPDATE import_sessions SET status = 'COMPLETED' WHERE id = ?",
        (session_id,),
    )
    conn.commit()
    conn.close()


# ---------------------------------------------------------
# 4. Giao diện ứng dụng Streamlit
# ---------------------------------------------------------
st.title("⚡ DA11 - Nhập Dữ Liệu Nền & Khôi Phục Tiến Trình")

tab1, tab2, tab3 = st.tabs(
    ["📥 Nhập File Mới", "🔄 Tiếp Tục Phiên Dở Dang", "📜 Lịch Sử & Báo Cáo"]
)

# --- TAB 1: NHẬP FILE MỚI ---
with tab1:
    st.subheader("Tải file CSV")
    uploaded_file = st.file_uploader(
        "Chọn file CSV (Cột yêu cầu: name, email, age)", type=["csv"]
    )

    if uploaded_file is not None:
        df = pd.read_csv(uploaded_file)
        st.write(f"📊 Đã nạp: **{len(df)} dòng**")

        req_cols = {"name", "email", "age"}
        if not req_cols.issubset(set(df.columns)):
            st.error(
                f"❌ Sai cấu trúc cột! Cần có đủ các cột: {', '.join(req_cols)}"
            )
        else:
            if st.button("🚀 Bắt đầu nhập dữ liệu (Xử lý nền)"):
                # Tạo phiên làm việc mới trong DB
                conn = sqlite3.connect(DB_NAME)
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT INTO import_sessions (filename, total_rows, status) VALUES (?, ?, 'RUNNING')",
                    (uploaded_file.name, len(df)),
                )
                session_id = cursor.lastrowid
                conn.commit()
                conn.close()

                # Khởi chạy Thread nền độc lập
                thread = threading.Thread(
                    target=process_data_background, args=(session_id, df, 0)
                )
                thread.start()

                st.success(
                    f"✅ Đã khởi chạy tiến trình xử lý nền cho Phiên #{session_id}!"
                )
                st.info(
                    "Bạn có thể chuyển sang Tab **'Tiếp Tục Phiên Dở Dang'** để theo dõi tiến độ thời gian thực."
                )

# --- TAB 2: KHÔI PHỤC & THEO DÕI TIẾN TRÌNH ---
with tab2:
    st.subheader("Theo dõi thời gian thực & Tiếp tục khi gián đoạn")

    conn = sqlite3.connect(DB_NAME)
    active_sessions = pd.read_sql_query(
        "SELECT * FROM import_sessions WHERE status != 'COMPLETED' ORDER BY id DESC",
        conn,
    )
    conn.close()

    if not active_sessions.empty:
        st.write("📌 **Các phiên chưa hoàn thành:**")
        st.dataframe(active_sessions, use_container_width=True)

        selected_id = st.selectbox(
            "Chọn ID phiên muốn theo dõi/khôi phục:", active_sessions["id"]
        )

        if selected_id:
            conn = sqlite3.connect(DB_NAME)
            session_info = pd.read_sql_query(
                f"SELECT * FROM import_sessions WHERE id = {selected_id}", conn
            ).iloc[0]
            conn.close()

            total = session_info["total_rows"]
            processed = session_info["processed_rows"]
            pct = processed / total if total > 0 else 0

            st.progress(pct)
            st.write(
                f"Tiến độ hiện tại: **{processed}/{total} dòng** ({pct*100:.1f}%)"
            )

            col1, col2 = st.columns(2)
            with col1:
                if st.button("🔄 Cập nhật tiến độ"):
                    st.rerun()
            with col2:
                if processed < total and st.button("▶️ Khôi phục / Tiếp tục chạy phiên này"):
                    # Tải lại file CSV để resume (yêu cầu người dùng chọn lại file nếu cần hoặc lấy từ cache)
                    st.warning("Vui lòng đảm bảo file CSV tương ứng đang được chọn ở Tab 1 để Resume.")
    else:
        st.info("Không có phiên nào bị gián đoạn. Tất cả dữ liệu đã xử lý hoàn tất!")

# --- TAB 3: LỊCH SỬ & TÁCH DÒNG LỖI ---
with tab3:
    st.subheader("Lịch sử các phiên nhập")
    conn = sqlite3.connect(DB_NAME)
    all_sessions = pd.read_sql_query(
        "SELECT * FROM import_sessions ORDER BY id DESC", conn
    )
    st.dataframe(all_sessions, use_container_width=True)

    session_to_view = st.number_input(
        "Nhập Session ID để xem chi tiết dòng lỗi:",
        min_value=1,
        step=1,
        value=1,
    )

    if st.button("🔍 Xem báo cáo lỗi"):
        errors_df = pd.read_sql_query(
            f"SELECT row_number AS Dòng, raw_data AS 'Dữ liệu gốc', error_reason AS 'Lý do lỗi' FROM invalid_data WHERE session_id = {session_to_view}",
            conn,
        )

        if not errors_df.empty:
            st.error(
                f"Tìm thấy {len(errors_df)} dòng lỗi trong Session #{session_to_view}:"
            )
            st.dataframe(errors_df, use_container_width=True)

            csv_data = errors_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Tải file danh sách dòng lỗi (CSV)",
                data=csv_data,
                file_name=f"errors_session_{session_to_view}.csv",
                mime="text/csv",
            )
        else:
            st.success(
                f"Phiên #{session_to_view} không có dòng lỗi nào (hoặc Session ID không tồn tại)."
            )

    conn.close()
