import sqlite3
import threading
import time
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="DA11 - Xử Lý Nền, Khôi Phục & Quản Lý Phiên", layout="wide"
)

DB_NAME = "da11_advanced.db"


# 1. Khởi tạo Cơ sở dữ liệu SQLite
def init_db():
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    cursor = conn.cursor()

    # Bảng quản lý phiên nhập (hỗ trợ lưu vết và khôi phục tiến trình)
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

    # Bảng lưu dữ liệu hợp lệ
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


# 2. Logic Kiểm tra cấu trúc & tính hợp lệ từng dòng
def validate_row(row):
    errors = []
    name = str(row.get("name", "")).strip()
    email = str(row.get("email", "")).strip()

    if not name or pd.isna(row.get("name")):
        errors.append("Tên bị trống")

    if not email or "@" not in email or pd.isna(row.get("email")):
        errors.append("Email sai định dạng (thiếu @)")

    try:
        age = int(row.get("age"))
        if age <= 0 or age > 120:
            errors.append("Tuổi phải từ 1 đến 120")
    except (ValueError, TypeError):
        errors.append("Tuổi không phải là số")

    return errors


# 3. Hàm Xử Lý Nền (Background Processing Worker)
def process_data_background(session_id, df, start_row=0):
    """Chạy ngầm trong một Thread riêng biệt, hỗ trợ resume từ start_row."""
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    cursor = conn.cursor()

    batch_valid = []
    batch_invalid = []

    try:
        for index in range(start_row, len(df)):
            row = df.iloc[index]
            row_num = index + 1
            errors = validate_row(row)

            if not errors:
                batch_valid.append(
                    (
                        session_id,
                        str(row["name"]),
                        str(row["email"]),
                        int(row["age"]),
                    )
                )
            else:
                batch_invalid.append(
                    (session_id, row_num, str(row.to_dict()), "; ".join(errors))
                )

            # Ghi batch vào DB mỗi 50 dòng hoặc khi kết thúc
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

                # Cập nhật tiến độ xử lý vào bảng phiên
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

            # Giả lập độ trễ nhẹ giúp hệ thống mượt mà
            time.sleep(0.001)

        # Đánh dấu phiên đã hoàn thành
        cursor.execute(
            "UPDATE import_sessions SET status = 'COMPLETED' WHERE id = ?",
            (session_id,),
        )
        conn.commit()
    except Exception as e:
        print(f"Lỗi tiến trình nền: {e}")
    finally:
        conn.close()


# 4. Giao diện Streamlit
st.title("⚡ DA11 - Xử Lý Nền, Khôi Phục & Quản Lý Phiên")

tab1, tab2, tab3 = st.tabs(
    ["📥 Nhập File Mới", "🔄 Quản Lý & Tiếp Tục Phiên", "📜 Lịch Sử & Báo Cáo"]
)

# --- TAB 1: NHẬP FILE MỚI ---
with tab1:
    st.subheader("Tải lên file CSV dữ liệu lớn")
    uploaded_file = st.file_uploader(
        "Chọn file CSV (Yêu cầu cột: name, email, age)", type=["csv"]
    )

    if uploaded_file is not None:
        df = pd.read_csv(uploaded_file)
        st.success(f"📊 Đã nạp file: **{uploaded_file.name}** ({len(df)} dòng)")

        req_cols = {"name", "email", "age"}
        if not req_cols.issubset(set(df.columns)):
            st.error(
                f"❌ Sai cấu trúc cột! File cần có các cột: {', '.join(req_cols)}"
            )
        else:
            if st.button("🚀 Bắt đầu xử lý nền"):
                # Tạo phiên mới trong cơ sở dữ liệu
                conn = sqlite3.connect(DB_NAME, check_same_thread=False)
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT INTO import_sessions (filename, total_rows, status) VALUES (?, ?, 'RUNNING')",
                    (uploaded_file.name, len(df)),
                )
                session_id = cursor.lastrowid
                conn.commit()
                conn.close()

                # Lưu DataFrame vào session_state để hỗ trợ tính năng Resume
                if "uploaded_dfs" not in st.session_state:
                    st.session_state["uploaded_dfs"] = {}
                st.session_state["uploaded_dfs"][session_id] = df

                # Khởi chạy luồng chạy ngầm (Background Thread)
                thread = threading.Thread(
                    target=process_data_background, args=(session_id, df, 0)
                )
                thread.daemon = True
                thread.start()

                st.success(
                    f"✅ Đã khởi chạy tiến trình xử lý nền cho Phiên #{session_id}!"
                )
                st.info(
                    "👉 Chuyển sang tab **'Quản Lý & Tiếp Tục Phiên'** để theo dõi thanh tiến độ thời gian thực."
                )

# --- TAB 2: QUẢN LÝ TIẾN TRÌNH & RESUME KHI GIÁN ĐOẠN ---
with tab2:
    st.subheader("🔄 Theo dõi tiến độ thời gian thực & Khôi phục")

    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    sessions_df = pd.read_sql_query(
        "SELECT * FROM import_sessions ORDER BY id DESC", conn
    )
    conn.close()

    if not sessions_df.empty:
        st.dataframe(sessions_df, use_container_width=True)

        session_id_input = st.selectbox(
            "Chọn ID Phiên để tương tác:", sessions_df["id"].tolist()
        )

        if session_id_input:
            conn = sqlite3.connect(DB_NAME, check_same_thread=False)
            row_info = pd.read_sql_query(
                f"SELECT * FROM import_sessions WHERE id = {session_id_input}",
                conn,
            ).iloc[0]
            conn.close()

            total = row_info["total_rows"]
            processed = row_info["processed_rows"]
            status = row_info["status"]
            pct = processed / total if total > 0 else 0

            st.progress(pct)
            st.write(
                f"Trạng thái: **{status}** | Đã xử lý: **{processed}/{total}** dòng ({(pct*100):.1f}%)"
            )

            col1, col2 = st.columns(2)
            with col1:
                if st.button("🔄 Làm mới trạng thái"):
                    st.rerun()

            with col2:
                if status != "COMPLETED" and processed < total:
                    if st.button("▶️ Tiếp tục chạy từ điểm dừng (Resume)"):
                        if (
                            "uploaded_dfs" in st.session_state
                            and session_id_input in st.session_state["uploaded_dfs"]
                        ):
                            df_cached = st.session_state["uploaded_dfs"][
                                session_id_input
                            ]
                            # Tiếp tục chạy từ vị trí processed hiện tại
                            thread = threading.Thread(
                                target=process_data_background,
                                args=(session_id_input, df_cached, processed),
                            )
                            thread.daemon = True
                            thread.start()
                            st.success(
                                f"Đã tiếp tục xử lý phiên #{session_id_input} từ dòng {processed + 1}!"
                            )
                            st.rerun()
                        else:
                            st.error(
                                "⚠️ Vui lòng chọn lại file CSV ở Tab 1 để đưa vào bộ nhớ tạm trước khi tiếp tục."
                            )
    else:
        st.info("Chưa có phiên nhập nào.")

# --- TAB 3: LỊCH SỬ & XUẤT BÁO CÁO LỖI ---
with tab3:
    st.subheader("📜 Lịch sử & Xuất báo cáo dòng lỗi")
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    all_sessions = pd.read_sql_query(
        "SELECT * FROM import_sessions ORDER BY id DESC", conn
    )
    st.dataframe(all_sessions, use_container_width=True)

    view_session_id = st.number_input(
        "Nhập ID phiên muốn kiểm tra dòng lỗi:",
        min_value=1,
        step=1,
        value=1,
    )

    if st.button("🔍 Lọc danh sách dòng lỗi"):
        error_df = pd.read_sql_query(
            f"SELECT row_number AS 'Dòng số', raw_data AS 'Dữ liệu gốc', error_reason AS 'Lý do lỗi' FROM invalid_data WHERE session_id = {view_session_id}",
            conn,
        )
        if not error_df.empty:
            st.error(f"Tìm thấy {len(error_df)} dòng lỗi:")
            st.dataframe(error_df, use_container_width=True)

            csv_err = error_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Tải file báo cáo dòng lỗi (CSV)",
                data=csv_err,
                file_name=f"bao_cao_loi_phien_{view_session_id}.csv",
                mime="text/csv",
            )
        else:
            st.success("Phiên này không có dòng lỗi nào!")
    conn.close()
