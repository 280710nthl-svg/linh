import sqlite3
import threading
import time
import pandas as pd
import streamlit as st

# Cấu hình giao diện trang web
st.set_page_config(
    page_title="DA11 - Hệ Thống Nhập Dữ Liệu Thông Minh",
    page_icon="📊",
    layout="wide",
)

# Tùy chỉnh giao diện đồ họa đẹp mắt bằng CSS (Pure Python injection)
st.markdown(
    """
    <style>
        .main { background-color: #f8f9fa; }
        .stButton>button {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border-radius: 8px;
            padding: 0.6rem 1.2rem;
            font-weight: bold;
            border: none;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .stButton>button:hover {
            opacity: 0.9;
        }
        .metric-card {
            background-color: white;
            padding: 20px;
            border-radius: 12px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            text-align: center;
        }
    </style>
""",
    unsafe_allow_html=True,
)

DB_NAME = "da11_database.db"


# ---------------------------------------------------------
# KHỞI TẠO CƠ SỞ DỮ LIỆU SQLITE
# ---------------------------------------------------------
def init_db():
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    cursor = conn.cursor()

    # Bảng quản lý phiên nhập (Hỗ trợ khôi phục khi gián đoạn)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS import_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            total_rows INTEGER,
            processed_rows INTEGER DEFAULT 0,
            valid_rows INTEGER DEFAULT 0,
            invalid_rows INTEGER DEFAULT 0,
            status TEXT DEFAULT 'RUNNING',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Bảng lưu trữ dữ liệu hợp lệ
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS valid_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER,
            name TEXT,
            email TEXT,
            age INTEGER
        )
    """)

    # Bảng lưu vết các dòng bị lỗi cấu trúc / dữ liệu
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
# KIỂM TRA TÍNH HỢP LỆ TỪNG DÒNG
# ---------------------------------------------------------
def validate_row(row):
    errors = []
    name = str(row.get("name", "")).strip()
    email = str(row.get("email", "")).strip()

    if not name or pd.isna(row.get("name")):
        errors.append("Tên bị để trống")

    if not email or "@" not in email or pd.isna(row.get("email")):
        errors.append("Email không hợp lệ (thiếu @)")

    try:
        age = int(row.get("age"))
        if age <= 0 or age > 120:
            errors.append("Tuổi phải từ 1 đến 120")
    except (ValueError, TypeError):
        errors.append("Tuổi không phải là số nguyên")

    return errors


# ---------------------------------------------------------
# XỬ LÝ NỀN (BACKGROUND THREADING & RESUME)
# ---------------------------------------------------------
def process_data_background(session_id, df, start_row=0):
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

            # Ghi batch vào DB mỗi 50 dòng để tối ưu hiệu suất xử lý 5000+ dòng
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

            time.sleep(0.0005)  # Giả lập độ trễ siêu nhỏ cho tiến trình mượt mà

        cursor.execute(
            "UPDATE import_sessions SET status = 'COMPLETED' WHERE id = ?",
            (session_id,),
        )
        conn.commit()
    except Exception as e:
        print(f"Lỗi hệ thống nền: {e}")
    finally:
        conn.close()


# ---------------------------------------------------------
# GIAO DIỆN CHÍNH (STREAMLIT UI)
# ---------------------------------------------------------
st.title("🚀 DA11 - Quản Lý Nhập Dữ Liệu & Theo Dõi Tiến Trình")
st.markdown(
    "Ứng dụng thông minh xử lý file lớn (5000+ dòng), kiểm tra cấu trúc nghiêm ngặt, xử lý nền độc lập và khôi phục khi gián đoạn."
)

tab1, tab2, tab3 = st.tabs(
    ["📥 Tải Lên & Nhập File", "🔄 Tiến Trình & Khôi Phục", "📜 Lịch Sử & Báo Cáo"]
)

# --- TAB 1: TẢI LÊN FILE CSV ---
with tab1:
    st.subheader("Tải lên file dữ liệu CSV")
    st.info(
        "💡 **Yêu cầu cấu trúc:** File bắt buộc phải chứa các cột: `name`, `email`, `age`."
    )

    uploaded_file = st.file_uploader(
        "Chọn file CSV của bạn tại đây", type=["csv"]
    )

    if uploaded_file is not None:
        df = pd.read_csv(uploaded_file)
        st.success(
            f"✅ Đã tải file thành công: **{uploaded_file.name}** (Tổng số dòng: **{len(df):,}**)"
        )

        req_cols = {"name", "email", "age"}
        if not req_cols.issubset(set(df.columns)):
            st.error(
                f"❌ **Lỗi cấu trúc file!** Thiếu các cột bắt buộc: `{req_cols - set(df.columns)}`"
            )
        else:
            with st.expander("🔍 Xem trước 5 dòng đầu tiên"):
                st.dataframe(df.head(), use_container_width=True)

            if st.button("🚀 Bắt Đầu Xử Lý Nền (Background Process)"):
                conn = sqlite3.connect(DB_NAME, check_same_thread=False)
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT INTO import_sessions (filename, total_rows, status) VALUES (?, ?, 'RUNNING')",
                    (uploaded_file.name, len(df)),
                )
                session_id = cursor.lastrowid
                conn.commit()
                conn.close()

                # Lưu DataFrame vào session_state để hỗ trợ resume khi cần
                if "cached_dfs" not in st.session_state:
                    st.session_state["cached_dfs"] = {}
                st.session_state["cached_dfs"][session_id] = df

                # Khởi chạy Thread chạy ngầm độc lập
                thread = threading.Thread(
                    target=process_data_background, args=(session_id, df, 0)
                )
                thread.daemon = True
                thread.start()

                st.success(
                    f"🎉 Đã khởi chạy tiến trình nền cho Phiên **#{session_id}**!"
                )
                st.warning(
                    "👉 Hãy chuyển sang tab **'Tiến Trình & Khôi Phục'** để theo dõi biểu đồ thời gian thực."
                )

# --- TAB 2: QUẢN LÝ TIẾN TRÌNH & RESUME ---
with tab2:
    st.subheader("📊 Theo Dõi Tiến Độ Thời Gian Thực & Khôi Phục Gián Đoạn")

    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    sessions_df = pd.read_sql_query(
        "SELECT * FROM import_sessions ORDER BY id DESC", conn
    )
    conn.close()

    if not sessions_df.empty:
        st.dataframe(sessions_df, use_container_width=True)

        selected_id = st.selectbox(
            "Chọn ID Phiên cần theo dõi/thao tác:", sessions_df["id"].tolist()
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
            pct = processed / total if total > 0 else 0

            # Hiển thị đồ họa tiến độ đẹp mắt
            st.progress(pct)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Tổng số dòng", f"{total:,}")
            c2.metric("Đã xử lý", f"{processed:,} ({pct*100:.1f}%)")
            c3.metric("Hợp lệ (Đã lưu)", f"{valid:,}")
            c4.metric("Dòng lỗi", f"{invalid:,}")

            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("🔄 Làm Mới Trạng Thái"):
                    st.rerun()

            with col_btn2:
                if status != "COMPLETED" and processed < total:
                    if st.button("▶️ Tiếp Tục Chạy (Resume khi gián đoạn)"):
                        if (
                            "cached_dfs" in st.session_state
                            and selected_id in st.session_state["cached_dfs"]
                        ):
                            df_res = st.session_state["cached_dfs"][selected_id]
                            t_resume = threading.Thread(
                                target=process_data_background,
                                args=(selected_id, df_res, processed),
                            )
                            t_resume.daemon = True
                            t_resume.start()
                            st.success(
                                f"Đã tiếp tục tiến trình cho Phiên #{selected_id} từ dòng {processed + 1}!"
                            )
                            st.rerun()
                        else:
                            st.error(
                                "⚠️ Vui lòng chọn lại file CSV tương ứng ở Tab 1 để bộ nhớ tạm ghi nhận dữ liệu resume."
                            )
    else:
        st.info("Chưa ghi nhận phiên nhập liệu nào trong hệ thống.")

# --- TAB 3: LỊCH SỬ & XUẤT BÁO CÁO LỖI ---
with tab3:
    st.subheader("📜 Lịch Sử Phiên Nhập & Tải Báo Cáo Lỗi")

    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    history_df = pd.read_sql_query(
        "SELECT * FROM import_sessions ORDER BY id DESC", conn
    )
    st.dataframe(history_df, use_container_width=True)

    view_id = st.number_input(
        "Nhập ID Phiên muốn xem danh sách dòng lỗi chi tiết:",
        min_value=1,
        step=1,
        value=1,
    )

    if st.button("🔍 Trích Xuất Báo Cáo Lỗi"):
        err_df = pd.read_sql_query(
            f"SELECT row_number AS 'Dòng số', raw_data AS 'Dữ liệu gốc', error_reason AS 'Lý do lỗi' FROM invalid_data WHERE session_id = {view_id}",
            conn,
        )

        if not err_df.empty:
            st.error(
                f"⚠️ Phát hiện tổng cộng **{len(err_df)}** dòng lỗi ở Phiên #{view_id}:"
            )
            st.dataframe(err_df, use_container_width=True)

            # Nút tải file báo cáo lỗi CSV
            csv_error_bytes = err_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Tải Xuống File Báo Cáo Lỗi (CSV)",
                data=csv_error_bytes,
                file_name=f"bao_cao_loi_phien_{view_id}.csv",
                mime="text/csv",
            )
        else:
            st.success(
                f"🎉 Phiên #{view_id} hoàn toàn sạch sẽ, không có dòng nào bị lỗi!"
            )

    conn.close()
