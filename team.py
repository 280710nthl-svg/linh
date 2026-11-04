import sqlite3
import time
import pandas as pd
import streamlit as st

# 1. Cấu hình trang Dashboard
st.set_page_config(
    page_title="DA11 - Quản Lý Dữ Liệu Thông Minh", page_icon="⚡", layout="wide"
)

# Sử dụng các component native của Streamlit để tạo giao diện đẹp mà không cần HTML thuần
DB_NAME = "da11_ui_surcharge.db"


def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS import_sessions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT,
            total_rows INTEGER,
            valid_rows INTEGER,
            invalid_rows INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS invalid_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            session_id INTEGER,
            row_number INTEGER,
            error_reason TEXT
        )
    """)
    conn.commit()
    conn.close()


init_db()

# 2. Tiêu đề Dashboard phong cách chuyên nghiệp
st.title("📊 Hệ Thống Nhập & Kiểm Tra Dữ Liệu Tự Động")
st.markdown(
    "Ứng dụng phân tích file linh hoạt, tách dòng lỗi thông minh, hiển thị tiến độ và biểu đồ trực quan."
)

# Tạo các tab phân vùng giao diện rõ ràng
tab1, tab2, tab3 = st.tabs(
    ["📥 Nhập File & Kiểm Tra", "📈 Biểu Đồ Thống Kê", "📜 Lịch Sử & Báo Cáo"]
)

# --- TAB 1: NHẬP FILE & XỬ LÝ ---
with tab1:
    st.subheader("Tải lên tệp dữ liệu của bạn")
    uploaded_file = st.file_uploader(
        "Hỗ trợ file CSV, Excel (.xlsx)", type=["csv", "xlsx", "xls"]
    )

    if uploaded_file is not None:
        # Đọc file tự động theo định dạng
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        st.success(
            f"Đã tải thành công file: **{uploaded_file.name}** ({len(df):,} dòng dữ liệu)"
        )

        with st.expander("🔍 Xem trước dữ liệu đầu vào"):
            st.dataframe(df.head(), use_container_width=True)

        # Chọn cột bắt buộc kiểm tra
        all_cols = df.columns.tolist()
        target_cols = st.multiselect(
            "Chọn các cột yêu cầu KHÔNG ĐƯỢC ĐỂ TRỐNG:",
            options=all_cols,
            default=all_cols[: min(2, len(all_cols))],
        )

        if st.button("🚀 Bắt Đầu Xử Lý Dữ Liệu"):
            progress_bar = st.progress(0)
            status_text = st.empty()

            valid_count = 0
            invalid_records = []
            total = len(df)

            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO import_sessions (filename, total_rows, valid_rows, invalid_rows) VALUES (?, ?, 0, 0)",
                (uploaded_file.name, total),
            )
            session_id = cursor.lastrowid
            conn.commit()

            # Vòng lặp xử lý từng dòng
            for idx, row in df.iterrows():
                row_num = idx + 1
                errors = []
                for col in target_cols:
                    val = row.get(col)
                    if pd.isna(val) or str(val).strip() == "":
                        errors.append(f"Cột '{col}' bị trống")

                if not errors:
                    valid_count += 1
                else:
                    err_msg = "; ".join(errors)
                    invalid_records.append((session_id, row_num, err_msg))

                # Cập nhật thanh tiến độ đồ họa mượt mà
                if row_num % max(1, total // 100) == 0 or row_num == total:
                    pct = row_num / total
                    progress_bar.progress(pct)
                    status_text.text(
                        f"Đang xử lý: {row_num}/{total} dòng ({(pct*100):.1f}%)"
                    )

            # Lưu vào Database
            if invalid_records:
                cursor.executemany(
                    "INSERT INTO invalid_data (session_id, row_number, error_reason) VALUES (?, ?, ?)",
                    invalid_records,
                )

            invalid_count = len(invalid_records)
            cursor.execute(
                "UPDATE import_sessions SET valid_rows = ?, invalid_rows = ? WHERE id = ?",
                (valid_count, invalid_count, session_id),
            )
            conn.commit()
            conn.close()

            st.success("🎉 Hoàn tất quá trình xử lý và kiểm tra dữ liệu!")

            # Hiển thị khối Thẻ Thống Kê (Metrics Cards) đẹp mắt bằng Python
            m1, m2, m3 = st.columns(3)
            m1.metric(
                label="📁 Tổng số dòng",
                value=f"{total:,}",
                delta="100% dữ liệu",
            )
            m2.metric(
                label="✅ Dòng hợp lệ",
                value=f"{valid_count:,}",
                delta=f"{(valid_count/total*100):.1f}%",
            )
            m3.metric(
                label="⚠️ Dòng lỗi",
                value=f"{invalid_count:,}",
                delta=f"-{(invalid_count/total*100):.1f}%",
                delta_value_color="inverse",
            )

# --- TAB 2: ĐỒ HỌA TRỰC QUAN ---
with tab2:
    st.subheader("📈 Biểu đồ trực quan hóa kết quả kiểm tra")
    conn = sqlite3.connect(DB_NAME)
    history_df = pd.read_sql_query("SELECT * FROM import_sessions", conn)
    conn.close()

    if not history_df.empty:
        # Sử dụng biểu đồ tích hợp sẵn cực kỳ sinh động của Streamlit (không cần HTML)
        st.write("### Tỷ lệ phân bổ giữa dòng Hợp lệ và Dòng lỗi qua các phiên")
        chart_data = history_df[["id", "valid_rows", "invalid_rows"]].set_index(
            "id"
        )
        st.bar_chart(chart_data)
    else:
        st.info(
            "Chưa có dữ liệu thống kê. Vui lòng nhập file ở Tab 1 để xem biểu đồ."
        )

# --- TAB 3: LỊCH SỬ & BÁO CÁO LỖI ---
with tab3:
    st.subheader("📜 Lịch sử các phiên nhập liệu & Xuất báo cáo lỗi")
    conn = sqlite3.connect(DB_NAME)
    sessions = pd.read_sql_query(
        "SELECT * FROM import_sessions ORDER BY id DESC", conn
    )
    st.dataframe(sessions, use_container_width=True)

    sid = st.number_input(
        "Nhập ID phiên để xem chi tiết lỗi:", min_value=1, step=1, value=1
    )
    if st.button("🔍 Xem Chi Tiết Lỗi"):
        err_df = pd.read_sql_query(
            f"SELECT row_number AS 'Dòng số', error_reason AS 'Lý do lỗi' FROM invalid_data WHERE session_id = {sid}",
            conn,
        )
        if not err_df.empty:
            st.error(f"Phát hiện {len(err_df)} dòng lỗi tại phiên #{sid}:")
            st.dataframe(err_df, use_container_width=True)

            # Nút tải file báo cáo lỗi CSV
            st.download_button(
                label="📥 Tải Báo Cáo Lỗi (CSV)",
                data=err_df.to_csv(index=False).encode("utf-8"),
                file_name=f"bao_cao_loi_phien_{sid}.csv",
                mime="text/csv",
            )
        else:
            st.success("Phiên này hoàn toàn sạch, không có lỗi!")
    conn.close()
