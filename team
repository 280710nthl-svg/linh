import sqlite3
import time
import pandas as pd
import streamlit as st

# Cấu hình giao diện Streamlit
st.set_page_config(
    page_title="DA11 - Nhập Dữ Liệu & Theo Dõi Tiến Trình", layout="wide"
)

DB_NAME = "data_import_da11.db"


# Khởi tạo cơ sở dữ liệu SQLite
def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    # Bảng lưu lịch sử phiên nhập
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
    conn.commit()
    conn.close()


init_db()


# Hàm kiểm tra tính hợp lệ từng dòng dữ liệu
def validate_row(row):
    errors = []

    # Kiểm tra tên
    if pd.isna(row.get("name")) or str(row.get("name")).strip() == "":
        errors.append("Tên không được để trống")

    # Kiểm tra email
    email = str(row.get("email"))
    if pd.isna(row.get("email")) or "@" not in email:
        errors.append("Email không hợp lệ (thiếu @)")

    # Kiểm tra tuổi
    try:
        age = int(row.get("age"))
        if age <= 0 or age > 120:
            errors.append("Tuổi phải từ 1 đến 120")
    except (ValueError, TypeError):
        errors.append("Tuổi phải là định dạng số nguyên")

    return errors


# Giao diện chính
st.title("📊 DA11 - Nhập Dữ Liệu & Theo Dõi Tiến Trình")
st.markdown(
    "Ứng dụng hỗ trợ nhập file lớn, kiểm tra cấu trúc, hiển thị tiến độ thời gian thực, tách dòng lỗi và lưu lịch sử phiên."
)

tab1, tab2 = st.tabs(["📥 Nhập Dữ Liệu & Tiến Trình", "📜 Lịch Sử Phiên Nhập"])

with tab1:
    st.subheader("Tải lên file CSV dữ liệu")
    st.info(
        "💡 **Yêu cầu cấu trúc file CSV:** Phải chứa các cột tối thiểu: `name`, `email`, `age`."
    )

    uploaded_file = st.file_uploader("Chọn file CSV của bạn", type=["csv"])

    if uploaded_file is not None:
        try:
            # Đọc file CSV bằng pandas
            df = pd.read_csv(uploaded_file)
            st.success(
                f"Đã tải thành công file: **{uploaded_file.name}** (Tổng số dòng: {len(df)})"
            )

            # Kiểm tra cấu trúc cột
            required_columns = {"name", "email", "age"}
            if not required_columns.issubset(set(df.columns)):
                st.error(
                    f"❌ **Sai cấu trúc file!** File thiếu các cột bắt buộc: {required_columns - set(df.columns)}"
                )
            else:
                # Hiển thị trước một phần dữ liệu
                with st.expander("🔍 Xem trước 5 dòng đầu của file"):
                    st.dataframe(df.head(), use_container_width=True)

                if st.button("🚀 Bắt đầu quá trình nhập dữ liệu"):
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    valid_records = []
                    invalid_records = []
                    total_rows = len(df)
                    start_time = time.time()

                    # Xử lý từng dòng, đảm bảo lỗi 1 dòng không làm mất dữ liệu dòng khác
                    for index, row in df.iterrows():
                        row_num = index + 1
                        errors = validate_row(row)

                        if not errors:
                            valid_records.append(
                                (
                                    str(row["name"]),
                                    str(row["email"]),
                                    int(row["age"]),
                                )
                            )
                        else:
                            invalid_records.append(
                                {
                                    "Dòng số": row_num,
                                    "Dữ liệu gốc": str(row.to_dict()),
                                    "Lý do lỗi": "; ".join(errors),
                                }
                            )

                        # Cập nhật thanh tiến độ trực quan
                        if row_num % max(1, total_rows // 100) == 0 or row_num == total_rows:
                            progress_percent = row_num / total_rows
                            progress_bar.progress(progress_percent)
                            status_text.text(
                                f"Đang xử lý: {row_num}/{total_rows} dòng ({(progress_percent * 100):.1f}%)"
                            )

                    # Lưu kết quả vào cơ sở dữ liệu SQLite
                    conn = sqlite3.connect(DB_NAME)
                    cursor = conn.cursor()

                    # Lưu thông tin phiên nhập
                    cursor.execute(
                        """
                        INSERT INTO import_sessions (filename, total_rows, valid_rows, invalid_rows)
                        VALUES (?, ?, ?, ?)
                    """,
                        (
                            uploaded_file.name,
                            total_rows,
                            len(valid_records),
                            len(invalid_records),
                        ),
                    )
                    session_id = cursor.lastrowid

                    # Lưu các bản ghi hợp lệ hàng loạt
                    valid_data_with_session = [
                        (session_id, r[0], r[1], r[2]) for r in valid_records
                    ]
                    cursor.executemany(
                        """
                        INSERT INTO valid_data (session_id, name, email, age)
                        VALUES (?, ?, ?, ?)
                    """,
                        valid_data_with_session,
                    )

                    conn.commit()
                    conn.close()

                    elapsed_time = time.time() - start_time

                    # Thông báo hoàn tất và báo cáo kết quả
                    st.success(
                        f"🎉 Hoàn thành nhập dữ liệu trong vòng {elapsed_time:.2f} giây!"
                    )

                    # Thống kê nhanh kết quả
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Tổng số dòng xử lý", total_rows)
                    col2.metric("Dòng hợp lệ (Đã lưu)", len(valid_records))
                    col3.metric("Dòng lỗi (Đã tách riêng)", len(invalid_records))

                    # Hiển thị và cung cấp tải xuống cho các dòng lỗi
                    if invalid_records:
                        st.warning(
                            "⚠️ Phát hiện một số dòng lỗi cấu trúc hoặc dữ liệu không hợp lệ:"
                        )
                        df_errors = pd.DataFrame(invalid_records)
                        st.dataframe(df_errors, use_container_width=True)

                        # Nút tải file báo cáo lỗi CSV
                        csv_error_data = df_errors.to_csv(
                            index=False
                        ).encode("utf-8")
                        st.download_button(
                            label="📥 Tải báo cáo danh sách dòng lỗi (CSV)",
                            data=csv_error_data,
                            file_name=f"bao_cao_loi_{uploaded_file.name}",
                            mime="text/csv",
                        )

        except Exception as e:
            st.error(f"❌ Đã xảy ra lỗi hệ thống khi xử lý file: {str(e)}")

with tab2:
    st.subheader("📜 Quản lý lịch sử các phiên nhập dữ liệu")
    conn = sqlite3.connect(DB_NAME)
    sessions_df = pd.read_sql_query(
        "SELECT * FROM import_sessions ORDER BY created_at DESC", conn
    )

    if not sessions_df.empty:
        st.dataframe(sessions_df, use_container_width=True)

        # Chọn phiên để xem chi tiết dữ liệu hợp lệ đã được lưu
        session_ids = sessions_df["id"].tolist()
        selected_session = st.selectbox(
            "🔍 Chọn ID phiên nhập để xem dữ liệu hợp lệ chi tiết:", session_ids
        )

        if selected_session:
            detail_df = pd.read_sql_query(
                f"SELECT name, email, age FROM valid_data WHERE session_id = {selected_session}",
                conn,
            )
            st.write(f"Dữ liệu hợp lệ của Phiên ID #{selected_session}:")
            st.dataframe(detail_df, use_container_width=True)
    else:
        st.info("Chưa có lịch sử phiên nhập nào được ghi nhận.")

    conn.close()
