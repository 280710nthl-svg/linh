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

# 1. Cấu hình giao diện trang web
st.set_page_config(
    page_title="DA11 - Hệ Thống Nhập Dữ Liệu & Quản Lý Phiên",
    page_icon="⚡",
    layout="wide",
)

DB_NAME = "da11_full_integrated.db"


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


# 3. Hàm xuất file Excel lịch sử các phiên (Giống hệt bảng mẫu tiếng Việt)
def export_history_to_excel(df_history):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Lich_Su_Phien"

    # Tiêu đề lớn
    ws.merge_cells("A1:H1")
    title_cell = ws["A1"]
    title_cell.value = "DANH SÁCH LỊCH SỬ CÁC PHIÊN NHẬP DỮ LIỆU - DA11"
    title_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(
        start_color="1E3C72", end_color="1E3C72", fill_type="solid"
    )
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    # Header tiếng Việt
    headers = [
        "Mã phiên",
        "Tên tệp dữ liệu",
        "Tổng số dòng",
        "Số dòng đã xử lý",
        "Số dòng hợp lệ",
        "Số dòng bị lỗi",
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
    for col_num in range(1, 9):
        cell = ws.cell(row=3, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    # Đổ dữ liệu vào các dòng
    for row_idx, row in enumerate(df_history.itertuples(), start=4):
        # Định dạng lại ID thành dạng #1, #2,...
        session_id_str = f"#{row.id}"
        row_data = [
            session_id_str,
            row.filename,
            row.total_rows,
            row.processed_rows,
            row.valid_rows,
            row.invalid_rows,
            row.status,
            str(row.created_at),
        ]
        ws.append(row_data)
        ws.row_dimensions[row_idx].height = 22

        for col_idx in range(1, 9):
            c = ws.cell(row=row_idx, column=col_idx)
            c.border = thin_border
            # Căn giữa cho Mã phiên, Số liệu, Trạng thái, Thời gian
            if col_idx in [1, 3, 4, 5, 6, 7, 8]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    # Độ rộng cột
    col_widths = [12, 32, 16, 18, 16, 16, 16, 22]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

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


# 7. Giao diện Streamlit chính
st.title("⚡ DA11 - Quản Lý Nhập Dữ Liệu & Theo Dõi Tiến Trình")
st.caption(
    "Ứng dụng xử lý tệp lớn, chạy độc lập mượt mà, khôi phục khi gián đoạn và xuất báo cáo tiếng Việt."
)

tab1, tab2, tab3 = st.tabs(
    [
        "📥 Tải File & Cấu Hình",
        "🔄 Quản Lý Phiên & Tiếp Tục (Resume)",
        "📜 Lịch Sử & Báo Cáo",
    ]
)

# --- TAB 1 ---
with tab1:
    st.subheader("1. Tải lên tệp dữ liệu (Hỗ trợ CSV, Excel, JSON)")
    uploaded_file = st.file_uploader(
        "Chọn tệp của bạn", type=["csv", "xlsx", "xls", "json"]
    )

    if uploaded_file is not None:
        df = load_uploaded_file(uploaded_file)
        if df is not None:
            st.success(
                f"📄 Đã tải tệp: **{uploaded_file.name}** | Tổng số dòng: **{len(df):,}** | Số cột: **{len(df.columns)}**"
            )

            with st.expander("🔍 Xem trước 5 dòng đầu tiên (Đã chuẩn hóa tên cột)"):
                st.dataframe(df.head(), use_container_width=True)

            st.subheader("2. Tùy chọn quy tắc kiểm tra lỗi")
            all_cols = df.columns.tolist()
            required_cols = st.multiselect(
                "Chọn các cột BẮT BUỘC phải có dữ liệu (Không được để trống):",
                options=all_cols,
                default=all_cols[: min(2, len(all_cols))],
            )

            if st.button("🚀 Bắt Đầu Xử Lý Dữ Liệu"):
                conn = sqlite3.connect(DB_NAME, check_same_thread=False)
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT INTO import_sessions (filename, total_rows, status) VALUES (?, ?, 'ĐANG CHẠY')",
                    (uploaded_file.name, len(df)),
                )
                session_id = cursor.lastrowid
                conn.commit()
                conn.close()

                if "df_cache" not in st.session_state:
                    st.session_state["df_cache"] = {}
                st.session_state["df_cache"][session_id] = (df, required_cols)

                thread = threading.Thread(
                    target=process_data_background,
                    args=(session_id, df, required_cols, 0),
                )
                thread.daemon = True
                thread.start()

                st.success(
                    f"🎉 Đã khởi tạo và bắt đầu xử lý thành công Phiên nhập **#{session_id}**!"
                )
                st.info(
                    "👉 Hãy chuyển sang Tab **'Quản Lý Phiên & Tiếp Tục (Resume)'** để theo dõi tiến trình thời gian thực."
                )

# --- TAB 2 ---
with tab2:
    st.subheader(
        "🔄 Theo dõi tiến độ thời gian thực & Tiếp tục khi bị gián đoạn"
    )
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    sessions_df = pd.read_sql_query(
        "SELECT * FROM import_sessions ORDER BY id DESC", conn
    )
    conn.close()

    if not sessions_df.empty:
        # Hiển thị bảng phiên dạng tiếng Việt trực quan
        display_df = sessions_df.copy()
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
        ]
        st.dataframe(display_df, use_container_width=True)

        selected_id = st.selectbox(
            "Chọn ID Phiên cần theo dõi hoặc thao tác:",
            sessions_df["id"].tolist(),
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

            st.progress(pct)

            m1, m2, m3, m4 = st.columns(4)
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

            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("🔄 Làm Mới Trạng Thái"):
                    st.rerun()
            with col_btn2:
                if status != "HOÀN THÀNH" and processed < total:
                    if st.button("▶️ Tiếp Tục Xử Lý (Resume từ điểm dừng)"):
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
                                f"Đã khôi phục và tiếp tục chạy Phiên #{selected_id} từ dòng số {processed + 1}!"
                            )
                            st.rerun()
                        else:
                            st.error(
                                "⚠️ Vui lòng tải lại tệp ở Tab 1 để bộ nhớ tạm ghi nhận dữ liệu phục hồi."
                            )
    else:
        st.info("Chưa có phiên nhập dữ liệu nào trong hệ thống.")

# --- TAB 3: LỊCH SỬ & XUẤT EXCEL BÁO CÁO PHIÊN ---
with tab3:
    st.subheader("📜 Danh Sách Lịch Sử Các Phiên Nhập Dữ Liệu")
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    history_df = pd.read_sql_query(
        "SELECT * FROM import_sessions ORDER BY id DESC", conn
    )

    if not history_df.empty:
        # Đổi tên cột hiển thị tiếng Việt trên Streamlit
        view_history = history_df.copy()
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
        ]
        st.dataframe(view_history, use_container_width=True)

        st.markdown("---")
        st.subheader("📊 Xuất Danh Sách Phiên Ra File Excel")
        st.write(
            "Bấm vào nút dưới đây để tải xuống bảng lịch sử các phiên nhập dưới định dạng file Excel chuyên nghiệp:"
        )

        excel_buffer = export_history_to_excel(history_df)
        st.download_button(
            label="📥 Tải Xuống Báo Cáo Lịch Sử Phiên (File Excel .xlsx)",
            data=excel_buffer,
            file_name="Danh_Sach_Lich_Su_Phien_DA11.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.info("Chưa có dữ liệu lịch sử phiên nhập nào.")

    st.markdown("---")
    view_id = st.number_input(
        "Nhập ID Phiên muốn xem danh sách dòng lỗi chi tiết:",
        min_value=1,
        step=1,
        value=1,
    )

    if st.button("🔍 Trích Xuất Danh Sách Dòng Lỗi"):
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
                label="📥 Tải Xuống File Báo Cáo Lỗi (CSV)",
                data=csv_bytes,
                file_name=f"bao_cao_loi_phien_{view_id}.csv",
                mime="text/csv",
            )
        else:
            st.success(
                f"🎉 Phiên #{view_id} hoàn toàn chính xác, không có dòng lỗi nào!"
            )
    conn.close()
