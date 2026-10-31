import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Kiểm tra dữ liệu", layout="wide")

st.title("Check file ")
st.caption("Trum CBN")

# 1. Upload File
uploaded_file = st.file_uploader("Tải lên file dữ liệu (.csv hoặc .xlsx)", type=["csv", "xlsx"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        st.subheader("📋 Xem trước dữ liệu gốc")
        st.dataframe(df.head(10))
        
        st.sidebar.header("⚙️ Cấu hình quy tắc kiểm tra")
        
        # Chọn quy tắc
        check_null = st.sidebar.checkbox("Kiểm tra ô trống (Null)", value=True)
        check_dup = st.sidebar.checkbox("Kiểm tra dòng trùng lặp (Duplicates)", value=True)
        check_email = st.sidebar.checkbox("Kiểm tra định dạng Email", value=False)

        # Chọn cột nếu kiểm tra email
        email_col = None
        if check_email:
            email_col = st.sidebar.selectbox("Chọn cột chứa Email", options=df.columns)

        if st.sidebar.button("🔍 Tiến hành kiểm tra"):
            errors = []

            # Quy tắc 1: Kiểm tra ô trống
            if check_null:
                for row_idx, row in df.iterrows():
                    for col_name in df.columns:
                        if pd.isna(row[col_name]):
                            errors.append({
                                "Dòng (Row)": row_idx + 2, # +2 tương ứng dòng thực tế trong Excel/CSV
                                "Cột (Column)": col_name,
                                "Giá trị (Value)": "Rỗng (Null)",
                                "Loại lỗi": "Thiếu dữ liệu",
                                "Gợi ý khắc phục": "Điền bổ sung thông tin thiếu"
                            })

            # Quy tắc 2: Kiểm tra trùng lặp
            if check_dup:
                dup_mask = df.duplicated(keep=False)
                for row_idx in df[dup_mask].index:
                    errors.append({
                        "Dòng (Row)": row_idx + 2,
                        "Cột (Column)": "Tất cả các cột",
                        "Giá trị (Value)": "Hàng trùng lặp",
                        "Loại lỗi": "Dữ liệu bị trùng",
                        "Gợi ý khắc phục": "Xóa bớt bản ghi bị lặp"
                    })

            # Quy tắc 3: Kiểm tra định dạng Email bằng Regex
            if check_email and email_col:
                email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
                for row_idx, value in df[email_col].items():
                    if pd.notna(value) and not re.match(email_regex, str(value)):
                        errors.append({
                            "Dòng (Row)": row_idx + 2,
                            "Cột (Column)": email_col,
                            "Giá trị (Value)": str(value),
                            "Loại lỗi": "Sai định dạng Email",
                            "Gợi ý khắc phục": "Sửa lại đúng cấu trúc email (VD: abc@gmail.com)"
                        })

            # Hiển thị kết quả
            st.divider()
            st.subheader("📈 Kết quả kiểm tra")
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Tổng số dòng", len(df))
            col2.metric("Số lỗi phát hiện", len(errors))
            col3.metric("Trạng thái", "⚠️ Có lỗi" if errors else "✅ Hợp lệ")

            if errors:
                err_df = pd.DataFrame(errors)
                st.error(f"Phát hiện tổng cộng {len(errors)} vấn đề cần xử lý:")
                st.dataframe(err_df, use_container_width=True)

                # Nút tải file báo cáo lỗi
                csv_data = err_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 Tải báo cáo lỗi (.csv)",
                    data=csv_data,
                    file_name="bao_cao_loi_dulieu.csv",
                    mime="text/csv"
                )
            else:
                st.success("Dữ liệu đạt tiêu chuẩn, không tìm thấy lỗi nào!")

    except Exception as e:
        st.error(f"Lỗi khi đọc file: {e}")
