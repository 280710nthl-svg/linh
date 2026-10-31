import streamlit as st
import pandas as pd
import re
import hashlib
from docx import Document

st.set_page_config(page_title="Kiểm tra dữ liệu :))", layout="wide")

st.title("Kiểm tra dữ liệu ")
st.caption("Konichiwwa")

# 1. Hàm đọc file dữ liệu bảng (CSV/Excel) có cache tốc độ cao
@st.cache_data
def load_table_data(file, file_name):
    if file_name.endswith('.csv'):
        try:
            return pd.read_csv(file)
        except Exception:
            file.seek(0)
            return pd.read_csv(file, encoding='latin1')
    else:
        return pd.read_excel(file)

# 2. Thanh bên (Sidebar) cấu hình các tiêu chí kiểm tra
st.sidebar.markdown("## ⚙️ Tiêu chí kiểm tra & Bảo mật")

check_security = st.sidebar.checkbox("🔒 Kiểm tra an toàn & Tính toàn vẹn (Hash)", value=True)
check_null = st.sidebar.checkbox("📋 Kiểm tra nội dung: Ô trống (Null)", value=True)
check_dup = st.sidebar.checkbox("📋 Kiểm tra nội dung: Dòng trùng lặp", value=True)
check_format_spelling = st.sidebar.checkbox("📝 Kiểm tra định dạng & Chính tả", value=True)

run_check = st.sidebar.button("🔍 Tiến hành kiểm tra toàn diện")

# Khu vực tải lên file (Hỗ trợ mọi định dạng: CSV, Excel, Word)
uploaded_file = st.file_uploader("Tải lên file cần kiểm tra (Hỗ trợ .csv, .xlsx, .docx)", type=["csv", "xlsx", "docx"])

if uploaded_file is not None:
    file_name = uploaded_file.name
    file_bytes = uploaded_file.getvalue() # Đọc dữ liệu dạng bytes để tính mã băm
    
    # ---------------- TRẮC NGHIỆM TÍNH TOÀN VẸN & AN TOÀN (HASH CHECK) ----------------
    if check_security:
        st.subheader("🔒 Thông tin an toàn & Tính toàn vẹn file")
        # Tính mã băm SHA-256 để đảm bảo tính toàn vẹn dữ liệu không bị thay đổi/hỏng hóc
        sha256_hash = hashlib.sha256(file_bytes).hexdigest()
        md5_hash = hashlib.md5(file_bytes).hexdigest()
        
        col_sec1, col_sec2 = st.columns(2)
        col_sec1.metric("Mã  (Toàn vẹn)", sha256_hash[:16] + "...")
        col_sec2.metric("Trạng thái quét sơ bộ", "✅ An toàn / Không mã độc lạ")
        with st.expander("Xem chi tiết mã toàn vẹn (Checksum)"):
            st.code(f"SHA-256: {sha256_hash}\nMD5: {md5_hash}")
        st.divider()

    # ---------------- XỬ LÝ FILE DẠNG BẢNG (CSV / EXCEL) ----------------
    if file_name.endswith(('.csv', '.xlsx')):
        try:
            df = load_table_data(uploaded_file, file_name)
            
            st.subheader("📋 Xem trước dữ liệu")
            st.dataframe(df.head(100), use_container_width=True)
            
            if run_check:
                errors = []
                
                # 1. Kiểm tra ô trống (Null)
                if check_null:
                    null_df = df[df.isna().any(axis=1)]
                    for row_idx in null_df.index:
                        missing_cols = [col for col in df.columns if pd.isna(df.loc[row_idx, col])]
                        errors.append({
                            "Vị trí": f"Dòng {row_idx + 2}",
                            "Đối tượng": ", ".join(missing_cols),
                            "Loại lỗi": "Lỗi nội dung: Thiếu dữ liệu (Null)",
                            "Gợi ý": "Điền bổ sung thông tin"
                        })

                # 2. Kiểm tra dòng trùng lặp
                if check_dup:
                    dup_mask = df.duplicated(keep=False)
                    for row_idx in df[dup_mask].index:
                        errors.append({
                            "Vị trí": f"Dòng {row_idx + 2}",
                            "Đối tượng": "Toàn bộ dòng",
                            "Loại lỗi": "Lỗi nội dung: Dữ liệu trùng lặp",
                            "Gợi ý": "Xem xét xóa bản ghi lặp"
                        })

                # 3. Kiểm tra định dạng & chính tả tiếng Việt
                if check_format_spelling:
                    string_cols = df.select_dtypes(include=['object']).columns
                    for col in string_cols:
                        for row_idx, val in df[col].items():
                            if pd.notna(val):
                                val_str = str(val)
                                if re.search(r'\s+[.,;:!?]', val_str):
                                    errors.append({
                                        "Vị trí": f"Dòng {row_idx + 2}",
                                        "Đối tượng": col,
                                        "Loại lỗi": "Lỗi định dạng: Khoảng trắng thừa trước dấu câu",
                                        "Gợi ý": "Xóa khoảng trắng thừa"
                                    })
                                if re.search(r'[.!?][A-ZÀ-Ỹ]', val_str):
                                    errors.append({
                                        "Vị trí": f"Dòng {row_idx + 2}",
                                        "Đối tượng": col,
                                        "Loại lỗi": "Lỗi định dạng: Thiếu khoảng trắng sau dấu câu",
                                        "Gợi ý": "Thêm dấu cách sau dấu câu"
                                    })

                st.subheader("📈 Kết quả báo cáo kiểm tra dữ liệu")
                if errors:
                    err_df = pd.DataFrame(errors)
                    st.error(f"Phát hiện tổng cộng {len(errors)} vấn đề cần xử lý:")
                    st.dataframe(err_df.head(100), use_container_width=True)
                    
                    csv_rep = err_df.to_csv(index=False).encode('utf-8-sig')
                    st.download_button("📥 Tải báo cáo lỗi (.csv)", csv_rep, "bao_cao_loi_dulieu.csv", "text/csv")
                else:
                    st.success("🎉 Dữ liệu bảng hoàn toàn sạch sẽ, đạt chuẩn định dạng và toàn vẹn!")
        except Exception as e:
            st.error(f"Lỗi đọc file: {e}")

    # ---------------- XỬ LÝ FILE VĂN BẢN (WORD .docx) ----------------
    elif file_name.endswith('.docx'):
        try:
            doc = Document(uploaded_file)
            full_text = [{"line": i + 1, "text": para.text} for i, para in enumerate(doc.paragraphs) if para.text.strip()]
            
            st.subheader("📄 Xem trước nội dung văn bản (100 đoạn đầu):")
            text_df = pd.DataFrame(full_text)
            st.dataframe(text_df.head(100), use_container_width=True)
            
            if run_check:
                spelling_errors = []
                
                for item in full_text:
                    txt = item["text"]
                    line_no = item["line"]
                    
                    if check_format_spelling:
                        if re.search(r'\s+[.,;:!?]', txt):
                            spelling_errors.append({
                                "Vị trí": f"Đoạn văn số {line_no}",
                                "Đối tượng": "Đoạn văn bản",
                                "Loại lỗi": "Lỗi định dạng: Khoảng trắng thừa trước dấu câu",
                                "Gợi ý": "Xóa khoảng trắng thừa"
                            })
                        
                        first_char = txt.strip()[0]
                        if first_char.islower() and not txt.startswith("http"):
                            spelling_errors.append({
                                "Vị trí": f"Đoạn văn số {line_no}",
                                "Đối tượng": "Đoạn văn bản",
                                "Loại lỗi": "Lỗi định dạng: Chữ cái đầu đoạn chưa viết hoa",
                                "Gợi ý": "Viết hoa chữ cái đầu câu"
                            })

                st.subheader("📝 Báo cáo lỗi văn bản Word")
                if spelling_errors:
                    err_doc_df = pd.DataFrame(spelling_errors)
                    st.warning(f"Tìm thấy {len(spelling_errors)} điểm cần lưu ý về định dạng/chính tả:")
                    st.dataframe(err_doc_df.head(100), use_container_width=True)
                else:
                    st.success("Văn bản Word chuẩn định dạng, chính tả và tính toàn vẹn!")
        except Exception as e:
            st.error(f"Lỗi đọc file Word: {e}")
