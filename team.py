# Custom CSS giao diện nền tối (Dark Mode)
st.markdown(
    """
    <style>
    /* Nền đen toàn trang */
    .stApp {
        background: #0e1117 !important;
        background-attachment: fixed;
    }

    /* Khung nội dung */
    div.block-container {
        background-color: transparent !important;
        padding: 2.5rem 3rem;
        border-radius: 16px;
        box-shadow: none;
        margin-top: 2rem;
        margin-bottom: 2rem;
    }

    /* Màu chữ */
    .stApp {
        color: #f1f3f5;
    }

    p, span, label {
        color: #f1f3f5;
    }

    /* Tiêu đề */
    h1, h2, h3, h4 {
        color: #ffffff !important;
    }

    /* Thanh Tab */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
        background-color: #161b22;
        padding: 8px 12px;
        border-radius: 10px;
    }

    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 10px 24px;
        font-weight: 600;
        color: #d1d5db;
        border: none;
        transition: all 0.3s ease;
    }

    .stTabs [aria-selected="true"] {
        background-color: #243b70 !important;
        color: #ffffff !important;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.35);
    }

    .stTabs [aria-selected="true"] p {
        color: #ffffff !important;
    }

    /* Thẻ Metric */
    div[data-testid="stMetric"] {
        background-color: #161b22;
        padding: 16px 20px;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.25);
        border-left: 5px solid #5277c7;
    }

    /* Nút bấm */
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        padding: 10px 24px;
        transition: all 0.3s ease;
    }

    /* Bảng dữ liệu */
    [data-testid="stDataFrame"] {
        background-color: #161b22;
    }

    </style>
    """,
    unsafe_allow_html=True,
)
