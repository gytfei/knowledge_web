import streamlit as st
from pathlib import Path
import mammoth

st.set_page_config(layout="wide")
st.markdown("""
<style>
h1 {
    font-size: 28px !important;
}

</style>
""", unsafe_allow_html=True)
st.markdown("""
<style>
/* 减小页面顶部留白 */
.block-container {
    padding-top: 1rem !important;
}

/* 让标题上边距更小 */
h1, h2, h3 {
    margin-top: 0rem !important;
}
</style>
""", unsafe_allow_html=True)
st.markdown("""
<style>

/* 隐藏默认的多页面导航标题 */
section[data-testid="stSidebar"] div[data-testid="stSidebarNav"] > ul {
    margin-top: 10px;
}

/* 改 sidebar 页面文字样式 */
section[data-testid="stSidebar"] div[data-testid="stSidebarNav"] span {
    font-size: 20px !important;
    font-weight: 700 !important;
}

</style>
""", unsafe_allow_html=True)
st.title("📄 文档预览")



doc_path = st.session_state.get("preview_doc_path")

if not doc_path or not Path(doc_path).exists():
    st.error("未找到文档")
else:
    with open(doc_path, "rb") as f:
        result = mammoth.convert_to_html(f)
        html = result.value

    st.components.v1.html(
        f"""
        <div style="
            background-color:white;
            padding:0px;
            max-width:900px;
            margin:auto;
            font-family:Arial;
            line-height:1.8;
        ">
        {html}
        </div>
        """,
        height=1000,
        scrolling=True
    )
