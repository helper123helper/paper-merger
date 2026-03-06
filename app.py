import streamlit as st
from docx import Document
from docx.shared import RGBColor
import io
import re

st.set_page_config(page_title="試卷答案自動合併工具", page_icon="📝")

st.title("🎓 試卷題目與答案自動合併工具")
st.markdown("### 2526 HSC 模擬考專用優化版")

col1, col2 = st.columns(2)
with col1:
    file_q = st.file_uploader("1. 上傳學生版 (Q)", type=["docx"])
with col2:
    file_ans = st.file_uploader("2. 上傳答案卷 (Ans)", type=["docx"])

if file_q and file_ans:
    if st.button("🪄 執行自動合併並修復走位", type="primary"):
        try:
            doc_q = Document(file_q)
            doc_ans = Document(file_ans)
            
            # 1. 精準抓取答案 (只抓取表格中「建議答案」那一欄)
            ans_data = []
            for table in doc_ans.tables:
                for row in table.rows:
                    cells = [c.text.strip() for c in row.cells]
                    if len(cells) >= 2:
                        # 排除標題列
                        if any(x in cells[0] for x in ["題號", "Part", "Part A"]): continue
                        
                        # 邏輯：如果第2格有內容，且不只是數字(分數)，就視為答案
                        ans_text = cells[1]
                        if ans_text and not ans_text.isdigit() and "建議答案" not in ans_text:
                            # 移除答案
