import streamlit as st
from docx import Document
from docx.shared import RGBColor
import io
import re

st.set_page_config(page_title="HSC 試卷合併工具", page_icon="📝")

st.title("🎓 HSC 模擬考教學檔生成器")
st.markdown("### 針對 2526 S6 Mock 格式優化")

col1, col2 = st.columns(2)
with col1:
    file_q = st.file_uploader("1. 上傳學生版 (Q)", type=["docx"])
with col2:
    file_ans = st.file_uploader("2. 上傳答案卷 (Ans)", type=["docx"])

if file_q and file_ans:
    if st.button("🪄 執行自動合併", type="primary"):
        try:
            doc_q = Document(file_q)
            doc_ans = Document(file_ans)
            
            # 1. 抓取答案：鎖定表格中間的建議答案欄
            ans_list = []
            for table in doc_ans.tables:
                for row in table.rows:
                    cells = [c.text.strip() for c in row.cells]
                    if len(cells) >= 2:
                        # 排除標題與個人資料列
                        if any(x in cells[0] for x in ["題號", "Part", "姓名", "班別", "成績"]):
                            continue
                        
                        # 答案通常在第 2 格 (索引 1)
                        content = cells[1]
                        if content and "建議答案" not in content and not content.isdigit():
                            # 清理掉答案中重複出現的分數標記
                            clean_text = re.sub(r'[\(（]\s*\d+\s*分\s*[\)）]', '', content).strip()
                            if clean_text:
                                ans_list.append(clean_text)

            # 2. 填入題目：尋找 (x 分) 標記
            ans
