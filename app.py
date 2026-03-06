import streamlit as st
from docx import Document
from docx.shared import RGBColor
import io
import re

st.set_page_config(page_title="試卷答案自動合併工具", page_icon="📝")

st.title("🎓 試卷題目與答案自動合併工具")
st.markdown("### 2526 HSC 模擬考專用 (修復走位版)")

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
            
            # 1. 抓取答案 (精準解析表格結構)
            ans_list = []
            for table in doc_ans.tables:
                for row in table.rows:
                    # 取得每一列的所有文字並清理
                    cells = [c.text.strip() for c in row.cells]
                    # 針對 HSC 答案格式：通常 [0] 是題號, [1] 是建議答案, [2] 是分數
                    if len(cells) >= 2:
                        ans_text = cells[1]
                        # 排除標題列、空行、或只有分數的行
                        if ans_text and not any(x in cells[0] for x in ["題號", "Part", "姓名", "班別"]):
                            if "建議答案" not in ans_text and not ans_text.isdigit():
                                # 清除答案中可能重複出現的 (x分)
                                clean_ans = re.sub(r'[\(（]\s*\d+\s*分\s*[\)）]', '', ans_text).strip()
                                if clean_ans:
                                    ans_list.append(clean_ans)

            # 2. 定位題目並填入
            ans_count = 0
            # 定義需要跳過關鍵字（封面資訊）
            skip_keywords = ["姓名", "班別", "學號", "成績", "年度", "模擬考試"]
