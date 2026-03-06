import streamlit as st
from docx import Document
from docx.shared import RGBColor
import io
import re

st.set_page_config(page_title="HSC 試卷合併工具", page_icon="📝")

st.title("🎓 HSC 模擬考專用：教學檔生成工具")
st.info("已針對 2526 S6 Mock 試卷優化：自動跳過封面、修復子題錯位。")

col1, col2 = st.columns(2)
with col1:
    file_q = st.file_uploader("1. 上傳學生版 (Q)", type=["docx"])
with col2:
    file_ans = st.file_uploader("2. 上傳答案卷 (Ans)", type=["docx"])

if file_q and file_ans:
    if st.button("🪄 執行完美對位合併", type="primary"):
        try:
            doc_q = Document(file_q)
            doc_ans = Document(file_ans)
            
            # --- 1. 抓取答案 ---
            ans_list = []
            for table in doc_ans.tables:
                for row in table.rows:
                    cells = [c.text.strip() for c in row.cells]
                    if len(cells) >= 2:
                        # 排除標題列
                        if any(x in cells[0] for x in ["題號", "Part", "姓名", "班別"]):
                            continue
                        
                        # 邏輯：答案通常在倒數第二格 (排除最後一格的分數)
                        # 如果該格包含文字且不是純數字
                        ans_text = cells[-2]
                        if ans_text and not ans_text.isdigit() and "建議答案" not in ans_text:
                            # 移除答案內重複的 (x分)
                            clean_ans = re.sub(r'[\(（]\s*\d+\s*分\s*[\)）]', '', ans_text).strip()
                            if clean_ans:
                                ans_list.append(clean_ans)

            # --- 2. 處理題目卷表格內的子題 (針對第一題表格) ---
            ans_idx = 0
            for table in doc_q.tables:
                for row in table.rows:
                    for cell in row.cells:
                        # 如果表格格子裡有特定標記或需要填寫的地方
                        # 針對 S6 Mock 第一題表格，我們尋找空格或特定欄位
                        # 此處採
