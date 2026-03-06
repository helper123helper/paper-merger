import streamlit as st
from docx import Document
from docx.shared import RGBColor
import io
import re

st.set_page_config(page_title="HSC 試卷完美對位工具", page_icon="🎓")

st.title("🎓 HSC 模擬考：教學檔生成工具 (完美版)")
st.info("已修復：表格內(a)(b)錯位問題，直接將表格答案內容 copy 到題目下方或表格內。")

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
            
            # --- 1. 抓取答案庫 (過濾題號與分數，只留核心文字) ---
            ans_list = []
            for table in doc_ans.tables:
                for row in table.rows:
                    cells = [c.text.strip() for c in row.cells]
                    if len(cells) >= 2:
                        # 排除標題列
                        if any(x in cells[0] for x in ["題號", "Part", "姓名", "班別"]):
                            continue
                        
                        # 從第二欄開始找真正的答案 (排除單個字母題號和純數字分數)
                        target_ans = ""
                        for candidate in cells[1:]:
                            clean_c = candidate.strip()
                            if clean_c and not clean_c.isdigit() and len(clean_c) > 1:
                                if "建議答案" not in clean_c:
                                    target_ans = clean_c
                                    break
                        
                        if target_ans:
                            # 移除答案內重複的分數標記
                            target_ans = re.sub(r'[\(（]\s*\d+\s*分\s*[\)）]', '', target_ans).strip()
                            ans_list.append(target_ans)

            # --- 2. 處理題目卷中的表格 (如第一題表格) ---
            ans_ptr = 0
            for q_table in doc_q.tables:
                for row in q_table.rows:
                    # 如果該列最後一格是空的，我們就填入答案
                    if ans_ptr < len(ans_list):
                        # 檢查是否為需要填寫的表格列 (通常最後兩格是空的)
                        if not row.cells[-1].text.strip():
                            row.cells[-1].text = ans_list[ans_ptr]
                            # 設為藍色加粗
                            for paragraph in row.cells[-1].paragraphs:
                                for run in paragraph.runs:
                                    run.font.bold = True
                                    run.font.color.rgb = RGBColor(0, 102, 204)
                            ans_ptr += 1

            # --- 3. 處理段
