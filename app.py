import streamlit as st
from docx import Document
from docx.shared import RGBColor
import io
import re

st.set_page_config(page_title="HSC 試卷對位工具", page_icon="🎓")

st.title("🎓 HSC 模擬考：教學檔生成工具")
st.info("已修復：跳過英文字母題號，直接對位表格答案內容。")

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
            
            # --- 1. 抓取答案庫 (從 Ans 表格提取真正的文字答案) ---
            ans_list = []
            for table in doc_ans.tables:
                for row in table.rows:
                    # 抓取該行所有格子的內容
                    cells_text = [c.text.strip() for c in row.cells if c.text.strip()]
                    
                    # 邏輯：跳過標題、跳過單個英文字母、跳過純數字分數
                    # 我們要找的是長度大於 1 的文字描述
                    found_ans = ""
                    for text in cells_text:
                        # 過濾題號(a, b...)、分數(1, 2...)、以及標題
                        if len(text) > 1 and not text.isdigit() and "建議答案" not in text and "題號" not in text:
                            # 移除 (x分) 標記
                            clean_text = re.sub(r'[\(（]\s*\d+\s*分\s*[\)）]', '', text).strip()
                            if clean_text:
                                found_ans = clean_text
                                break
                    
                    if found_ans:
                        ans_list.append(found_ans)

            # --- 2. 處理題目卷 (Q) 的表格填充 (針對第一題那種表格) ---
            ans_idx = 0
            for table in doc_q.tables:
                for row in table.rows:
                    # 如果最後一格是空的且我們有答案，就填進去
                    if ans_idx < len(ans_list) and not row.cells[-1].text.strip():
                        target_cell = row.cells[-1]
                        target_cell.text = ans_list[ans_idx]
                        # 設為藍色加粗
                        for p in target_cell.paragraphs:
                            for r in p.runs:
                                r.font.bold = True
                                r.font.color.rgb = RGBColor(0, 102, 204)
                        ans_idx += 1

            # --- 3. 處理段落題目 (跳過封面) ---
            start_merging = False
            cover_keywords = ["姓名", "班別", "學號", "成績", "年度"]
            
            for para in doc_q.paragraphs:
                p_text = para.text.strip()
                
                # 看到「甲部」才開始填答案，避免封面受
