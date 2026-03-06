import streamlit as st
from docx import Document
from docx.shared import RGBColor
import io
import re

st.set_page_config(page_title="HSC 專業對位工具", page_icon="🎓")

st.title("🎓 HSC 模擬考：全自動表格對位工具")
st.markdown("### 修正：直接抓取文字答案，跳過英文字母題號")

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
            
            # --- 1. 抓取答案 (精準過濾英文字母題號) ---
            ans_list = []
            for table in doc_ans.tables:
                for row in table.rows:
                    cells = [c.text.strip() for c in row.cells]
                    if len(cells) >= 2:
                        # 排除標題列
                        if any(x in cells[0] for x in ["題號", "Part", "姓名", "班別"]):
                            continue
                        
                        # 在該行的儲存格中尋找真正的文字答案
                        # 邏輯：長度大於 1 且不是純數字，避開 a, b, c 和分數
                        for cell_content in cells:
                            c = cell_content.strip()
                            if len(c) > 1 and not c.isdigit() and "建議答案" not in c:
                                # 移除 (分) 標記
                                clean_ans = re.sub(r'[\(（]\s*\d+\s*分\s*[\)）]', '', c).strip()
                                if clean_ans:
                                    ans_list.append(clean_ans)
                                    break # 每一行只抓一個核心答案

            # --- 2. 處理題目卷 (Q) ---
            ans_ptr = 0
            start_merging = False
            # 封面關鍵字
            cover_keywords = ["姓名", "班別", "學號", "成績", "年度", "考試時間"]
            
            for para in doc_q.paragraphs:
                text = para.text.strip()
                
                # 只有看到「甲部」才開始，徹底避開封面走位
                if "甲部" in text:
                    start_merging = True
                
                if not start_merging:
                    continue
                
                # 偵測題目：結尾有 (分)
                if re.search(r'[\(（]\s*\d+\s*分\s*[\)）]$', text):
                    # 確保不是封面欄位
                    if not any(k in text for k in cover_keywords):
                        if ans_ptr < len(ans_list):
                            if "【建議答案】" not in para.text:
                                run = para.add_run(f"\n【建議答案】：{ans_list[ans_ptr]}")
                                run.font.bold = True
                                run.font.color.rgb = RGBColor(0,
