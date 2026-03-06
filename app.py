import streamlit as st
from docx import Document
from docx.shared import RGBColor
import io
import re

st.set_page_config(page_title="試卷答案自動合併工具", page_icon="📝")

st.title("🎓 試卷題目與答案自動合併工具")
st.markdown("### HSC 模擬試卷專用版\n1. 上傳 **學生版 (Q)**\n2. 上傳 **Answer (Ans)**")

col1, col2 = st.columns(2)
with col1:
    file_q = st.file_uploader("1. 上傳題目卷 (Q)", type=["docx"])
with col2:
    file_ans = st.file_uploader("2. 上傳答案卷 (Ans)", type=["docx"])

if file_q and file_ans:
    if st.button("🪄 執行自動合併", type="primary"):
        try:
            doc_q = Document(file_q)
            doc_ans = Document(file_ans)
            
            # 1. 精準抓取答案卷表格內容
            ans_data = []
            for table in doc_ans.tables:
                for row in table.rows:
                    cells = [c.text.strip() for c in row.cells]
                    # HSC 答案格式通常是：左邊題號，中間答案，右邊分數
                    if len(cells) >= 2:
                        # 排除標題列
                        if "題號" in cells[0] or "建議答案" in cells[1]:
                            continue
                        
                        # 真正的答案通常在第 2 格 (索引 1)
                        main_ans = cells[1].replace('\n', ' ')
                        if main_ans:
                            ans_data.append(main_ans)

            # 2. 定位題目並填入
            ans_idx = 0
            for para in doc_q.paragraphs:
                text = para.text.strip()
                
                # 偵測 HSC 題目結尾常見的 (x 分) 標記
                if re.search(r'[\(（]\s*\d+\s*分\s*[\)）]', text):
                    if ans_idx < len(ans_data):
                        # 在題目段落最後一個 Run 之後加入答案，避免破壞原段落格式
                        run = para.add_run(f"\n【建議答案】：{ans_data[ans_idx]}")
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(0, 102, 204) # 藍色
                        ans_idx += 1

            # 3. 匯出檔案
            bio = io.BytesIO()
            doc_q.save(bio)
            
            if ans_idx > 0:
                st.success(f"✅ 修正成功！已根據 HSC 格式填入 {ans_idx} 個答案。")
                st.download_button(
                    label="📥 下載修正後的教學 Word 檔",
                    data=bio.getvalue(),
                    file_name="HSC_教學檔_修正版.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.warning("未能偵測到題目，請檢查題目卷末尾是否有「(分)」字樣。")
                
        except Exception as e:
            st.error(f"發生錯誤：{str(e)}")

st.divider()
st.caption("已針對 2526 HSC 模擬考格式優化")
