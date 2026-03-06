import streamlit as st
from docx import Document
from docx.shared import RGBColor
import io

# 設定網頁標題
st.set_page_config(page_title="試卷答案自動合併工具", page_icon="📝")

st.title("🎓 試卷題目與答案自動合併工具")
st.markdown("### 使用說明：\n1. 上傳 **題目卷 (Q)**\n2. 上傳 **答案卷 (Ans)**\n3. 點擊按鈕生成結果。")

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
            
            ans_data = []
            for table in doc_ans.tables:
                for row in table.rows:
                    cells_text = [c.text.strip() for c in row.cells if c.text.strip()]
                    if len(cells_text) > 1:
                        ans_data.append(" / ".join(cells_text[1:]))

            ans_idx = 0
            for para in doc_q.paragraphs:
                if "（" in para.text and "分）" in para.text:
                    if ans_idx < len(ans_data):
                        run = para.add_run(f"\n【建議答案】：{ans_data[ans_idx]}")
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(0, 102, 204)
                        ans_idx += 1

            bio = io.BytesIO()
            doc_q.save(bio)
            
            st.success(f"✅ 合併成功！已處理 {ans_idx} 題。")
            st.download_button(
                label="📥 下載完成的教學 Word 檔",
                data=bio.getvalue(),
                file_name="教學檔_自動生成.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"處理出錯：{str(e)}")

st.divider()
st.caption("專為教育工作者設計 | 自動化流程")
