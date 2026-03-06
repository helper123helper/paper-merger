import streamlit as st
from docx import Document
from docx.shared import RGBColor
import io
import re

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
            
            # 1. 改進的答案抓取邏輯 (針對 HSC 表格結構)
            ans_data = []
            for table in doc_ans.tables:
                for row in table.rows:
                    cells_text = [c.text.strip() for c in row.cells if c.text.strip()]
                    # 如果該列有內容，且不只是題號
                    if len(cells_text) >= 2:
                        # 排除掉標題列（如：題號、建議答案、分數 等字眼）
                        if not any(x in cells_text[0] for x in ["題號", "項目", "部分"]):
                            # 抓取中間的答案部分
                            ans_content = " / ".join(cells_text[1:])
                            ans_data.append(ans_content)

            # 2. 強化的題目偵測邏輯
            ans_idx = 0
            for para in doc_q.paragraphs:
                text = para.text.strip()
                # 偵測標準：只要段落結尾包含 (x分) 或 （x分），不論全形半形
                if re.search(r'[\(（]\s*\d+\s*分[\)）]', text):
                    if ans_idx < len(ans_data):
                        # 在題目後方換行並加入藍色答案
                        run = para.add_run(f"\n【建議答案】：{ans_data[ans_idx]}")
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(0, 102, 204)
                        ans_idx += 1

            # 3. 匯出檔案
            bio = io.BytesIO()
            doc_q.save(bio)
            
            if ans_idx > 0:
                st.success(f"✅ 合併成功！已自動填入 {ans_idx} 題答案。")
                st.download_button(
                    label="📥 下載完成的教學 Word 檔",
                    data=bio.getvalue(),
                    file_name="教學檔_自動生成.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.warning("⚠️ 偵測到 0 題。請檢查題目卷是否包含如「（2分）」的標記，或答案卷是否有表格。")
                
        except Exception as e:
            st.error(f"處理過程中發生錯誤：{str(e)}")

st.divider()
st.caption("自動化流程修復版 - 增強了格式兼容性")
