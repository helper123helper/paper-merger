import streamlit as st
from docx import Document
from docx.shared import RGBColor
import io
import re

st.set_page_config(page_title="試卷答案自動合併工具", page_icon="📝")

st.title("🎓 試卷題目與答案自動合併工具")
st.markdown("### 2526 HSC 模擬考專用 (對位校正版)")

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
            
            # 1. 精準抓取答案 (針對 HSC 表格結構優化)
            ans_list = []
            for table in doc_ans.tables:
                for row in table.rows:
                    cells = [c.text.strip() for c in row.cells]
                    # 邏輯：答案通常在倒數第二格（排除最後一格的分數）
                    if len(cells) >= 2:
                        ans_text = cells[-2] # 抓取倒數第二格，通常是建議答案
                        # 排除標題列、空行或封面資訊
                        if ans_text and not any(x in cells[0] for x in ["題號", "Part", "姓名", "班別"]):
                            if "建議答案" not in ans_text and not ans_text.isdigit():
                                # 清除答案內可能存在的 (分)
                                clean_ans = re.sub(r'[\(（]\s*\d+\s*分\s*[\)）]', '', ans_text).strip()
                                if clean_ans:
                                    ans_list.append(clean_ans)

            # 2. 定位題目並填入
            ans_idx = 0
            # 定義封面必須跳過的關鍵字
            cover_keywords = ["姓名", "班別", "學號", "成績", "年度", "模擬考試", "考試時間"]
            
            for para in doc_q.paragraphs:
                p_text = para.text.strip()
                
                # 判斷是否為題目：結尾有 (x 分) 且長度大於 10 且不是封面
                is_q = re.search(r'[\(（]\s*\d+\s*分\s*[\)）]$', p_text)
                is_not_cover = not any(k in p_text for k in cover_keywords)
                
                if is_q and is_not_cover:
                    if ans_idx < len(ans_list):
                        # 確保不重複填入
                        if "【建議答案】" not in para.text:
                            # 插入藍色加粗答案
                            run = para.add_run(f"\n【建議答案】：{ans_list[ans_idx]}")
                            run.font.bold = True
                            run.font.color.rgb = RGBColor(0, 102, 204)
                            ans_idx += 1

            # 3. 匯出檔案
            bio = io.BytesIO()
            doc_q.save(bio)
            
            if ans_idx > 0:
                st.success(f"✅ 處理完成！已成功對位 {ans_idx} 題答案。")
                st.download_button(
                    label="📥 下載 HSC 教學修正檔",
                    data=bio.getvalue(),
                    file_name="HSC_教學檔_對位完美版.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.warning("未能自動匹配題目，請檢查題目卷格式。")

        except Exception as e:
            st.error(f"發生程式錯誤：{str(e)}")

st.divider()
st.caption("已修正 SyntaxError 並針對 S6 Mock 試卷優化對位邏輯")
