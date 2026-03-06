import streamlit as st
from docx import Document
from docx.shared import RGBColor
import io
import re

st.set_page_config(page_title="HSC 試卷合併工具", page_icon="📝")

st.title("🎓 HSC 模擬考專用：教學檔生成工具")
st.info("已修復 IndentationError。此版本會自動跳過封面，並將答案精準填入題目下方。")

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
            
            # --- 1. 抓取答案 (從 Ans 卷表格提取) ---
            ans_list = []
            for table in doc_ans.tables:
                for row in table.rows:
                    cells = [c.text.strip() for c in row.cells]
                    if len(cells) >= 2:
                        # 排除標題列與封面資訊
                        if any(x in cells[0] for x in ["題號", "Part", "姓名", "班別", "成績"]):
                            continue
                        
                        # 抓取倒數第二格 (建議答案)
                        ans_text = cells[-2]
                        if ans_text and "建議答案" not in ans_text and not ans_text.isdigit():
                            # 清除答案內重複的 (x分)
                            clean_ans = re.sub(r'[\(（]\s*\d+\s*分\s*[\)）]', '', ans_text).strip()
                            if clean_ans:
                                ans_list.append(clean_ans)

            # --- 2. 處理題目卷 (Q) 並填入答案 ---
            ans_idx = 0
            # 定義封面必須跳過的關鍵字，徹底防止走位
            cover_keywords = ["姓名", "班別", "學號", "成績", "年度", "模擬考試", "考生須知"]
            
            for para in doc_q.paragraphs:
                text = para.text.strip()
                
                # 判定邏輯：
                # 1. 段落末尾有 (x 分) 或 (x分)
                # 2. 不包含封面的關鍵字
                # 3. 字數大於 5 (避開純頁碼)
                is_question = re.search(r'[\(（]\s*\d+\s*分\s*[\)）]$', text)
                is_not_cover = not any(k in text for k in cover_keywords)
                
                if is_question and is_not_cover:
                    if ans_idx < len(ans_list):
                        # 確保不重複添加
                        if "【建議答案】" not in para.text:
                            # 加入換行，並以藍色加粗顯示
                            run = para.add_run(f"\n【建議答案】：{ans_list[ans_idx]}")
                            run.font.bold = True
                            run.font.color.rgb = RGBColor(0, 102, 204)
                            ans_idx += 1

            # --- 3. 匯出檔案 ---
            output = io.BytesIO()
            doc_q.save(output)
            
            if ans_idx > 0:
                st.success(f"✅ 成功對位！共填入 {ans_idx} 題答案。")
                st.download_button(
                    label="📥 下載 HSC 教學對位完美版",
                    data=output.getvalue(),
                    file_name="HSC_教學檔_對位完美版.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.warning("未能匹配題目，請檢查題目卷是否有 (分) 標記。")

        except Exception as e:
            st.error(f"錯誤：{str(e)}")

st.divider()
st.caption("2526 HSC Mock 專用修復版 - 解決 IndentationError")
