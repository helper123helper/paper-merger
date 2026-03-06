import streamlit as st
from docx import Document
from docx.shared import RGBColor
import io
import re

st.set_page_config(page_title="HSC 試卷完美合併工具", page_icon="🎓")

st.title("🎓 HSC 模擬考：教學檔生成工具 (完美對位版)")
st.info("已修復：自動跳過英文字母題號，直接將文字答案填入題目下方。")

col1, col2 = st.columns(2)
with col1:
    file_q = st.file_uploader("1. 上傳學生版 (Q)", type=["docx"])
with col2:
    file_ans = st.file_uploader("2. 上傳答案卷 (Ans)", type=["docx"])

if file_q and file_ans:
    if st.button("🪄 執行完美合併", type="primary"):
        try:
            doc_q = Document(file_q)
            doc_ans = Document(file_ans)
            
            # --- 1. 抓取答案庫 (過濾題號，只抓核心答案) ---
            ans_list = []
            for table in doc_ans.tables:
                for row in table.rows:
                    cells = [c.text.strip() for c in row.cells]
                    if len(cells) >= 2:
                        # 排除標題列與封面
                        if any(x in cells[0] for x in ["題號", "Part", "姓名", "班別"]):
                            continue
                        
                        # 重點修復：判斷哪一格才是真正的答案
                        # 通常答案在 cells[1] 或 cells[2]，且不應該是單個英文字母
                        for cell_content in cells[1:]:
                            content = cell_content.strip()
                            # 排除掉：空的、純數字(分數)、或只有一個英文字母(題號)的內容
                            if content and not content.isdigit() and len(content) > 1:
                                if "建議答案" not in content:
                                    # 清理重複的分數標記
                                    clean_ans = re.sub(r'[\(（]\s*\d+\s*分\s*[\)）]', '', content).strip()
                                    if clean_ans:
                                        ans_list.append(clean_ans)
                                        break # 抓到這一行的主要答案後就跳下一行

            # --- 2. 處理題目卷 (Q) ---
            ans_idx = 0
            start_merging = False
            # 封面過濾關鍵字
            cover_keywords = ["姓名", "班別", "學號", "成績", "考試時間"]
            
            for para in doc_q.paragraphs:
                text = para.text.strip()
                
                # 只有看到「甲部」或「回答所有問題」才開始，避開封面
                if "甲部" in text or "回答" in text:
                    start_merging = True
                
                if not start_merging:
                    continue
                
                # 偵測題目特徵：段落末尾有 (x 分)
                if re.search(r'[\(（]\s*\d+\s*分\s*[\)）]$', text):
                    # 再次確保不是封面欄位
                    if not any(k in text for k in cover_keywords):
                        if ans_idx < len(ans_list):
                            if "【建議答案】" not in para.text:
                                # 加入換行、藍色、加粗
                                run = para.add_run(f"\n【建議答案】：{ans_list[ans_idx]}")
                                run.font.bold = True
                                run.font.color.rgb = RGBColor(0, 102, 204)
                                ans_idx += 1

            # --- 3. 匯出 ---
            output = io.BytesIO()
            doc_q.save(output)
            
            if ans_idx > 0:
                st.success(f"✅ 成功對位！共填入 {ans_idx} 題文字答案。")
                st.download_button(
                    label="📥 下載 HSC 教學檔 (完美版)",
                    data=output.getvalue(),
                    file_name="HSC_教學檔_完美修復版.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.warning("未能匹配題目，請檢查格式是否包含「(x 分)」。")

        except Exception as e:
            st.error(f"發生錯誤：{str(e)}")

st.divider()
st.caption("針對 2526 HSC Mock 題號位移問題優化")
