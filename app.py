import streamlit as st
from docx import Document
from docx.shared import RGBColor
import io
import re

st.set_page_config(page_title="HMSC 試卷對位工具", page_icon="🎓")

st.title("🎓 HMSC 模擬考：教學檔生成工具 (理想成果版)")
st.info("修復：自動過濾題號字母(a,b,c)，直接抓取文字答案並填入正確位置。")

col1, col2 = st.columns(2)
with col1:
    file_q = st.file_uploader("1. 上傳學生版 (Q)", type=["docx"])
with col2:
    file_ans = st.file_uploader("2. 上傳答案卷 (Ans)", type=["docx"])

if file_q and file_ans:
    if st.button("🪄 生成完美對位教學檔", type="primary"):
        try:
            doc_q = Document(file_q)
            doc_ans = Document(file_ans)
            
            # --- 1. 抓取答案庫 (過濾掉題號、分數和空行) ---
            ans_pool = []
            for table in doc_ans.tables:
                for row in table.rows:
                    cells_text = [c.text.strip() for c in row.cells if c.text.strip()]
                    for t in cells_text:
                        # 過濾：長度需大於1(避開a,b,c)、非純數字(避開分數)、非固定標題
                        if len(t) > 1 and not t.isdigit() and "建議答案" not in t and "題號" not in t:
                            # 移除可能存在的 (x分) 標記
                            clean_text = re.sub(r'[\(（]\s*\d+\s*分\s*[\)）]', '', t).strip()
                            if clean_text:
                                ans_pool.append(clean_text)
                                break # 每一列只抓一個最像答案的內容

            ans_idx = 0
            
            # --- 2. 處理題目卷中的表格 (例如第一題的發展理論表) ---
            for table in doc_q.tables:
                for row in table.rows:
                    # 如果最後一格是空的且有答案，就填入
                    if ans_idx < len(ans_pool) and not row.cells[-1].text.strip():
                        target_cell = row.cells[-1]
                        target_cell.text = ans_pool[ans_idx]
                        # 設為藍色
                        for p in target_cell.paragraphs:
                            for r in p.runs:
                                r.font.bold = True
                                r.font.color.rgb = RGBColor(0, 102, 204)
                        ans_idx += 1

            # --- 3. 處理段落題目 (跳過封面) ---
            start_merging = False
            for para in doc_q.paragraphs:
                text = para.text.strip()
                
                # 只有看到「甲部」才開始，避免破壞封面
                if "甲部" in text:
                    start_merging = True
                if not start_merging:
                    continue
                
                # 偵測題目：結尾包含 (x 分)
                if re.search(r'[\(（]\s*\d+\s*分\s*[\)）]$', text):
                    if ans_idx < len(ans_pool):
                        if "【建議答案】" not in para.text:
                            run = para.add_run(f"\n【建議答案】：{ans_pool[ans_idx]}")
                            run.font.bold = True
                            run.font.color.rgb = RGBColor(0, 102, 204)
                            ans_idx += 1

            # --- 4. 匯出 ---
            output = io.BytesIO()
            doc_q.save(output)
            st.success(f"✅ 處理完成！已成功對位並填入 {ans_idx} 處答案。")
            st.download_button(
                label="📥 下載 HMSC 教學檔 (完美修復版)",
                data=output.getvalue(),
                file_name="HMSC_教學檔_完美對位版.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"發生錯誤：{str(e)}")

st.divider()
st.caption("專為 2526 HMSC S6 Mock 格式優化")
