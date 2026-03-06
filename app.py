import streamlit as st
from docx import Document
from docx.shared import RGBColor
import io
import re

st.set_page_config(page_title="HSC 專業對位工具", page_icon="🎓")

st.title("🎓 HSC 模擬考：教學檔生成工具 (修正版)")
st.info("已修復 SyntaxError 並優化表格對位。")

file_q = st.file_uploader("1. 上傳學生版 (Q)", type=["docx"])
file_ans = st.file_uploader("2. 上傳答案卷 (Ans)", type=["docx"])

if file_q and file_ans:
    if st.button("🪄 執行完美合併", type="primary"):
        try:
            doc_q = Document(file_q)
            doc_ans = Document(file_ans)
            
            # --- 1. 抓取答案 (精準過濾題號與分數) ---
            ans_pool = []
            for table in doc_ans.tables:
                for row in table.rows:
                    # 抓取該行所有文字
                    texts = [c.text.strip() for c in row.cells if c.text.strip()]
                    for t in texts:
                        # 過濾條件：長度大於1、非純數字、非標題
                        if len(t) > 1 and not t.isdigit() and "建議答案" not in t and "題號" not in t:
                            # 清理 (分) 標記
                            clean = re.sub(r'[\(（]\s*\d+\s*分\s*[\)）]', '', t).strip()
                            if clean:
                                ans_pool.append(clean)
                                break

            ans_idx = 0
            # --- 2. 優先處理題目卷中的表格 (同步填入) ---
            for table in doc_q.tables:
                for row in table.rows:
                    # 如果最後一格是空的，填入答案
                    if ans_idx < len(ans_pool) and not row.cells[-1].text.strip():
                        cell = row.cells[-1]
                        cell.text = ans_pool[ans_idx]
                        for p in cell.paragraphs:
                            for r in p.runs:
                                r.font.bold = True
                                r.font.color.rgb = RGBColor(0, 102, 204)
                        ans_idx += 1

            # --- 3. 處理剩餘的段落題目 (跳過封面) ---
            start_flag = False
            for para in doc_q.paragraphs:
                txt = para.text.strip()
                
                # 看到「甲部」才開始，防止弄壞封面
                if "甲部" in txt:
                    start_flag = True
                if not start_flag:
                    continue
                
                # 偵測題目：結尾是 (x 分)
                if re.search(r'[\(（]\s*\d+\s*分\s*[\)）]$', txt):
                    if ans_idx < len(ans_pool):
                        if "【建議答案】" not in para.text:
                            run = para.add_run(f"\n【建議答案】：{ans_pool[ans_idx]}")
                            run.font.bold = True
                            run.font.color.rgb = RGBColor(0, 102, 204)
                            ans_idx += 1

            # --- 4. 匯出 ---
            output = io.BytesIO()
            doc_q.save(output)
            st.success(f"✅ 對位完成！共填入 {ans_idx} 處答案。")
            st.download_button(
                label="📥 下載 HSC 完美教學檔",
                data=output.getvalue(),
                file_name="HSC_Final_Teaching_File.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        except Exception as e:
            st.error(f"執行出錯：{str(e)}")

st.divider()
st.caption("針對 2526 HSC S6 Mock 格式優化")
