import streamlit as st
import pandas as pd
from docx import Document
import io
import json
import os
import csv

# 設定網頁標題與寬度
st.set_page_config(page_title="值班日誌自動生成器", layout="wide")

# ================= 系統設定與資料庫初始化 =================
TEMPLATE_PATH = "template.docx"
DB_FILE = "handovers.json"

# 載入歷史交班紀錄 (實現關閉不重置)
def load_handovers():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r", encoding="utf-8") as f:
            try:
                return json.load(f)
            except:
                return []
    return []

# 儲存交班紀錄
def save_handovers(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

# 初始化 Session State
if 'handovers' not in st.session_state:
    st.session_state.handovers = load_handovers()

st.title("🏥 醫師病房值班日誌自動生成器")

# ================= 區塊 1：上傳 CSV 數據 =================
st.header("1. 上傳 HIS 系統匯出檔案")
st.info("系統已內建 `template.docx` 樣板，您現在只需上傳數據即可。")
csv_file = st.file_uploader("上傳值班日誌數據 (.csv)", type=['csv'])

# ================= 區塊 2：交班事項登錄表單 =================
st.header("2. 交班事項登錄")

with st.form("handover_form", clear_on_submit=True):
    st.subheader("新增交班紀錄")
    c1, c2, c3 = st.columns(3)
    with c1:
        is_er = st.checkbox("🚨 ER (急診)")
        name = st.text_input("病人姓名 (必填)")
        age = st.text_input("年紀")
    with c2:
        gender = st.selectbox("性別", ["", "男", "女"])
        med_record = st.text_input("病歷號")
        time_occurred = st.time_input("狀況發生時間")
    with c3:
        attending_doc = st.text_input("主治醫師")
        diagnosis = st.text_input("診斷")
        
    content = st.text_area("交班內容 (必填)")
    submitted = st.form_submit_button("確認新增交班")
    
    if submitted:
        if not name or not content:
            st.error("「病人姓名」與「交班內容」為必填欄位！")
        else:
            new_record = {
                "is_er": is_er, "name": name, "age": age,
                "gender": gender, "med_record": med_record,
                "attending_doc": attending_doc, "diagnosis": diagnosis,
                "time_occurred": time_occurred.strftime("%H:%M"),
                "content": content
            }
            st.session_state.handovers.append(new_record)
            save_handovers(st.session_state.handovers) # 立即存檔
            st.success(f"已新增 {name} 的交班紀錄！")
            st.rerun()

# ================= 區塊 3：已登錄交班事項 =================
st.header("3. 已登錄交班事項")

col1, col2 = st.columns([8, 2])
with col2:
    if st.button("🗑️ 清空所有交班紀錄", type="secondary", help="值班結束後點此清空"):
        st.session_state.handovers = []
        save_handovers([])
        st.rerun()

if not st.session_state.handovers:
    st.info("目前尚無交班紀錄。")
else:
    for idx, h in enumerate(st.session_state.handovers):
        er_tag = "🔴 [ER] " if h["is_er"] else ""
        title = f"{er_tag}{h['name']} - {h['time_occurred']}"
        
        with st.expander(title):
            details = [f"{k}: {h[v]}" for k, v in zip(
                ["年紀", "性別", "病歷號", "主治醫師", "診斷"], 
                ["age", "gender", "med_record", "attending_doc", "diagnosis"]
            ) if h[v]]
            
            st.write(" | ".join(details))
            st.markdown(f"**交班內容：**\n{h['content']}")
            
            if st.button(f"刪除 {h['name']}", key=f"del_{idx}"):
                st.session_state.handovers.pop(idx)
                save_handovers(st.session_state.handovers)
                st.rerun()

# ================= 區塊 4：解析與生成文件 =================
st.header("4. 確認與輸出")

def process_csv_and_word(csv_data, handovers):
    # 檢查樣板是否存在
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"找不到樣板檔案 '{TEMPLATE_PATH}'，請確認檔案名稱與位置。")
        
    doc = Document(TEMPLATE_PATH)
    
    # 1. 處理交班事項文字生成
    handover_text = ""
    for h in handovers:
        er_mark = "🚨[ER] " if h["is_er"] else ""
        handover_text += f"【{er_mark}{h['name']}】 {h['time_occurred']}\n"
        meta = [h[k] for k in ['age', 'gender', 'med_record', 'attending_doc', 'diagnosis'] if h[k]]
        if meta: handover_text += f"基本資料：{' / '.join(meta)}\n"
        handover_text += f"交班內容：{h['content']}\n"
        handover_text += "-" * 20 + "\n"

    # 寫入交班事項到「病房特殊狀況及處理：」下方
    for paragraph in doc.paragraphs:
        if "病房特殊狀況及處理：" in paragraph.text:
            paragraph.add_run("\n" + handover_text)

    # 2. 如果有上傳 CSV，解析並填入表格
    if csv_data is not None:
        lines = csv_data.getvalue().decode("utf-8").splitlines()
        reader = list(csv.reader(lines))
        
        # 尋找區塊索引
        idx_station = idx_new = idx_out = -1
        for i, row in enumerate(reader):
            if not row: continue
            if row[0] == "護理站": idx_station = i
            elif row[0] == "新入院病人": idx_new = i
            elif row[0] == "出院病人": idx_out = i
            
        # 填入護理站人數 (假設是 Word 中的第 1 個表格)
        if idx_station != -1 and len(doc.tables) >= 1:
            table = doc.tables[0]
            # 對應 CSV: 護理站(急診,二樓,三樓,四樓,五樓,總人數) 在 idx_station + 1 到 + 6
            for r_offset in range(1, 7):
                csv_row = reader[idx_station + r_offset]
                # Word表格的 row 從 1 開始填 (避開標題)
                word_row_idx = r_offset 
                if len(csv_row) >= 4 and word_row_idx < len(table.rows):
                    table.cell(word_row_idx, 1).text = csv_row[1] # 男
                    table.cell(word_row_idx, 2).text = csv_row[2] # 女
                    table.cell(word_row_idx, 3).text = csv_row[3] # 總數

        # 填入新入院病人 (假設是 Word 中的第 2 個表格)
        if idx_new != -1 and len(doc.tables) >= 2:
            table = doc.tables[1]
            # 找到實際資料起始點 (跳過「新入院病人」與「欄位名稱」兩行)
            data_start = idx_new + 2 
            for i in range(data_start, len(reader)):
                row = reader[i]
                if not row or not row[0].strip() or row[0] == "出院病人": break # 遇到空白或下一區塊停止
                
                # Word 中預設有空白行，計算目前要填入的 row
                word_row_idx = i - data_start + 1
                if word_row_idx < len(table.rows):
                    table.cell(word_row_idx, 0).text = row[0] # 姓名
                    table.cell(word_row_idx, 1).text = row[1] # 病歷號
                    table.cell(word_row_idx, 2).text = row[2] # 床號
                    table.cell(word_row_idx, 3).text = row[3] # 性別
                    table.cell(word_row_idx, 4).text = row[4] # 年齡
                    table.cell(word_row_idx, 5).text = row[5] # 診斷
                    if len(row) > 6:
                        table.cell(word_row_idx, 7).text = row[6] # 燈號 (依您的欄位設定放在燈號欄)

        # 填入出院病人 (假設是 Word 中的第 3 個表格)
        if idx_out != -1 and len(doc.tables) >= 3:
            table = doc.tables[2]
            data_start = idx_out + 2
            for i in range(data_start, len(reader)):
                row = reader[i]
                if not row or not row[0].strip(): break
                
                word_row_idx = i - data_start + 1
                if word_row_idx < len(table.rows):
                    table.cell(word_row_idx, 0).text = row[0] # 姓名
                    table.cell(word_row_idx, 1).text = row[1] # 病歷號
                    table.cell(word_row_idx, 2).text = row[2] # 床號
                    table.cell(word_row_idx, 3).text = row[3] # 性別
                    table.cell(word_row_idx, 4).text = row[4] # 年齡
                    table.cell(word_row_idx, 5).text = row[5] # 診斷
                    if len(row) > 6:
                        table.cell(word_row_idx, 6).text = row[6] # 出院動態

    # 存入記憶體供下載
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

if st.button("🚀 確認輸出 Word 檔案", type="primary"):
    try:
        final_doc_stream = process_csv_and_word(csv_file, st.session_state.handovers)
        st.success("檔案生成成功！")
        st.download_button(
            label="📥 點擊下載今日值班日誌",
            data=final_doc_stream,
            file_name="今日值班日誌_已完成.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"生成檔案時發生錯誤: {e}")
