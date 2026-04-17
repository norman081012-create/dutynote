import streamlit as st
import pandas as pd
from docx import Document
import io
import json
import os
import csv
import re

# 設定網頁標題與寬度
st.set_page_config(page_title="值班日誌自動生成器-穩定版", layout="wide")

# ================= 系統設定與資料庫初始化 =================
TEMPLATE_PATH = "template.docx"
DB_FILE = "handovers.json"

def load_handovers():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r", encoding="utf-8") as f:
            try: return json.load(f)
            except: return []
    return []

def save_handovers(data):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

if 'handovers' not in st.session_state:
    st.session_state.handovers = load_handovers()

st.title("🏥 醫師病房值班日誌自動生成器 (穩定版)")

# ================= 區塊 1：上傳 CSV 數據 =================
st.header("1. 上傳 HIS 系統匯出檔案")
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
            st.session_state.handovers.append({
                "is_er": is_er, "name": name, "age": age,
                "gender": gender, "med_record": med_record,
                "attending_doc": attending_doc, "diagnosis": diagnosis,
                "time_occurred": time_occurred.strftime("%H:%M"),
                "content": content
            })
            save_handovers(st.session_state.handovers)
            st.success(f"已新增 {name} 的交班紀錄！")
            st.rerun()

# ================= 區塊 3：已登錄交班事項預覽 =================
st.header("3. 已登錄交班事項")
col1, col2 = st.columns([8, 2])
with col2:
    if st.button("🗑️ 清空所有交班紀錄", type="secondary"):
        st.session_state.handovers = []
        save_handovers([])
        st.rerun()

if not st.session_state.handovers:
    st.info("目前尚無交班紀錄。")
else:
    for idx, h in enumerate(st.session_state.handovers):
        title = f"{'🔴[ER] ' if h['is_er'] else ''}{h['name']} - {h['time_occurred']}"
        with st.expander(title):
            st.markdown(f"**詳細資料：** {h['age']}歲/{h['gender']} | 病歷：{h['med_record']} | 主治：{h['attending_doc']}")
            st.markdown(f"**診斷：** {h['diagnosis']}")
            st.markdown(f"**交班內容：**\n{h['content']}")
            if st.button(f"刪除 {h['name']}", key=f"del_{idx}"):
                st.session_state.handovers.pop(idx)
                save_handovers(st.session_state.handovers)
                st.rerun()

# ================= 區塊 4：解析與生成核心邏輯 =================
def process_data(csv_data, handovers):
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"找不到 {TEMPLATE_PATH}。請確認已上傳 template.docx 到 GitHub。")
    
    doc = Document(TEMPLATE_PATH)
    
    # 1. 解析 CSV 內容
    lines = []
    if csv_data:
        decoded = csv_data.getvalue().decode("utf-8")
        lines = list(csv.reader(io.StringIO(decoded)))
    
    # 2. 處理日期與基本欄位比對索引
    idx_map = {}
    for i, row in enumerate(lines):
        row_str = "".join(row)
        if "值班起始日期" in row_str:
            match = re.search(r'(\d{4})(\d{2})(\d{2})', row_str)
            if match:
                y, m, d = match.groups()
                date_text = f"日期：  {int(y)-1911} 年  {m} 月  {d} 日"
                for p in doc.paragraphs:
                    if "日期" in p.text: p.text = date_text
        if "護理站" in row_str: idx_map["station"] = i
        if "新入院病人" in row_str: idx_map["new"] = i
        if "出院病人" in row_str: idx_map["out"] = i

    # 3. 填寫人數統計表 (Word 表格 0)
    if "station" in idx_map and len(doc.tables) >= 1:
        tb = doc.tables[0]
        for offset in range(1, 7): # 急診到總人數共6列
            csv_row = lines[idx_map["station"] + offset]
            if offset < len(tb.rows):
                tb.cell(offset, 1).text = csv_row[1] # 男
                tb.cell(offset, 2).text = csv_row[2] # 女
                tb.cell(offset, 3).text = csv_row[3] # 總數

    # 4. 填寫新入院 (Word 表格 1) - 自動增行
    if "new" in idx_map and len(doc.tables) >= 2:
        tb = doc.tables[1]
        data_rows = []
        for i in range(idx_map["new"] + 2, len(lines)):
            if not lines[i] or not lines[i][0].strip() or "出院" in lines[i][0]: break
            data_rows.append(lines[i])
        
        for i, csv_row in enumerate(data_rows):
            w_idx = i + 1 # 跳過標題列
            if w_idx >= len(tb.rows): tb.add_row()
            tb.cell(w_idx, 0).text = csv_row[0] # 姓名
            tb.cell(w_idx, 1).text = csv_row[1] # 病歷
            tb.cell(w_idx, 2).text = csv_row[2] # 床號
            tb.cell(w_idx, 3).text = csv_row[3] # 性別
            tb.cell(w_idx, 4).text = csv_row[4] # 年齡
            tb.cell(w_idx, 5).text = csv_row[5] # 診斷
            if len(csv_row) > 6: tb.cell(w_idx, 7).text = csv_row[6] # 燈號

    # 5. 填寫出院 (Word 表格 2) - 自動增行
    if "out" in idx_map and len(doc.tables) >= 3:
        tb = doc.tables[2]
        data_rows = []
        for i in range(idx_map["out"] + 2, len(lines)):
            if not lines[i] or not lines[i][0].strip(): break
            data_rows.append(lines[i])
        
        for i, csv_row in enumerate(data_rows):
            w_idx = i + 1
            if w_idx >= len(tb.rows): tb.add_row()
            tb.cell(w_idx, 0).text = csv_row[0]
            tb.cell(w_idx, 1).text = csv_row[1]
            tb.cell(w_idx, 2).text = csv_row[2]
            tb.cell(w_idx, 3).text = csv_row[3]
            tb.cell(w_idx, 4).text = csv_row[4]
            tb.cell(w_idx, 5).text = csv_row[5]
            tb.cell(w_idx, 6).text = csv_row[6] # 出院動態

    # 6. 填寫交班事項 (核心修復)
    h_text = ""
    for h in handovers:
        h_text += f"\n【{'🔴ER' if h['is_er'] else ''}{h['name']}】 {h['time_occurred']}\n"
        h_text += f"資料：{h['age']}歲/{h['gender']}/{h['med_record']} (主治:{h['attending_doc']})\n"
        h_text += f"交班：{h['content']}\n" + ("-"*30)

    # 搜尋所有段落與表格儲存格，確保抓到位置
    target_found = False
    for p in doc.paragraphs:
        if "病房特殊狀況及處理" in p.text.replace(" ", ""):
            p.add_run(h_text)
            target_found = True
    
    if not target_found: # 如果段落找不到，找表格內
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if "病房特殊狀況及處理" in cell.text:
                        cell.text += h_text

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

if st.button("🚀 生成並下載 Word 檔案", type="primary"):
    try:
        final_file = process_data(csv_file, st.session_state.handovers)
        st.success("檔案已備妥！")
        st.download_button(
            label="📥 點擊下載",
            data=final_file,
            file_name="值班日誌_輸出版本.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"錯誤詳情: {e}")
