import streamlit as st
import pandas as pd
from docx import Document
import io
import json
import os
import re
import datetime

st.set_page_config(page_title="值班日誌自動生成器", layout="wide")

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
if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0

st.title("🏥 醫師病房值班日誌自動生成器")

# ================= 區塊 1：全局控制與資料輸入 =================
col_title, col_btn = st.columns([8, 2])
with col_btn:
    if st.button("🔄 刷新並清空所有資料", type="secondary", use_container_width=True):
        st.session_state.handovers = []
        save_handovers([])
        st.session_state.uploader_key += 1
        st.rerun()

st.header("1. 貼上 HIS 系統匯出資料")
col_date, col_text = st.columns([2, 8])
with col_date:
    # 讓使用者直接選擇日期，自動轉民國年寫入 Word
    duty_date = st.date_input("📅 選擇值班日期", datetime.date.today())

with col_text:
    # 改為 Text Area 接收直接貼上的文字
    raw_text_input = st.text_area(
        "📝 在此貼上資料 (支援直接從 Excel 複製貼上)", 
        height=200, 
        key=f"text_input_{st.session_state.uploader_key}",
        help="請連同標題（如護理站、新入院病人等）一起複製貼上。"
    )

# ================= 區塊 2：交班事項登錄表單 =================
st.header("2. 交班事項登錄")
with st.form("handover_form", clear_on_submit=True):
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
        attending_doc = st.selectbox("主治醫師", ["", "鍾", "張", "劉", "謝", "成"])
        diagnosis = st.text_input("診斷")
    content = st.text_area("交班內容 (必填)")
    
    if st.form_submit_button("確認新增交班"):
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
st.header("3. 已登錄交班事項 (依 ER 與時間排序)")

if not st.session_state.handovers:
    st.info("目前尚無交班紀錄。")
else:
    sorted_view = sorted(st.session_state.handovers, key=lambda x: (not x['is_er'], x['time_occurred']))
    for h in sorted_view:
        original_idx = st.session_state.handovers.index(h) 
        title = f"{'🚨[ER] ' if h['is_er'] else ''}{h['name']} - {h['time_occurred']}"
        with st.expander(title):
            st.markdown(f"**詳細資料：** {h['age']}歲/{h['gender']} | 病歷：{h['med_record']} | 主治：{h['attending_doc']}")
            st.markdown(f"**診斷：** {h['diagnosis']}")
            st.markdown(f"**交班內容：**\n{h['content']}")
            if st.button(f"刪除 {h['name']}", key=f"del_{original_idx}"):
                st.session_state.handovers.pop(original_idx)
                save_handovers(st.session_state.handovers)
                st.rerun()

# ================= 區塊 4：解析與生成核心邏輯 =================
def process_data(raw_text, handovers, selected_date):
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"找不到 {TEMPLATE_PATH}。請確認已上傳樣板。")
    
    doc = Document(TEMPLATE_PATH)
    
    # --- 0. 填寫日期 ---
    roc_year = selected_date.year - 1911
    date_str = f"日期：  {roc_year} 年  {selected_date.month:02d} 月  {selected_date.day:02d} 日"
    for p in doc.paragraphs:
        if "日期" in p.text:
            p.text = date_str
            break

    # --- 1. 將貼上的文字轉換為結構化資料 (List of Lists) ---
    lines = []
    if raw_text:
        raw_lines = raw_text.splitlines()
        for line in raw_lines:
            if not line.strip(): continue # 略過完全空白的行
            
            # 從 Excel 貼上通常會用 Tab 分隔，如果沒有 Tab 則用多個空白分隔
            if '\t' in line:
                row = [cell.strip() for cell in line.split('\t')]
            else:
                row = [cell.strip() for cell in re.split(r'\s{2,}', line)]
            lines.append(row)
    
    # --- 2. 尋找各區塊表頭的索引位置 ---
    idx_station = idx_new = idx_out = -1
    for i, row in enumerate(lines):
        row_str = "".join(row).replace(" ", "")
        
        # 尋找關鍵字
        if "護理站" in row_str and "病人總數" in row_str: idx_station = i
        elif "病患姓名" in row_str and "入院燈號" in row_str: idx_new = i
        elif "病患姓名" in row_str and "出院動態" in row_str: idx_out = i

    # --- 3. 填寫人數統計表 ---
    if idx_station != -1 and len(doc.tables) >= 1:
        tb = doc.tables[0]
        w_row = 1
        for i in range(idx_station + 1, len(lines)):
            row = lines[i]
            if not "".join(row).strip(): continue
            if "病患姓名" in "".join(row) or "新入院" in "".join(row): break # 遇到下一個區塊就停止
            
            if w_row < len(tb.rows) and len(row) >= 4:
                # 確保只抓數字，避免抓到雜訊
                tb.cell(w_row, 1).text = str(row[1]).strip()
                tb.cell(w_row, 2).text = str(row[2]).strip()
                tb.cell(w_row, 3).text = str(row[3]).strip()
                w_row += 1
            if len(row) > 0 and "總人數" in row[0]: break

    # --- 4. 填寫新入院 ---
    if idx_new != -1 and len(doc.tables) >= 2:
        tb = doc.tables[1]
        w_row = 1
        for i in range(idx_new + 1, len(lines)):
            row = lines[i]
            if not "".join(row).strip(): continue
            if "出院" in "".join(row) or "病患姓名" in "".join(row): break
            if w_row >= len(tb.rows): tb.add_row()
            
            if len(row) > 0: tb.cell(w_row, 0).text = str(row[0]).strip()
            if len(row) > 1: tb.cell(w_row, 1).text = str(row[1]).strip()
            if len(row) > 2: tb.cell(w_row, 2).text = str(row[2]).strip()
            if len(row) > 3: tb.cell(w_row, 3).text = str(row[3]).strip()
            if len(row) > 4: tb.cell(w_row, 4).text = str(row[4]).strip()
            if len(row) > 5: tb.cell(w_row, 5).text = str(row[5]).strip()
            if len(row) > 6: tb.cell(w_row, 7).text = str(row[6]).strip() # 寫入入院燈號
            w_row += 1

    # --- 5. 填寫出院 ---
    if idx_out != -1 and len(doc.tables) >= 3:
        tb = doc.tables[2]
        w_row = 1
        for i in range(idx_out + 1, len(lines)):
            row = lines[i]
            if not "".join(row).strip(): continue
            if w_row >= len(tb.rows): tb.add_row()
            
            if len(row) > 0: tb.cell(w_row, 0).text = str(row[0]).strip()
            if len(row) > 1: tb.cell(w_row, 1).text = str(row[1]).strip()
            if len(row) > 2: tb.cell(w_row, 2).text = str(row[2]).strip()
            if len(row) > 3: tb.cell(w_row, 3).text = str(row[3]).strip()
            if len(row) > 4: tb.cell(w_row, 4).text = str(row[4]).strip()
            if len(row) > 5: tb.cell(w_row, 5).text = str(row[5]).strip()
            if len(row) > 6: tb.cell(w_row, 6).text = str(row[6]).strip() # 寫入出院動態
            w_row += 1

    # --- 6. 填寫交班事項 ---
    sorted_handovers = sorted(handovers, key=lambda x: (not x['is_er'], x['time_occurred']))
    h_text = ""
    for h in sorted_handovers:
        h_text += f"\n【{'🚨ER ' if h['is_er'] else ''}{h['name']}】 {h['time_occurred']}\n"
        h_text += f"資料：{h['age']}歲/{h['gender']}/{h['med_record']} (主治:{h['attending_doc']})\n"
        h_text += f"交班：{h['content']}\n" + ("-"*30)

    inserted = False
    for p in doc.paragraphs:
        if "病房特殊狀況及處理" in p.text.replace(" ", ""):
            p.add_run(h_text)
            inserted = True
            break
            
    if not inserted:
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if "病房特殊狀況及處理" in cell.text.replace(" ", ""):
                        cell.text += h_text
                        inserted = True
                        break
                if inserted: break
            if inserted: break

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

st.header("4. 確認與輸出")
if st.button("🚀 生成並下載 Word 檔案", type="primary"):
    try:
        # 將輸入的文字與選擇的日期傳入函數
        final_file = process_data(raw_text_input, st.session_state.handovers, duty_date)
        st.success("檔案已更新並備妥！")
        st.download_button(
            label="📥 點擊下載最新版值班日誌",
            data=final_file,
            file_name=f"值班日誌_{duty_date.strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"錯誤詳情: {e}")
