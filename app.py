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

st.title("🏥 醫師病房值班日誌自動生成器 (無痕格式版)")

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
    duty_date = st.date_input("📅 選擇值班日期", datetime.date.today())

with col_text:
    raw_text_input = st.text_area(
        "📝 在此貼上資料 (從 Excel 複製包含護理站與病人的區塊)", 
        height=200, 
        key=f"text_input_{st.session_state.uploader_key}"
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
st.header("3. 已登錄交班事項")
if not st.session_state.handovers:
    st.info("目前尚無交班紀錄。")
else:
    sorted_view = sorted(st.session_state.handovers, key=lambda x: (not x['is_er'], x['time_occurred']))
    for h in sorted_view:
        original_idx = st.session_state.handovers.index(h) 
        title = f"{'🚨[ER] ' if h['is_er'] else ''}{h['name']} - {h['time_occurred']}"
        with st.expander(title):
            st.markdown(f"**資料：** {h['age']}歲/{h['gender']} | 病歷：{h['med_record']} | 主治：{h['attending_doc']}")
            st.markdown(f"**診斷：** {h['diagnosis']}")
            st.markdown(f"**交班內容：**\n{h['content']}")
            if st.button(f"刪除 {h['name']}", key=f"del_{original_idx}"):
                st.session_state.handovers.pop(original_idx)
                save_handovers(st.session_state.handovers)
                st.rerun()

# ================= 區塊 4：解析與無痕生成核心邏輯 =================
def process_data(raw_text, handovers, selected_date):
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"找不到 {TEMPLATE_PATH}。請確認已上傳樣板。")
    
    doc = Document(TEMPLATE_PATH)
    
    # --- 0. 填寫日期 (無痕替換) ---
    roc_year = selected_date.year - 1911
    date_str = f"日期：  {roc_year} 年  {selected_date.month:02d} 月  {selected_date.day:02d} 日"
    for p in doc.paragraphs:
        if "日期" in p.text:
            p.text = date_str
            break

    # --- 1. 文字解析引擎 (精準提取數據) ---
    parsed_stations = {}
    parsed_new = []
    parsed_out = []
    
    if raw_text:
        lines = raw_text.splitlines()
        current_section = None
        for line in lines:
            line = line.strip()
            if not line: continue
            
            # 狀態切換
            if "護理站" in line and "病人總數" in line: current_section = "station"
            elif "新入院" in line or "新住院" in line: current_section = "new"
            elif "出院病人" in line: current_section = "out"
            
            # 切割資料 (支援 Tab 或多重空白)
            parts = [p.strip() for p in line.split('\t')]
            if len(parts) < 2:
                parts = [p.strip() for p in re.split(r'\s{2,}', line)]
            
            # 裝載資料
            if current_section == "station":
                st_name = parts[0].replace(" ", "")
                if ("護理站" in st_name or "總人數" in st_name) and len(parts) >= 4:
                    # 抓取對應的 男、女、總數 數字
                    parsed_stations[st_name] = parts[1:4] 
            elif current_section == "new":
                if "姓名" in parts[0] or "病患" in parts[0]: continue
                if len(parts) >= 5: parsed_new.append(parts)
            elif current_section == "out":
                if "姓名" in parts[0] or "病患" in parts[0]: continue
                if len(parts) >= 5: parsed_out.append(parts)

    # --- 2. Word 無痕填表引擎 (絕對不增刪行) ---
    new_idx = 0
    out_idx = 0
    
    for table in doc.tables:
        fill_mode = None
        for row in table.rows:
            cells = row.cells
            if not cells: continue
            
            c0_text = cells[0].text.replace(" ", "").strip()
            
            # A. 護理站人數對位填入
            if c0_text in parsed_stations:
                nums = parsed_stations[c0_text]
                if len(cells) >= 4:
                    cells[1].text = str(nums[0]) # 男
                    cells[2].text = str(nums[1]) # 女
                    cells[3].text = str(nums[2]) # 總人數
                continue
            
            # B. 判斷是否為病人表頭 (例如: "姓名" 且旁邊是 "病歷號")
            if "姓名" in c0_text and len(cells) > 1 and "病歷" in cells[1].text.replace(" ", ""):
                header_text = "".join([c.text for c in cells]).replace(" ", "")
                if "燈號" in header_text or "危險" in header_text:
                    fill_mode = "new"
                elif "動態" in header_text or "出院" in header_text:
                    fill_mode = "out"
                continue # 標題列跳過，下一行開始填寫
            
            # C. 遇到其他無關段落，關閉填寫模式
            if fill_mode and ("出院病人" in c0_text or "危險評估" in c0_text or "病房特殊" in c0_text):
                fill_mode = None
            
            # D. 在預留的空格中填入病人資料
            if fill_mode == "new":
                # 只有當我們還有新病人資料，且該行是空的(或只有底線/空格)時才填寫
                if new_idx < len(parsed_new) and (not c0_text or c0_text == "" or "_" in c0_text):
                    p_data = parsed_new[new_idx]
                    cells[0].text = p_data[0] # 抓取名字(例如張Ｏ月娥)
                    cells[1].text = p_data[1] if len(p_data) > 1 else ""
                    cells[2].text = p_data[2] if len(p_data) > 2 else ""
                    cells[3].text = p_data[3] if len(p_data) > 3 else ""
                    cells[4].text = p_data[4] if len(p_data) > 4 else ""
                    cells[5].text = p_data[5] if len(p_data) > 5 else ""
                    if len(p_data) > 6 and len(cells) > 7:
                        cells[7].text = p_data[6] # 燈號
                    new_idx += 1
            
            elif fill_mode == "out":
                if out_idx < len(parsed_out) and (not c0_text or c0_text == "" or "_" in c0_text):
                    p_data = parsed_out[out_idx]
                    cells[0].text = p_data[0] # 抓取名字
                    cells[1].text = p_data[1] if len(p_data) > 1 else ""
                    cells[2].text = p_data[2] if len(p_data) > 2 else ""
                    cells[3].text = p_data[3] if len(p_data) > 3 else ""
                    cells[4].text = p_data[4] if len(p_data) > 4 else ""
                    cells[5].text = p_data[5] if len(p_data) > 5 else ""
                    if len(p_data) > 6 and len(cells) > 6:
                        cells[6].text = p_data[6] # 出院動態
                    out_idx += 1

    # --- 3. 填寫交班事項 ---
    sorted_handovers = sorted(handovers, key=lambda x: (not x['is_er'], x['time_occurred']))
    h_text = ""
    for h in sorted_handovers:
        h_text += f"\n【{'🚨ER ' if h['is_er'] else ''}{h['name']}】{h['time_occurred']} | {h['age']}歲/{h['gender']} | 病歷:{h['med_record']}({h['attending_doc']})\n"
        h_text += f"交班：{h['content']}"

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
