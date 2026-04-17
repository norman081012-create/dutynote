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
    duty_date = st.date_input("📅 選擇值班日期", datetime.date.today())

with col_text:
    raw_text_input = st.text_area(
        "📝 在此貼上資料 (支援直接從 Excel 複製貼上)", 
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

    # --- 1. 建立高容錯的文字解析器 ---
    parsed_data = {
        "stations": {},
        "new_patients": [],
        "out_patients": []
    }
    
    if raw_text:
        lines = raw_text.splitlines()
        current_section = None
        for line in lines:
            line = line.strip()
            if not line: continue
            
            # 判斷目前讀取到哪個區塊
            if "急診護理站" in line or "二樓護理站" in line:
                current_section = "station"
            elif "新入院" in line or "新住院" in line:
                current_section = "new"
            elif "出院病人" in line:
                current_section = "out"
            
            # 切割資料 (支援 Tab 或多重空白)
            parts = [p.strip() for p in line.split('\t')]
            if len(parts) < 2:
                parts = [p.strip() for p in re.split(r'\s{2,}', line)]
            
            # 存入對應的資料庫
            if current_section == "station":
                st_name = parts[0].replace(" ", "")
                if "護理站" in st_name or "總人數" in st_name:
                    if len(parts) >= 4:
                        parsed_data["stations"][st_name] = parts[1:4]
            elif current_section == "new":
                if "姓名" in parts[0] or "病患姓名" in parts[0]: continue
                if len(parts) >= 6:
                    parsed_data["new_patients"].append(parts)
            elif current_section == "out":
                if "姓名" in parts[0] or "病患姓名" in parts[0]: continue
                if len(parts) >= 6:
                    parsed_data["out_patients"].append(parts)

    # --- 2. Word 智慧填寫與刪除引擎 (無差別掃描所有表格) ---
    for table in doc.tables:
        row_idx = 0
        fill_mode = None
        p_idx = 0
        
        while row_idx < len(table.rows):
            row = table.rows[row_idx]
            cells = row.cells
            if not cells:
                row_idx += 1
                continue
            
            # 取得該行第一格的文字 (去空白)
            c0_text = cells[0].text.replace(" ", "").strip()
            
            # A. 遇到護理站：直接填寫人數
            if c0_text in parsed_data["stations"]:
                vals = parsed_data["stations"][c0_text]
                if len(cells) > 3:
                    cells[1].text = str(vals[0])
                    cells[2].text = str(vals[1])
                    cells[3].text = str(vals[2])
            
            # B. 遇到病人表頭：開啟對應的填寫模式
            elif "姓名" in c0_text and len(cells) > 1 and "病歷" in cells[1].text.replace(" ", ""):
                header_full = "".join([c.text for c in cells]).replace(" ", "")
                if "燈號" in header_full or "危險" in header_full:
                    fill_mode = "new"
                    p_idx = 0
                elif "動態" in header_full or "出院" in header_full:
                    fill_mode = "out"
                    p_idx = 0
                row_idx += 1
                continue
            
            # C. 遇到下一個區塊：關閉填寫模式
            if fill_mode and ("出院病人" in c0_text or "危險評估" in c0_text or "病房特殊" in c0_text):
                fill_mode = None
            
            # D. 正在填寫病人資料
            if fill_mode:
                target_list = parsed_data["new_patients"] if fill_mode == "new" else parsed_data["out_patients"]
                
                if p_idx < len(target_list):
                    # 把病人資料寫進去
                    p_data = target_list[p_idx]
                    cells[0].text = p_data[0] # 姓名
                    cells[1].text = p_data[1] # 病歷號
                    cells[2].text = p_data[2] # 床號
                    cells[3].text = p_data[3] # 性別
                    cells[4].text = p_data[4] # 年齡
                    cells[5].text = p_data[5] # 診斷
                    
                    if fill_mode == "new" and len(p_data) > 6 and len(cells) > 7:
                        cells[7].text = p_data[6] # 燈號
                    elif fill_mode == "out" and len(p_data) > 6 and len(cells) > 6:
                        cells[6].text = p_data[6] # 動態
                    
                    p_idx += 1
                    row_idx += 1
                else:
                    # 病人已經填完了，如果這是一行預留的空白行 -> 直接刪除瘦身！
                    if not c0_text or c0_text == "":
                        row._element.getparent().remove(row._element)
                        # 刪除後，下一行會自動補上來，所以 row_idx "不要" +1
                    else:
                        fill_mode = None
                        row_idx += 1
            else:
                row_idx += 1

    # --- 3. 填寫交班事項 (極致壓縮排版防爆頁) ---
    sorted_handovers = sorted(handovers, key=lambda x: (not x['is_er'], x['time_occurred']))
    h_text = ""
    for h in sorted_handovers:
        # 將資料濃縮成兩行，大幅節省高度
        h_text += f"\n【{'🚨ER ' if h['is_er'] else ''}{h['name']}】{h['time_occurred']} | {h['age']}歲/{h['gender']} | 病歷:{h['med_record']}({h['attending_doc']})\n"
        h_text += f"交班：{h['content']}"

    # 尋找填寫位置
    inserted = False
    for p in doc.paragraphs:
        if "病房特殊狀況及處理" in p.text.replace(" ", ""):
            p.add_run(h_text)
            inserted = True
            break
            
    # 如果段落在表格裡面，就去表格裡面找
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
