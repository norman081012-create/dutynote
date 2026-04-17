import streamlit as st
import pandas as pd
import docx
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_BREAK
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

st.title("🏥 醫師病房值班日誌自動生成器 (原廠排版鎖定版)")

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

# ================= 區塊 4：無痕對位與精準填表引擎 =================
def safe_fill_cell(cell, text):
    """安全填寫儲存格：縮小字體以防撐破表格"""
    if not text: return
    # 清空原本儲存格內的隱藏空白
    p = cell.paragraphs[0]
    p.text = ""
    # 建立新文字並強制縮小字體 (10pt 幾乎能塞進所有小格子)
    run = p.add_run(str(text))
    run.font.size = Pt(10)

def process_data(raw_text, handovers, selected_date):
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"找不到 {TEMPLATE_PATH}。請確認已上傳樣板。")
    
    doc = Document(TEMPLATE_PATH)
    
    # --- 0. 日期無痕替換 ---
    roc_year = selected_date.year - 1911
    date_str = f"日期： {roc_year} 年 {selected_date.month:02d} 月 {selected_date.day:02d} 日"
    
    for p in doc.paragraphs:
        if "日期" in p.text.replace(" ", ""): p.text = date_str
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if "日期" in p.text.replace(" ", ""): p.text = date_str

    # --- 1. 文字解析引擎 (遇到危險評估直接忽略) ---
    parsed_stations = {}
    parsed_new = []
    parsed_out = []
    
    if raw_text:
        lines = raw_text.splitlines()
        current_section = None
        for line in lines:
            line = line.strip()
            if not line: continue
            
            parts = [p.strip() for p in line.split('\t')]
            if len(parts) < 2:
                parts = [p.strip() for p in re.split(r'\s{2,}', line)]
            
            row_str = "".join(parts).replace(" ", "")
            
            # --- 嚴格控制區塊切換 ---
            if "危險評估" in row_str or "自殺顧慮" in row_str: 
                current_section = "ignore" # 啟動無視模式
                continue
                
            if "護理站" in row_str or "急診護理站" in row_str: current_section = "station"
            elif "新入院" in row_str or "新住院" in row_str: current_section = "new"
            elif "出院病人" in row_str: current_section = "out"
            
            # --- 裝載資料 ---
            if current_section == "station":
                st_name = parts[0].replace(" ", "")
                if st_name in ["急診護理站", "二樓護理站", "三樓護理站", "四樓護理站", "五樓護理站", "總人數"]:
                    if len(parts) >= 4:
                        parsed_stations[st_name] = parts[1:4] 
            elif current_section == "new":
                if "姓名" in parts[0] or "病患" in parts[0] or "新住院" in parts[0]: continue
                if len(parts) >= 5: parsed_new.append(parts)
            elif current_section == "out":
                if "姓名" in parts[0] or "病患" in parts[0] or "出院" in parts[0]: continue
                if len(parts) >= 5: parsed_out.append(parts)

    # --- 2. Word 無痕對號入座 (利用 safe_fill_cell 防撐破) ---
    new_idx = 0
    out_idx = 0
    
    for table in doc.tables:
        fill_mode = None
        for row in table.rows:
            cells = row.cells
            if not cells: continue
            
            c0_text = cells[0].text.replace(" ", "").replace("　", "").replace("\xa0", "").strip()
            row_text_concat = "".join([c.text for c in cells]).replace(" ", "").replace("　", "").replace("\xa0", "")
            
            # A. 護理站人數
            if c0_text in ["急診護理站", "二樓護理站", "三樓護理站", "四樓護理站", "五樓護理站", "總人數"]:
                if c0_text in parsed_stations and len(cells) >= 4:
                    nums = parsed_stations[c0_text]
                    safe_fill_cell(cells[1], nums[0])
                    safe_fill_cell(cells[2], nums[1])
                    safe_fill_cell(cells[3], nums[2])
                continue
            
            # B. 判斷是否為病人表頭
            if "姓名" in row_text_concat and "病歷" in row_text_concat:
                if "燈號" in row_text_concat or "危險" in row_text_concat:
                    fill_mode = "new"
                elif "動態" in row_text_concat or "出院" in row_text_concat:
                    fill_mode = "out"
                continue
            
            # C. 遇到無關段落或危險評估，關閉填寫模式
            if fill_mode and ("出院病人" in c0_text or "危險評估" in c0_text or "自殺顧慮" in c0_text or "病房特殊" in c0_text):
                fill_mode = None
            
            is_empty_cell = (c0_text == "" or c0_text == "_" or c0_text == "0")
            
            # D. 填入病人
            if fill_mode == "new" and is_empty_cell:
                if new_idx < len(parsed_new):
                    p_data = parsed_new[new_idx]
                    safe_fill_cell(cells[0], p_data[0])
                    if len(cells) > 1 and len(p_data) > 1: safe_fill_cell(cells[1], p_data[1])
                    if len(cells) > 2 and len(p_data) > 2: safe_fill_cell(cells[2], p_data[2])
                    if len(cells) > 3 and len(p_data) > 3: safe_fill_cell(cells[3], p_data[3])
                    if len(cells) > 4 and len(p_data) > 4: safe_fill_cell(cells[4], p_data[4])
                    if len(cells) > 5 and len(p_data) > 5: safe_fill_cell(cells[5], p_data[5])
                    if len(cells) > 7 and len(p_data) > 6: safe_fill_cell(cells[7], p_data[6])
                    new_idx += 1
            
            elif fill_mode == "out" and is_empty_cell:
                if out_idx < len(parsed_out):
                    p_data = parsed_out[out_idx]
                    safe_fill_cell(cells[0], p_data[0])
                    if len(cells) > 1 and len(p_data) > 1: safe_fill_cell(cells[1], p_data[1])
                    if len(cells) > 2 and len(p_data) > 2: safe_fill_cell(cells[2], p_data[2])
                    if len(cells) > 3 and len(p_data) > 3: safe_fill_cell(cells[3], p_data[3])
                    if len(cells) > 4 and len(p_data) > 4: safe_fill_cell(cells[4], p_data[4])
                    if len(cells) > 5 and len(p_data) > 5: safe_fill_cell(cells[5], p_data[5])
                    if len(cells) > 6 and len(p_data) > 6: safe_fill_cell(cells[6], p_data[6])
                    out_idx += 1

    # --- 3. 填寫交班事項 (並強制換頁) ---
    sorted_handovers = sorted(handovers, key=lambda x: (not x['is_er'], x['time_occurred']))
    h_text = ""
    for h in sorted_handovers:
        h_text += f"【{'🚨ER ' if h['is_er'] else ''}{h['name']}】{h['time_occurred']} | {h['age']}歲/{h['gender']} | 病歷:{h['med_record']}({h['attending_doc']})\n"
        h_text += f"交班：{h['content']}\n"
        h_text += "-"*30 + "\n"

    inserted = False
    
    # 搜尋是否在表格內
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if "病房特殊狀況及處理" in p.text.replace(" ", ""):
                        # 強制加入「分頁符號」，確保文字出現在第二頁開頭
                        p.insert_paragraph_before().add_run().add_break(WD_BREAK.PAGE)
                        p.add_run("\n" + h_text)
                        inserted = True
                        break
                if inserted: break
            if inserted: break
        if inserted: break
        
    # 如果不在表格內，而在普通段落
    if not inserted:
        for p in doc.paragraphs:
            if "病房特殊狀況及處理" in p.text.replace(" ", ""):
                p.insert_paragraph_before().add_run().add_break(WD_BREAK.PAGE)
                p.add_run("\n" + h_text)
                break

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
