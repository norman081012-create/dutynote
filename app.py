import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_BREAK
import copy
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

st.title("🏥 醫師病房值班日誌自動生成器 (終極 XML 複製版)")

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

# ================= 區塊 4：底層 XML 複製與填寫引擎 =================
def safe_fill_cell(cell, text):
    """安全填入文字並設定 10pt 大小防撐破表格"""
    if text is None or text == "": return
    # 清空原有段落
    for p in cell.paragraphs:
        p.text = ""
    # 寫入新段落
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    run = p.add_run(str(text))
    run.font.size = Pt(10)

def clear_row_cells(row):
    """清空一整列的內容 (用於複製出來的新列)"""
    for cell in row.cells:
        for p in cell.paragraphs:
            p.text = ""

def process_data(raw_text, handovers, selected_date):
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"找不到 {TEMPLATE_PATH}。請確認已將樣板放在同一資料夾中。")
    
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

    # --- 1. 文字解析引擎 ---
    parsed_stations = {}
    parsed_new = []
    parsed_out = []
    
    if raw_text:
        for line in raw_text.splitlines():
            line = line.strip()
            if not line: continue
            
            parts = [p.strip() for p in line.split('\t')]
            if len(parts) < 2:
                parts = [p.strip() for p in re.split(r'\s{2,}', line)]
            
            row_str = "".join(parts).replace(" ", "")
            
            # --- 危險評估靜默迴避 ---
            if "危險評估" in row_str or "自殺顧慮" in row_str: continue
                
            if "護理站" in row_str or "急診護理站" in row_str: 
                st_name = parts[0].replace(" ", "")
                if st_name in ["急診護理站", "二樓護理站", "三樓護理站", "四樓護理站", "五樓護理站", "總人數"] and len(parts) >= 4:
                    parsed_stations[st_name] = parts[1:4]
            elif len(parts) >= 5 and "姓名" not in parts[0] and "病患" not in parts[0]:
                if len(parts) >= 7 and ("紅" in row_str or "黃" in row_str or "綠" in row_str or len(parts[6]) < 4):
                    # 簡單判斷：有燈號特徵的歸入新病人
                    parsed_new.append(parts)
                else:
                    # 有出院動態長文字特徵的歸入出院病人
                    parsed_out.append(parts)

    # --- 2. Word 原地 XML 複製與填表 ---
    for table in doc.tables:
        i = 0
        while i < len(table.rows):
            row = table.rows[i]
            if not row.cells: 
                i += 1
                continue
                
            c0_text = row.cells[0].text.replace(" ", "").strip()
            row_text = "".join([c.text for c in row.cells]).replace(" ", "")
            
            # A. 護理站填寫
            if c0_text in ["急診護理站", "二樓護理站", "三樓護理站", "四樓護理站", "五樓護理站", "總人數"]:
                if c0_text in parsed_stations and len(row.cells) >= 4:
                    nums = parsed_stations[c0_text]
                    safe_fill_cell(row.cells[1], nums[0])
                    safe_fill_cell(row.cells[2], nums[1])
                    safe_fill_cell(row.cells[3], nums[2])
                i += 1
                continue
            
            # B. 新住院病人處理
            if "姓名" in row_text and "燈號" in row_text:
                i += 1 # 跳到標題的下一行 (第一行空白欄位)
                for p_data in parsed_new:
                    current_row = table.rows[i]
                    c_check = current_row.cells[0].text.replace(" ", "").strip()
                    
                    # 如果這行是空的，直接寫入
                    if not c_check or c_check == "_" or c_check == "0":
                        safe_fill_cell(current_row.cells[0], p_data[0])
                        if len(current_row.cells) > 1 and len(p_data) > 1: safe_fill_cell(current_row.cells[1], p_data[1])
                        if len(current_row.cells) > 2 and len(p_data) > 2: safe_fill_cell(current_row.cells[2], p_data[2])
                        if len(current_row.cells) > 3 and len(p_data) > 3: safe_fill_cell(current_row.cells[3], p_data[3])
                        if len(current_row.cells) > 4 and len(p_data) > 4: safe_fill_cell(current_row.cells[4], p_data[4])
                        if len(current_row.cells) > 5 and len(p_data) > 5: safe_fill_cell(current_row.cells[5], p_data[5])
                        if len(current_row.cells) > 7 and len(p_data) > 6: safe_fill_cell(current_row.cells[7], p_data[6])
                        i += 1
                    else:
                        # 空間不夠了！直接複製上一行 (XML Deepcopy) 插入
                        prev_row = table.rows[i-1]
                        new_tr = copy.deepcopy(prev_row._tr)
                        prev_row._tr.addnext(new_tr)
                        
                        # 抓取剛複製出來的新列
                        new_row = table.rows[i]
                        clear_row_cells(new_row) # 清除複製過來的舊資料
                        
                        # 填寫
                        safe_fill_cell(new_row.cells[0], p_data[0])
                        if len(new_row.cells) > 1 and len(p_data) > 1: safe_fill_cell(new_row.cells[1], p_data[1])
                        if len(new_row.cells) > 2 and len(p_data) > 2: safe_fill_cell(new_row.cells[2], p_data[2])
                        if len(new_row.cells) > 3 and len(p_data) > 3: safe_fill_cell(new_row.cells[3], p_data[3])
                        if len(new_row.cells) > 4 and len(p_data) > 4: safe_fill_cell(new_row.cells[4], p_data[4])
                        if len(new_row.cells) > 5 and len(p_data) > 5: safe_fill_cell(new_row.cells[5], p_data[5])
                        if len(new_row.cells) > 7 and len(p_data) > 6: safe_fill_cell(new_row.cells[7], p_data[6])
                        i += 1
                continue

            # C. 出院病人處理
            if "姓名" in row_text and "動態" in row_text:
                i += 1
                for p_data in parsed_out:
                    current_row = table.rows[i]
                    c_check = current_row.cells[0].text.replace(" ", "").strip()
                    
                    if not c_check or c_check == "_" or c_check == "0":
                        safe_fill_cell(current_row.cells[0], p_data[0])
                        if len(current_row.cells) > 1 and len(p_data) > 1: safe_fill_cell(current_row.cells[1], p_data[1])
                        if len(current_row.cells) > 2 and len(p_data) > 2: safe_fill_cell(current_row.cells[2], p_data[2])
                        if len(current_row.cells) > 3 and len(p_data) > 3: safe_fill_cell(current_row.cells[3], p_data[3])
                        if len(current_row.cells) > 4 and len(p_data) > 4: safe_fill_cell(current_row.cells[4], p_data[4])
                        if len(current_row.cells) > 5 and len(p_data) > 5: safe_fill_cell(current_row.cells[5], p_data[5])
                        if len(current_row.cells) > 6 and len(p_data) > 6: safe_fill_cell(current_row.cells[6], p_data[6])
                        i += 1
                    else:
                        prev_row = table.rows[i-1]
                        new_tr = copy.deepcopy(prev_row._tr)
                        prev_row._tr.addnext(new_tr)
                        
                        new_row = table.rows[i]
                        clear_row_cells(new_row)
                        
                        safe_fill_cell(new_row.cells[0], p_data[0])
                        if len(new_row.cells) > 1 and len(p_data) > 1: safe_fill_cell(new_row.cells[1], p_data[1])
                        if len(new_row.cells) > 2 and len(p_data) > 2: safe_fill_cell(new_row.cells[2], p_data[2])
                        if len(new_row.cells) > 3 and len(p_data) > 3: safe_fill_cell(new_row.cells[3], p_data[3])
                        if len(new_row.cells) > 4 and len(p_data) > 4: safe_fill_cell(new_row.cells[4], p_data[4])
                        if len(new_row.cells) > 5 and len(p_data) > 5: safe_fill_cell(new_row.cells[5], p_data[5])
                        if len(new_row.cells) > 6 and len(p_data) > 6: safe_fill_cell(new_row.cells[6], p_data[6])
                        i += 1
                continue

            i += 1 # 移至下一列繼續檢查

    # --- 3. 填寫交班事項 (並強制換頁) ---
    sorted_handovers = sorted(handovers, key=lambda x: (not x['is_er'], x['time_occurred']))
    h_text = ""
    for h in sorted_handovers:
        h_text += f"【{'🚨ER ' if h['is_er'] else ''}{h['name']}】{h['time_occurred']} | {h['age']}歲/{h['gender']} | 病歷:{h['med_record']}({h['attending_doc']})\n"
        h_text += f"交班：{h['content']}\n"
        h_text += "-"*30 + "\n"

    inserted = False
    
    # 在表格內尋找
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if "病房特殊狀況及處理" in p.text.replace(" ", ""):
                        # 強制加入分頁符號，把這行推到第二頁
                        p.insert_paragraph_before().add_run().add_break(WD_BREAK.PAGE)
                        p.add_run("\n" + h_text)
                        inserted = True
                        break
                if inserted: break
            if inserted: break
        if inserted: break
        
    # 如果不在表格內
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
