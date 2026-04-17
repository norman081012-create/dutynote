import streamlit as st
from docx import Document
from docx.shared import Pt
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

st.title("🏥 醫師病房值班日誌自動生成器 (絕對不動格式版)")

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

# === 即時解析引擎 ===
def parse_his_data(raw_text):
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
            
            # 迴避危險評估區塊
            if "危險評估" in row_str or "自殺顧慮" in row_str: continue
            
            # 抓取護理站與總人數
            matched_station = False
            for key_name in ["急診護理站", "二樓護理站", "三樓護理站", "四樓護理站", "五樓護理站", "總人數"]:
                if key_name in row_str and len(parts) >= 4:
                    for idx, p in enumerate(parts):
                        if key_name in p.replace(" ", ""):
                            if idx + 3 < len(parts):
                                parsed_stations[key_name] = parts[idx+1 : idx+4]
                            matched_station = True
                            break
                    if matched_station: break
            if matched_station: continue
            
            # 抓取病患
            if len(parts) >= 5 and "姓名" not in row_str and "病患" not in row_str:
                if len(parts) >= 7 and ("紅" in row_str or "黃" in row_str or "綠" in row_str or len(parts[6]) < 4):
                    parsed_new.append(parts)
                else:
                    parsed_out.append(parts)
    return parsed_stations, parsed_new, parsed_out

with col_text:
    raw_text_input = st.text_area(
        "📝 在此貼上資料 (從 Excel 複製包含護理站與病人的區塊)", 
        height=200, 
        key=f"text_input_{st.session_state.uploader_key}"
    )
    
    parsed_stations, parsed_new, parsed_out = parse_his_data(raw_text_input)
    if parsed_stations or parsed_new or parsed_out:
        with st.expander("👀 點擊查看系統成功抓取的 HIS 資料 (確保沒漏抓)"):
            st.write(f"✅ 成功抓取 **{len(parsed_stations)}** 個護理站/總人數數據、**{len(parsed_new)}** 位新住院、**{len(parsed_out)}** 位出院")
            st.json({"護理站統計": parsed_stations, "新住院病人": parsed_new, "出院病人": parsed_out})

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
    sorted_view = sorted(st.session_state.handovers, key=lambda x: (not x.get('is_er', False), x.get('time_occurred', x.get('time', ''))))
    for h in sorted_view:
        original_idx = st.session_state.handovers.index(h) 
        h_er = h.get('is_er', False)
        h_name = h.get('name', '')
        h_time = h.get('time_occurred', h.get('time', ''))
        h_age = h.get('age', '')
        h_gender = h.get('gender', '')
        h_med = h.get('med_record', '')
        h_att = h.get('attending_doc', h.get('attending', ''))
        h_content = h.get('content', '')
        
        title = f"{'🚨[ER] ' if h_er else ''}{h_name} - {h_time}"
        with st.expander(title):
            st.markdown(f"**資料：** {h_age}歲/{h_gender} | 病歷：{h_med} | 主治：{h_att}")
            st.markdown(f"**交班內容：**\n{h_content}")
            if st.button(f"刪除 {h_name}", key=f"del_{original_idx}"):
                st.session_state.handovers.pop(original_idx)
                save_handovers(st.session_state.handovers)
                st.rerun()

# ================= 區塊 4：純文字無痕對位填表引擎 =================
def get_unique_cells(row):
    """【核心修復】過濾掉 Word 中的合併儲存格，只抓取肉眼可見的獨立格子"""
    unique_cells = []
    for cell in row.cells:
        if cell not in unique_cells:
            unique_cells.append(cell)
    return unique_cells

def safe_fill_cell(cell, text):
    if text is None or text == "": return
    for p in cell.paragraphs: p.text = "" 
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    run = p.add_run(str(text))
    run.font.size = Pt(10)

def build_word_document(parsed_stations, parsed_new, parsed_out, handovers, selected_date):
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
            u_cells = get_unique_cells(row)
            for cell in u_cells:
                for p in cell.paragraphs:
                    if "日期" in p.text.replace(" ", ""): p.text = date_str

    # --- 1. 導彈級自動對位填寫 (絕對不增刪行) ---
    new_idx = 0
    out_idx = 0
    
    for table in doc.tables:
        fill_mode = None
        name_col_idx = 0 
        
        for row in table.rows:
            # 取得真正獨立的格子，避開合併儲存格的陷阱
            u_cells = get_unique_cells(row)
            if not u_cells: continue
                
            row_text_all = "".join([c.text for c in u_cells]).replace(" ", "").replace("　", "").replace("\xa0", "")
            
            # A. 護理站導彈鎖定填寫
            matched_station = None
            for key_name in ["急診護理站", "二樓護理站", "三樓護理站", "四樓護理站", "五樓護理站", "總人數"]:
                if key_name in row_text_all:
                    matched_station = key_name
                    break
                    
            if matched_station and matched_station in parsed_stations:
                target_idx = -1
                for idx, cell in enumerate(u_cells):
                    clean_cell = re.sub(r'[\r\n\t]', '', cell.text.replace(" ", "").replace("　", ""))
                    if matched_station in clean_cell:
                        target_idx = idx
                        break
                # 精準填入右邊三個格子 (男, 女, 總數)
                if target_idx != -1 and target_idx + 3 < len(u_cells):
                    nums = parsed_stations[matched_station]
                    safe_fill_cell(u_cells[target_idx+1], nums[0])
                    safe_fill_cell(u_cells[target_idx+2], nums[1])
                    safe_fill_cell(u_cells[target_idx+3], nums[2])
                continue
            
            # B. 啟動病人填寫模式並鎖定欄位
            if "姓名" in row_text_all and "病歷" in row_text_all:
                if "燈號" in row_text_all or "強制" in row_text_all:
                    fill_mode = "new"
                elif "動態" in row_text_all or "出院" in row_text_all:
                    fill_mode = "out"
                
                for idx, cell in enumerate(u_cells):
                    clean_cell = re.sub(r'[\r\n\t]', '', cell.text.replace(" ", "").replace("　", ""))
                    if "姓名" in clean_cell:
                        name_col_idx = idx
                        break
                continue
                
            # C. 關閉填寫模式
            if fill_mode and ("出院病人" in row_text_all or "危險評估" in row_text_all or "自殺顧慮" in row_text_all or "病房特殊" in row_text_all):
                fill_mode = None
                
            # 判斷這格是不是原廠預留的空行
            if fill_mode and name_col_idx < len(u_cells):
                c_name = u_cells[name_col_idx].text.replace(" ", "").replace("_", "").replace("0", "").strip()
                c_name = re.sub(r'[\r\n\t]', '', c_name)
                is_empty_cell = (c_name == "")
                
                # D. 填入新住院病人 (跳過 index 6 危險評估)
                if fill_mode == "new" and is_empty_cell and new_idx < len(parsed_new):
                    p_data = parsed_new[new_idx]
                    safe_fill_cell(u_cells[name_col_idx], p_data[0])
                    if len(u_cells) > name_col_idx+1 and len(p_data) > 1: safe_fill_cell(u_cells[name_col_idx+1], p_data[1])
                    if len(u_cells) > name_col_idx+2 and len(p_data) > 2: safe_fill_cell(u_cells[name_col_idx+2], p_data[2])
                    if len(u_cells) > name_col_idx+3 and len(p_data) > 3: safe_fill_cell(u_cells[name_col_idx+3], p_data[3])
                    if len(u_cells) > name_col_idx+4 and len(p_data) > 4: safe_fill_cell(u_cells[name_col_idx+4], p_data[4])
                    if len(u_cells) > name_col_idx+5 and len(p_data) > 5: safe_fill_cell(u_cells[name_col_idx+5], p_data[5])
                    if len(u_cells) > name_col_idx+7 and len(p_data) > 6: safe_fill_cell(u_cells[name_col_idx+7], p_data[6]) # 燈號
                    new_idx += 1
                        
                # E. 填入出院病人
                elif fill_mode == "out" and is_empty_cell and out_idx < len(parsed_out):
                    p_data = parsed_out[out_idx]
                    safe_fill_cell(u_cells[name_col_idx], p_data[0])
                    if len(u_cells) > name_col_idx+1 and len(p_data) > 1: safe_fill_cell(u_cells[name_col_idx+1], p_data[1])
                    if len(u_cells) > name_col_idx+2 and len(p_data) > 2: safe_fill_cell(u_cells[name_col_idx+2], p_data[2])
                    if len(u_cells) > name_col_idx+3 and len(p_data) > 3: safe_fill_cell(u_cells[name_col_idx+3], p_data[3])
                    if len(u_cells) > name_col_idx+4 and len(p_data) > 4: safe_fill_cell(u_cells[name_col_idx+4], p_data[4])
                    if len(u_cells) > name_col_idx+5 and len(p_data) > 5: safe_fill_cell(u_cells[name_col_idx+5], p_data[5])
                    if len(u_cells) > name_col_idx+6 and len(p_data) > 6: safe_fill_cell(u_cells[name_col_idx+6], p_data[6]) # 動態
                    out_idx += 1

    # --- 3. 填寫交班事項 ---
    sorted_handovers = sorted(handovers, key=lambda x: (not x.get('is_er', False), x.get('time_occurred', x.get('time', ''))))
    h_text = "\n"
    for h in sorted_handovers:
        h_er = h.get('is_er', False)
        h_name = h.get('name', '')
        h_time = h.get('time_occurred', h.get('time', ''))
        h_age = h.get('age', '')
        h_gender = h.get('gender', '')
        h_med = h.get('med_record', '')
        h_att = h.get('attending_doc', h.get('attending', ''))
        h_content = h.get('content', '')
        
        h_text += f"【{'🚨ER ' if h_er else ''}{h_name}】{h_time} | {h_age}歲/{h_gender} | 病歷:{h_med}({h_att})\n"
        h_text += f"交班：{h_content}\n"
        h_text += "-"*30 + "\n"

    inserted = False
    for table in doc.tables:
        for row in table.rows:
            u_cells = get_unique_cells(row)
            for cell in u_cells:
                for p in cell.paragraphs:
                    if "病房特殊狀況及處理" in p.text.replace(" ", ""):
                        run = p.add_run(h_text)
                        run.font.size = Pt(11)
                        inserted = True
                        break
                if inserted: break
            if inserted: break
        if inserted: break
        
    if not inserted:
        for p in doc.paragraphs:
            if "病房特殊狀況及處理" in p.text.replace(" ", ""):
                run = p.add_run(h_text)
                run.font.size = Pt(11)
                break

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

st.header("4. 確認與輸出")
if st.button("🚀 生成並下載 Word 檔案", type="primary"):
    try:
        final_file = build_word_document(parsed_stations, parsed_new, parsed_out, st.session_state.handovers, duty_date)
        st.success("檔案已更新並備妥！(已啟動儲存格過濾引擎，完美破解合併儲存格問題)")
        st.download_button(
            label="📥 點擊下載最新版值班日誌",
            data=final_file,
            file_name=f"值班日誌_{duty_date.strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"錯誤詳情: {e}")
