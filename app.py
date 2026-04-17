import streamlit as st
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_BREAK, WD_ALIGN_PARAGRAPH
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

st.title("🏥 醫師病房值班日誌自動生成器 (終極混合排版版)")

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
            
            if "危險評估" in row_str or "自殺顧慮" in row_str: continue
            
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

# ================= 區塊 4：終極混合填表引擎 =================
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
            for cell in row.cells:
                for p in cell.paragraphs:
                    if "日期" in p.text.replace(" ", ""): p.text = date_str

    # --- 1. 護理站終極解法：刪除原表，就地重建 ---
    station_table_idx = -1
    for idx, table in enumerate(doc.tables):
        for row in table.rows:
            if not row.cells: continue
            row_text_all = "".join([c.text for c in row.cells]).replace(" ", "").replace("　", "").replace("\xa0", "")
            # 找到標題列包含男、女、總數的表格
            if "男" in row_text_all and "女" in row_text_all and "病人總數" in row_text_all and "急診護理站" in row_text_all:
                station_table_idx = idx
                break
        if station_table_idx != -1: break

    if station_table_idx != -1:
        # A. 獲取原表格
        old_table = doc.tables[station_table_idx]
        
        # B. 建立新表格並直接插入到舊表格的上方
        new_table = doc.add_table(rows=7, cols=4)
        new_table.style = 'Table Grid'
        
        # 移動新表格的位置
        tbl_element = new_table._tbl
        old_table._tbl.addprevious(tbl_element)
        
        # C. 填寫新表格標題
        headers = ["護理站", "男", "女", "病人總數"]
        for i, h in enumerate(headers):
            new_table.cell(0, i).text = h
            new_table.cell(0, i).paragraphs[0].runs[0].font.size = Pt(11)
            
        # D. 填寫各護理站數據
        station_names = ["急診護理站", "二樓護理站", "三樓護理站", "四樓護理站", "五樓護理站", "總人數"]
        for row_idx, st_name in enumerate(station_names, start=1):
            new_table.cell(row_idx, 0).text = st_name
            if st_name in parsed_stations:
                nums = parsed_stations[st_name]
                new_table.cell(row_idx, 1).text = str(nums[0])
                new_table.cell(row_idx, 2).text = str(nums[1])
                new_table.cell(row_idx, 3).text = str(nums[2])
            else:
                new_table.cell(row_idx, 1).text = ""
                new_table.cell(row_idx, 2).text = ""
                new_table.cell(row_idx, 3).text = ""
                
        # E. 刪除舊的、有問題的護理站表格
        old_table._element.getparent().remove(old_table._element)

    # --- 2. 病人資料：無痕對位填寫 ---
    new_idx = 0
    out_idx = 0
    
    for table in doc.tables:
        fill_mode = None
        name_col_idx = 0 
        
        for row in table.rows:
            if not row.cells: continue
                
            row_text_all = "".join([c.text for c in row.cells]).replace(" ", "").replace("　", "").replace("\xa0", "")
            
            # A. 啟動病人填寫模式並鎖定欄位
            if "姓名" in row_text_all and "病歷" in row_text_all:
                if "燈號" in row_text_all or "強制" in row_text_all:
                    fill_mode = "new"
                elif "動態" in row_text_all or "出院" in row_text_all:
                    fill_mode = "out"
                
                for idx, cell in enumerate(row.cells):
                    clean_cell = re.sub(r'[\r\n\t]', '', cell.text.replace(" ", "").replace("　", ""))
                    if "姓名" in clean_cell:
                        name_col_idx = idx
                        break
                continue
                
            # B. 關閉填寫模式
            if fill_mode and ("出院病人" in row_text_all or "危險評估" in row_text_all or "自殺顧慮" in row_text_all or "病房特殊" in row_text_all):
                fill_mode = None
                
            # 判斷這格是不是原廠預留的空行
            if fill_mode and name_col_idx < len(row.cells):
                c_name = row.cells[name_col_idx].text.replace(" ", "").replace("_", "").replace("0", "").strip()
                is_empty_cell = (c_name == "")
                
                # C. 填入新住院病人
                if fill_mode == "new" and is_empty_cell and new_idx < len(parsed_new):
                    p_data = parsed_new[new_idx]
                    safe_fill_cell(row.cells[name_col_idx], p_data[0])
                    if len(row.cells) > name_col_idx+1 and len(p_data) > 1: safe_fill_cell(row.cells[name_col_idx+1], p_data[1])
                    if len(row.cells) > name_col_idx+2 and len(p_data) > 2: safe_fill_cell(row.cells[name_col_idx+2], p_data[2])
                    if len(row.cells) > name_col_idx+3 and len(p_data) > 3: safe_fill_cell(row.cells[name_col_idx+3], p_data[3])
                    if len(row.cells) > name_col_idx+4 and len(p_data) > 4: safe_fill_cell(row.cells[name_col_idx+4], p_data[4])
                    if len(row.cells) > name_col_idx+5 and len(p_data) > 5: safe_fill_cell(row.cells[name_col_idx+5], p_data[5])
                    if len(row.cells) > name_col_idx+7 and len(p_data) > 6: safe_fill_cell(row.cells[name_col_idx+7], p_data[6])
                    new_idx += 1
                        
                # D. 填入出院病人
                elif fill_mode == "out" and is_empty_cell and out_idx < len(parsed_out):
                    p_data = parsed_out[out_idx]
                    safe_fill_cell(row.cells[name_col_idx], p_data[0])
                    if len(row.cells) > name_col_idx+1 and len(p_data) > 1: safe_fill_cell(row.cells[name_col_idx+1], p_data[1])
                    if len(row.cells) > name_col_idx+2 and len(p_data) > 2: safe_fill_cell(row.cells[name_col_idx+2], p_data[2])
                    if len(row.cells) > name_col_idx+3 and len(p_data) > 3: safe_fill_cell(row.cells[name_col_idx+3], p_data[3])
                    if len(row.cells) > name_col_idx+4 and len(p_data) > 4: safe_fill_cell(row.cells[name_col_idx+4], p_data[4])
                    if len(row.cells) > name_col_idx+5 and len(p_data) > 5: safe_fill_cell(row.cells[name_col_idx+5], p_data[5])
                    if len(row.cells) > name_col_idx+6 and len(p_data) > 6: safe_fill_cell(row.cells[name_col_idx+6], p_data[6])
                    out_idx += 1

    # --- 3. 填寫交班事項 (強制分頁至第二頁) ---
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
            for cell in row.cells:
                for p in cell.paragraphs:
                    if "病房特殊狀況及處理" in p.text.replace(" ", ""):
                        # 強制分頁，確保內容在第二頁
                        p.insert_paragraph_before().add_run().add_break(WD_BREAK.PAGE)
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
                p.insert_paragraph_before().add_run().add_break(WD_BREAK.PAGE)
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
        st.success("檔案已更新並備妥！(已完美套用原廠樣板與防呆機制)")
        st.download_button(
            label="📥 點擊下載最新版值班日誌",
            data=final_file,
            file_name=f"值班日誌_{duty_date.strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"錯誤詳情: {e}")
