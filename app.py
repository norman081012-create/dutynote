import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from copy import deepcopy
import unicodedata
import io
import json
import os
import re
import datetime
from datetime import timezone, timedelta

# 設定台灣時區 (UTC+8)，解決雲端伺服器時間誤差問題
tw_tz = timezone(timedelta(hours=8))

st.set_page_config(page_title="值班日誌自動生成器", layout="wide")

TEMPLATE_PATH = "template.docx"
DB_FILE = "handovers.json"

ATTENDING_DOCS = ["", "鍾偉倫", "張志華", "成毓賢", "劉俊麟", "謝金村"]
DIAG_CHOICES = ["", "Schizophrenia", "bipolar", "depression", "其他 (請於下方輸入)"]

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

# 【修改點 3】移除 "(預覽與防呆 V11)"
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

# 取得台灣當下時間
now_tw = datetime.datetime.now(tw_tz)

with col_date:
    # 【修改點 1】使用台灣時間的日期
    duty_date = st.date_input("📅 選擇值班日期", now_tw.date())

def parse_his_data(raw_text):
    parsed_stations = {}
    parsed_new = []
    parsed_out = []
    if raw_text:
        for line in raw_text.splitlines():
            line = line.strip()
            if not line: continue
            parts = [p.strip() for p in re.split(r'\t|\s{2,}', line)]
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
    raw_text_input = st.text_area("📝 貼上 HIS 內容", height=150, key=f"text_input_{st.session_state.uploader_key}")
    parsed_stations, parsed_new, parsed_out = parse_his_data(raw_text_input)

# ================= 區塊 2：交班事項登錄表單 =================
# 【修改點 3】移除 "(完美流水帳格式)"
st.header("2. 交班事項登錄")
with st.form("handover_form", clear_on_submit=True):
    c1, c2 = st.columns(2)
    with c1:
        location = st.selectbox("單位/病房 (預設此)", ["病房", "急診", "二樓病房", "三樓病房", "四樓病房", "五樓病房"])
        name = st.text_input("病人姓名 (必填)")
        age = st.text_input("年紀")
        gender = st.selectbox("性別", ["", "男", "女"])
        med_record = st.text_input("病歷號")
        history_input = st.text_area("內外科病史輸入", height=60)
        
    with c2:
        # 【修改點 1】使用台灣時間的時間
        time_occurred = st.time_input("狀況發生時間", value=now_tw.time())
        attending_doc = st.selectbox("主治醫師", ATTENDING_DOCS)
        diag_choice = st.selectbox("診斷快速選項", DIAG_CHOICES)
        diag_manual = st.text_input("手動輸入診斷 (若選其他)")
        
    content = st.text_area("交班內容 (必填)")
    
    if st.form_submit_button("確認新增交班"):
        if not name or not content:
            st.error("「姓名」與「內容」為必填！")
        else:
            diag_final = diag_manual if not diag_choice or diag_choice == "其他 (請於下方輸入)" else diag_choice
            st.session_state.handovers.append({
                "location": location, "name": name, "age": age, "gender": gender,
                "med_record": med_record, "attending_doc": attending_doc,
                "time_occurred": time_occurred.strftime("%H:%M"), "content": content,
                "diagnosis": diag_final, "history": history_input,
                "is_er": (location == "急診") 
            })
            save_handovers(st.session_state.handovers)
            st.rerun()

# ================= 區塊 3：已登錄交班預覽 =================
st.header("3. 已登錄交班事項")
if st.session_state.handovers:
    sorted_view = sorted(st.session_state.handovers, key=lambda x: (x.get('location') != '急診', x.get('time_occurred')))
    for h in sorted_view:
        idx = st.session_state.handovers.index(h)
        h_age_disp = h['age'] if h.get('age') else "?"
        h_gen_disp = f"{h['gender']}性" if h.get('gender') else ""
        
        with st.expander(f"[{h['location']}] {h['name']} ({h_age_disp}歲{h_gen_disp}) - {h['time_occurred']}"):
            st.write(f"主治：{h['attending_doc']} | 病史：{h['history']} | 診斷：{h['diagnosis']}")
            st.write(f"內容：{h['content']}")
            if st.button(f"刪除 {h['name']}", key=f"del_{idx}"):
                st.session_state.handovers.pop(idx)
                save_handovers(st.session_state.handovers)
                st.rerun()

# ================= 核心工具函數 =================
def get_unique_cells(row):
    unique_cells = []
    for cell in row.cells:
        if cell not in unique_cells: unique_cells.append(cell)
    return unique_cells

def safe_fill_cell(cell, text, font_size=12):
    if text is None: text = ""
    for p in cell.paragraphs: p.text = "" 
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    run = p.add_run(str(text).strip())
    run.font.size = Pt(font_size)
    run.bold = False
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.left_indent = Pt(0)
    p.paragraph_format.first_line_indent = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)

def get_text_width(text):
    width = 0
    for char in text:
        if unicodedata.east_asian_width(char) in ('F', 'W', 'A'):
            width += 2
        else: width += 1
    return width

def visual_smart_chunker(text, max_visual_width=78):
    if not text: return []
    tokens = re.findall(r'[a-zA-Z0-9.\-\_]+|.', text)
    chunks = []
    current_chunk = ""
    current_width = 0
    
    for token in tokens:
        token_width = get_text_width(token)
        if current_width + token_width > max_visual_width:
            if current_chunk: chunks.append(current_chunk.strip())
            current_chunk = token.lstrip()
            current_width = get_text_width(current_chunk)
        else:
            current_chunk += token
            current_width += token_width
            
    if current_chunk: chunks.append(current_chunk.strip())
    return chunks

# ================= 區塊 4：預覽與輸出 =================
st.header("4. 預覽與輸出")

# --- 1. 產生預覽文字 (供畫面顯示) ---
preview_lines = []
sorted_h = sorted(st.session_state.handovers, key=lambda x: (x.get('location') != '急診', x.get('time_occurred')))

for h in sorted_h:
    h_loc = h.get('location', '病房')
    h_name = h.get('name', '').strip()
    h_age = h.get('age', '').strip()
    h_gen = h.get('gender', '').strip()
    h_med = h.get('med_record', '').strip()
    h_att = h.get('attending_doc', '').strip()
    h_diag = h.get('diagnosis', '').strip()
    h_his = h.get('history', '').strip()
    h_time = h.get('time_occurred', '').strip()
    h_content = h.get('content', '').replace('\n', ' ').strip()

    h_age_display = h_age if h_age else "?"
    h_gen_display = f"{h_gen}性" if h_gen else ""
    age_gen_part = f"，{h_age_display}歲{h_gen_display}"

    med_part = f"病歷號:{h_med} " if h_med else ""
    pt_part = f"({h_loc}){med_part}姓名:{h_name}{age_gen_part}"
    
    ward_tag = f"({h_loc[0:2]})" if h_loc not in ["急診", "病房"] else ""
    
    # 【修改點 2】解決未填入醫師會殘留病人的問題：如果沒有主治醫師，則該段直接留白
    doc_part = f"{h_att}醫師{ward_tag}病人" if h_att else ""
    
    his_part = f"內外科病史:{h_his}" if h_his else ""
    diag_part = f"診斷:{h_diag}" if h_diag else ""
    time_part = f"{h_time}時" if h_time else ""
    
    diag_time = ""
    if diag_part and time_part: diag_time = f"{diag_part} {time_part}"
    elif diag_part: diag_time = diag_part
    elif time_part: diag_time = time_part
        
    components = [pt_part, doc_part, his_part, diag_time, h_content]
    components = [c for c in components if c.strip()]
    full_line = "，".join(components)
    preview_lines.append(full_line)

# 顯示網頁預覽區
if preview_lines:
    with st.expander("👀 點擊展開：最終交班文字預覽 (與 Word 輸出內容相同)", expanded=True):
        st.info("💡 提示：因雲端伺服器未安裝微軟 Word，無法直接預覽 PDF。請確認下方文字與排版無誤後，下載 Word 檔再自行另存為 PDF。")
        preview_text = "\n\n".join(preview_lines) 
        st.text_area("即將寫入 Word 的文字：", value=preview_text, height=250, disabled=True)

# --- 2. 生成 Word 檔案 ---
def build_word_document(p_stations, p_new, p_out, handovers, selected_date):
    if not os.path.exists(TEMPLATE_PATH): raise FileNotFoundError(f"找不到 {TEMPLATE_PATH}。")
    doc = Document(TEMPLATE_PATH)
    
    # 填寫日期
    roc_year = selected_date.year - 1911
    date_str = f"日期： {roc_year} 年 {selected_date.month:02d} 月 {selected_date.day:02d} 日"
    for p in doc.paragraphs:
        if "日期" in p.text.replace(" ", ""): p.text = date_str
    
    # HIS表格填寫
    new_idx, out_idx = 0, 0
    for table in doc.tables:
        fill_mode, name_col_idx = None, 0
        for row in table.rows:
            u_cells = get_unique_cells(row)
            row_txt = "".join([c.text for c in u_cells]).replace(" ", "")
            
            matched_st = None
            for kn in ["急診護理站", "二樓護理站", "三樓護理站", "四樓護理站", "五樓護理站", "總人數"]:
                if kn in row_txt: matched_st = kn; break
            if matched_st and matched_st in p_stations:
                for idx, c in enumerate(u_cells):
                    clean_cell = re.sub(r'[\r\n\t]', '', c.text.replace(" ", "").replace("　", ""))
                    if matched_st in clean_cell:
                        if idx+3 < len(u_cells):
                            safe_fill_cell(u_cells[idx+1], p_stations[matched_st][0], font_size=10)
                            safe_fill_cell(u_cells[idx+2], p_stations[matched_st][1], font_size=10)
                            safe_fill_cell(u_cells[idx+3], p_stations[matched_st][2], font_size=10)
                continue

            if "姓名" in row_txt and "病歷" in row_txt:
                fill_mode = "new" if ("燈號" in row_txt or "強制" in row_txt) else "out"
                for idx, c in enumerate(u_cells):
                    if "姓名" in c.text.replace(" ",""): name_col_idx = idx; break
                continue

            if fill_mode and name_col_idx < len(u_cells):
                c_name = re.sub(r'[\r\n\t\s_0]', '', u_cells[name_col_idx].text)
                if c_name == "":
                    if fill_mode == "new" and new_idx < len(p_new):
                        pd = p_new[new_idx]
                        for k in range(min(len(pd), len(u_cells))):
                            safe_fill_cell(u_cells[name_col_idx+k if k<6 else name_col_idx+k+1], pd[k], font_size=10)
                        new_idx += 1
                    elif fill_mode == "out" and out_idx < len(p_out):
                        pd = p_out[out_idx]
                        for k in range(min(len(pd), len(u_cells))):
                            safe_fill_cell(u_cells[name_col_idx+k], pd[k], font_size=10)
                        out_idx += 1

    # 危險評估瘦身：保留 4 列
    for table in doc.tables:
        header_row_idx = -1
        for i, row in enumerate(table.rows):
            row_txt = "".join([c.text for c in get_unique_cells(row)]).replace(" ", "")
            if "自殺顧慮" in row_txt and "哽塞顧慮" in row_txt:
                header_row_idx = i; break
        if header_row_idx != -1:
            target_row_count = header_row_idx + 5
            while len(table.rows) > target_row_count:
                row_to_del = table.rows[-1]
                row_to_del._element.getparent().remove(row_to_del._element)
            break

    # 交班字串組裝與斷行
    all_chunks_to_fill = []
    for i, line in enumerate(preview_lines):
        chunks = visual_smart_chunker(line, max_visual_width=78)
        all_chunks_to_fill.extend(chunks)
        if i < len(preview_lines) - 1:
            all_chunks_to_fill.append("")

    # 逐列填入表格
    target_table = None
    start_row_idx = -1
    discuss_row_idx = -1
    
    for table in doc.tables:
        for idx, row in enumerate(table.rows):
            u_cells = get_unique_cells(row)
            if not u_cells: continue
            row_txt = u_cells[0].text.replace(" ", "")
            if "病房特殊狀況及處理" in row_txt:
                target_table = table; start_row_idx = idx + 1 
            if "討論與講評" in row_txt and start_row_idx != -1:
                discuss_row_idx = idx; break
        if target_table and discuss_row_idx != -1: break

    if target_table and start_row_idx != -1 and discuss_row_idx != -1:
        current_row_idx = start_row_idx
        for chunk_text in all_chunks_to_fill:
            if current_row_idx < discuss_row_idx:
                target_cell = get_unique_cells(target_table.rows[current_row_idx])[0]
                safe_fill_cell(target_cell, chunk_text, font_size=12)
                current_row_idx += 1
            else:
                ref_row = target_table.rows[discuss_row_idx]
                blank_tr = deepcopy(target_table.rows[discuss_row_idx - 1]._tr)
                ref_row._tr.addprevious(blank_tr)
                discuss_row_idx += 1
                target_cell = get_unique_cells(target_table.rows[current_row_idx])[0]
                safe_fill_cell(target_cell, chunk_text, font_size=12)
                current_row_idx += 1
                
        while current_row_idx < discuss_row_idx:
            target_cell = get_unique_cells(target_table.rows[current_row_idx])[0]
            safe_fill_cell(target_cell, "", font_size=12)
            current_row_idx += 1

    stream = io.BytesIO(); doc.save(stream); stream.seek(0)
    return stream

if st.button("🚀 生成下載 Word", type="primary"):
    try:
        f = build_word_document(parsed_stations, parsed_new, parsed_out, st.session_state.handovers, duty_date)
        # 【修改點 3】移除後方贅字
        st.success("✅ 檔案已更新並備妥！")
        st.download_button("📥 點擊下載", f, f"值班日誌_{duty_date.strftime('%Y%m%d')}.docx")
    except Exception as e:
        st.error(f"錯誤: {e}")
