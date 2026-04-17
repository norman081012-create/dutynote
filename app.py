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

# 固定主治醫師全名清單
ATTENDING_DOCS = ["", "鍾偉倫", "張志華", "成毓賢", "劉俊麟", "謝金村"]
# 診斷快速選項
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

st.title("🏥 醫師病房值班日誌自動生成器 (標準流水帳 V3)")

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
st.header("2. 交班事項登錄 (完美流水帳格式)")
with st.form("handover_form", clear_on_submit=True):
    c1, c2 = st.columns(2)
    with c1:
        # 修正：病房選項正確預設
        location = st.selectbox("單位/病房 (預設此)", ["病房", "急診", "二樓病房", "三樓病房", "四樓病房", "五樓病房"])
        name = st.text_input("病人姓名 (必填)")
        age = st.text_input("年紀")
        gender = st.selectbox("性別", ["", "男", "女"])
        med_record = st.text_input("病歷號")
        # 增加病史輸入
        history_input = st.text_area("內外科病史輸入", height=60)
        
    with c2:
        # 狀況發生時間預設抓取當時時間
        time_occurred = st.time_input("狀況發生時間", value=datetime.datetime.now().time())
        # 主治醫師全名選項
        attending_doc = st.selectbox("主治醫師", ATTENDING_DOCS)
        # 診斷選項連動
        diag_choice = st.selectbox("診斷快速選項", DIAG_CHOICES)
        diag_manual = st.text_input("手動輸入診斷 (若選其他)")
        
    content = st.text_area("交班內容 (必填)")
    
    if st.form_submit_button("確認新增交班"):
        if not name or not content:
            st.error("「姓名」與「內容」為必填！")
        else:
            diag_final = diag_manual if diag_choice == "其他 (請於下方輸入)" else diag_choice
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
        with st.expander(f"[{h['location']}] {h['name']} - {h['time_occurred']}"):
            st.write(f"主治：{h['attending_doc']} | 病史：{h['history']} | 診斷：{h['diagnosis']}")
            st.write(f"內容：{h['content']}")
            if st.button(f"刪除 {h['name']}", key=f"del_{idx}"):
                st.session_state.handovers.pop(idx)
                save_handovers(st.session_state.handovers)
                st.rerun()

# ================= 區塊 4：純文字無痕填空引擎 =================
def get_unique_cells(row):
    unique_cells = []
    for cell in row.cells:
        if cell not in unique_cells: unique_cells.append(cell)
    return unique_cells

def safe_fill_cell(cell, text):
    if text is None or text == "": return
    for p in cell.paragraphs: p.text = "" 
    run = (cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()).add_run(str(text))
    run.font.size = Pt(10)

def build_word_document(p_stations, p_new, p_out, handovers, selected_date):
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"找不到 {TEMPLATE_PATH}。請確認已將樣板放在同一資料夾中。")
    
    doc = Document(TEMPLATE_PATH)
    
    # 民國年日期填寫
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
                            safe_fill_cell(u_cells[idx+1], p_stations[matched_st][0])
                            safe_fill_cell(u_cells[idx+2], p_stations[matched_st][1])
                            safe_fill_cell(u_cells[idx+3], p_stations[matched_st][2])
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
                            safe_fill_cell(u_cells[name_col_idx+k if k<6 else name_col_idx+k+1], pd[k])
                        new_idx += 1
                    elif fill_mode == "out" and out_idx < len(p_out):
                        pd = p_out[out_idx]
                        for k in range(min(len(pd), len(u_cells))):
                            safe_fill_cell(u_cells[name_col_idx+k], pd[k])
                        out_idx += 1

    # --- 【重要】刪除「討論與講評：」表格以騰出空間防止跑版 ---
    deleted_discuss = False
    discuss_table_idx = -1
    for idx, table in enumerate(doc.tables):
        first_row_txt = "".join([c.text for c in table.rows[0].cells]).replace(" ","")
        if "討論與講評" in first_row_txt:
            discuss_table_idx = idx
            break
            
    if discuss_table_idx != -1:
        table_to_del = doc.tables[discuss_table_idx]
        # 在刪除前插入一小段空白，保持視覺距離
        p_before = table_to_del.paragraphs[0] if table_to_del.paragraphs else None
        if p_before: p_before.insert_paragraph_before()
        # 執行刪除
        table_to_del._element.getparent().remove(table_to_del._element)
        deleted_discuss = True

    # --- 【重要】刪除第一張圖最下面一欄 (危險評估) ---
    deleted_risk = False
    risk_table_idx = -1
    for idx, table in enumerate(doc.tables):
        first_row_txt = "".join([c.text for c in table.rows[0].cells]).replace(" ","")
        if "危險評估" in first_row_txt:
            risk_table_idx = idx
            break
            
    if risk_table_idx != -1:
        table_to_del = doc.tables[risk_table_idx]
        table_to_del._element.getparent().remove(table_to_del._element)
        deleted_risk = True

    # --- 填寫標準流水帳格式交班 ---
    sorted_h = sorted(handovers, key=lambda x: (x.get('location') != '急診', x.get('time_occurred')))
    
    h_lines = []
    for h in sorted_h:
        h_loc = h.get('location', '病房')
        h_name = h.get('name', '')
        h_age = h.get('age', '')
        h_gen = h.get('gender', '')
        h_med = h.get('med_record', '')
        h_att = h.get('attending_doc', '')
        h_diag = h.get('diagnosis', '')
        h_his = h.get('history', '')
        h_time = h.get('time_occurred', '')
        # 將內容中的換行取代為空白，強迫不換段
        h_content = h.get('content', '').replace('\n', ' ')

        # 客製化字串防呆與格式化
        med_str = f"病歷號:{h_med} " if h_med else ""
        age_str = f"{h_age}歲" if h_age else ""
        gen_str = f"{h_gen}性" if h_gen else ""
        age_gen_combined = f" {age_str}{gen_str}" if (h_age or h_gen) else ""
        
        ward_tag = f"({h_loc[0:2]})" if h_loc != "急診" else ""
        doc_str = f"{h_att}醫師{ward_tag}病人" if h_att else f"{ward_tag}病人"
        
        diag_str = f"，診斷{h_diag}" if h_diag else ""
        his_str = f"，內外科病史:{h_his}" if h_his else ""
        time_str = f" {h_time}時" if h_time else ""

        # 最終流水帳格式組裝
        line = f"({h_loc})病歷號:{h_name}{age_gen_combined}，{doc_str}{his_str}{diag_str}{time_str}，{h_content}\n"
        h_lines.append(line)

    final_h_text = "".join(h_lines)

    # 寫入 Word，使用標準12pt字體且不加粗
    inserted = False
    for table in doc.tables:
        for row in table.rows:
            for cell in get_unique_cells(row):
                for p in cell.paragraphs:
                    # 修正：那一欄不可以有字，從下一行填入
                    if "病房特殊狀況及處理" in p.text.replace(" ",""):
                        run = p.add_run(final_h_text)
                        run.font.size = Pt(12)  # 標準流水帳字體大一些
                        run.bold = False        # 取消粗體
                        # 確保文字貼齊格子
                        p.paragraph_format.space_before = Pt(0)
                        p.paragraph_format.space_after = Pt(0)
                        inserted = True; break
                if inserted: break
            if inserted: break
        if inserted: break
        
    if not inserted:
        for p in doc.paragraphs:
            if "病房特殊狀況及處理" in p.text.replace(" ",""):
                run = p.add_run(final_h_text)
                run.font.size = Pt(12)
                run.bold = False
                break

    stream = io.BytesIO(); doc.save(stream); stream.seek(0)
    return stream

st.header("4. 確認與輸出")
if st.button("🚀 生成下載 Word", type="primary"):
    try:
        f = build_word_document(parsed_stations, parsed_new, parsed_out, st.session_state.handovers, duty_date)
        st.success(f"檔案已備妥！(已自動刪除刪除{'危險評估' if deleted_risk else ''}{'討論表格' if deleted_discuss else ''}以腾出空間)")
        st.download_button("📥 點擊下載", f, f"值班日誌_{duty_date.strftime('%Y%m%d')}.docx")
    except Exception as e:
        st.error(f"錯誤: {e}")
