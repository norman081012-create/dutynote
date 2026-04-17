import streamlit as st
import docx
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import json
import os
import re
import datetime

st.set_page_config(page_title="值班日誌自動生成器", layout="wide")

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

st.title("🏥 醫師病房值班日誌自動生成器 (全自動無樣板版)")

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

# ================= 區塊 4：全自動生成文件引擎 =================
def build_word_document(raw_text, handovers, selected_date):
    # --- 1. 從貼上的文字萃取資料 ---
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
            if "護理站" in row_str or "急診護理站" in row_str: current_section = "station"
            if "新入院" in row_str or "新住院" in row_str: current_section = "new"
            if "出院病人" in row_str: current_section = "out"
            
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

    # --- 2. 從零開始建立 Word 檔案 ---
    doc = Document()
    
    # 設定邊界 (適應 A4 盡量放多一點內容)
    for section in doc.sections:
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)

    # A. 標題與日期
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("財團法人台灣省私立高雄仁愛之家附設慈惠醫院\n醫師病房值班日誌")
    title_run.bold = True
    title_run.font.size = Pt(14)
    
    roc_year = selected_date.year - 1911
    date_p = doc.add_paragraph(f"日期：  {roc_year} 年  {selected_date.month:02d} 月  {selected_date.day:02d} 日")
    date_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # B. 護理站統計表
    doc.add_paragraph("護理站")
    t1 = doc.add_table(rows=7, cols=4)
    t1.style = 'Table Grid'
    t1_headers = ["護理站", "男", "女", "病人總數"]
    for i, h in enumerate(t1_headers): t1.cell(0, i).text = h
    
    stations = ["急診護理站", "二樓護理站", "三樓護理站", "四樓護理站", "五樓護理站", "總人數"]
    for idx, st_name in enumerate(stations):
        t1.cell(idx+1, 0).text = st_name
        if st_name in parsed_stations:
            nums = parsed_stations[st_name]
            t1.cell(idx+1, 1).text = str(nums[0])
            t1.cell(idx+1, 2).text = str(nums[1])
            t1.cell(idx+1, 3).text = str(nums[2])

    doc.add_paragraph("") # 空行分隔

    # C. 新住院病人
    doc.add_paragraph("新住院病人:")
    t2_rows = max(7, len(parsed_new) + 1) # 至少留 6 個空行
    t2 = doc.add_table(rows=t2_rows, cols=9)
    t2.style = 'Table Grid'
    t2_headers = ["姓   名", "病歷號碼", "床 號", "性別", "年齡", "診   斷", "危險評估", "燈 號", "強制"]
    for i, h in enumerate(t2_headers): t2.cell(0, i).text = h
    
    for idx, p_data in enumerate(parsed_new):
        if idx + 1 < t2_rows:
            t2.cell(idx+1, 0).text = p_data[0]
            if len(p_data) > 1: t2.cell(idx+1, 1).text = p_data[1]
            if len(p_data) > 2: t2.cell(idx+1, 2).text = p_data[2]
            if len(p_data) > 3: t2.cell(idx+1, 3).text = p_data[3]
            if len(p_data) > 4: t2.cell(idx+1, 4).text = p_data[4]
            if len(p_data) > 5: t2.cell(idx+1, 5).text = p_data[5]
            if len(p_data) > 6: t2.cell(idx+1, 7).text = p_data[6] # 燈號填入第 7 格

    doc.add_paragraph("")

    # D. 出院病人
    doc.add_paragraph("出院病人:")
    t3_rows = max(7, len(parsed_out) + 1)
    t3 = doc.add_table(rows=t3_rows, cols=7)
    t3.style = 'Table Grid'
    t3_headers = ["姓   名", "病歷號碼", "床 號", "性別", "年齡", "診   斷", "出 院 動 態"]
    for i, h in enumerate(t3_headers): t3.cell(0, i).text = h
    
    for idx, p_data in enumerate(parsed_out):
        if idx + 1 < t3_rows:
            t3.cell(idx+1, 0).text = p_data[0]
            if len(p_data) > 1: t3.cell(idx+1, 1).text = p_data[1]
            if len(p_data) > 2: t3.cell(idx+1, 2).text = p_data[2]
            if len(p_data) > 3: t3.cell(idx+1, 3).text = p_data[3]
            if len(p_data) > 4: t3.cell(idx+1, 4).text = p_data[4]
            if len(p_data) > 5: t3.cell(idx+1, 5).text = p_data[5]
            if len(p_data) > 6: t3.cell(idx+1, 6).text = p_data[6]

    doc.add_paragraph("")

    # E. 危險評估 (純空表)
    doc.add_paragraph("危險評估")
    t4 = doc.add_table(rows=6, cols=8)
    t4.style = 'Table Grid'
    t4_headers = ["自殺顧慮", "哽塞顧慮", "身體顧慮", "暴力顧慮", "跌倒顧慮", "過度飲水", "私自離院", "其他"]
    for i, h in enumerate(t4_headers): t4.cell(0, i).text = h

    doc.add_paragraph("")

    # F. 病房特殊狀況及處理 (填入交班內容)
    p_special = doc.add_paragraph("病房特殊狀況及處理：")
    p_special.runs[0].bold = True
    
    sorted_handovers = sorted(handovers, key=lambda x: (not x['is_er'], x['time_occurred']))
    h_text = ""
    for h in sorted_handovers:
        h_text += f"\n【{'🚨ER ' if h['is_er'] else ''}{h['name']}】{h['time_occurred']} | {h['age']}歲/{h['gender']} | 病歷:{h['med_record']} (主治:{h['attending_doc']})\n"
        h_text += f"交班：{h['content']}\n"
        h_text += "-" * 40 + "\n"
        
    if not h_text:
        h_text = "\n\n\n\n" # 如果沒交班，留白幾行
        
    doc.add_paragraph(h_text)

    # G. 結尾簽名區塊
    p_discuss = doc.add_paragraph("討論與講評:")
    p_discuss.runs[0].bold = True
    doc.add_paragraph("\n\n") # 留空給講評
    
    # 底部簽核表 (無框線排版)
    t_sig = doc.add_table(rows=2, cols=3)
    
    # 第一列
    t_sig.cell(0, 0).text = "值班醫師："
    t_sig.cell(0, 1).text = "晨會主持人："
    t_sig.cell(0, 2).text = "呈核             批示\n\n□副院長︰\n\n□院  長："
    
    # 第二列
    t_sig.cell(1, 0).text = "\n專科護理師："
    t_sig.cell(1, 1).text = "\n精神部主任："

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

st.header("4. 確認與輸出")
if st.button("🚀 生成並下載 Word 檔案", type="primary"):
    try:
        final_file = build_word_document(raw_text_input, st.session_state.handovers, duty_date)
        st.success("🎉 檔案已從零開始完美生成！")
        st.download_button(
            label="📥 點擊下載最新版值班日誌",
            data=final_file,
            file_name=f"值班日誌_{duty_date.strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"錯誤詳情: {e}")
