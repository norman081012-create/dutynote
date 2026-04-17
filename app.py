import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_BREAK
import io
import json
import os
import re
import datetime
import pandas as pd

st.set_page_config(page_title="醫師值班日誌自動化系統", layout="wide")

# 固定護理站的背景格式
STATIONS_LIST = ["急診護理站", "二樓護理站", "三樓護理站", "四樓護理站", "五樓護理站", "總人數"]
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

st.title("🏥 醫師病房值班日誌 (護理站格式鎖定版)")

# ================= 區塊 1：資料輸入與格式預覽 =================
st.header("1. 建立背景格式與輸入資料")

col_ctrl, col_input = st.columns([3, 7])

with col_ctrl:
    duty_date = st.date_input("📅 選擇值班日期", datetime.date.today())
    if st.button("🔄 清空所有資料 (含交班與貼上內容)"):
        st.session_state.handovers = []
        save_handovers([])
        st.session_state.uploader_key += 1
        st.rerun()

# 解析函數：專攻 6 個護理站與病人資料
def parse_his_content(text):
    data = {s: ["0", "0", "0"] for s in STATIONS_LIST} # 預設背景格式
    new_pts, out_pts = [], []
    
    if text:
        lines = text.splitlines()
        current_sec = None
        for line in lines:
            line = line.strip()
            if not line: continue
            
            # 分割資料 (Tab 或多空格)
            parts = [p.strip() for p in re.split(r'\t|\s{2,}', line)]
            row_str = "".join(parts).replace(" ", "")
            
            # A. 匹配 6 個護理站
            for st_name in STATIONS_LIST:
                if st_name in row_str and len(parts) >= 4:
                    # 尋找關鍵字所在格子，抓取後方三個數字
                    for i, p in enumerate(parts):
                        if st_name in p.replace(" ", ""):
                            data[st_name] = parts[i+1 : i+4]
                            break
            
            # B. 判定病人區塊
            if "新入院" in row_str or "新住院" in row_str: current_sec = "new"
            elif "出院病人" in row_str: current_sec = "out"
            elif "危險評估" in row_str: current_sec = None
            
            if len(parts) >= 5 and "姓名" not in row_str and "病患" not in row_str:
                if current_sec == "new": new_pts.append(parts)
                elif current_sec == "out": out_pts.append(parts)
                
    return data, new_pts, out_pts

with col_input:
    raw_text = st.text_area(
        "📝 在此貼上 HIS 資料 (包含護理站與病人名單)", 
        height=200, 
        key=f"input_{st.session_state.uploader_key}"
    )
    
    # 建立正確的背景格式預覽
    st.subheader("📊 護理站背景格式預覽")
    p_stations, p_new, p_out = parse_his_content(raw_text)
    
    # 轉為 DataFrame 顯示，讓使用者確認 6 個欄位都正確
    df_preview = pd.DataFrame.from_dict(
        p_stations, orient='index', columns=['男', '女', '總人數']
    )
    st.table(df_preview)
    
    if p_new or p_out:
        st.caption(f"✅ 偵測到新住院 {len(p_new)} 人 / 出院 {len(p_out)} 人")

# ================= 區塊 2：交班事項 =================
st.header("2. 登錄交班事項 (將產出於第二頁)")
with st.form("handover_form", clear_on_submit=True):
    c1, c2, c3 = st.columns(3)
    with c1:
        is_er = st.checkbox("🚨 ER 紅色標記")
        name = st.text_input("病人姓名")
        age = st.text_input("年紀 (若無則不顯示)")
    with c2:
        gender = st.selectbox("性別", ["", "男", "女"])
        med_record = st.text_input("病歷號")
        time_occ = st.time_input("狀況發生時間")
    with c3:
        attending = st.selectbox("主治醫師", ["", "鍾", "張", "劉", "謝", "成"])
        diag = st.text_input("診斷")
    content = st.text_area("交班內容")
    
    if st.form_submit_button("確認新增交班"):
        if name and content:
            st.session_state.handovers.append({
                "is_er": is_er, "name": name, "age": age, "gender": gender,
                "med_record": med_record, "attending": attending, "diag": diag,
                "time": time_occ.strftime("%H:%M"), "content": content
            })
            save_handovers(st.session_state.handovers)
            st.rerun()

# 顯示已登錄交班 (點開可編輯/刪除)
for idx, h in enumerate(st.session_state.handovers):
    title = f"{'🚨[ER] ' if h['is_er'] else ''}{h['name']} ({h['time']})"
    with st.expander(title):
        st.write(f"病歷號: {h['med_record']} | 診斷: {h['diag']} | 主治: {h['attending']}")
        st.write(f"內容: {h['content']}")
        if st.button(f"刪除 {h['name']}", key=f"del_{idx}"):
            st.session_state.handovers.pop(idx)
            save_handovers(st.session_state.handovers)
            st.rerun()

# ================= 區塊 3：文件生成引擎 =================
def safe_fill(cell, text):
    if not text: return
    for p in cell.paragraphs: p.text = ""
    run = (cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()).add_run(str(text))
    run.font.size = Pt(10)

def generate_report(date_obj, stations_data, new_pts, out_pts, handovers):
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError("請確保 template.docx 已上傳至 GitHub。")
    
    doc = Document(TEMPLATE_PATH)
    
    # 1. 填寫日期 (民國年)
    roc_year = date_obj.year - 1911
    date_str = f"日期： {roc_year} 年 {date_obj.month:02d} 月 {date_obj.day:02d} 日"
    for p in doc.paragraphs:
        if "日期" in p.text: p.text = date_str
    
    # 2. 填寫表格 (護理站、新住院、出院)
    new_p_idx, out_p_idx = 0, 0
    for table in doc.tables:
        mode = None
        for row in table.rows:
            txt = row.cells[0].text.replace(" ", "").replace("　", "").strip()
            all_txt = "".join([c.text for c in row.cells]).replace(" ", "")
            
            # A. 護理站背景格式填寫
            if txt in STATIONS_LIST:
                nums = stations_data[txt]
                if len(row.cells) >= 4:
                    safe_fill(row.cells[1], nums[0])
                    safe_fill(row.cells[2], nums[1])
                    safe_fill(row.cells[3], nums[2])
                continue
            
            # B. 病人區塊判定
            if "姓名" in all_txt and "病歷" in all_txt:
                mode = "new" if "燈號" in all_txt else "out"
                continue
            if mode and ("出院" in txt or "危險" in txt or "特殊" in txt):
                mode = None
            
            # C. 填入病人
            is_empty = (txt == "" or txt == "_" or txt == "0")
            if mode == "new" and is_empty and new_p_idx < len(new_pts):
                p = new_pts[new_p_idx]
                safe_fill(row.cells[0], p[0]) # 姓名
                for k in range(1, min(len(p), len(row.cells))):
                    safe_fill(row.cells[k], p[k])
                new_p_idx += 1
            elif mode == "out" and is_empty and out_p_idx < len(out_pts):
                p = out_pts[out_p_idx]
                safe_fill(row.cells[0], p[0])
                for k in range(1, min(len(p), len(row.cells))):
                    safe_fill(row.cells[k], p[k])
                out_p_idx += 1

    # 3. 填寫交班事項 (強制分頁至第二頁)
    sorted_h = sorted(handovers, key=lambda x: (not x['is_er'], x['time']))
    h_text = "\n"
    for h in sorted_h:
        er_tag = "🚨[ER] " if h['is_er'] else ""
        age_str = f"{h['age']}歲 / " if h['age'] else ""
        h_text += f"【{er_tag}{h['name']}】{h['time']} | {age_str}{h['gender']} | 病歷:{h['med_record']}({h['attending']})\n"
        h_text += f"交班內容：{h['content']}\n" + "-"*30 + "\n"

    for p in doc.paragraphs:
        if "病房特殊狀況及處理" in p.text.replace(" ", ""):
            # 插入分頁符號，確保內容在第二頁
            p.insert_paragraph_before().add_run().add_break(WD_BREAK.PAGE)
            p.add_run("\n" + h_text)
            break

    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

st.header("3. 確認與輸出")
if st.button("🚀 生成下載 Word 檔案", type="primary"):
    try:
        final_doc = generate_report(duty_date, p_stations, p_new, p_out, st.session_state.handovers)
        st.success("✅ 檔案已依照背景格式生成成功！")
        st.download_button(
            label="📥 點擊下載值班日誌",
            data=final_doc,
            file_name=f"值班日誌_{duty_date.strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"發生錯誤: {e}")
