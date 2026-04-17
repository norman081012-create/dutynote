import streamlit as st
import pandas as pd
from docx import Document
import io
import json
import os
import csv
import re

# 設定網頁標題與寬度
st.set_page_config(page_title="值班日誌自動生成器", layout="wide")

# ================= 系統設定與資料庫 =================
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

# 狀態初始化
if 'handovers' not in st.session_state:
    st.session_state.handovers = load_handovers()
if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0  # 用來重置上傳區塊

st.title("🏥 醫師病房值班日誌自動生成器")

# ================= 區塊 1：全局控制與上傳 =================
col_title, col_btn = st.columns([8, 2])
with col_btn:
    if st.button("🔄 刷新並清空所有資料", type="secondary", use_container_width=True):
        st.session_state.handovers = []
        save_handovers([])
        st.session_state.uploader_key += 1 # 強制刷新檔案上傳器
        st.rerun()

st.header("1. 上傳 HIS 系統匯出檔案")
# 透過 key 綁定 session_state，改變 key 就能清空上傳的檔案
csv_file = st.file_uploader("上傳值班日誌數據 (.csv)", type=['csv'], key=f"csv_uploader_{st.session_state.uploader_key}")

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
        attending_doc = st.text_input("主治醫師")
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
st.header("3. 已登錄交班事項 (依 ER 與時間排序)")

if not st.session_state.handovers:
    st.info("目前尚無交班紀錄。")
else:
    # 核心排序邏輯：先按 ER (True 優先於 False)，再按時間排序
    sorted_view = sorted(st.session_state.handovers, key=lambda x: (not x['is_er'], x['time_occurred']))
    
    for h in sorted_view:
        # 尋找這個項目在原始列表中的真實 index 以便刪除
        original_idx = st.session_state.handovers.index(h) 
        
        title = f"{'🚨[ER] ' if h['is_er'] else ''}{h['name']} - {h['time_occurred']}"
        with st.expander(title):
            st.markdown(f"**詳細資料：** {h['age']}歲/{h['gender']} | 病歷：{h['med_record']} | 主治：{h['attending_doc']}")
            st.markdown(f"**診斷：** {h['diagnosis']}")
            st.markdown(f"**交班內容：**\n{h['content']}")
            if st.button(f"刪除 {h['name']}", key=f"del_{original_idx}"):
                st.session_state.handovers.pop(original_idx)
                save_handovers(st.session_state.handovers)
                st.rerun()

# ================= 區塊 4：解析與生成核心邏輯 =================
def process_data(csv_data, handovers):
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"找不到 {TEMPLATE_PATH}。請確認已上傳樣板。")
    
    doc = Document(TEMPLATE_PATH)
    lines = []
    if csv_data:
        decoded = csv_data.getvalue().decode("utf-8")
        lines = list(csv.reader(io.StringIO(decoded)))
    
    # --- 1. 關鍵字尋找表頭 ---
    idx_station = idx_new = idx_out = -1
    for i, row in enumerate(lines):
        row_str = "".join(row).replace(" ", "")
        if "值班起始日期" in row_str:
            match = re.search(r'(\d{4})(\d{2})(\d{2})', row_str)
            if match:
                y, m, d = match.groups()
                for p in doc.paragraphs:
                    if "日期" in p.text: p.text = f"日期：  {int(y)-1911} 年  {m} 月  {d} 日"
        # 嚴格匹配特徵表頭
        if "護理站" in row_str and "病人總數" in row_str: idx_station = i
        elif "病患姓名" in row_str and "入院燈號" in row_str: idx_new = i
        elif "病患姓名" in row_str and "出院動態" in row_str: idx_out = i

    # --- 2. 填寫人數統計表 (Word 表格 0) ---
    if idx_station != -1 and len(doc.tables) >= 1:
        tb = doc.tables[0]
        w_row = 1 # Word 表格從第 1 列開始填 (避開標題)
        for i in range(idx_station + 1, len(lines)):
            row = lines[i]
            if not row or not row[0].strip(): continue # 略過空行
            if "病患姓名" in "".join(row): break # 遇到下一個大表就停止
            
            if w_row < len(tb.rows) and len(row) >= 4:
                tb.cell(w_row, 1).text = row[1] # 男
                tb.cell(w_row, 2).text = row[2] # 女
                tb.cell(w_row, 3).text = row[3] # 總數
                w_row += 1
            if "總人數" in row[0]: break

    # --- 3. 填寫新入院 (Word 表格 1) ---
    if idx_new != -1 and len(doc.tables) >= 2:
        tb = doc.tables[1]
        w_row = 1
        for i in range(idx_new + 1, len(lines)):
            row = lines[i]
            if not row or not row[0].strip(): continue
            if "出院" in "".join(row) or "病患姓名" in "".join(row): break # 遇到下一個大表停止
            
            if w_row >= len(tb.rows): tb.add_row()
            tb.cell(w_row, 0).text = row[0] # 姓名
            tb.cell(w_row, 1).text = row[1] # 病歷
            tb.cell(w_row, 2).text = row[2] # 床號
            tb.cell(w_row, 3).text = row[3] # 性別
            tb.cell(w_row, 4).text = row[4] # 年齡
            tb.cell(w_row, 5).text = row[5] # 診斷
            if len(row) > 6: tb.cell(w_row, 7).text = row[6] # 入院燈號 (第7格)
            w_row += 1

    # --- 4. 填寫出院 (Word 表格 2) ---
    if idx_out != -1 and len(doc.tables) >= 3:
        tb = doc.tables[2]
        w_row = 1
        for i in range(idx_out + 1, len(lines)):
            row = lines[i]
            if not row or not row[0].strip(): continue
            
            if w_row >= len(tb.rows): tb.add_row()
            tb.cell(w_row, 0).text = row[0]
            tb.cell(w_row, 1).text = row[1]
            tb.cell(w_row, 2).text = row[2]
            tb.cell(w_row, 3).text = row[3]
            tb.cell(w_row, 4).text = row[4]
            tb.cell(w_row, 5).text = row[5]
            if len(row) > 6: tb.cell(w_row, 6).text = row[6] # 出院動態 (第6格)
            w_row += 1

    # --- 5. 填寫交班事項 (排序與防重複寫入) ---
    sorted_handovers = sorted(handovers, key=lambda x: (not x['is_er'], x['time_occurred']))
    h_text = ""
    for h in sorted_handovers:
        h_text += f"\n【{'🚨ER ' if h['is_er'] else ''}{h['name']}】 {h['time_occurred']}\n"
        h_text += f"資料：{h['age']}歲/{h['gender']}/{h['med_record']} (主治:{h['attending_doc']})\n"
        h_text += f"交班：{h['content']}\n" + ("-"*30)

    inserted = False
    # 尋找段落
    for p in doc.paragraphs:
        if "病房特殊狀況及處理" in p.text.replace(" ", "") and not inserted:
            p.add_run(h_text)
            inserted = True
            break # 找到並寫入後立刻跳出，防止重複寫入

    # 如果段落找不到，尋找表格內的儲存格
    if not inserted:
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if "病房特殊狀況及處理" in cell.text.replace(" ", "") and not inserted:
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
        final_file = process_data(csv_file, st.session_state.handovers)
        st.success("檔案已備妥！")
        st.download_button(
            label="📥 點擊下載",
            data=final_file,
            file_name="值班日誌_輸出版本.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        st.error(f"錯誤詳情: {e}")
