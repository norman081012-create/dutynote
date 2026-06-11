import streamlit as st
import datetime
from utils import (
    tw_tz, ATTENDING_DOCS_GLOBAL, ATTENDING_DOCS_FORM, DIAG_CHOICES_FORM,
    load_handovers, save_handovers, parse_his_data, parse_prn_data,
    get_sort_key, build_word_and_check_overflow
)

st.set_page_config(page_title="值班日誌自動生成器", layout="wide")

# 年齡選單 (往上 49~1，預設太忙，往下 50~110)
age_options = [str(i) for i in range(1, 50)] + ["太忙了沒時間問"] + [str(i) for i in range(50, 111)]
default_age_idx = age_options.index("太忙了沒時間問")

# --- CSS 樣式注入 ---
st.markdown("""
<style>
div[data-baseweb="input"] input { text-align: center !important; }
div[data-baseweb="select"] div { text-align: center !important; justify-content: center !important; }
div[data-testid="stAlert"] { margin-bottom: 0px !important; padding-top: 10px !important; padding-bottom: 10px !important; }
h2 { padding-top: 0.5rem !important; }
</style>
""", unsafe_allow_html=True)

# ================= 表單狀態與全域設定初始化 =================
if 'handovers' not in st.session_state: st.session_state.handovers = load_handovers()
if 'uploader_key' not in st.session_state: st.session_state.uploader_key = 0

now_tw = datetime.datetime.now(tw_tz)
if "f_duty_date" not in st.session_state: st.session_state.f_duty_date = now_tw.date()
    
if "f_loc" not in st.session_state:
    st.session_state.update({
        "f_loc": "病房", "f_name": "", "f_age": "太忙了沒時間問", "f_gen": "",
        "f_med": "", "f_hist": "", "f_time": now_tw.time(),
        "f_doc": "太忙了沒時間問", "f_diag_c": "太忙了沒時間問", "f_diag_m": "", "f_content": "",
        "f_special": False, "add_error": False
    })

# ================= Callback =================
def clear_form():
    st.session_state.update({
        "f_loc": "病房", "f_name": "", "f_age": "太忙了沒時間問", "f_gen": "",
        "f_med": "", "f_hist": "", "f_time": datetime.datetime.now(tw_tz).time(),
        "f_doc": "太忙了沒時間問", "f_diag_c": "太忙了沒時間問", "f_diag_m": "", 
        "f_content": "", "f_special": False, "add_error": False
    })

def load_form(h):
    st.session_state.f_loc = h.get("location", "病房")
    st.session_state.f_name = h.get("name", "")
    age = h.get("age", "")
    st.session_state.f_age = "太忙了沒時間問" if age == "" else age
    st.session_state.f_gen = h.get("gender", "")
    st.session_state.f_med = h.get("med_record", "")
    st.session_state.f_hist = h.get("history", "")
    try: st.session_state.f_time = datetime.datetime.strptime(h.get("time_occurred", "00:00"), "%H:%M").time()
    except: st.session_state.f_time = datetime.datetime.now(tw_tz).time()
        
    doc = h.get("attending_doc", "")
    st.session_state.f_doc = "太忙了沒時間問" if doc == "" else doc
    
    diag = h.get("diagnosis", "")
    if diag in ["Schizophrenia", "bipolar", "depression"]:
        st.session_state.f_diag_c = diag
        st.session_state.f_diag_m = ""
    elif diag == "":
        st.session_state.f_diag_c = "太忙了沒時間問"
        st.session_state.f_diag_m = ""
    else:
        st.session_state.f_diag_c = "其他 (請於下方輸入)"
        st.session_state.f_diag_m = diag
        
    st.session_state.f_content = h.get("content", "")
    st.session_state.f_special = h.get("is_special", False)

def cb_refresh():
    st.session_state.handovers = []
    save_handovers([])
    clear_form()
    st.session_state.uploader_key += 1
    st.session_state.f_duty_date = datetime.datetime.now(tw_tz).date()

def cb_add():
    if not st.session_state.f_name or not st.session_state.f_content:
        st.session_state.add_error = True
    else:
        st.session_state.add_error = False
        diag_c_val = "" if st.session_state.f_diag_c == "太忙了沒時間問" else st.session_state.f_diag_c
        diag_final = st.session_state.f_diag_m if not diag_c_val or diag_c_val == "其他 (請於下方輸入)" else diag_c_val
        age_val = "" if st.session_state.f_age == "太忙了沒時間問" else st.session_state.f_age
        doc_val = "" if st.session_state.f_doc == "太忙了沒時間問" else st.session_state.f_doc

        st.session_state.handovers.append({
            "location": st.session_state.f_loc, "name": st.session_state.f_name, 
            "age": age_val, "gender": st.session_state.f_gen,
            "med_record": st.session_state.f_med, "attending_doc": doc_val,
            "time_occurred": st.session_state.f_time.strftime("%H:%M"), "content": st.session_state.f_content,
            "diagnosis": diag_final, "history": st.session_state.f_hist,
            "is_er": (st.session_state.f_loc == "急診"),
            "is_special": st.session_state.f_special
        })
        save_handovers(st.session_state.handovers)
        clear_form()

def cb_edit(idx, h):
    load_form(h)
    st.session_state.handovers.pop(idx)
    save_handovers(st.session_state.handovers)

def cb_delete(idx):
    st.session_state.handovers.pop(idx)
    save_handovers(st.session_state.handovers)


# ================= UI 畫面區 =================
st.title("🏥 醫師病房值班日誌自動生成器")

col_warn, col_btn = st.columns([8, 2], vertical_alignment="center")
with col_warn:
    st.info("⚠️ **溫馨提示：** 新值班醫師接班時，請務必點擊右方的「🔄 刷新並清空所有資料」，否則將會讀取到前一位醫師的設定檔與暫存資料喔！")
with col_btn:
    st.button("🔄 刷新並清空所有資料", type="secondary", use_container_width=True, on_click=cb_refresh)

st.header("1. 貼上系統匯出資料")
c_date, c_his, c_prn = st.columns([2, 4, 4])

with c_date:
    st.date_input("📅 選擇值班日期", key="f_duty_date")
    # (已移除選擇值班醫師選單)

with c_his:
    raw_his = st.text_area("📝 貼上 HIS 內容 (人數/出入院)", height=150, key=f"his_{st.session_state.uploader_key}")
    parsed_stations, parsed_new, parsed_out = parse_his_data(raw_his)

with c_prn:
    raw_prn = st.text_area("💊 貼上 PRN 藥物清單 (選填)", height=150, key=f"prn_{st.session_state.uploader_key}")
    prn_summary = parse_prn_data(raw_prn)

# ================= 區塊 2：交班事項登錄表單 =================
st.header("2. 交班事項登錄")
c1, c2 = st.columns(2)
with c1:
    st.selectbox("單位/病房 (預設此)", ["病房", "急診", "二樓病房", "三樓病房", "四樓病房", "五樓病房"], key="f_loc")
    st.text_input("病人姓名 (必填)", key="f_name")
    st.selectbox("年紀", age_options, index=default_age_idx, key="f_age")
    st.selectbox("性別", ["", "男", "女"], key="f_gen")
    st.text_input("病歷號", key="f_med")
    st.text_area("內外科病史輸入", height=60, key="f_hist")
    
with c2:
    st.time_input("狀況發生時間", key="f_time")
    st.selectbox("主治醫師", ATTENDING_DOCS_FORM, key="f_doc")
    st.selectbox("診斷快速選項", DIAG_CHOICES_FORM, key="f_diag_c")
    st.text_input("手手動輸入診斷 (若選其他)", key="f_diag_m")
    st.checkbox("🚨 特別交班", key="f_special")
    
st.text_area("交班內容 (必填)", key="f_content")

btn_col1, btn_col2, btn_col3 = st.columns([2, 1, 1])
with btn_col1:
    st.button("✅ 確認新增交班", type="primary", use_container_width=True, on_click=cb_add)
    if st.session_state.add_error: st.error("「姓名」與「內容」為必填！")
with btn_col2:
    st.button("🔄 重新輸入", use_container_width=True, on_click=clear_form)

# ================= 區塊 3：已登錄交班預覽 =================
st.header("3. 已登錄交班事項")
if st.session_state.handovers:
    sorted_view = sorted(st.session_state.handovers, key=get_sort_key)
    for h in sorted_view:
        idx = st.session_state.handovers.index(h)
        h_age_disp = h['age'] if h.get('age') else "?"
        h_gen_disp = f"{h['gender']}性" if h.get('gender') else ""
        sp_tag = " [🚨特別交班]" if h.get('is_special') else ""
        
        with st.expander(f"[{h['location']}] {h['name']} ({h_age_disp}歲{h_gen_disp}) - {h['time_occurred']}{sp_tag}"):
            h_diag_disp = h['diagnosis'] if h.get('diagnosis') else "??"
            st.write(f"主治：{h['attending_doc']} | 病史：{h['history']} | 診斷：{h_diag_disp}")
            st.write(f"內容：{h['content']}")
            
            c_edit, c_del, c_empty = st.columns([1.5, 1.5, 7])
            with c_edit: st.button(f"✏️ 修改 {h['name']}", key=f"edit_{idx}", on_click=cb_edit, args=(idx, h))
            with c_del: st.button(f"🗑️ 刪除 {h['name']}", key=f"del_{idx}", on_click=cb_delete, args=(idx,))

# ================= 工具與輸出 =================
st.header("4. 預覽與輸出")

# 生成最終預覽文字
preview_lines = []
sorted_h = sorted(st.session_state.handovers, key=get_sort_key)
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
    doc_part = f"{h_att}醫師{ward_tag}病人" if h_att else ""
    his_part = f"內外科病史:{h_his}" if h_his else ""
    if not h_diag: h_diag = "??"
    diag_part = f"診斷:{h_diag}"
    time_part = f"約{h_time}時" if h_time else ""
    
    diag_time = ""
    if diag_part and time_part: diag_time = f"{diag_part} {time_part}"
    elif diag_part: diag_time = diag_part
    elif time_part: diag_time = time_part
        
    components = [c for c in [pt_part, doc_part, his_part, diag_time, h_content] if c.strip()]
    preview_lines.append("，".join(components))

# 附加 PRN 藥物
if prn_summary:
    preview_lines.append("")
    preview_lines.extend(prn_summary.splitlines())

if preview_lines:
    with st.expander("👀 點擊展開：最終交班文字預覽 (與 Word 輸出內容相同)", expanded=True):
        st.text_area("即將寫入 Word 的文字：", value="\n\n".join(preview_lines), height=250, disabled=True)

if st.button("🚀 生成下載 Word", type="primary"):
    try:
        # 已移除 selected_doc 參數
        f_stream, overflow = build_word_and_check_overflow(
            parsed_stations, parsed_new, parsed_out, 
            preview_lines, 
            st.session_state.f_duty_date
        )
        if overflow:
            st.info("ℹ️ 交班內容較長，系統已自動為您排版新分頁，並確保『新版簽章區塊與勾選框』置於最後一頁的底部不跑位！")
        else:
            st.success("✅ 檔案已更新並備妥！")
            
        st.download_button("📥 點擊下載", f_stream, f"值班日誌_{st.session_state.f_duty_date.strftime('%Y%m%d')}.docx")
    except Exception as e:
        st.error(f"錯誤: {e}")
