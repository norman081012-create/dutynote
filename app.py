import streamlit as st
import pandas as pd
from docx import Document
import io

# 設定網頁標題與寬度
st.set_page_config(page_title="值班日誌自動生成器", layout="wide")

# 初始化 Session State 來儲存交班事項
if 'handovers' not in st.session_state:
    st.session_state.handovers = []

st.title("🏥 醫師病房值班日誌自動生成器")

# ================= 區塊 1：檔案上傳 =================
st.header("1. 上傳檔案")
col1, col2 = st.columns(2)
with col1:
    template_file = st.file_uploader("上傳 Word 樣板 (.docx)", type=['docx'])
with col2:
    csv_file = st.file_uploader("上傳值班日誌數據 (.csv)", type=['csv'])

# ================= 區塊 2：交班事項登錄表單 =================
st.header("2. 交班事項登錄")

with st.form("handover_form", clear_on_submit=True):
    st.subheader("新增交班紀錄")
    
    # 排版：使用 columns 讓表單更緊湊
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
    
    submitted = st.form_submit_button("確認新增交班")
    
    if submitted:
        if not name or not content:
            st.error("「病人姓名」與「交班內容」為必填欄位！")
        else:
            # 將表單資料存入 Session State
            st.session_state.handovers.append({
                "is_er": is_er,
                "name": name,
                "age": age,
                "gender": gender,
                "med_record": med_record,
                "attending_doc": attending_doc,
                "diagnosis": diagnosis,
                "time_occurred": time_occurred.strftime("%H:%M"),
                "content": content
            })
            st.success(f"已新增 {name} 的交班紀錄！")

# ================= 區塊 3：已登錄交班事項預覽 =================
st.header("3. 已登錄交班事項")

if not st.session_state.handovers:
    st.info("目前尚無交班紀錄。")
else:
    for idx, handover in enumerate(st.session_state.handovers):
        # 標題格式化：如果有勾選 ER，則加上紅色標記
        er_tag = "🔴 [ER] " if handover["is_er"] else ""
        title = f"{er_tag}{handover['name']} - {handover['time_occurred']}"
        
        # 使用展開面板 (expander) 顯示詳細內容
        with st.expander(title):
            details = []
            if handover['age']: details.append(f"年紀: {handover['age']}")
            if handover['gender']: details.append(f"性別: {handover['gender']}")
            if handover['med_record']: details.append(f"病歷號: {handover['med_record']}")
            if handover['attending_doc']: details.append(f"主治醫師: {handover['attending_doc']}")
            if handover['diagnosis']: details.append(f"診斷: {handover['diagnosis']}")
            
            st.write(" | ".join(details))
            st.markdown(f"**交班內容：**\n{handover['content']}")
            
            # 提供刪除按鈕
            if st.button(f"刪除 {handover['name']} 的紀錄", key=f"del_{idx}"):
                st.session_state.handovers.pop(idx)
                st.rerun()

# ================= 區塊 4：生成 Word 檔案 =================
st.header("4. 確認與輸出")

def generate_word_doc(template, handovers):
    """
    此函數負責處理 Word 樣板的寫入。
    實務上，您需要根據您 Word 樣板中表格的確切 index 來填入 CSV 數據。
    這裡先示範如何將「交班事項」寫入文件末端。
    """
    doc = Document(template)
    
    # 處理交班事項文字生成
    handover_text = ""
    for h in handovers:
        er_mark = "[ER] " if h["is_er"] else ""
        handover_text += f"【{er_mark}{h['name']}】 {h['time_occurred']}\n"
        
        # 組合有填寫的選填欄位
        meta = []
        if h['age']: meta.append(h['age'])
        if h['gender']: meta.append(h['gender'])
        if h['med_record']: meta.append(h['med_record'])
        if h['attending_doc']: meta.append(h['attending_doc'])
        if h['diagnosis']: meta.append(h['diagnosis'])
        
        if meta:
            handover_text += f"基本資料：{' / '.join(meta)}\n"
        handover_text += f"交班內容：{h['content']}\n"
        handover_text += "-" * 20 + "\n"

    # 尋找「病房特殊狀況及處理：」並將交班紀錄補在其後
    for paragraph in doc.paragraphs:
        if "病房特殊狀況及處理：" in paragraph.text:
            paragraph.add_run("\n" + handover_text)
            
    # [進階功能預留] 
    # 若要解析 CSV 並填入 Word 表格中，需在此處讀取 csv_file
    # 並透過 doc.tables[0], doc.tables[1] 逐格 (cell.text) 替換。

    # 將生成的 Word 存入記憶體中以供下載
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

if st.button("確認輸出 Word 檔案", type="primary"):
    if template_file is None:
        st.error("請先上傳 Word 樣板檔案！")
    else:
        try:
            # 呼叫生成函數
            final_doc_stream = generate_word_doc(template_file, st.session_state.handovers)
            
            st.success("檔案生成成功！請點擊下方按鈕下載。")
            st.download_button(
                label="📥 下載完成的值班日誌",
                data=final_doc_stream,
                file_name="今日值班日誌_已完成.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"生成檔案時發生錯誤: {e}")
