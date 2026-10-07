"""
台灣健保降血脂藥物給付規範｜門診速查（表一：ASCVD 風險分級）
依據：衛福部健保署「全民健康保險降血脂藥物給付規定」修訂對照表（115/9/1 生效）
本工具為臨床速查摘要；實際申報以健保署最新公告及健保藥品代碼為準。
執行：pip install streamlit && streamlit run ldl_app.py
"""
import calendar
from datetime import date, timedelta

import streamlit as st

WEEKDAY = "一二三四五六日"


def add_months(d: date, n: int) -> date:
    m = d.month - 1 + n
    y, m = d.year + m // 12, m % 12 + 1
    return date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def fmt(d: date) -> str:
    return f"{d:%Y/%m/%d}（{WEEKDAY[d.weekday()]}）"


def window(base: date, lo, hi, unit: str) -> str:
    """unit: 'w' 週 / 'm' 月；回傳「起日 – 迄日」"""
    f = (lambda n: base + timedelta(weeks=n)) if unit == "w" else (lambda n: add_months(base, n))
    return f"{fmt(f(lo))} – {fmt(f(hi))}" if lo != hi else fmt(f(lo))


st.set_page_config(page_title="降血脂給付速查", page_icon="🩺", layout="wide")

# ---------- 規範資料（表一） ----------
LEVELS = {
    "極高": dict(color="#B3392F", start=55, ldl=55, nonhdl=85, drug_first=True,
               tx="改善風險因子＋中至高強度 statin，可合併 ezetimibe；6–8 週複查。",
               fu="達標後每 6 個月追蹤。"),
    "非常高": dict(color="#D2622A", start=70, ldl=70, nonhdl=100, drug_first=True,
                tx="改善風險因子＋中至高強度 statin，可合併 ezetimibe；6–8 週複查。",
                fu="達標後每 6 個月追蹤。"),
    "高": dict(color="#C98A1A", start=100, ldl=100, nonhdl=130, drug_first=True,
              tx="生活調整與藥物並行；中至高強度 statin，可合併 ezetimibe；6–8 週複查。",
              fu="達標後每 6 個月追蹤。"),
    "中": dict(color="#22857A", start=115, ldl=115, nonhdl=145, drug_first=False,
              tx="先生活型態調整 3–6 個月；未達標開始中強度 statin，6–8 週複查。",
              fu="達標後每 6–12 個月追蹤。"),
    "低": dict(color="#3F7D4E", start=130, ldl=130, nonhdl=160, drug_first=False,
              tx="先生活型態調整 3–6 個月；未達標開始中強度 statin，6–8 週複查。",
              fu="達標後每 6–12 個月追蹤。"),
    "0項": dict(color="#235E8C", start=160, ldl=160, nonhdl=None, drug_first=False,
               tx="依中、低風險流程：先生活型態調整 3–6 個月，未達標再開始藥物治療。",
               fu="依中、低風險：達標後每 6–12 個月追蹤。"),
}

# 極高：兩款，皆須「主診斷」＋合併其一
EXTREME = {
    "（一）冠狀動脈疾病": [
        "一年內曾經歷心肌梗塞",
        "兩次（含）以上心肌梗塞病史",
        "多支冠狀動脈阻塞",
        "急性冠心症合併糖尿病",
        "周邊動脈疾病或頸動脈狹窄",
    ],
    "（二）周邊動脈疾病": [
        "冠狀動脈疾病",
        "頸動脈狹窄",
    ],
}
# 非常高：(勾選文字, 結果卡簡稱)
VERY_HIGH = [
    ("急性冠心症病史 :gray[經臨床檢查確診為動脈硬化心血管疾病]", "急性冠心症病史"),
    ("接受血管再通術 :gray[心導管介入治療或外科冠狀動脈繞道手術]", "曾接受血管再通術"),
    ("缺血性中風或短暫性腦缺血發作 :gray[合併動脈硬化相關疾病或病史]", "缺血性中風／TIA"),
    ("周邊動脈疾病 :gray[曾接受血管再通術、有肢體缺血相關症狀或截肢]", "周邊動脈疾病"),
    ("影像檢查確認顯著斑塊負擔（≧50% 直徑狹窄率） "
     ":gray[冠狀動脈血管攝影、冠狀動脈或周邊血管電腦斷層攝影、頸動脈或周邊血管超音波]",
     "影像確認狹窄 ≧50%"),
]

# ---------- 樣式 ----------
st.markdown("""
<style>
.block-container {padding-top: 1.6rem; max-width: 1200px;}
.result {border-left: 10px solid var(--c); background: color-mix(in srgb, var(--c) 7%, transparent);
         padding: 1rem 1.2rem; border-radius: 4px; margin-bottom: .8rem;}
.result h2 {margin: 0; color: var(--c); font-size: 2.1rem;}
.result .why {font-size: .9rem; opacity: .8; margin-top: .2rem;}
.nums {display: flex; gap: 2rem; margin: .8rem 0 .2rem; flex-wrap: wrap;}
.nums div span {display:block; font-size: .8rem; opacity: .7;}
.nums div b {font-size: 1.35rem;}
.ok {color: #2E7D4F; font-weight: 600;}
.ng {color: #B3392F; font-weight: 600;}
</style>
""", unsafe_allow_html=True)

st.title("降血脂藥物給付速查")

left, right = st.columns([3, 2], gap="large")

# ---------- 輸入 ----------
with left:
    st.subheader("1 病人資料")
    c1, c2, c3 = st.columns(3)
    sex = c1.radio("性別", ["男", "女"], horizontal=True)
    age = c2.number_input("年齡", min_value=18, max_value=110, value=None, step=1, placeholder="歲")
    visit = c3.date_input("本次就診日", value=date.today(), format="YYYY/MM/DD")

    c1, c2, c3 = st.columns(3)
    ldl = c1.number_input("LDL-C", min_value=0.0, value=None, step=1.0, placeholder="mg/dL")
    tc = c2.number_input("TC", min_value=0.0, value=None, step=1.0, placeholder="mg/dL")
    hdl = c3.number_input("HDL-C", min_value=0.0, value=None, step=1.0, placeholder="mg/dL")
    nonhdl = (tc - hdl) if (tc is not None and hdl is not None) else None
    if nonhdl is not None:
        st.caption(f"non-HDL-C（TC − HDL）= {nonhdl:.0f} mg/dL")

    st.subheader("2 極高風險條件")
    st.caption("條文分兩款，兩款都要先有主診斷再合併其一。只勾主診斷不構成極高風險。")
    with st.container(border=True):
        ext_hit = []
        for main, subs in EXTREME.items():
            if st.checkbox(f"**{main}**再合併下列任一項", key=f"em_{main}"):
                sub_cols = st.columns([1, 20])
                with sub_cols[1]:
                    hits = [x for x in subs if st.checkbox(x, key=f"es_{main}_{x}")]
                if hits:
                    ext_hit += [f"{main[3:]}＋{x}" for x in hits]
                else:
                    sub_cols[1].caption("尚未勾選合併條件，不列入極高風險。")

    st.subheader("3 非常高風險條件")
    st.caption("臨床確診 ASCVD，或影像確認顯著斑塊負擔。符合任一項即成立。")
    with st.container(border=True):
        vh_hit = [short for label, short in VERY_HIGH if st.checkbox(label, key=f"v_{short}")]

    st.subheader("4 高風險條件")
    st.caption("四款其一即成立。")
    with st.container(border=True):
        h_hit = []
        if st.checkbox("糖尿病"):
            h_hit.append("糖尿病")
        if st.checkbox("慢性腎臟病進入透析治療前 "
                       ":gray[UACR ≧30 mg/g 或 eGFR <60 mL/min/1.73m²，至少持續 3 個月]"):
            h_hit.append("透析前 CKD")
        if ldl is not None:
            ldl190 = ldl >= 190
            st.checkbox("LDL-C ≧190 mg/dL :gray[填入 LDL-C 達 190 時自動成立]",
                        value=ldl190, disabled=True)
        else:
            ldl190 = st.checkbox("LDL-C ≧190 mg/dL :gray[填入 LDL-C 達 190 時自動成立]")
        if ldl190:
            h_hit.append("LDL-C ≧190")
        if st.checkbox("冠狀動脈鈣化分數（CAC）≧400"):
            h_hit.append("CAC ≧400")

    st.subheader("5 心血管風險因子計數")
    st.caption("未符合上述高風險條件時，以此處的因子數量評估。2 項以上為中風險，1 項為低風險，"
               "0 項為「0 項心血管風險因子」。")
    with st.container(border=True):
        rf = []
        if st.checkbox("高血壓"):
            rf.append("高血壓")

        age_label = "男性 ≧45 歲，女性 ≧55 歲 :gray[填入性別與年齡時自動判定]"
        if age is not None:
            age_rf = age >= (45 if sex == "男" else 55)
            st.checkbox(age_label, value=age_rf, disabled=True)
        else:
            age_rf = st.checkbox(age_label)
        if age_rf:
            rf.append("年齡")

        if st.checkbox("早發性冠心病家族史 :gray[男性 ≦55 歲、女性 ≦65 歲]"):
            rf.append("早發性冠心病家族史")

        hdl_label = "HDL-C 偏低 :gray[男性 <40 mg/dL，女性 <50 mg/dL；填入性別與 HDL-C 時自動判定]"
        if hdl is not None:
            low_hdl = hdl < (40 if sex == "男" else 50)
            st.checkbox(hdl_label, value=low_hdl, disabled=True, key="rf_hdl")
        else:
            low_hdl = st.checkbox(hdl_label, key="rf_hdl")
        if low_hdl:
            rf.append("HDL-C 偏低")

        if st.checkbox("抽菸"):
            rf.append("抽菸")

        st.markdown("代謝性症候群 :gray[符合下列至少三項]")
        sub_cols = st.columns([1, 20])
        with sub_cols[1]:
            ms = [
                st.checkbox("腹部肥胖 :gray[男性 ≧90 cm，女性 ≧80 cm]"),
                st.checkbox("血壓偏高 :gray[≧130/85 mmHg 或使用高血壓藥物]"),
                st.checkbox("空腹血糖偏高 :gray[≧100 mg/dL 或使用糖尿病藥物]"),
                st.checkbox("空腹 TG 偏高 :gray[≧150 mg/dL 或使用治療 TG 血脂藥物]"),
            ]
            st.checkbox("HDL-C 偏低 :gray[男性 <40 mg/dL，女性 <50 mg/dL；與上方 HDL-C 項連動]",
                        value=low_hdl, disabled=True, key=f"ms_hdl_{low_hdl}")
            n_ms = sum(ms) + int(low_hdl)
            if n_ms >= 3:
                st.markdown(f"<span class='ng'>{n_ms}/5 項，代謝性症候群成立</span>",
                            unsafe_allow_html=True)
            else:
                st.caption(f"目前 {n_ms}/5 項")
        if n_ms >= 3:
            rf.append("代謝性症候群")


# ---------- 判定 ----------
if ext_hit:
    level, why = "極高", ext_hit
elif vh_hit:
    level, why = "非常高", vh_hit
elif h_hit:
    level, why = "高", h_hit
elif len(rf) >= 2:
    level, why = "中", rf
elif len(rf) == 1:
    level, why = "低", rf
else:
    level, why = "0項", ["未符合任何高風險條件或心血管風險因子"]

L = LEVELS[level]

# ---------- 結果 ----------
with right:
    why_txt = "、".join(why)
    if level in ("中", "低"):
        why_txt = f"心血管風險因子 {len(rf)} 項：{why_txt}"
    nonhdl_target = f"&lt; {L['nonhdl']}" if L["nonhdl"] else "未列"
    st.markdown(f"""
<div class="result" style="--c:{L['color']}">
  <h2>{level + '風險' if level != '0項' else '0 項心血管風險因子'}</h2>
  <div class="why">{why_txt}</div>
  <div class="nums">
    <div><span>起始給付 LDL-C</span><b>≥ {L['start']}</b></div>
    <div><span>目標 LDL-C</span><b>&lt; {L['ldl']}</b></div>
    <div><span>次要目標 non-HDL-C</span><b>{nonhdl_target}</b></div>
  </div>
</div>
""", unsafe_allow_html=True)

    # 個案判讀
    with st.container(border=True):
        st.markdown("**本次數值**")
        if ldl is None:
            st.write("輸入 LDL-C 後顯示是否符合給付與達標。")
        else:
            if ldl >= L["start"]:
                st.markdown(f"<span class='ng'>LDL-C {ldl:.0f} ≥ {L['start']}：符合起始給付門檻／未達標</span>",
                            unsafe_allow_html=True)
                if not L["drug_first"]:
                    st.write("此級需先生活型態調整 3–6 個月，仍未達標再開始用藥。")
            else:
                st.markdown(f"<span class='ok'>LDL-C {ldl:.0f} &lt; {L['ldl']}：主要目標已達標</span>",
                            unsafe_allow_html=True)
                st.caption("未使用降脂藥者即未達起始給付門檻。")

            if L["nonhdl"] and nonhdl is not None:
                if ldl < L["ldl"]:
                    if nonhdl < L["nonhdl"]:
                        st.markdown(f"<span class='ok'>non-HDL-C {nonhdl:.0f} &lt; {L['nonhdl']}：次要目標達標</span>",
                                    unsafe_allow_html=True)
                    else:
                        st.markdown(f"<span class='ng'>non-HDL-C {nonhdl:.0f} ≥ {L['nonhdl']}：次要目標未達</span>",
                                    unsafe_allow_html=True)
                else:
                    st.caption(f"non-HDL-C {nonhdl:.0f}：LDL-C 達標後再評估次要目標。")

    with st.container(border=True):
        st.markdown("**起始處理**")
        st.write(L["tx"])

        st.markdown(f"**追蹤日期**（以 {fmt(visit)} 起算）")
        if L["drug_first"]:
            rows = [
                ("開始用藥後複查（6–8 週）", window(visit, 6, 8, "w")),
                ("更動藥物後複查血脂（1–3 個月）", window(visit, 1, 3, "m")),
                ("達標後定期追蹤（每 6 個月）", f"下次 {window(visit, 6, 6, 'm')}"),
            ]
        else:
            rows = [
                ("生活型態調整後複查（3–6 個月）", window(visit, 3, 6, "m")),
                ("若開始 statin，用藥後複查（6–8 週）", window(visit, 6, 8, "w")),
                ("更動藥物後複查血脂（1–3 個月）", window(visit, 1, 3, "m")),
                ("達標後定期追蹤（每 6–12 個月）", f"下次 {window(visit, 6, 12, 'm')}"),
            ]
        for label, d in rows:
            st.markdown(f"{label}  \n**{d}**")
        st.caption("就診日即為開始用藥、更動藥物或確認達標的那一天時，直接看對應列。")

        st.markdown("**仍未達標**")
        st.write("檢視服藥；調至高強度或最大耐受 statin，必要時合併其他降脂藥。")

    with st.expander("完整分級表"):
        for name, v in LEVELS.items():
            nh = f"< {v['nonhdl']}" if v["nonhdl"] else "未列"
            mark = "◀" if name == level else ""
            st.markdown(f"**{name}**　起始 ≥{v['start']}　目標 <{v['ldl']} / {nh} {mark}")

    # ---------- 衛教資訊 ----------
    st.subheader("衛教資訊")
    lv_name = level + "風險" if level != "0項" else "0 項心血管風險因子"

    with st.expander("為什麼同樣的 LDL-C，處理方式不一樣", expanded=True):
        st.markdown(f"""
健保先評估您將來發生心肌梗塞、中風或周邊動脈疾病的機會，分成六級，每一級各有開始用藥的數值和要降到的目標。
報告上的參考範圍，無法回答「要不要吃藥」。

您目前屬於 **{lv_name}**：LDL-C 達 **{L['start']} mg/dL** 以上可開始用藥，目標是降到 **{L['ldl']} mg/dL** 以下。
""")

    with st.expander("治療怎麼開始"):
        if L["drug_first"]:
            st.markdown("""
您的風險等級可以**用藥和生活調整同時開始**。第一線通常是中至高強度 statin，必要時合併 ezetimibe，
實際選擇依醫師評估您的血脂、過去用藥和身體耐受度。
""")
        else:
            st.markdown("""
您的風險等級要**先調整生活 3–6 個月**，期滿再驗一次完整血脂；仍未達標，才開始中強度 statin。
""")
        st.markdown("""
生活調整會評估：血壓、糖化血色素（HbA1c）、體重、抽菸、喝酒與作息。保健食品無法取代完整評估和處方藥。
""")

    with st.expander("抽血時間怎麼安排"):
        if L["drug_first"]:
            st.markdown("""
- 開始治療後 **6–8 週**驗血脂
- 達標：維持原用藥，之後**每 6 個月**追蹤
- 治療有更動：**1–3 個月內**再驗一次
""")
        else:
            st.markdown("""
- 生活調整 **3–6 個月**後驗血脂
- 若開始中強度 statin：**6–8 週**後確認是否達標
- 達標：維持，之後**每 6–12 個月**追蹤
- 治療有更動：**1–3 個月內**再驗一次
""")
        st.markdown("剛開始或剛換藥時抽血較密，穩定達標後間隔才拉長。回診時帶著過去半年到一年的血脂報告，比較容易看出變化。")

    with st.expander("吃藥後還沒達標"):
        st.markdown("""
第一步是確認有沒有按時服藥。接著醫師可能：

- 改為高強度 statin，或調到可耐受的最大劑量
- 合併其他降膽固醇藥物：ezetimibe、PCSK9 單株抗體、siRNA、ATP citrate lyase 抑制劑

請由醫師調整，不要自行加量或合併多種藥物。
""")

    with st.expander("Ezetimibe 與複方藥的規定"):
        st.markdown("""
**適用疾病**：原發性高膽固醇血症、同型接合子家族性高膽固醇血症、同型接合子性植物脂醇血症。

**可合併 ezetimibe 的情況**
- 無法耐受 statin 副作用，例如嚴重肌肉痠痛或肌炎
- 單用 statin **6–8 週**仍未達標（舊制為 3 個月）

**仍需單用 statin 滿 3 個月的指定品項**：Ezetity tablets 10mg、Ezzicad 10mg、Ezta 10mg、Ezetimibe Sandoz 10mg。

**Ezetimibe＋statin 複方**：一般為單用 statin 6–8 週未達標可使用；指定複方品項仍需 3 個月。複方藥**不可與 gemfibrozil 併用**。
""")

    with st.expander("需要特別注意的情況"):
        st.markdown("""
- **家族性高膽固醇血症線索**：膽固醇非常高、肌腱黃色瘤、年輕就發生心血管疾病，或家族中有人如此。建議依台灣診斷標準篩檢，不能只靠生活調整。
- **極高與非常高風險**：應驗完整血脂；急性發作住院者，應在入院 24 小時內完成。
- **數值低於門檻不要自行停藥**：目前漂亮的數字，很可能正是藥物壓下來的結果。
""")

    with st.expander("飲食衛教"):
        st.markdown("""
**少吃：飽和脂肪與反式脂肪**（最直接推高 LDL-C）
- 肥肉、五花肉、雞皮、豬油、牛油、奶油
- 炸物、酥皮糕餅、奶精、人造奶油、加工肉品（香腸、培根、火腿）

**換成：好的油脂與蛋白質**
- 炒菜用植物油（橄欖油、芥花油、苦茶油），少用動物油
- 蛋白質優先選豆腐、豆漿、魚、去皮雞肉，紅肉選瘦的部位
- 每週吃 2 次魚，一天一小把無調味堅果

**多吃：可溶性纖維**（幫助降低 LDL-C）
- 燕麥、糙米等全穀，豆類，每餐半盤蔬菜，水果每天 2 份

**三酸甘油脂偏高者另外注意**
- 少喝含糖飲料、少吃甜點與精緻澱粉（白飯、麵包、麵條不過量）
- 盡量不喝酒

**和藥物有關**
- **葡萄柚／柚子**：simvastatin、lovastatin 受影響最大，atorvastatin 次之；pravastatin、rosuvastatin、pitavastatin 影響很小。吃哪一種藥請問醫師或藥師。
- **紅麴保健食品**含有和 lovastatin 相同的成分，正在吃 statin 時不要自行加吃。
- 飲食調整是在藥物之外加分，不能取代處方藥，也不要因為吃得清淡就自行停藥。
""")

    st.caption("內容依「藥品給付規定」修訂對照表第 2 節（115/9/1 生效）重點整理。")

st.divider()
st.caption("資料來源：衛福部中央健康保險署「全民健康保險降血脂藥物給付規定」修訂對照表（自 115/9/1 生效）。"
           "本工具為臨床速查摘要；實際申報以健保署最新公告及健保藥品代碼為準。")
