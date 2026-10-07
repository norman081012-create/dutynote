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
VERY_HIGH = [
    "ACS 病史",
    "曾血管再通術",
    "特定缺血性中風／TIA",
    "症狀性或曾介入／截肢之 PAD",
    "影像顯示血管狹窄 ≥50%",
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
    st.subheader("病人資料")
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

    st.subheader("1. 臨床 ASCVD")
    with st.container(border=True):
        st.markdown("**極高**：須先有主診斷，再合併下列任一項。只勾主診斷不構成極高風險。")
        ext_hit = []
        for main, subs in EXTREME.items():
            if st.checkbox(f"{main}，再合併下列任一項", key=f"em_{main}"):
                sub_cols = st.columns([1, 20])
                with sub_cols[1]:
                    hits = [x for x in subs if st.checkbox(x, key=f"es_{main}_{x}")]
                if hits:
                    ext_hit += [f"{main[3:]}＋{x}" for x in hits]
                else:
                    sub_cols[1].caption("尚未勾選合併條件，不列入極高風險。")
    with st.container(border=True):
        st.markdown("**非常高**：臨床 ASCVD")
        vh_hit = [x for x in VERY_HIGH if st.checkbox(x, key=f"v_{x}")]

    st.subheader("2. 高風險條件")
    with st.container(border=True):
        h_hit = []
        if st.checkbox("糖尿病"):
            h_hit.append("糖尿病")
        if st.checkbox("透析前 CKD：UACR ≥30 mg/g 或 eGFR <60，持續 ≥3 個月"):
            h_hit.append("透析前 CKD")
        st.caption("已進入透析者不直接列入此項。")
        if st.checkbox("CAC ≥400"):
            h_hit.append("CAC ≥400")
        if ldl is not None and ldl >= 190:
            h_hit.append(f"LDL-C {ldl:.0f} ≥190")
            st.markdown(f"<span class='ng'>LDL-C {ldl:.0f} ≥190，自動列入高風險</span>",
                        unsafe_allow_html=True)

    st.subheader("3. 心血管危險因子")
    with st.container(border=True):
        rf = []
        if st.checkbox("高血壓"):
            rf.append("高血壓")

        age_cut = 45 if sex == "男" else 55
        if age is not None:
            age_rf = age >= age_cut
            st.checkbox(f"年齡（{sex} ≥{age_cut}）", value=age_rf, disabled=True,
                        help="依上方年齡自動判定")
        else:
            age_rf = st.checkbox(f"年齡（{sex} ≥{age_cut}）")
        if age_rf:
            rf.append("年齡")

        if st.checkbox("早發性冠心病家族史（男 ≤55、女 ≤65 歲發病）"):
            rf.append("早發 CHD 家族史")

        hdl_cut = 40 if sex == "男" else 50
        if hdl is not None:
            low_hdl = hdl < hdl_cut
            st.checkbox(f"HDL-C 偏低（{sex} <{hdl_cut}）", value=low_hdl, disabled=True,
                        help="依上方 HDL-C 自動判定")
        else:
            low_hdl = st.checkbox(f"HDL-C 偏低（{sex} <{hdl_cut}）")
        if low_hdl:
            rf.append("HDL-C 偏低")

        if st.checkbox("抽菸"):
            rf.append("抽菸")

        with st.expander("代謝性症候群（5 項中 ≥3 項）"):
            waist = 90 if sex == "男" else 80
            ms = [
                st.checkbox(f"腰圍 {sex} ≥{waist} cm"),
                st.checkbox("BP ≥130/85 或用藥"),
                st.checkbox("空腹血糖 ≥100 或用藥"),
                st.checkbox("TG ≥150 或用藥"),
                low_hdl,
            ]
            st.caption(f"HDL-C {sex} <{hdl_cut}：{'是' if low_hdl else '否'}（同上方 HDL 項）")
            n_ms = sum(ms)
            st.write(f"目前 {n_ms}/5 項")
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
    level, why = "0項", ["無心血管危險因子"]

L = LEVELS[level]

# ---------- 結果 ----------
with right:
    why_txt = "、".join(why)
    if level in ("中", "低"):
        why_txt = f"危險因子 {len(rf)} 項：{why_txt}"
    nonhdl_target = f"&lt; {L['nonhdl']}" if L["nonhdl"] else "未列"
    st.markdown(f"""
<div class="result" style="--c:{L['color']}">
  <h2>{level}{'風險' if level != '0項' else '危險因子'}</h2>
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

st.divider()
st.caption("資料來源：衛福部中央健康保險署「全民健康保險降血脂藥物給付規定」修訂對照表（自 115/9/1 生效）。"
           "本工具為臨床速查摘要；實際申報以健保署最新公告及健保藥品代碼為準。")
