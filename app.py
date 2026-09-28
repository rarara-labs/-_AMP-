import streamlit as st
import pandas as pd
import io
import requests
from datetime import date

# ── GitHub 데이터 파일 URL (업로드 후 수정) ──
DATA_URL = "https://raw.githubusercontent.com/ons-jinju/광레벨_AMP점검/main/26%EB%85%84_%EC%A7%84%EC%A3%BC_%EA%B4%91%EB%A0%88%EB%B2%A8_AMP%EC%A0%90%EA%B2%80.xlsx"

st.set_page_config(
    page_title="광레벨·AMP 점검",
    page_icon="🔦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
html,body,[class*="css"]{font-family:'Noto Sans KR',sans-serif;}

.stTabs [data-baseweb="tab-list"]{gap:1px;flex-wrap:wrap;}
.stTabs [data-baseweb="tab"]{padding:5px 10px;font-size:12px;white-space:nowrap;}

.st{font-size:.88rem;font-weight:700;margin:12px 0 6px;padding-bottom:3px;
    border-bottom:2px solid rgba(128,128,128,.2);}

div[data-testid="metric-container"]{
  background:rgba(128,128,128,.06);
  border:1px solid rgba(128,128,128,.14);
  border-radius:8px;padding:10px 8px;text-align:center;
}
[data-testid="stMetricValue"]{font-size:1.55rem!important;line-height:1.15!important;}
[data-testid="stMetricLabel"]{font-size:.72rem!important;opacity:.75!important;}
[data-testid="stMetricDelta"]{display:none!important;}

.block-container{padding-top:2.5rem!important;}
@media(max-width:768px){
  .block-container{padding:.4rem .5rem!important;padding-top:3rem!important;}
  [data-testid="stMetricValue"]{font-size:1.25rem!important;}
  div[data-testid="metric-container"]{padding:8px 5px;}
}
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════
#  유틸 함수
# ════════════════════════════════════════════
def safe(val):
    if val is None or (isinstance(val, float) and pd.isna(val)): return ''
    return str(val).strip()

def kpi_grid(items):
    n = len(items)
    labels = ''.join(
        f'<div style="font-size:.68rem;opacity:.55;font-weight:500;'
        f'padding-bottom:5px;border-bottom:1px solid rgba(128,128,128,.15)">'
        f'{lbl}</div>'
        for lbl, _, _ in items
    )
    values = ''.join(
        f'<div style="font-size:1.45rem;font-weight:700;color:{clr};padding-top:6px;line-height:1.1">'
        f'{val}</div>'
        for _, val, clr in items
    )
    return (
        f'<div style="background:rgba(128,128,128,.05);border:1px solid rgba(128,128,128,.15);'
        f'border-radius:10px;padding:12px 10px;margin-bottom:10px">'
        f'<div style="display:grid;grid-template-columns:repeat({n},1fr);text-align:center">'
        f'{labels}{values}</div></div>'
    )

def overdue_note(df, date_col, name_col, days=5):
    if df is None or df.empty: return
    if date_col not in df.columns or name_col not in df.columns: return
    today_ts = pd.Timestamp(date.today())
    sample = df[date_col].dropna().iloc[0] if not df[date_col].dropna().empty else None
    if sample is None: return
    try:
        df2 = df.copy()
        df2['_dt'] = pd.to_datetime(df[date_col], errors='coerce')
        df2['_days'] = (today_ts - df2['_dt']).dt.days
    except Exception:
        return
    long = df2[df2['_days'] >= days].copy()
    if long.empty: return
    parts = []
    for _, r in long.head(15).iterrows():
        nm = safe(r.get(name_col, ''))
        d  = int(r['_days']) if pd.notna(r['_days']) else 0
        parts.append(f"{nm}({d}일)")
    note = f"  ＊{days}일↑ 경과: " + " / ".join(parts)
    if len(long) > 15:
        note += f" 외 {len(long)-15}건"
    st.markdown(
        f'<div style="font-size:11px;opacity:.55;margin:-4px 0 8px;line-height:1.6">{note}</div>',
        unsafe_allow_html=True
    )


# ════════════════════════════════════════════
#  데이터 로딩
# ════════════════════════════════════════════
@st.cache_data(ttl=600)
def load_data(url):
    response = requests.get(url)
    response.raise_for_status()
    xl = pd.ExcelFile(io.BytesIO(response.content))

    data = {}
    try:    data['lte_optical'] = pd.read_excel(xl, sheet_name='LTE광레벨불량')
    except: data['lte_optical'] = pd.DataFrame()

    try:    data['5g_optical'] = pd.read_excel(xl, sheet_name='5G광레벨불량')
    except: data['5g_optical'] = pd.DataFrame()

    try:    data['mibos_amp'] = pd.read_excel(xl, sheet_name='MIBOS AMP 미사용 점검', header=1)
    except: data['mibos_amp'] = pd.DataFrame()

    return data


# ════════════════════════════════════════════
#  헤더
# ════════════════════════════════════════════
st.markdown("##### 🏢 진주품질개선팀")
st.markdown("## 광레벨·AMP 점검")

with st.spinner("📡 데이터 로딩 중..."):
    try:
        data = load_data(DATA_URL)
    except Exception as e:
        st.error(f"❌ 데이터 로드 실패: {e}")
        st.stop()

col_r, col_b = st.columns([6, 1])
with col_b:
    if st.button("🔄 새로고침"):
        st.cache_data.clear()
        st.rerun()

st.divider()

tab1, tab2, tab3 = st.tabs(["🔦 LTE 광레벨", "🔦 5G 광레벨", "🔧 MIBOS AMP 점검"])


# ══════════════════════════════════════════════════
#  TAB1  LTE 광레벨 불량
# ══════════════════════════════════════════════════
with tab1:
    df_lopt = data.get('lte_optical', pd.DataFrame())
    if df_lopt.empty:
        st.info("LTE광레벨불량 시트가 없거나 데이터가 없습니다.")
    else:
        df_lopt.columns = [str(c).strip() for c in df_lopt.columns]
        LM = ['RU명','DU TX','DU RX','DU TXBS','RU TX','RU RX','RU TXBS']
        LD = ['DU 파장','du_sfp_vendor','RU 파장','ru_sfp_vendor']
        lm = [c for c in LM if c in df_lopt.columns]
        ld = [c for c in LD if c in df_lopt.columns]

        st.markdown(kpi_grid([("전체", len(df_lopt), '#1971c2')]), unsafe_allow_html=True)

        if '진단결과' in df_lopt.columns:
            opts = ['전체'] + sorted(df_lopt['진단결과'].dropna().astype(str).unique().tolist())
            sel  = st.selectbox("진단결과 필터", opts, key="lopt_diag")
            df_lf = df_lopt if sel == '전체' else df_lopt[df_lopt['진단결과'].astype(str) == sel]
        else:
            df_lf = df_lopt

        # 검색
        s_lte = st.text_input("검색", key="lte_s", placeholder="🔍 RU명 검색", label_visibility="collapsed")
        if s_lte:
            df_lf = df_lf[df_lf.apply(lambda r: s_lte.lower() in str(r).lower(), axis=1)]

        st.caption(f"총 {len(df_lf)}건")

        for _, row in df_lf.iterrows():
            ru_nm = safe(row.get('RU명', '')) or '(이름없음)'
            vals  = '  '.join(
                f'{c}: {safe(row.get(c,""))}'
                for c in lm[1:] if safe(row.get(c,''))
            )
            label_l = f'{ru_nm}  {vals}' if vals else ru_nm

            with st.expander(label_l, expanded=False):
                num_l = []
                for c in lm[1:]:
                    v = safe(row.get(c,''))
                    try:
                        vc = '#f03e3e' if float(v) <= -19 else 'inherit'
                    except:
                        vc = 'inherit'
                    num_l.append(
                        f'<span style="display:inline-block;margin:2px 12px 2px 0;font-size:13px">'
                        f'<span style="opacity:.5;font-size:11px">{c}</span>&nbsp;'
                        f'<b style="color:{vc}">{v}</b></span>'
                    )
                if num_l:
                    st.markdown('<div style="margin-bottom:8px">' + ''.join(num_l) + '</div>', unsafe_allow_html=True)
                if ld:
                    parts_l = []
                    for c in ld:
                        v = safe(row.get(c,'—'))
                        parts_l.append(
                            f'<div style="margin:4px 0;font-size:13px">'
                            f'<span style="opacity:.5;font-size:11px">{c}</span>&nbsp;&nbsp;'
                            f'<b>{v}</b></div>'
                        )
                    st.markdown(''.join(parts_l), unsafe_allow_html=True)


# ══════════════════════════════════════════════════
#  TAB2  5G 광레벨 불량
# ══════════════════════════════════════════════════
with tab2:
    df_5opt = data.get('5g_optical', pd.DataFrame())
    if df_5opt.empty:
        st.info("5G광레벨불량 시트가 없거나 데이터가 없습니다.")
    else:
        df_5opt.columns = [str(c).strip() for c in df_5opt.columns]
        if 'DUL명' not in df_5opt.columns and len(df_5opt.columns) > 3:
            df_5opt = df_5opt.copy()
            df_5opt['DUL명'] = df_5opt.iloc[:, 3].astype(str)

        FM = ['DUL명','DUH TX','DUH RX','DUL TX','DUL RX']
        FD = ['DUH SFP 제조사','DUH SFP 파장','DUL SFP 제조사','DUL SFP 파장']
        fm = [c for c in FM if c in df_5opt.columns]
        fd = [c for c in FD if c in df_5opt.columns]

        st.markdown(kpi_grid([("전체", len(df_5opt), '#7048e8')]), unsafe_allow_html=True)

        if '진단결과' in df_5opt.columns:
            opts5 = ['전체'] + sorted(df_5opt['진단결과'].dropna().astype(str).unique().tolist())
            sel5  = st.selectbox("진단결과 필터", opts5, key="5opt_diag")
            df_5f = df_5opt if sel5 == '전체' else df_5opt[df_5opt['진단결과'].astype(str) == sel5]
        else:
            df_5f = df_5opt

        s_5g = st.text_input("검색", key="5g_s", placeholder="🔍 DUL명 검색", label_visibility="collapsed")
        if s_5g:
            df_5f = df_5f[df_5f.apply(lambda r: s_5g.lower() in str(r).lower(), axis=1)]

        st.caption(f"총 {len(df_5f)}건")

        for _, row in df_5f.iterrows():
            dul_nm = safe(row.get('DUL명','')) or '(이름없음)'
            vals5  = '  '.join(
                f'{c}: {safe(row.get(c,""))}'
                for c in fm[1:] if safe(row.get(c,''))
            )
            label5 = f'{dul_nm}  {vals5}' if vals5 else dul_nm

            with st.expander(label5, expanded=False):
                num5 = []
                for c in fm[1:]:
                    v = safe(row.get(c,''))
                    try:
                        vc = '#f03e3e' if float(v) <= -19 else 'inherit'
                    except:
                        vc = 'inherit'
                    num5.append(
                        f'<span style="display:inline-block;margin:2px 12px 2px 0;font-size:13px">'
                        f'<span style="opacity:.5;font-size:11px">{c}</span>&nbsp;'
                        f'<b style="color:{vc}">{v}</b></span>'
                    )
                if num5:
                    st.markdown('<div style="margin-bottom:8px">' + ''.join(num5) + '</div>', unsafe_allow_html=True)
                if fd:
                    parts5 = []
                    for c in fd:
                        v = safe(row.get(c,'—'))
                        parts5.append(
                            f'<div style="margin:4px 0;font-size:13px">'
                            f'<span style="opacity:.5;font-size:11px">{c}</span>&nbsp;&nbsp;'
                            f'<b>{v}</b></div>'
                        )
                    st.markdown(''.join(parts5), unsafe_allow_html=True)


# ══════════════════════════════════════════════════
#  TAB3  MIBOS AMP 미사용 점검
# ══════════════════════════════════════════════════
with tab3:
    df_amp = data.get('mibos_amp', pd.DataFrame())

    if df_amp is None or df_amp.empty:
        st.warning("MIBOS AMP 미사용 점검 데이터가 없습니다.")
    else:
        df_amp = df_amp.dropna(how='all').copy()
        total_amp = len(df_amp)

        kpi_items = [("전체", total_amp, '#1971c2')]

        _nk_col = next((c for c in df_amp.columns if '누적' in str(c)), None)
        if _nk_col:
            nk_vals = df_amp[_nk_col].astype(str)
            _nk_counts = nk_vals.value_counts()
            for _v, _clr in [('대상','#f03e3e'),('사용전환','#2f9e44'),('제외','#868e96')]:
                if _v in _nk_counts:
                    kpi_items.append((_v, int(_nk_counts[_v]), _clr))

        _date_col = next((c for c in df_amp.columns if '조치일' in str(c) and '예' not in str(c)), None)
        if _date_col:
            _done = int(df_amp[_date_col].notna().sum())
            _todo = total_amp - _done
            kpi_items.append(("조치완료", _done, '#2f9e44'))
            kpi_items.append(("미조치",   _todo, '#e67700'))

        st.markdown(kpi_grid(kpi_items), unsafe_allow_html=True)

        if _date_col:
            _nm_col = next((c for c in df_amp.columns if '중계기' in str(c) and '명' in str(c)), None) or \
                      next((c for c in df_amp.columns if '명' in str(c)), None)
            if _nm_col:
                overdue_note(df_amp[df_amp[_date_col].isna()], _date_col, _nm_col, days=5)

        cm3, cf3 = st.columns([3, 1])
        search_amp = cm3.text_input("검색", key="t3_s",
                                    placeholder="🔍 중계기명·조치내역 검색",
                                    label_visibility="collapsed")
        flt_opts = ['전체']
        if _nk_col:
            flt_opts += sorted(df_amp[_nk_col].dropna().astype(str).unique().tolist())
        flt_amp = cf3.selectbox("누적관리", flt_opts, key="t3_f", label_visibility="collapsed")

        df_amp_f = df_amp.copy()
        if search_amp:
            df_amp_f = df_amp_f[df_amp_f.apply(
                lambda r: search_amp.lower() in str(r).lower(), axis=1)]
        if flt_amp != '전체' and _nk_col:
            df_amp_f = df_amp_f[df_amp_f[_nk_col].astype(str) == flt_amp]

        _want_cols = ['운용팀', '중계기 명', '중계기명', '중계기 모델', '중계기모델',
                      '서비스주파수', '조치내역', '조치일', '조치예정일', '누적관리']
        show_cols_amp = [c for c in df_amp_f.columns if c in _want_cols]
        if not show_cols_amp:
            show_cols_amp = [c for c in df_amp_f.columns if not str(c).startswith('Unnamed')]

        st.caption(f"표시 {len(df_amp_f)}건")
        st.dataframe(df_amp_f[show_cols_amp], use_container_width=True, height=500)
