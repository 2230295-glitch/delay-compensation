"""
지체보상금 대시보드 v8 — 클린 화이트 디자인
"""
import io, re, calendar, zipfile

try:
    from docx import Document as _DocxDoc
    from docx.shared import Pt as _Pt, Cm as _Cm, RGBColor as _RGB
    from docx.enum.text import WD_ALIGN_PARAGRAPH as _WDA
    from docx.oxml.ns import qn as _dqn
    from docx.oxml import OxmlElement as _dxml
    _HAS_DOCX = True
except ImportError:
    _HAS_DOCX = False
from datetime import date
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as gcl

st.set_page_config(page_title="지체보상금 대시보드", layout="wide", page_icon="📊",
                   initial_sidebar_state="collapsed")

# Pretendard — 프리텐다드 (한국 기업용 표준 폰트)
st.markdown("""
<link rel="preconnect" href="https://cdn.jsdelivr.net">
<link rel="stylesheet" as="style" crossorigin
  href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css">
""", unsafe_allow_html=True)

st.markdown("""<style>
/* ── 폰트 전역 ── */
*, *::before, *::after,
html, body, [class*="css"],
.stApp, button, input, select, textarea,
[data-testid], [data-baseweb],
.stMarkdown, .stDataFrame, p, span, div, th, td, label {
  font-family: 'Pretendard', 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif !important;
  -webkit-font-smoothing: antialiased;
  letter-spacing: -.015em;
}

/* ── 기본 레이아웃 ── */
.stApp { background: #f4f6fb !important; }
header[data-testid="stHeader"]  { background: transparent !important; height: 0 !important; min-height: 0 !important; }
[data-testid="stToolbar"]        { display: none !important; }
[data-testid="stDecoration"]     { display: none !important; }
[data-testid="stStatusWidget"]   { display: none !important; }
#MainMenu                        { display: none !important; }
section[data-testid="stSidebar"] { display: none !important; }
.stMainBlockContainer, .block-container {
  max-width: 100% !important;
  padding: 0 2rem 4rem !important;
}

/* ── 최상단 헤더 ── */
.page-header {
  background: #ffffff;
  border-bottom: 1px solid #e8eaf0;
  padding: .85rem 2rem; margin: 0 -2rem 1.5rem;
  display: flex; align-items: center; justify-content: space-between;
}
.ph-left { display: flex; flex-direction: column; gap: .18rem; }
.ph-brand {
  display: flex; align-items: baseline; gap: .5rem;
}
.ph-company { font-size: .82rem; font-weight: 700; color: #1a56db; letter-spacing: 0; }
.ph-title   { font-size: 1.12rem; font-weight: 800; color: #111827; }
.ph-channel {
  display: inline-block; background: #fff7ed; color: #c2410c;
  font-size: .7rem; font-weight: 700; padding: .12rem .5rem;
  border-radius: 4px; border: 1px solid #fed7aa; margin-left: .3rem;
  vertical-align: middle;
}
.ph-sub { font-size: .72rem; color: #9ca3af; }
.ph-right { display: flex; align-items: center; gap: .6rem; }
.ph-date {
  font-size: .75rem; color: #6b7280;
  background: #f9fafb; border: 1px solid #e5e7eb;
  padding: .28rem .75rem; border-radius: 5px;
}

/* ── 섹션 제목 ── */
.sec-title {
  display: flex; align-items: center; gap: .5rem;
  font-size: .88rem; font-weight: 700; color: #111827;
  margin: 1.1rem 0 .7rem; padding-bottom: .55rem;
  border-bottom: 1.5px solid #e5e7eb;
}
.sec-accent { width: 3px; height: 14px; border-radius: 2px; background: #f59e0b; flex-shrink: 0; }
.sec-sub { font-size: .72rem; color: #9ca3af; font-weight: 400; margin-left: .3rem; }

/* ── KPI 카드 ── */
.kpi-grid { display: grid; grid-template-columns: repeat(4,1fr); gap: .75rem; margin-bottom: 1.1rem; }
.kpi-card {
  background: #fff; border-radius: 8px;
  padding: 1rem 1.2rem .9rem;
  border: 1px solid #e8eaf0;
  border-bottom: 3px solid #e8eaf0;
  transition: box-shadow .15s;
}
.kpi-card:hover { box-shadow: 0 4px 14px rgba(0,0,0,.09); }
.kpi-card.c-amber  { border-bottom-color: #f5a000; }
.kpi-card.c-red    { border-bottom-color: #ef4444; }
.kpi-card.c-blue   { border-bottom-color: #3b82f6; }
.kpi-card.c-green  { border-bottom-color: #10b981; }
.kpi-card.c-indigo { border-bottom-color: #6366f1; }
.kpi-card.c-violet { border-bottom-color: #8b5cf6; }

.kpi-lbl  { font-size: .7rem; font-weight: 600; color: #9ca3af; margin-bottom: .35rem; letter-spacing: .02em; text-transform: uppercase; }
.kpi-val  { font-size: 2.6rem; font-weight: 800; color: #111827; line-height: 1; letter-spacing: -.03em; }
.kpi-val.c-amber { color: #f5a000; }
.kpi-val.c-red   { color: #ef4444; }
.kpi-val.c-blue  { color: #3b82f6; }
.kpi-unit { font-size: 1.1rem; font-weight: 700; margin-left: .1rem; }
.kpi-delta { font-size: .72rem; margin-top: .4rem; color: #9ca3af; line-height: 1.5; }
.up   { color: #ef4444 !important; font-weight: 700; }
.dn   { color: #3b82f6 !important; font-weight: 700; }
.neu  { color: #9ca3af !important; }

/* ── 탭 ── */
[data-baseweb="tab-list"] {
  background: #fff !important;
  border: 1px solid #e5e7eb !important; border-bottom: none !important;
  border-radius: 10px 10px 0 0 !important;
  padding: 0 .6rem !important; gap: 0 !important;
}
[data-baseweb="tab"] {
  font-size: .87rem !important; font-weight: 600 !important;
  padding: .72rem 1.5rem !important; color: #9ca3af !important;
  border-bottom: 3px solid transparent !important;
  background: transparent !important;
  font-family: 'Pretendard', sans-serif !important;
}
[data-baseweb="tab"][aria-selected="true"] {
  color: #f59e0b !important;
  border-bottom-color: #f59e0b !important;
  font-weight: 800 !important;
}
[data-baseweb="tab-panel"] {
  background: #fff !important;
  border: 1px solid #e5e7eb !important; border-top: none !important;
  border-radius: 0 0 10px 10px !important;
  padding: 1.3rem 1.5rem 1.5rem !important;
}

/* ── 연도/월 선택기 ── */
.picker-wrap {
  display: flex; align-items: flex-start; gap: 1.5rem;
  background: #f9fafb; border: 1px solid #e5e7eb;
  border-radius: 8px; padding: .8rem 1rem; margin-bottom: 1.1rem;
}
.picker-label {
  font-size: .68rem; font-weight: 700; color: #9ca3af;
  letter-spacing: .07em; text-transform: uppercase;
  margin-bottom: .38rem;
}

/* 연도/월 버튼 스타일 덮어쓰기 */
.stButton > button {
  font-family: 'Pretendard', sans-serif !important;
  font-weight: 600 !important; border-radius: 6px !important;
  font-size: .83rem !important; transition: all .12s !important;
}
button[kind="primary"] {
  background: #f59e0b !important; border-color: #f59e0b !important;
  color: #fff !important; font-weight: 700 !important;
}
button[kind="secondary"] {
  background: #fff !important; border-color: #e5e7eb !important;
  color: #374151 !important;
}
button[kind="secondary"]:hover {
  background: #fef9ee !important; border-color: #f59e0b !important;
  color: #92400e !important;
}

/* ── 다운로드 버튼 ── */
.stDownloadButton > button {
  background: #1f2937 !important; color: #fff !important;
  border: none !important; border-radius: 6px !important;
  font-weight: 600 !important; font-size: .82rem !important;
  font-family: 'Pretendard', sans-serif !important;
  padding: .5rem 1rem !important;
  transition: background .12s !important;
}
.stDownloadButton > button:hover { background: #374151 !important; }

/* ── 변동 카드 ── */
.chg-block {
  background: #fff; border-radius: 8px;
  border: 1px solid #e5e7eb;
  overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,.04);
}
.chg-block .chg-inner { padding: .85rem 1rem; }
.chg-block.new  .chg-head { background: #fef2f2; border-bottom: 1px solid #fecaca; padding: .65rem 1rem; }
.chg-block.add  .chg-head { background: #fffbeb; border-bottom: 1px solid #fde68a; padding: .65rem 1rem; }
.chg-block.done .chg-head { background: #f0fdf4; border-bottom: 1px solid #bbf7d0; padding: .65rem 1rem; }
.chg-head {
  display: flex; align-items: baseline; gap: .4rem;
}
.chg-name { font-size: .85rem; font-weight: 700; }
.chg-cnt  { font-size: .73rem; color: #6b7280; }
.chg-item {
  display: flex; justify-content: space-between; align-items: center;
  padding: .34rem 0; border-bottom: 1px solid #f3f4f6;
  font-size: .82rem; gap: .5rem;
}
.chg-nm { color: #374151; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; flex: 1; }
.chg-org { font-size: .7rem; color: #9ca3af; white-space: nowrap; }
.chg-amt { font-weight: 700; white-space: nowrap; flex-shrink: 0; }
.chg-none { color: #d1d5db; font-size: .82rem; padding: 1.2rem 0; text-align: center; }

/* ── 랭킹 ── */
.rank-wrap { padding: .1rem 0; }
.rank-row {
  display: grid; grid-template-columns: 1.5rem 1fr 5.5rem;
  align-items: center; gap: .65rem;
  padding: .5rem 0; border-bottom: 1px solid #f3f4f6;
}
.rank-no {
  width: 1.3rem; height: 1.3rem; border-radius: 50%;
  display: flex; align-items: center; justify-content: center;
  font-size: .7rem; font-weight: 800; color: #fff; flex-shrink: 0;
}
.rank-name { font-size: .83rem; font-weight: 600; color: #111827; }
.rank-org  { font-size: .69rem; color: #9ca3af; margin-top: .1rem; }
.rank-bar-bg { background: #f3f4f6; border-radius: 2px; height: 5px; margin-top: .22rem; overflow: hidden; }
.rank-bar-fg { height: 100%; border-radius: 2px; }
.rank-amt { text-align: right; font-size: .86rem; font-weight: 700; }

/* ── 집계 테이블 ── */
.agg-tbl { width: 100%; border-collapse: collapse; font-size: .83rem; }
.agg-tbl thead tr { background: #1f2937; }
.agg-tbl th {
  background: #1f2937; font-weight: 700; font-size: .73rem;
  color: #e5e7eb; padding: .7rem .9rem; text-align: right;
  white-space: nowrap;
}
.agg-tbl th:first-child { text-align: left; border-radius: 6px 0 0 0; }
.agg-tbl th:last-child  { border-radius: 0 6px 0 0; }
.agg-tbl td {
  padding: .55rem .9rem; border-bottom: 1px solid #f3f4f6;
  color: #374151; text-align: right; white-space: nowrap;
}
.agg-tbl td:first-child { text-align: left; font-weight: 600; color: #111827; }
.agg-tbl tr:last-child td {
  background: #1f2937; font-weight: 800; color: #fff; border-top: none;
}
.agg-tbl tr:nth-child(even) td { background: #fafafa; }
.agg-tbl tr:hover td { background: #fff7ed !important; }
.agg-tbl .has-val { color: #c2410c; font-weight: 700; }
.agg-tbl .zero    { color: #d1d5db; }

/* ── dataframe ── */
[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
[data-testid="stDataFrame"] th {
  background: #f9fafb !important; font-size: .78rem !important;
  font-weight: 700 !important; color: #6b7280 !important;
}

/* ── 배지 ── */
.tag {
  display: inline-block; font-size: .68rem; font-weight: 600;
  padding: .08rem .42rem; border-radius: 3px; margin-left: .3rem;
  vertical-align: middle;
}
.tag-gray   { background: #f3f4f6; color: #6b7280; }
.tag-amber  { background: #fff7ed; color: #c2410c; border: 1px solid #fed7aa; }
.tag-blue   { background: #eff6ff; color: #1d4ed8; border: 1px solid #bfdbfe; }

/* ── 구분선 ── */
.divider { height: 1px; background: #e5e7eb; margin: 1rem 0; }
</style>""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────
ALLOW_CH       = {"DW04","DW07","DW08","DW09","DW10","DW11","DW17"}
EXCLUDE_CODES  = {"10014461","C4900","C2000","C3100"}
EXCLUDE_BU     = {"기타사업부","수탁사업부","완제사업부"}
EXCLUDE_OFFICE = {"DW_B2B","DW_CH신사업팀"}
EXCLUDE_MGR    = {"1202150261"}
FONT           = "Pretendard,Apple SD Gothic Neo,Malgun Gothic,sans-serif"

def _num(v):
    try: return float(str(v).replace(",","").strip())
    except: return float("nan")

def fmt_won(v, short=False):
    v = int(v)
    if short:
        if abs(v) >= 100_000_000: return f"{v/100_000_000:.1f}억"
        if abs(v) >= 10_000:      return f"{round(v/10_000):,}만"
        return f"{v:,}"
    if abs(v) >= 100_000_000: return f"{v/100_000_000:.2f}억원"
    if abs(v) >= 10_000:      return f"{round(v/10_000):,}만원"
    return f"{v:,}원"

def mo_label(s):
    y, m = s.split("-")
    return f"{y[2:]}년 {int(m)}월"

def _norm(s):
    s = re.sub(r'[\s\(\)\.\-·\']', '', str(s)).upper()
    for p in ['의료법인','재단법인','학교법인','사회복지법인']:
        if s.startswith(p): s = s[len(p):]
    return s

# ── 비즈니스 로직 함수들 (변경 없음) ─────────────────────────
def load_contract(file_bytes, sap_df):
    try:    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name='계약서', dtype=str, header=5)
    except: df = pd.read_excel(io.BytesIO(file_bytes), dtype=str, header=5)
    cols = ['No','병원명','계약유형','사업자번호','계약일자','계약병상','계약금액','계약기간',
            '_8','대금지급','지체보상금조항','비고','_12']
    df.columns = cols + [f'_e{i}' for i in range(max(0, len(df.columns)-len(cols)))]
    df = df[df['No'].notna() & df['No'].astype(str).str.match(r'^\d+$', na=False)].copy()
    df['지체여부'] = df['지체보상금조항'].apply(lambda x: '지체보상금' in str(x))
    def _rate(x):
        m = re.search(r'\[(\d+)/(\d+)\]', str(x))
        return float(m.group(1)) / float(m.group(2)) if m else 1/1000
    df['요율'] = df['지체보상금조항'].apply(_rate)
    df['_norm'] = df['병원명'].apply(_norm)
    s2 = sap_df.copy(); s2['_norm'] = s2['판매처명'].apply(_norm)
    sm = dict(zip(s2['_norm'], s2['판매처']))
    df['코드'] = df['_norm'].map(sm)
    for idx, row in df[df['코드'].isna()].iterrows():
        cn = row['_norm']
        if len(cn) < 4: continue
        for sn, sc in zip(s2['_norm'], s2['판매처']):
            if len(sn) >= 4 and (cn in sn or sn in cn):
                df.at[idx, '코드'] = sc; break
    result = {}
    for _, r in df[df['코드'].notna()].iterrows():
        c = r['코드']
        if c not in result or not r['지체여부']:
            result[c] = {"지체여부": r['지체여부'], "요율": r['요율'],
                         "계약유형": str(r.get('계약유형', '')).strip(),
                         "사업자번호": str(r.get('사업자번호', '')).strip()}
    n_yes = sum(1 for v in result.values() if v["지체여부"])
    n_no  = sum(1 for v in result.values() if not v["지체여부"])
    return result, n_yes, n_no

def load_base(file_bytes, cmap=None, think_rotations=None):
    df = pd.read_excel(io.BytesIO(file_bytes), dtype=str).dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]
    out = []
    for _, r in df.iterrows():
        code = str(r.get("판매처","")).strip(); name = str(r.get("판매처명","")).strip()
        base = _num(r.get("기준회전일",0)); yak = _num(r.get("약정회전일",0))
        exc  = _num(r.get("예외회전일",0))
        if not code or code == "nan" or pd.isna(base) or base <= 0: continue
        eff = (int(exc) if not pd.isna(exc) and exc > 0
               else (int(yak) if not pd.isna(yak) and yak > 0 else int(base)))
        info = (cmap or {}).get(code, {})
        use = info.get("지체여부", True); rate = info.get("요율", 1/1000)
        out.append({"코드":code,"명칭":name,"기준회전일":int(base),
                    "유효기준회전일":eff,"지체여부":use,"요율":rate})
    existing = {r["코드"] for r in out}
    if think_rotations:
        for code, rot in think_rotations.items():
            if code not in existing:
                info = (cmap or {}).get(code, {})
                use = info.get("지체여부", True); rate = info.get("요율", 1/1000)
                out.append({"코드":code,"명칭":"","기준회전일":rot,"유효기준회전일":rot,"지체여부":use,"요율":rate})
    return pd.DataFrame(out)

def _parse_rot(v):
    s = str(v).strip()
    if '개월' in s:
        try: return int(s.replace('개월','').strip()) * 30
        except: return 30
    if '일' in s:
        try: return int(s.replace('일','').strip())
        except: return 30
    try: return int(s)
    except: return 30

def load_think(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name='씽크현황', header=None, dtype=str)
    codes = {}
    for _, r in df.iloc[4:].iterrows():
        code = str(r.iloc[2]).strip()
        if not code or code in ('nan', '판매처', 'None'): continue
        codes[code] = _parse_rot(r.iloc[8])
    return codes

def load_think_monthly(file_bytes):
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name='씽크현황', header=None, dtype=str)
    row2 = df.iloc[2].tolist(); row3 = df.iloc[3].tolist()
    mo_cols = {}; cur = None
    for i, (r2, r3) in enumerate(zip(row2, row3)):
        s2 = str(r2).strip()
        if '.' in s2 and s2 not in ('nan','None'):
            cur = s2; mo_cols[cur] = {}
        if cur and str(r3).strip() in ('매출','수금') and str(r3).strip() not in mo_cols[cur]:
            mo_cols[cur][str(r3).strip()] = i
    valid = {k: v for k, v in mo_cols.items() if '합계' not in k and '계약' not in k and len(k) == 7}
    months_dict = {}
    for _, row in df.iloc[4:].iterrows():
        code = str(row.iloc[2]).strip(); name = str(row.iloc[3]).strip()
        if not code or code in ('nan','None','판매처'): continue
        기준 = _parse_rot(row.iloc[8])
        cum_sale = cum_coll = 0.0; prev_현회전일 = 0
        for mo_str, cidx in sorted(valid.items()):
            try: yr, mo = int(mo_str.split('.')[0]), int(mo_str.split('.')[1])
            except: continue
            def _v(col):
                try:
                    s = str(row.iloc[col]).replace(',','').strip()
                    if s.lower() in ('nan','none',''): return 0.0
                    return float(s) * 1_000_000
                except: return 0.0
            sale = _v(cidx['매출']) if '매출' in cidx else 0.0
            coll = _v(cidx['수금']) if '수금' in cidx else 0.0
            cum_sale += sale; cum_coll += coll; 잔고 = cum_sale - cum_coll
            days = calendar.monthrange(yr, mo)[1]
            if 잔고 <= 0: prev_현회전일 = 0; continue
            현회전일 = 1 if sale > 0 else prev_현회전일 + days
            prev_현회전일 = 현회전일
            key = (yr, mo)
            if key not in months_dict: months_dict[key] = []
            months_dict[key].append({"연":yr,"월":mo,"코드":code,"명칭":name,"채널":"씽크",
                                     "잔고":int(잔고),"현회전일":현회전일,
                                     "_기준회전일_override":기준,"사무소":"","사업부":"","담당자":""})
    return [{"yr":yr,"mo":mo,"data":pd.DataFrame(rows)} for (yr,mo),rows in sorted(months_dict.items())]

def load_monthly(file_bytes, think_codes=None):
    df = pd.read_excel(io.BytesIO(file_bytes), dtype=str).dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]
    def _ym(v):
        s = str(v).strip()
        return (int(s.split(".")[0]), int(s.split(".")[1])) if "." in s else (None, None)
    df["_yr"],df["_mo"] = zip(*df["년월"].map(_ym))
    df = df[df["_yr"].notna()].copy()
    df["_yr"]=df["_yr"].astype(int); df["_mo"]=df["_mo"].astype(int)
    df["_잔고"]=df["순실잔고"].apply(_num); df["_일수"]=df["순실회전일"].apply(_num)
    df=df.dropna(subset=["_잔고","_일수"]).copy(); df["_일수"]=df["_일수"].astype(int)
    df=df[df["판매처"].notna()&(df["판매처"].str.strip()!="")&(df["판매처"]!="nan")]
    df=df[df["채널"].isin(ALLOW_CH)]; df=df[df["_잔고"]!=0]
    if think_codes and "채널" in df.columns:
        df=df[~(df["채널"]=="DW07")|df["판매처"].isin(think_codes)]
    if "담당자"   in df.columns: df=df[~df["담당자"].isin(EXCLUDE_MGR)]
    if "사업부명" in df.columns: df=df[~df["사업부명"].apply(lambda x:any(k in str(x) for k in EXCLUDE_BU))]
    df=df[~df["판매처"].isin(EXCLUDE_CODES)]
    if "사무소명" in df.columns: df=df[~df["사무소명"].isin(EXCLUDE_OFFICE)]
    months=[]
    for (yr,mo),grp in df.groupby(["_yr","_mo"]):
        rows=[{"연":int(yr),"월":int(mo),"코드":str(r["판매처"]).strip(),
               "명칭":str(r.get("판매처명","")).strip(),"채널":str(r.get("채널","")).strip(),
               "잔고":r["_잔고"],"현회전일":r["_일수"],
               "사무소":str(r.get("사무소명","")).strip() if "사무소명" in r.index else "",
               "사업부":str(r.get("사업부명","")).strip() if "사업부명" in r.index else "",
               "담당자":str(r.get("담당자명","")).strip() if "담당자명" in r.index else "",
               } for _,r in grp.iterrows()]
        months.append({"yr":int(yr),"mo":int(mo),"data":pd.DataFrame(rows)})
    months.sort(key=lambda x:(x["yr"],x["mo"])); return months

def calculate(months, base_df):
    bmap={r["코드"]:r for _,r in base_df.iterrows()}
    cum:dict={}; results=[]
    for m in months:
        yr,mo=m["yr"],m["mo"]
        for _,row in m["data"].iterrows():
            code=row["코드"]
            if code not in bmap: continue
            b=bmap[code]
            if not b["지체여부"]: continue
            기준=int(row["_기준회전일_override"]) if "_기준회전일_override" in row and pd.notna(row["_기준회전일_override"]) else b["유효기준회전일"]
            요율=b["요율"]; 현=row["현회전일"]; 잔=row["잔고"]
            지연=max(0,현-기준); prev=cum.get(code,0)
            if 지연<=0: cum[code]=0; 증분=0
            else: 증분=max(0,지연-prev); cum[code]=max(prev,지연)
            보상=round(잔*증분*요율) if 증분>0 else 0
            rate_str=f"1/{int(round(1/요율))}" if 요율>0 else "-"
            results.append({"기준월":f"{yr}-{mo:02d}","판매처코드":code,"판매처명":row["명칭"],
                "채널":row["채널"],"사무소":row.get("사무소",""),"사업부":row.get("사업부",""),"담당자":row.get("담당자",""),
                "현 미수금":int(잔),"기준회전일(SAP)":b["기준회전일"],"유효기준회전일":기준,
                "현 회전일":현,"지연일수(누적)":지연,"청구 증분일수":증분,"지체보상금":보상,"요율":rate_str,
                "비고":"청구" if 보상>0 else("정상화" if 지연==0 and prev>0 else "")})
    return pd.DataFrame(results)

def make_excel(df, title="전체내역"):
    wb=openpyxl.Workbook(); HDR=PatternFill("solid",fgColor="1f2937")
    def _h(ws):
        for c in ws[1]:
            c.font=Font(bold=True,color="FFFFFF",size=10)
            c.fill=HDR; c.alignment=Alignment(horizontal="center",vertical="center")
    w={"기준월":10,"판매처코드":12,"판매처명":22,"채널":10,"현 미수금":14,
       "기준회전일(SAP)":12,"유효기준회전일":12,"현 회전일":10,"지연일수(누적)":12,
       "청구 증분일수":12,"지체보상금":14,"요율":8,"비고":8}
    ws=wb.active; ws.title=title
    cols=list(df.columns); ws.append(cols)
    for r in df.itertuples(index=False): ws.append(list(r))
    _h(ws)
    for ci,cn in enumerate(cols,1):
        for ri in range(2,ws.max_row+1):
            c=ws.cell(ri,ci)
            if cn in {"현 미수금","지체보상금"}: c.number_format="#,##0"
            c.alignment=Alignment(horizontal="center")
        ws.column_dimensions[gcl(ci)].width=w.get(cn,12)
    ws.freeze_panes="A2"
    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

def make_full_excel(df):
    wb=openpyxl.Workbook(); HDR=PatternFill("solid",fgColor="1f2937")
    def _h(ws):
        for c in ws[1]:
            c.font=Font(bold=True,color="FFFFFF",size=10)
            c.fill=HDR; c.alignment=Alignment(horizontal="center",vertical="center")
    w={"기준월":10,"판매처코드":12,"판매처명":22,"채널":10,"현 미수금":14,
       "기준회전일(SAP)":12,"유효기준회전일":12,"현 회전일":10,"지연일수(누적)":12,
       "청구 증분일수":12,"지체보상금":14,"요율":8,"비고":8}
    ws1=wb.active; ws1.title="전체내역"
    cols=list(df.columns); ws1.append(cols)
    for r in df.itertuples(index=False): ws1.append(list(r))
    _h(ws1)
    for ci,cn in enumerate(cols,1):
        for ri in range(2,ws1.max_row+1):
            c=ws1.cell(ri,ci)
            if cn in {"현 미수금","지체보상금"}: c.number_format="#,##0"
            c.alignment=Alignment(horizontal="center")
        ws1.column_dimensions[gcl(ci)].width=w.get(cn,12)
    ws1.freeze_panes="A2"
    ws2=wb.create_sheet("거래처별누계")
    pv=df.pivot_table(index=["판매처코드","판매처명"],columns="기준월",
                      values="지체보상금",aggfunc="sum",fill_value=0).reset_index()
    pc=list(pv.columns); ws2.append(pc)
    for r in pv.itertuples(index=False): ws2.append(list(r))
    _h(ws2)
    for row in ws2.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value,(int,float)) and c.value!=0: c.number_format="#,##0"
            c.alignment=Alignment(horizontal="center")
    ws2.column_dimensions["A"].width=12; ws2.column_dimensions["B"].width=22
    for i in range(3,len(pc)+1): ws2.column_dimensions[gcl(i)].width=14
    ws2.freeze_panes="C2"
    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

def _공문_doc_setup(doc):
    sec0=doc.sections[0]
    sec0.top_margin=_Cm(1.3); sec0.bottom_margin=_Cm(1.0)
    sec0.left_margin=_Cm(1.8); sec0.right_margin=_Cm(1.8)
    def _rf(run,name='맑은 고딕'):
        rPr=run._r.get_or_add_rPr()
        for t in list(rPr):
            if t.tag==_dqn('w:rFonts'): rPr.remove(t)
        rf=_dxml('w:rFonts')
        rf.set(_dqn('w:ascii'),name); rf.set(_dqn('w:eastAsia'),name); rf.set(_dqn('w:hAnsi'),name)
        rPr.insert(0,rf)
    def _r(para,text,bold=False,size=10,color=None,name='맑은 고딕'):
        run=para.add_run(text); run.bold=bold; run.font.size=_Pt(size); run.font.name=name
        if color: run.font.color.rgb=_RGB.from_string(color)
        _rf(run,name); return run
    def _p(text='',bold=False,size=10,align=_WDA.LEFT,sa=2,color=None):
        p=doc.add_paragraph(); p.alignment=align
        p.paragraph_format.space_after=_Pt(sa); p.paragraph_format.space_before=_Pt(0)
        if text: _r(p,text,bold=bold,size=size,color=color)
        return p
    def _bline(side='bottom',sa=6):
        p=doc.add_paragraph()
        p.paragraph_format.space_after=_Pt(sa); p.paragraph_format.space_before=_Pt(0)
        pPr=p._p.get_or_add_pPr(); pBdr=_dxml('w:pBdr'); bdr=_dxml(f'w:{side}')
        bdr.set(_dqn('w:val'),'single'); bdr.set(_dqn('w:sz'),'18')
        bdr.set(_dqn('w:space'),'1'); bdr.set(_dqn('w:color'),'1A2A6C')
        pBdr.append(bdr); pPr.append(pBdr)
    def _tc(cell,text,bold=False,size=9.5,color=None,fill=None,align=_WDA.LEFT):
        para=cell.paragraphs[0]; run=para.add_run(str(text))
        run.bold=bold; run.font.size=_Pt(size); run.font.name='맑은 고딕'
        if color: run.font.color.rgb=_RGB.from_string(color)
        _rf(run); para.alignment=align
        if fill:
            tcPr=cell._tc.get_or_add_tcPr()
            for tag in list(tcPr):
                if tag.tag==_dqn('w:shd'): tcPr.remove(tag)
            shd=_dxml('w:shd')
            shd.set(_dqn('w:val'),'clear'); shd.set(_dqn('w:color'),'auto'); shd.set(_dqn('w:fill'),fill)
            tcPr.append(shd)
    def _sig():
        p1=_p(sa=0,align=_WDA.RIGHT); _r(p1,'서울시 강남구 봉은사로 114길 12',size=9.5)
        p2=_p(sa=0,align=_WDA.RIGHT); _r(p2,'주 식 회 사  대 웅 제 약',size=10)
        p3=_p(sa=3,align=_WDA.RIGHT); _r(p3,'대 표 이 사   이 창 재',size=10)
    def _footer():
        _bline('top',sa=4)
        _p('DAEWOONG PHARMACEUTICAL CO., LTD.',bold=True,size=10,align=_WDA.CENTER,sa=0)
        _p('대웅제약 주식회사',size=8.5,align=_WDA.CENTER,sa=0)
    def _letterhead(판매처명,제목,doc_no_suffix):
        from datetime import date as _d; today=_d.today()
        p_lh=_p(sa=1); _r(p_lh,'(주)  대 웅 제 약',bold=True,size=20)
        _bline('bottom',sa=3)
        th=doc.add_table(rows=1,cols=2)
        for cell in th.rows[0].cells: cell._tc.get_or_add_tcPr()
        _tc(th.rows[0].cells[0],f'문서번호 : {doc_no_suffix}')
        _tc(th.rows[0].cells[1],today.strftime("%Y년 %m월 %d일"),align=_WDA.RIGHT)
        _p(sa=2)
        p=_p(sa=1); _r(p,'수  신',bold=True,size=10.5); _r(p,f'  :  {판매처명} 귀중',size=10.5)
        p=_p(sa=3); _r(p,'제  목',bold=True,size=10.5); _r(p,f'  :  {제목}',bold=True,size=10.5)
        _bline('top',sa=5); return today
    return _r,_p,_bline,_tc,_sig,_footer,_letterhead

def make_공문_word(row):
    if not _HAS_DOCX: raise ImportError("pip install python-docx")
    doc=_DocxDoc()
    _r,_p,_bline,_tc,_sig,_footer,_letterhead=_공문_doc_setup(doc)
    판매처명=str(row.get("판매처명","")); 기준월=str(row.get("기준월",""))
    _letterhead(판매처명,"지체보상금 청구의 건",f"지체-{기준월.replace('-','')}-{판매처명[:4]}")
    _p('1.',bold=True,size=10,sa=2); p=_p(sa=3)
    _r(p,'귀원의 무궁한 발전을 기원합니다.\n',size=10)
    _r(p,'당사 대웅제약과 귀 병원 간에 체결된 thync™ 상품공급 계약서에 따라 이루어진 거래와 관련하여 ',size=10)
    _r(p,기준월,bold=True,size=10); _r(p,' 기준 지체보상금이 아래와 같이 발생하였기에 청구드립니다.',size=10)
    _p('─ 청구 내역',bold=True,size=10,sa=2)
    details=[('· 기준월',기준월),('· 현 미수금 잔액',f'{int(row["현 미수금"]):,} 원'),
             ('· 유효 기준회전일',f'{int(row["유효기준회전일"])}일'),('· 현 회전일',f'{int(row["현 회전일"])}일'),
             ('· 지연일수 (누적)',f'{int(row["지연일수(누적)"])}일'),('· 청구 증분일수',f'{int(row["청구 증분일수"])}일'),
             ('· 요율',str(row["요율"])),('· 지체보상금',f'{int(row["지체보상금"]):,} 원'),
             ('· 계산 근거',str(row.get("계산 근거","")))]
    tbl=doc.add_table(rows=len(details),cols=2); tbl.style='Table Grid'
    for i,(lbl,val) in enumerate(details):
        _tc(tbl.rows[i].cells[0],lbl,bold=True,fill='EEEEF5')
        _tc(tbl.rows[i].cells[1],val,color='C0392B' if '지체보상금' in lbl else None)
    for r_obj in tbl.rows: r_obj.cells[0].width=_Cm(5.5); r_obj.cells[1].width=_Cm(9.5)
    사업자번호=str(row.get("사업자번호","")).strip(); 계약유형=str(row.get("계약유형","")).strip()
    parts=[x for x in [f'사업자번호: {사업자번호}' if 사업자번호 not in("","nan") else "",
                       f'계약유형: {계약유형}' if 계약유형 not in("","nan") else ""] if x]
    if parts: _p(f'※ {" | ".join(parts)}',size=8.5,sa=3,color='666666')
    else: _p(sa=3)
    _p('2.',bold=True,size=10,sa=2); p=_p(sa=5)
    _r(p,'상기 내용에 대해 확인하시어, 가능하신 ',size=10)
    _r(p,'______년  ______월  ______일',size=10)
    _r(p,' 일정에 맞추어 회신을 부탁드립니다.',size=10)
    _p('거래처명 :                                                                                   (인)',size=10,sa=6)
    _sig(); _footer()
    buf=io.BytesIO(); doc.save(buf); return buf.getvalue()

def make_공문_word_cumulative(vendor_df):
    if not _HAS_DOCX: raise ImportError("pip install python-docx")
    doc=_DocxDoc()
    _r,_p,_bline,_tc,_sig,_footer,_letterhead=_공문_doc_setup(doc)
    first=vendor_df.sort_values("기준월").iloc[0]
    months=sorted(vendor_df["기준월"].unique()); period=f"{months[0]} ~ {months[-1]}"
    판매처명=str(first.get("판매처명","")); total=int(vendor_df["지체보상금"].sum())
    from datetime import date as _d; today=_d.today()
    _letterhead(판매처명,"지체보상금 청구의 건 (전체 기간)",f"지체-{today.strftime('%Y')}-전체-{판매처명[:4]}")
    _p('1.',bold=True,size=10,sa=2); p=_p(sa=3)
    _r(p,'귀원의 무궁한 발전을 기원합니다.\n',size=10)
    _r(p,'당사 대웅제약과 귀 병원 간에 체결된 thync™ 상품공급 계약서에 따라 이루어진 거래와 관련하여 ',size=10)
    _r(p,period,bold=True,size=10); _r(p,' 기간 동안 발생한 지체보상금 합계를 아래와 같이 청구드립니다.',size=10)
    _p('─ 월별 청구 내역',bold=True,size=10,sa=2)
    hdrs=["기준월","현 미수금","기준회전일","현 회전일","청구 증분일수","요율","지체보상금"]
    sorted_df=vendor_df.sort_values("기준월"); n=len(sorted_df)
    tbl=doc.add_table(rows=n+2,cols=len(hdrs)); tbl.style='Table Grid'
    for ci,h in enumerate(hdrs):
        _tc(tbl.rows[0].cells[ci],h,bold=True,color='FFFFFF',fill='1A2A6C',align=_WDA.CENTER)
    for ri,(_,r) in enumerate(sorted_df.iterrows(),start=1):
        bg='F0F0F8' if ri%2==0 else None
        vals=[r["기준월"],f'{int(r["현 미수금"]):,}',f'{int(r["유효기준회전일"])}',
              f'{int(r["현 회전일"])}',f'{int(r["청구 증분일수"])}',r["요율"],f'{int(r["지체보상금"]):,}']
        for ci,v in enumerate(vals): _tc(tbl.rows[ri].cells[ci],v,size=9.5,fill=bg,align=_WDA.CENTER)
    last=tbl.rows[n+1]
    _tc(last.cells[0],"합  계",bold=True,fill='E8EDF5',align=_WDA.CENTER)
    for ci in range(1,len(hdrs)-1):
        tcPr=last.cells[ci]._tc.get_or_add_tcPr()
        for tag in list(tcPr):
            if tag.tag==_dqn('w:shd'): tcPr.remove(tag)
        shd=_dxml('w:shd'); shd.set(_dqn('w:val'),'clear'); shd.set(_dqn('w:color'),'auto'); shd.set(_dqn('w:fill'),'E8EDF5')
        tcPr.append(shd)
    _tc(last.cells[-1],f'{total:,}',bold=True,color='C0392B',fill='E8EDF5',align=_WDA.CENTER)
    widths=[_Cm(2.2),_Cm(3.0),_Cm(1.8),_Cm(1.8),_Cm(2.2),_Cm(1.5),_Cm(2.5)]
    for r_obj in tbl.rows:
        for ci,w in enumerate(widths): r_obj.cells[ci].width=w
    _p(sa=2); p_tot=_p(sa=5,align=_WDA.RIGHT)
    _r(p_tot,'※ 합계 청구금액 : ',bold=True,size=11); _r(p_tot,f'{total:,} 원',bold=True,size=13,color='C0392B')
    _p('2.',bold=True,size=10,sa=2); p=_p(sa=5)
    _r(p,'상기 내용에 대해 확인하시어, 가능하신 ',size=10)
    _r(p,'______년  ______월  ______일',size=10)
    _r(p,' 일정에 맞추어 회신을 부탁드립니다.',size=10)
    _p('거래처명 :                                                                                   (인)',size=10,sa=6)
    _sig(); _footer()
    buf=io.BytesIO(); doc.save(buf); return buf.getvalue()

def make_공문_zip(df, month_label):
    buf=io.BytesIO()
    def _cs(r):
        if r["청구 증분일수"]<=0: return "—"
        return f'{int(r["현 미수금"]):,} × {int(r["청구 증분일수"])}일 × {r["요율"]} = {int(r["지체보상금"]):,}원'
    with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as zf:
        for _,row in df[df["지체보상금"]>0].sort_values("지체보상금",ascending=False).iterrows():
            row=row.copy(); row["계산 근거"]=_cs(row)
            safe=re.sub(r'[\\/:*?"<>|]','_',str(row["판매처명"]).strip())
            zf.writestr(f"지체보상금청구_{safe}_{row['기준월']}.docx",make_공문_word(row))
    return buf.getvalue()

def make_공문_zip_cumulative(result_df):
    charge_df=result_df[result_df["지체보상금"]>0]
    buf=io.BytesIO()
    with zipfile.ZipFile(buf,"w",zipfile.ZIP_DEFLATED) as zf:
        for code,grp in charge_df.groupby("판매처코드"):
            name=grp["판매처명"].iloc[0]
            safe=re.sub(r'[\\/:*?"<>|]','_',str(name).strip())
            zf.writestr(f"지체보상금청구_{safe}_전체기간.docx",make_공문_word_cumulative(grp))
    return buf.getvalue()

# ──────────────────────────────────────────────────────────────
BASE_PATH = Path("input/기준회전일/SAP 거래처별 기준회전일_DW11_0722.xlsx")
MO_FOLDER = Path("input/월별잔고회전일")
THINK_PATH = next(iter(sorted(Path("input").glob("씽크*.xlsx"))+sorted(Path("input").glob("씽크*.XLSX"))),None)
CT_LIST   = sorted(Path("input/계약정보").glob("*.xlsx"))+sorted(Path("input/계약정보").glob("*.xls"))
CT_PATH   = CT_LIST[0] if CT_LIST else None

@st.cache_data(show_spinner=False)
def _read_mo_folder(folder_str:str, _mtime:float=0):
    folder=Path(folder_str)
    files=sorted(folder.glob("*.xlsx"))+sorted(folder.glob("*.xls"))
    if not files: return None
    dfs=[pd.read_excel(f,dtype=str).dropna(how="all") for f in files]
    combined=pd.concat(dfs,ignore_index=True)
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine="openpyxl") as w: combined.to_excel(w,index=False)
    return buf.getvalue(),[f.name for f in files]

@st.cache_data(show_spinner=False)
def _read_bytes(path_str:str):
    p=Path(path_str); return p.read_bytes() if p.exists() else None

_mo_files=(sorted(MO_FOLDER.glob("*.xlsx"))+sorted(MO_FOLDER.glob("*.xls"))) if MO_FOLDER.exists() else []
_mo_result=_read_mo_folder(str(MO_FOLDER),_mtime=max((f.stat().st_mtime for f in _mo_files if f.exists()),default=0)) if MO_FOLDER.exists() else None
m_b,_mo_names=_mo_result if _mo_result else (None,[])
b_b=_read_bytes(str(BASE_PATH))
c_b=_read_bytes(str(CT_PATH)) if CT_PATH else None
t_b=_read_bytes(str(THINK_PATH)) if THINK_PATH else None

def _build_org_map(m_b):
    if not m_b: return {}
    try:
        df=pd.read_excel(io.BytesIO(m_b),dtype=str).dropna(how="all")
        df.columns=[str(c).strip() for c in df.columns]
        org_map={}
        for _,r in df.iterrows():
            code=str(r.get("판매처","")).strip()
            if not code or code=="nan": continue
            org_map[code]={"사무소":str(r.get("사무소명","")).strip() if "사무소명" in r.index else "",
                           "사업부":str(r.get("사업부명","")).strip() if "사업부명" in r.index else "",
                           "담당자":str(r.get("담당자명","")).strip() if "담당자명" in r.index else ""}
        return org_map
    except: return {}

@st.cache_data(show_spinner=False)
def run_calc(m_b,b_b,c_b,t_b):
    sap_raw=pd.read_excel(io.BytesIO(b_b),dtype=str).dropna(how="all")
    sap_raw.columns=[str(c).strip() for c in sap_raw.columns]
    cmap=n_yes=n_no=None
    if c_b: cmap,n_yes,n_no=load_contract(c_b,sap_raw)
    think_rot=load_think(t_b) if t_b else None
    bdf=load_base(b_b,cmap,think_rot)
    months=load_think_monthly(t_b) if t_b else load_monthly(m_b)
    if t_b and m_b:
        sap_months=load_monthly(m_b); sap_bal:dict={}
        for _me in sap_months:
            for _,_r in _me["data"].iterrows():
                _k=(str(_r["코드"]).strip(),_me["yr"],_me["mo"])
                sap_bal[_k]=(int(_r["잔고"]),int(_r["현회전일"]))
        _씽크ch={"씽크","DW11","DW12"}
        for _me in months:
            _df=_me["data"].copy()
            for _i,_r in _df[~_df["채널"].isin(_씽크ch)].iterrows():
                _k=(str(_r["코드"]).strip(),_me["yr"],_me["mo"])
                if _k in sap_bal: _df.at[_i,"잔고"]=sap_bal[_k][0]; _df.at[_i,"현회전일"]=sap_bal[_k][1]
                else: _df.at[_i,"잔고"]=0
            _me["data"]=_df[_df["잔고"]>0].reset_index(drop=True)
    org_map=_build_org_map(m_b)
    result=calculate(months,bdf)
    if org_map and not result.empty:
        result["사무소"]=result["판매처코드"].map(lambda c:org_map.get(c,{}).get("사무소",""))
        result["사업부"]=result["판매처코드"].map(lambda c:org_map.get(c,{}).get("사업부",""))
        result["담당자"]=result["판매처코드"].map(lambda c:org_map.get(c,{}).get("담당자",""))
    if cmap and not result.empty:
        result["계약유형"]=result["판매처코드"].map(lambda c:(cmap or {}).get(c,{}).get("계약유형",""))
        result["사업자번호"]=result["판매처코드"].map(lambda c:(cmap or {}).get(c,{}).get("사업자번호",""))
    else:
        if "계약유형"  not in result.columns: result["계약유형"]=""
        if "사업자번호" not in result.columns: result["사업자번호"]=""
    return result,cmap,n_yes,n_no

# ══════════════════════════════════════════════════════════════
# 헤더
# ══════════════════════════════════════════════════════════════
today_str=date.today().strftime("%Y. %m. %d")
st.markdown(f"""
<div class="page-header">
  <div class="ph-left">
    <div class="ph-brand">
      <span class="ph-company">대웅제약</span>
      <span class="ph-title">지체보상금 대시보드</span>
      <span class="ph-channel">DW07 · DW11</span>
    </div>
    <div class="ph-sub">월말 순실잔고 × 증분 지연일수 × 1/1000 | 계약서 지체보상금 조항 보유 거래처 한정</div>
  </div>
  <div class="ph-right">
    <span class="ph-date">기준일: {today_str}</span>
  </div>
</div>""", unsafe_allow_html=True)

# ── 데이터 계산
if not (m_b and b_b):
    st.warning("input 폴더에 월별잔고 또는 기준회전일 파일이 없습니다."); st.stop()

with st.spinner("데이터 계산 중..."):
    result_df, contract_map, n_yes, n_no = run_calc(m_b, b_b, c_b, t_b)

if result_df is None or result_df.empty:
    st.warning("계산 결과가 없습니다."); st.stop()

months_sorted = sorted(result_df["기준월"].unique())
mo_sum = result_df.groupby("기준월")["지체보상금"].sum()

# ══════════════════════════════════════════════════════════════
# 탭
# ══════════════════════════════════════════════════════════════
tab_agg, tab_detail = st.tabs(["  📊  집계표  ", "  🔍  월별 상세  "])

# ──────────────────────────────────────────────────────────────
# 탭 1: 집계표
# ──────────────────────────────────────────────────────────────
with tab_agg:
    total_all   = int(result_df["지체보상금"].sum())
    total_vend  = result_df[result_df["지체보상금"]>0]["판매처코드"].nunique()
    total_mo_cnt= len([m for m in months_sorted if mo_sum.get(m,0)>0])
    best_mo     = mo_sum.idxmax() if len(mo_sum)>0 else "-"
    best_mo_amt = int(mo_sum.max()) if len(mo_sum)>0 else 0
    cur_yr      = str(date.today().year)
    ann_total   = int(result_df[result_df["기준월"].str.startswith(cur_yr)]["지체보상금"].sum())

    st.markdown(f"""
    <div class="kpi-grid">
      <div class="kpi-card c-amber">
        <div class="kpi-lbl">전체 기간 지체보상금</div>
        <div class="kpi-val c-amber">{fmt_won(total_all,short=True)}</div>
        <div class="kpi-delta">{months_sorted[0]} ~ {months_sorted[-1]}</div>
      </div>
      <div class="kpi-card c-indigo">
        <div class="kpi-lbl">{cur_yr}년 누계</div>
        <div class="kpi-val">{fmt_won(ann_total,short=True)}</div>
        <div class="kpi-delta">{cur_yr}년 1월 이후 발생 합계</div>
      </div>
      <div class="kpi-card c-red">
        <div class="kpi-lbl">전체 기간 청구 거래처</div>
        <div class="kpi-val c-red">{total_vend}<span class="kpi-unit">개</span></div>
        <div class="kpi-delta">지체보상금 발생 고유 거래처 수</div>
      </div>
      <div class="kpi-card c-green">
        <div class="kpi-lbl">최고 발생 월</div>
        <div class="kpi-val" style="font-size:1.4rem">{mo_label(best_mo) if best_mo!="-" else "-"}</div>
        <div class="kpi-delta">{fmt_won(best_mo_amt)}</div>
      </div>
    </div>""", unsafe_allow_html=True)

    left_col, right_col = st.columns([1, 1], gap="large")

    with left_col:
        st.markdown('<div class="sec-title"><span class="sec-accent"></span>월별 집계표</div>', unsafe_allow_html=True)

        # HTML 집계 테이블 직접 렌더링
        rows_html = ""
        for mo in months_sorted:
            amt = int(mo_sum.get(mo, 0))
            n_c = int((result_df[result_df["기준월"]==mo]["지체보상금"]>0).sum())
            yr_part = mo.split("-")[0][2:]; mo_part = int(mo.split("-")[1])
            label = f"{yr_part}년 {mo_part}월"
            amt_cls = "has-val" if amt > 0 else "zero"
            amt_str = f"{amt:,}" if amt > 0 else "—"
            n_str = str(n_c) if n_c > 0 else "—"
            rows_html += f"""<tr>
              <td>{label}</td>
              <td class="{amt_cls}">{amt_str}</td>
              <td class="{amt_cls}">{n_str}</td>
            </tr>"""
        total_n = int((result_df["지체보상금"]>0).sum())
        rows_html += f"""<tr>
          <td>합계</td>
          <td>{total_all:,}</td>
          <td>{total_n}</td>
        </tr>"""
        st.markdown(f"""
        <div style="overflow-y:auto;max-height:520px;border:1px solid #e5e7eb;border-radius:8px">
        <table class="agg-tbl">
          <thead><tr>
            <th style="text-align:left">기준월</th>
            <th>지체보상금 (원)</th>
            <th>청구 건수</th>
          </tr></thead>
          <tbody>{rows_html}</tbody>
        </table>
        </div>""", unsafe_allow_html=True)

    with right_col:
        st.markdown('<div class="sec-title"><span class="sec-accent"></span>전체 기간 추이</div>', unsafe_allow_html=True)

        x_all = [mo_label(m) for m in months_sorted]
        y_all = [int(mo_sum.get(m,0))/10000 for m in months_sorted]
        cum_all = []; run = 0
        for v in y_all: run+=v; cum_all.append(round(run,1))
        last_idx = max((i for i,v in enumerate(y_all) if v>0), default=-1)
        cum_line = [cum_all[i] if i<=last_idx else None for i in range(len(months_sorted))]
        bar_colors = ["#f59e0b" if v>0 else "#f3f4f6" for v in y_all]

        fig=go.Figure()
        fig.add_trace(go.Scatter(x=x_all,y=cum_line,name="_fill",mode="lines",
            line=dict(color="rgba(0,0,0,0)",width=0),fill="tozeroy",
            fillcolor="rgba(245,158,11,0.08)",hoverinfo="skip",yaxis="y2",showlegend=False))
        fig.add_trace(go.Bar(x=x_all,y=y_all,name="월별",marker_color=bar_colors,marker_line_width=0,
            text=[f"{v:,.0f}" if v>0 else "" for v in y_all],textposition="outside",
            cliponaxis=False,textfont=dict(size=9,color="#6b7280"),
            hovertemplate="<b>%{x}</b><br>%{y:,.0f}만원<extra></extra>",yaxis="y1"))
        fig.add_trace(go.Scatter(x=x_all,y=cum_line,name="누적",mode="lines+markers",
            line=dict(color="#f59e0b",width=2),
            marker=dict(size=5,color="#f59e0b",line=dict(color="#fff",width=2)),
            hovertemplate="<b>%{x}</b><br>누적 %{y:,.0f}만원<extra></extra>",yaxis="y2"))
        mo_max=max(y_all) if max(y_all)>0 else 1
        cu_max=max(v for v in cum_all if v>0) if any(v>0 for v in cum_all) else 1
        fig.update_layout(
            plot_bgcolor="#ffffff",paper_bgcolor="#ffffff",
            margin=dict(t=15,b=5,l=50,r=60),height=440,bargap=0.35,
            legend=dict(orientation="h",x=0.5,xanchor="center",y=1.06,
                       font=dict(size=11,family=FONT),bgcolor="rgba(0,0,0,0)"),
            yaxis=dict(gridcolor="#f3f4f6",tickformat=",d",zeroline=False,
                      range=[0,mo_max*1.6],tickfont=dict(size=10,color="#9ca3af"),title=None,showline=False),
            yaxis2=dict(overlaying="y",side="right",range=[0,cu_max*1.25],tickformat=",d",
                       tickfont=dict(size=10,color="#d97706"),showgrid=False,title=None,showline=False),
            xaxis=dict(showgrid=False,tickfont=dict(size=10,color="#6b7280"),showline=False,tickangle=-40),
            font=dict(family=FONT))
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar":False})

    # ── 거래처별 누계표
    st.markdown('<div class="sec-title"><span class="sec-accent"></span>거래처별 누계 <span class="sec-sub">전체 기간 · 지체보상금 발생 거래처만</span></div>', unsafe_allow_html=True)

    pv_all = result_df.pivot_table(index=["판매처코드","판매처명"],columns="기준월",
                                   values="지체보상금",aggfunc="sum",fill_value=0).reset_index()
    pv_all.columns=[mo_label(c) if c not in["판매처코드","판매처명"] else c for c in pv_all.columns]
    pv_all["합계"]=pv_all.iloc[:,2:].sum(axis=1)
    pv_all=pv_all[pv_all["합계"]>0].sort_values("합계",ascending=False).reset_index(drop=True)
    pv_all.index=pv_all.index+1
    mo_cols_pv=[c for c in pv_all.columns if c not in["판매처코드","판매처명"]]
    def _heat(s):
        mx=s.max() if s.max()>0 else 1
        return [f"background:rgba(245,158,11,{v/mx*0.45:.2f})" for v in s]
    pv_event=st.dataframe(
        pv_all.style.format({c:"{:,.0f}" for c in mo_cols_pv}).apply(_heat,subset=["합계"]),
        use_container_width=True,height=400,on_select="rerun",selection_mode="multi-row",key="pv_all_df")
    _pv_pos=pv_event.selection.rows if pv_event.selection.rows else []
    ba,bb,_=st.columns([2,2,4])
    if _pv_pos:
        _pv_codes=[pv_all.iloc[i]["판매처코드"] for i in _pv_pos if i<len(pv_all)]
        _pv_df=result_df[result_df["판매처코드"].isin(_pv_codes)&(result_df["지체보상금"]>0)]
        n_pv=_pv_df["판매처코드"].nunique()
        with ba: st.download_button(f"📁 선택 전체기간 공문 ({n_pv}개사)",
            data=make_공문_zip_cumulative(_pv_df) if n_pv else b"",
            file_name=f"지체보상금_공문_전체기간_{date.today():%Y%m%d}.zip",
            mime="application/zip",disabled=(n_pv==0),use_container_width=True,key="btn_pv_docx")
        with bb: st.download_button("⬇ 전체 기간 엑셀",data=make_full_excel(result_df),
            file_name=f"지체보상금_전체_{date.today():%Y%m%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,key="btn_agg_full_excel")
    else:
        st.caption("행 선택 → 공문 다운로드")


# ──────────────────────────────────────────────────────────────
# 탭 2: 월별 상세
# ──────────────────────────────────────────────────────────────
with tab_detail:
    years_avail = sorted(set(m.split("-")[0] for m in months_sorted))
    if "sel_mo" not in st.session_state or st.session_state.sel_mo not in months_sorted:
        st.session_state.sel_mo = months_sorted[-1]
    if "sel_yr" not in st.session_state or st.session_state.sel_yr not in years_avail:
        st.session_state.sel_yr = st.session_state.sel_mo.split("-")[0]

    # ── 연도 + 월 선택기
    yr_label_col, yr_btn_cols = st.columns([1, 10])
    with yr_label_col:
        st.markdown('<div style="padding-top:.55rem;font-size:.72rem;font-weight:700;color:#9ca3af">연도</div>', unsafe_allow_html=True)
    yr_cols = yr_btn_cols.columns(len(years_avail) + max(0, 5-len(years_avail)))
    for i, yr in enumerate(years_avail):
        with yr_cols[i]:
            if st.button(f"{yr}년", key=f"yr_{yr}",
                         type="primary" if yr==st.session_state.sel_yr else "secondary",
                         use_container_width=True):
                st.session_state.sel_yr = yr
                yr_months = [m for m in months_sorted if m.startswith(yr)]
                if st.session_state.sel_mo not in yr_months:
                    st.session_state.sel_mo = yr_months[-1]
                st.rerun()

    yr_months = [m for m in months_sorted if m.startswith(st.session_state.sel_yr)]
    yr_months_set = set(yr_months)

    mo_label_col, mo_btn_area = st.columns([1, 10])
    with mo_label_col:
        st.markdown('<div style="padding-top:.55rem;font-size:.72rem;font-weight:700;color:#9ca3af">월</div>', unsafe_allow_html=True)
    row1, row2 = mo_btn_area.columns(6), mo_btn_area.columns(6)
    for i in range(6):
        mo=i+1; m=f"{st.session_state.sel_yr}-{mo:02d}"
        with row1[i]:
            if m in yr_months_set:
                if st.button(f"{mo}월",key=f"mo_{m}",type="primary" if m==st.session_state.sel_mo else "secondary",use_container_width=True):
                    st.session_state.sel_mo=m; st.rerun()
            else: st.button(f"{mo}월",key=f"mo_dis_{m}",disabled=True,use_container_width=True)
    for i in range(6):
        mo=i+7; m=f"{st.session_state.sel_yr}-{mo:02d}"
        with row2[i]:
            if m in yr_months_set:
                if st.button(f"{mo}월",key=f"mo_{m}",type="primary" if m==st.session_state.sel_mo else "secondary",use_container_width=True):
                    st.session_state.sel_mo=m; st.rerun()
            else: st.button(f"{mo}월",key=f"mo_dis_{m}",disabled=True,use_container_width=True)

    st.markdown('<div style="margin:.6rem 0"></div>', unsafe_allow_html=True)

    sel=st.session_state.sel_mo; sel_idx=months_sorted.index(sel)
    prev=months_sorted[sel_idx-1] if sel_idx>0 else None
    sel_df=result_df[result_df["기준월"]==sel]
    prev_df=result_df[result_df["기준월"]==prev] if prev else pd.DataFrame()
    sel_total=int(sel_df["지체보상금"].sum())
    sel_n=int((sel_df["지체보상금"]>0).sum())
    sel_cu=sel_df[sel_df["지체보상금"]>0]["판매처코드"].nunique()
    prev_total=int(prev_df["지체보상금"].sum()) if not prev_df.empty else 0
    prev_n=int((prev_df["지체보상금"]>0).sum()) if not prev_df.empty else 0
    prev_cu=prev_df[prev_df["지체보상금"]>0]["판매처코드"].nunique() if not prev_df.empty else 0
    _sel_yr=sel.split("-")[0]
    ann_sel=int(result_df[result_df["기준월"].str.startswith(_sel_yr)]["지체보상금"].sum())

    def _d(cur,prv,unit="원"):
        if prv==0: return '<span class="neu">— 전월 없음</span>'
        d=cur-prv
        if d==0: return '<span class="neu">± 0 전월 동일</span>'
        arr="▲" if d>0 else "▼"; cls="up" if d>0 else "dn"
        val=fmt_won(abs(d)) if unit=="원" else f"{abs(d)}{unit}"
        return f'<span class="{cls}">{arr} {val} vs 전월</span>'

    d_amt=sel_total-prev_total
    d_html=(f'<span class="{"up" if d_amt>0 else "dn"}">{"▲" if d_amt>0 else "▼"} {fmt_won(abs(d_amt))} vs 전월</span>'
            if d_amt!=0 and prev_total>0 else '<span class="neu">— 전월 없음</span>')

    # ── 월 KPI
    st.markdown(f"""
    <div class="kpi-grid">
      <div class="kpi-card c-amber">
        <div class="kpi-lbl">{mo_label(sel)} 지체보상금</div>
        <div class="kpi-val c-amber">{fmt_won(sel_total,short=True)}</div>
        <div class="kpi-delta">{d_html}</div>
      </div>
      <div class="kpi-card c-red">
        <div class="kpi-lbl">청구 건수</div>
        <div class="kpi-val c-red">{sel_n}<span class="kpi-unit">건</span></div>
        <div class="kpi-delta">{_d(sel_n,prev_n,"건")}</div>
      </div>
      <div class="kpi-card c-blue">
        <div class="kpi-lbl">청구 거래처</div>
        <div class="kpi-val">{sel_cu}<span class="kpi-unit">개</span></div>
        <div class="kpi-delta">{_d(sel_cu,prev_cu,"개")}</div>
      </div>
      <div class="kpi-card c-indigo">
        <div class="kpi-lbl">{_sel_yr}년 누계</div>
        <div class="kpi-val">{fmt_won(ann_sel,short=True)}</div>
        <div class="kpi-delta">{_sel_yr}년 1월 ~ {mo_label(sel)}</div>
      </div>
    </div>""", unsafe_allow_html=True)

    # ── 신규/추가/수금 카드
    _prev_map={r["판매처코드"]:r for _,r in prev_df.iterrows()}
    _curr_map={r["판매처코드"]:r for _,r in sel_df.iterrows()}
    _신규,_추가,_수금=[],[],[]
    for _code in set(_prev_map)|set(_curr_map):
        _c=_curr_map.get(_code); _p=_prev_map.get(_code)
        _cj=int(_c["지체보상금"]) if _c is not None else 0
        _pj=int(_p["지체보상금"]) if _p is not None else 0
        _cp=int(_c["현 미수금"]) if _c is not None else 0
        _pp=int(_p["현 미수금"]) if _p is not None else 0
        _nm=(_c if _c is not None else _p)["판매처명"]
        _org=str((_c if _c is not None else _p).get("사무소","")).strip() or str((_c if _c is not None else _p).get("사업부","")).strip()
        _org=_org if _org and _org not in("nan","") else ""
        if _cj>0 and _pj==0: _신규.append({"명":_nm,"금액":_cj,"org":_org})
        elif _cj>0 and _pj>0: _추가.append({"명":_nm,"금액":_cj,"org":_org})
        if _pj>0 and _cj==0: _수금.append({"명":_nm,"금액":max(_pp-_cp,0) or _pp,"org":_org})
    for _lst in (_신규,_추가,_수금): _lst.sort(key=lambda x:-x["금액"])

    def _card(lst,cls,title,color):
        total=sum(x["금액"] for x in lst)
        items=""
        for i,x in enumerate(lst[:5]):
            org_span=f'<span class="chg-org">{x["org"]}</span>' if x["org"] else ""
            items+=(f'<div class="chg-item">'
                    f'<span class="chg-nm">{x["명"]}{org_span}</span>'
                    f'<span class="chg-amt" style="color:{color}">{fmt_won(x["금액"],short=True)}</span>'
                    f'</div>')
        if len(lst)>5:
            etc=len(lst)-5; etc_amt=sum(x["금액"] for x in lst[5:])
            items+=f'<div style="font-size:.71rem;color:#9ca3af;padding:.3rem 0">그 외 {etc}처 ({fmt_won(etc_amt,short=True)})</div>'
        if not lst: items=f'<div class="chg-none">해당 없음</div>'
        cnt_badge=f'<span class="chg-cnt">{len(lst)}건 · {fmt_won(total,short=True)}</span>'
        return (f'<div class="chg-block {cls}">'
                f'<div class="chg-head"><span class="chg-name" style="color:{color}">{title}</span>{cnt_badge}</div>'
                f'<div class="chg-inner">{items}</div></div>')

    cc1,cc2,cc3=st.columns(3,gap="medium")
    with cc1: st.markdown(_card(_신규,"new",f"{mo_label(sel)} 신규 연체","#ef4444"),unsafe_allow_html=True)
    with cc2: st.markdown(_card(_추가,"add",f"{mo_label(sel)} 추가 연체","#d97706"),unsafe_allow_html=True)
    with cc3: st.markdown(_card(_수금,"done",f"{mo_label(sel)} 수금 완료","#059669"),unsafe_allow_html=True)

    st.markdown('<div style="margin:.8rem 0"></div>', unsafe_allow_html=True)

    # ── 랭킹 + 차트
    rank_col, chart_col = st.columns([1, 1], gap="large")
    rank_colors=["#ef4444","#f97316","#f59e0b","#84cc16","#22c55e","#14b8a6","#3b82f6","#6366f1","#8b5cf6","#ec4899"]

    with rank_col:
        st.markdown(f'<div class="sec-title"><span class="sec-accent"></span>{mo_label(sel)} 청구 발생 거래처</div>', unsafe_allow_html=True)
        charge=sel_df[sel_df["지체보상금"]>0].sort_values("지체보상금",ascending=False).reset_index(drop=True)
        all_max=int(charge["지체보상금"].max()) if not charge.empty else 1
        if charge.empty:
            st.markdown('<div style="text-align:center;color:#d1d5db;padding:3rem 0;font-size:.88rem">해당 월 청구 발생 없음</div>',unsafe_allow_html=True)
        else:
            rows=""
            for i,row in charge.iterrows():
                v=int(row["지체보상금"]); pct=min(v/all_max*100,100)
                color=rank_colors[i] if i<len(rank_colors) else "#9ca3af"
                org=str(row.get("사무소","")).strip() or str(row.get("사업부","")).strip()
                org=org if org and org not in("nan","") else ""
                org_txt=f'<div class="rank-org">{org}</div>' if org else ""
                rows+=(f'<div class="rank-row">'
                       f'<span class="rank-no" style="background:{color}">{i+1}</span>'
                       f'<div><div class="rank-name">{row["판매처명"]}</div>{org_txt}'
                       f'<div class="rank-bar-bg"><div class="rank-bar-fg" style="background:{color};width:{pct:.1f}%"></div></div></div>'
                       f'<div class="rank-amt" style="color:{color}">{fmt_won(v,short=True)}</div>'
                       f'</div>')
            st.markdown(f'<div class="rank-wrap">{rows}</div>',unsafe_allow_html=True)

    with chart_col:
        st.markdown(f'<div class="sec-title"><span class="sec-accent"></span>{_sel_yr}년 월별 추이</div>', unsafe_allow_html=True)
        all_12=[f"{_sel_yr}-{mo:02d}" for mo in range(1,13)]
        x_12=[f"{int(m.split('-')[1])}월" for m in all_12]
        y_12=[int(mo_sum.get(m,0))/10000 for m in all_12]
        bar_col_12=["#f59e0b" if m==sel else("#fde68a" if mo_sum.get(m,0)>0 else "#f3f4f6") for m in all_12]
        cum_12=[]; run=0
        for v in y_12: run+=v; cum_12.append(round(run,1))
        last_idx=max((i for i,v in enumerate(y_12) if v>0),default=-1)
        cum_line_12=[cum_12[i] if i<=last_idx else None for i in range(12)]

        fig2=go.Figure()
        fig2.add_trace(go.Scatter(x=x_12,y=cum_line_12,name="_f",mode="lines",
            line=dict(color="rgba(0,0,0,0)",width=0),fill="tozeroy",
            fillcolor="rgba(245,158,11,0.07)",hoverinfo="skip",yaxis="y2",showlegend=False))
        fig2.add_trace(go.Bar(x=x_12,y=y_12,name="월별",marker_color=bar_col_12,marker_line_width=0,
            text=[f"{v:,.0f}" if v>0 else "" for v in y_12],textposition="outside",
            cliponaxis=False,textfont=dict(size=9.5,color="#6b7280"),
            hovertemplate="<b>%{x}</b><br>%{y:,.0f}만원<extra></extra>",yaxis="y1"))
        fig2.add_trace(go.Scatter(x=x_12,y=cum_line_12,name="누적",mode="lines+markers",
            line=dict(color="#f59e0b",width=2),marker=dict(size=5,color="#f59e0b",line=dict(color="#fff",width=2)),
            hovertemplate="<b>%{x}</b><br>누적 %{y:,.0f}만원<extra></extra>",yaxis="y2"))
        mo_max2=max(y_12) if max(y_12)>0 else 1
        cu_max2=max(v for v in cum_12 if v>0) if any(v>0 for v in cum_12) else 1
        fig2.update_layout(
            plot_bgcolor="#ffffff",paper_bgcolor="#ffffff",
            margin=dict(t=15,b=5,l=50,r=60),height=360,bargap=0.4,
            legend=dict(orientation="h",x=0.5,xanchor="center",y=1.06,
                       font=dict(size=11,family=FONT),bgcolor="rgba(0,0,0,0)"),
            yaxis=dict(gridcolor="#f3f4f6",tickformat=",d",zeroline=False,range=[0,mo_max2*1.6],
                      tickfont=dict(size=10,color="#9ca3af"),title=None,showline=False),
            yaxis2=dict(overlaying="y",side="right",range=[0,cu_max2*1.3],tickformat=",d",
                       tickfont=dict(size=10,color="#d97706"),showgrid=False,title=None,showline=False),
            xaxis=dict(showgrid=False,tickfont=dict(size=11,color="#6b7280"),showline=False),
            font=dict(family=FONT))
        st.plotly_chart(fig2,use_container_width=True,config={"displayModeBar":False})

    # ── 다운로드 버튼
    st.markdown('<div class="sec-title"><span class="sec-accent"></span>다운로드</div>', unsafe_allow_html=True)
    charge_only=sel_df[sel_df["지체보상금"]>0].sort_values("지체보상금",ascending=False)
    n_charge=len(charge_only); all_charge_df=result_df[result_df["지체보상금"]>0]; n_vendors=all_charge_df["판매처코드"].nunique()
    dc1,dc2,dc3,dc4=st.columns(4)
    with dc1: st.download_button(f"⬇ {mo_label(sel)} 발생현황",data=make_excel(charge_only,f"{mo_label(sel)}"),
        file_name=f"지체보상금_{sel.replace('-','')}_발생현황.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True,key="btn_mo_xl")
    with dc2: st.download_button("⬇ 전체 기간 엑셀",data=make_full_excel(result_df),
        file_name=f"지체보상금_전체_{date.today():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True,key="btn_full_xl")
    with dc3: st.download_button(f"📄 {mo_label(sel)} 공문 ({n_charge}건)" if n_charge else "📄 공문 없음",
        data=make_공문_zip(charge_only,mo_label(sel)) if n_charge else b"",
        file_name=f"지체보상금_공문_{sel.replace('-','')}.zip",mime="application/zip",
        disabled=(n_charge==0),use_container_width=True,key="btn_mo_docx")
    with dc4: st.download_button(f"📁 전체기간 공문 ({n_vendors}개사)",
        data=make_공문_zip_cumulative(result_df) if n_vendors else b"",
        file_name=f"지체보상금_공문_전체기간_{date.today():%Y%m%d}.zip",mime="application/zip",
        disabled=(n_vendors==0),use_container_width=True,key="btn_all_docx")

    # ── 상세 테이블
    st.markdown(f'<div class="sec-title"><span class="sec-accent"></span>{mo_label(sel)} 상세 내역</div>', unsafe_allow_html=True)
    fc1,fc2=st.columns([6,1])
    kw=fc1.text_input("거래처명 검색",placeholder="거래처명으로 검색",label_visibility="collapsed",key="kw1")
    only_c=fc2.checkbox("청구만",key="oc1")
    view=sel_df.copy()
    if kw: view=view[view["판매처명"].str.contains(kw,na=False)]
    if only_c: view=view[view["지체보상금"]>0]
    view=view.sort_values("지체보상금",ascending=False).reset_index(drop=True)
    def _cs(r):
        if r["청구 증분일수"]<=0: return "—"
        return f'{int(r["현 미수금"]):,} × {int(r["청구 증분일수"])}일 × {r["요율"]} = {int(r["지체보상금"]):,}원'
    view["계산 근거"]=view.apply(_cs,axis=1)
    _all_cols=["판매처코드","판매처명","사업자번호","계약유형","채널","사업부","사무소","담당자",
               "현 미수금","유효기준회전일","현 회전일","지연일수(누적)","청구 증분일수","요율","지체보상금","계산 근거","비고"]
    disp=view[[c for c in _all_cols if c in view.columns]].rename(columns={"판매처코드":"거래처코드","채널":"거래채널"})
    disp.index=disp.index+1
    def _sty(r):
        if r["지체보상금"]>0: return ["background:#fffbeb;font-weight:600"]*len(r)
        return ["color:#d1d5db"]*len(r)
    t1_event=st.dataframe(
        disp.style.apply(_sty,axis=1).format({"현 미수금":"{:,.0f}","지체보상금":"{:,.0f}"}),
        use_container_width=True,height=440,on_select="rerun",selection_mode="multi-row",key="tab1_df",
        column_config={
            "거래처코드":     st.column_config.TextColumn("거래처코드",width=115),
            "판매처명":       st.column_config.TextColumn("거래처명",width=210),
            "사업자번호":     st.column_config.TextColumn("사업자번호",width=110),
            "계약유형":       st.column_config.TextColumn("계약유형",width=105),
            "거래채널":       st.column_config.TextColumn("채널",width=70),
            "사업부":         st.column_config.TextColumn("사업부",width=140),
            "사무소":         st.column_config.TextColumn("사무소",width=130),
            "담당자":         st.column_config.TextColumn("담당자",width=80),
            "현 미수금":      st.column_config.NumberColumn("현 미수금",width=120,format="%,.0f"),
            "유효기준회전일": st.column_config.NumberColumn("기준회전일",width=80),
            "현 회전일":      st.column_config.NumberColumn("현회전일",width=75),
            "지연일수(누적)": st.column_config.NumberColumn("지연일수",width=75),
            "청구 증분일수":  st.column_config.NumberColumn("청구일수",width=75),
            "요율":           st.column_config.TextColumn("요율",width=65),
            "지체보상금":     st.column_config.NumberColumn("지체보상금",width=125,format="%,.0f"),
            "계산 근거":      st.column_config.TextColumn("계산 근거",width=280),
            "비고":           st.column_config.TextColumn("비고",width=70)})
    _sel_pos=t1_event.selection.rows if t1_event.selection.rows else []
    if _sel_pos:
        _sel_codes=[view.iloc[i]["판매처코드"] for i in _sel_pos if i<len(view)]
        _sel_mo_df=view[view["판매처코드"].isin(_sel_codes)&(view["지체보상금"]>0)]
        _sel_all_df=result_df[result_df["판매처코드"].isin(_sel_codes)&(result_df["지체보상금"]>0)]
        n_mo=len(_sel_mo_df); n_all=_sel_all_df["판매처코드"].nunique()
        st.markdown(f'<div style="font-size:.73rem;color:#6b7280;margin:.4rem 0 .3rem">'
                    f'선택: <b style="color:#111827">{len(_sel_pos)}개사</b> · 이번달 {n_mo}건 · 전체기간 {n_all}개사</div>',unsafe_allow_html=True)
        _ba,_bb,_=st.columns([2,2,4])
        with _ba: st.download_button(f"📄 이번달 공문 ({n_mo}건)",
            data=make_공문_zip(_sel_mo_df,mo_label(sel)) if n_mo else b"",
            file_name=f"지체보상금_선택_{sel.replace('-','')}.zip",
            mime="application/zip",disabled=(n_mo==0),use_container_width=True,key="btn_sel_mo")
        with _bb: st.download_button(f"📁 전체기간 공문 ({n_all}개사)",
            data=make_공문_zip_cumulative(_sel_all_df) if n_all else b"",
            file_name=f"지체보상금_선택_전체기간_{date.today():%Y%m%d}.zip",
            mime="application/zip",disabled=(n_all==0),use_container_width=True,key="btn_sel_cum")
    else:
        st.caption("행을 클릭하면 선택 → 공문 다운로드 버튼이 표시됩니다")
