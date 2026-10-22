"""
지체보상금 대시보드 v7
실행: streamlit run app.py
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

st.markdown(
    '<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700;900&display=swap" rel="stylesheet">',
    unsafe_allow_html=True,
)

st.markdown("""<style>
html, body, [class*="css"], .stApp, button, input, select, textarea,
[data-testid], [data-baseweb], .stMarkdown, .stDataFrame, p, span, div, th, td, label {
  font-family: 'Noto Sans KR', 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif !important;
  -webkit-font-smoothing: antialiased;
}

/* ── 레이아웃 ── */
.stApp { background: #F5F5F5; }
header[data-testid="stHeader"] { background: transparent !important; height: 0 !important; min-height: 0 !important; }
[data-testid="stToolbar"] { display: none !important; }
[data-testid="stDecoration"] { display: none !important; }
[data-testid="stStatusWidget"] { display: none !important; }
#MainMenu { display: none !important; }
section[data-testid="stSidebar"] { display: none !important; }
.stMainBlockContainer, .block-container {
  max-width: 100% !important;
  padding: 0 1.8rem 3rem !important;
  padding-top: 0 !important;
}

/* ── 헤더 (레퍼런스: 흰 배경, 검정 타이틀) ── */
.top-header {
  background: #FFFFFF;
  padding: .9rem 1.8rem; margin: 0 -1.8rem 1.2rem;
  display: flex; align-items: center; justify-content: space-between;
  border-bottom: 1px solid #E0E0E0;
}
.hd-left { display: flex; flex-direction: column; gap: .15rem; }
.hd-title { font-size: 1.18rem; font-weight: 700; color: #212121; letter-spacing: -.02em; }
.hd-sub   { font-size: .72rem; color: #9E9E9E; margin-top: .05rem; }
.hd-badge {
  background: #FFF3E0; color: #E65100;
  border: 1px solid #FFCC80;
  font-size: .72rem; font-weight: 700; padding: .25rem .7rem; border-radius: 4px;
}

/* ── 탭 ── */
[data-baseweb="tab-list"] {
  background: #fff !important;
  border-radius: 10px 10px 0 0 !important;
  border: 1px solid #E0E0E0 !important;
  border-bottom: none !important;
  padding: 0 .5rem !important;
  gap: 0 !important;
}
[data-baseweb="tab"] {
  font-weight: 500 !important; font-size: .88rem !important;
  padding: .65rem 1.4rem !important; color: #9E9E9E !important;
  border-bottom: 3px solid transparent !important;
  background: transparent !important;
  font-family: 'Noto Sans KR', 'Malgun Gothic', sans-serif !important;
}
[data-baseweb="tab"][aria-selected="true"] {
  color: #F57C00 !important;
  border-bottom: 3px solid #F57C00 !important;
  font-weight: 700 !important;
}
[data-baseweb="tab-panel"] {
  background: #fff !important;
  border: 1px solid #E0E0E0 !important;
  border-top: none !important;
  border-radius: 0 0 10px 10px !important;
  padding: 1.2rem 1.4rem !important;
}

/* ── KPI 카드 ── */
.kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: .8rem; margin-bottom: 1.2rem; }
.kpi-card {
  background: #fff; border-radius: 10px;
  padding: .95rem 1.1rem .9rem 1rem;
  border: 1px solid #E8E8E8;
  border-left: 4px solid #E0E0E0;
  box-shadow: 0 1px 4px rgba(0,0,0,.06);
}
.kpi-card.orange { border-left-color: #F57C00; }
.kpi-card.red    { border-left-color: #E53E3E; }
.kpi-card.blue   { border-left-color: #1565C0; }
.kpi-card.teal   { border-left-color: #00796B; }
.kpi-card.purple { border-left-color: #6D28D9; }
.kpi-card.gray   { border-left-color: #9E9E9E; }
.kpi-lbl  { font-size: .69rem; color: #9E9E9E; font-weight: 500; letter-spacing: .01em; margin-bottom: .38rem; }
.kpi-val  { font-size: 1.85rem; font-weight: 700; color: #212121; line-height: 1.05; letter-spacing: -.02em; }
.kpi-val.orange { color: #F57C00; }
.kpi-val.red    { color: #E53E3E; }
.kpi-val.blue   { color: #1565C0; }
.kpi-val.teal   { color: #00796B; }
.kpi-val.purple { color: #6D28D9; }
.kpi-unit { font-size: .88rem; font-weight: 500; color: #757575; margin-left: .08rem; }
.kpi-sub  { font-size: .7rem; color: #9E9E9E; margin-top: .42rem; line-height: 1.7; }
.kpi-up   { color: #E53E3E; font-weight: 600; }
.kpi-dn   { color: #1565C0; font-weight: 600; }

/* ── 섹션 헤더 ── */
.sec-hd {
  display: flex; align-items: center; gap: .4rem;
  font-size: .97rem; font-weight: 700; color: #212121;
  padding-bottom: .55rem; margin: 1rem 0 .7rem;
  border-bottom: 1.5px solid #E0E0E0;
}
.sec-bar { width: 3px; height: 16px; border-radius: 2px; background: #F57C00; flex-shrink: 0; }
.sec-sub { font-size: .7rem; color: #9E9E9E; font-weight: 400; margin-left: .3rem; }

/* ── 버튼 ── */
.stButton > button {
  font-family: 'Noto Sans KR', 'Malgun Gothic', sans-serif !important;
  font-weight: 500 !important; border-radius: 999px !important;
  font-size: .82rem !important; transition: all .12s !important;
  padding: .35rem .9rem !important;
}
button[kind="primary"] {
  background: #F57C00 !important; border-color: #F57C00 !important;
  color: #fff !important; font-weight: 700 !important;
}
button[kind="secondary"] {
  background: #fff !important; border-color: #BDBDBD !important; color: #424242 !important;
}
button[kind="secondary"]:hover {
  border-color: #F57C00 !important; color: #E65100 !important;
}
.stDownloadButton > button {
  background: #37474F !important; color: #fff !important;
  border: none !important; border-radius: 6px !important;
  font-weight: 500 !important; font-size: .82rem !important;
  font-family: 'Noto Sans KR', 'Malgun Gothic', sans-serif !important;
  padding: .5rem 1rem !important;
}
.stDownloadButton > button:hover { background: #455A64 !important; }

/* ── 월 선택기 ── */
.yr-mo-area {
  background: #FAFAFA; border-radius: 6px;
  border: 1px solid #E0E0E0; padding: .7rem .9rem; margin-bottom: 1rem;
  display: flex; align-items: flex-start; gap: 1.2rem;
}
.yr-label { font-size: .67rem; font-weight: 700; color: #9E9E9E; letter-spacing: .06em;
            text-transform: uppercase; margin-bottom: .35rem; }

/* ── 변동 카드 (신규/추가/수금) ── */
.chg-card {
  background: #fff; border-radius: 10px;
  border: 1px solid #E0E0E0;
  overflow: hidden;
}
/* 카드 헤더: .chg-title 이 직접 배경색 적용 */
.chg-card .chg-title {
  font-size: .83rem; font-weight: 700;
  padding: .6rem .9rem;
  border-bottom: 1px solid #E8E8E8;
  display: block;
}
.chg-card.red   .chg-title { background: #FFF5F5; color: #C53030; }
.chg-card.amber .chg-title { background: #FFFBEB; color: #B45309; }
.chg-card.green .chg-title { background: #F0FDF4; color: #276749; }
/* 카드 바디: 아이템 행 */
.chg-row {
  display: flex; align-items: center; gap: .4rem;
  padding: .42rem .9rem; border-bottom: 1px solid #F5F5F5;
  font-size: .82rem;
}
.chg-row:last-child { border-bottom: none; }
.chg-num  { font-size: .68rem; color: #BDBDBD; width: 1rem; flex-shrink: 0; text-align: center; }
.chg-name { color: #424242; flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.chg-amt  { font-weight: 700; white-space: nowrap; flex-shrink: 0; }

/* ── 랭킹 ── */
.rank-item {
  display: grid; grid-template-columns: 2rem 1fr 5.5rem;
  align-items: center; gap: .7rem;
  padding: .5rem 0; border-bottom: 1px solid #F5F5F5;
}
.rank-badge {
  width: 1.4rem; height: 1.4rem; border-radius: 50%;
  display: inline-flex; align-items: center; justify-content: center;
  font-size: .7rem; font-weight: 700; color: #fff; flex-shrink: 0;
}
.rank-bar-bg { background: #F5F5F5; border-radius: 2px; height: 5px; margin-top: .22rem; overflow: hidden; }
.rank-bar-fg { height: 100%; border-radius: 2px; }

/* ── 랭킹 카드 래퍼 ── */
.rank-card {
  background: #fff; border: 1px solid #E0E0E0; border-radius: 10px;
  padding: .5rem .85rem; overflow-y: auto; max-height: 420px;
}
.rank-empty {
  color: #BDBDBD; text-align: center; font-size: .85rem; padding: 3rem 0;
}

/* ── 검색창 ── */
[data-testid="stTextInput"] > div > div {
  background: #fff !important;
  border: 1.5px solid #E0E0E0 !important;
  border-radius: 6px !important;
}
[data-testid="stTextInput"] > div > div:focus-within {
  border-color: #F57C00 !important;
  box-shadow: 0 0 0 3px rgba(245,124,0,.12) !important;
}
[data-testid="stTextInput"] input {
  font-size: .85rem !important;
  color: #212121 !important;
}

/* ── 다운로드 섹션 ── */
.dl-card {
  background: #fff; border: 1px solid #E0E0E0; border-radius: 10px;
  padding: .9rem 1rem; margin-bottom: .5rem;
}
.dl-card-title {
  font-size: .72rem; font-weight: 700; color: #9E9E9E;
  letter-spacing: .04em; text-transform: uppercase; margin-bottom: .6rem;
}

/* ── 집계 테이블 ── */
.agg-tbl { width: 100%; border-collapse: collapse; font-size: .83rem; }
.agg-tbl thead tr th {
  background: #37474F; color: #ECEFF1;
  font-weight: 700; font-size: .72rem;
  padding: .65rem .85rem; text-align: right; white-space: nowrap;
}
.agg-tbl thead tr th:first-child { text-align: left; }
.agg-tbl tbody td {
  padding: .5rem .85rem; border-bottom: 1px solid #F5F5F5;
  color: #424242; text-align: right; white-space: nowrap;
}
.agg-tbl tbody td:first-child { text-align: left; font-weight: 600; color: #212121; }
.agg-tbl tbody tr:last-child td {
  background: #37474F; color: #ECEFF1; font-weight: 700; border-top: none;
}
.agg-tbl tbody tr:hover td { background: #FFF8F0; }
.agg-tbl .has-val { color: #E65100; font-weight: 700; }
.agg-tbl .zero    { color: #BDBDBD; }

/* ── dataframe ── */
[data-testid="stDataFrame"] { border-radius: 6px; overflow: hidden; }
[data-testid="stDataFrame"] th {
  background: #37474F !important; color: #ECEFF1 !important;
  font-size: .78rem !important; font-weight: 700 !important;
}

/* ── 배지 ── */
.badge {
  display: inline-block; font-size: .68rem; font-weight: 600;
  padding: .08rem .42rem; border-radius: 3px; margin-left: .3rem; vertical-align: middle;
}
.badge-orange { background: #FFF3E0; color: #E65100; border: 1px solid #FFCC80; }
.badge-gray   { background: #F5F5F5; color: #757575; }
.badge-blue   { background: #E3F2FD; color: #1565C0; }
</style>""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────
ALLOW_CH       = {"DW07","DW11"}   # 씽크(DW11) + 오프라인(DW07) 만 대상
EXCLUDE_CODES  = {"10014461","C4900","C2000","C3100"}
EXCLUDE_BU     = {"기타사업부","수탁사업부","완제사업부"}
EXCLUDE_OFFICE = {"DW_B2B","DW_CH신사업팀"}
EXCLUDE_MGR    = {"1202150261"}
FONT           = "Noto Sans KR,Apple SD Gothic Neo,Malgun Gothic,sans-serif"

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

# ──────────────────────────────────────────────────────────────
def load_contract(file_bytes, sap_df):
    import openpyxl as _xl
    # SAP 기준파일 이름→코드 매핑 (이름 매칭 fallback용)
    s2 = sap_df.copy()
    _name_col = next((c for c in s2.columns if '판매처명' in c or '명칭' in c), None)
    if _name_col:
        s2['_norm'] = s2[_name_col].apply(_norm)
        sm = dict(zip(s2['_norm'], s2.get('판매처', s2.iloc[:, 0])))
        sn_list = list(s2['_norm']); sc_list = list(sm.values())
    else:
        sm = {}; sn_list = []; sc_list = []

    try:
        wb = _xl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
        ws = wb['계약서'] if '계약서' in wb.sheetnames else wb.active
    except Exception:
        return {}, 0, 0

    # 헤더 행 자동 탐지: 3가지 형식 지원
    # 'full'   : 거래처코드 컬럼 있음 (로컬 계약서)
    # 'summary': 판매처명+계약유형만 있음, 코드 없음 (서버 계약서)
    # 'raw'    : 실납처 형식
    FIXED = {}
    fmt = None
    data_start = 12
    for ridx, hrow in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), start=1):
        cells = [str(v or '').replace('\n', '').strip() for v in hrow]
        if any('거래처코드' in c for c in cells):
            fmt = 'full'; data_start = ridx + 1
            for i, c in enumerate(cells):
                if '거래처코드' in c: FIXED['code'] = i
                if '계약유형'  in c:  FIXED['type'] = i
                if '사업자' in c:     FIXED['biz'] = i
                if '지체보상금' in c or ('지체' in c and '보상' in c): FIXED['jiche'] = i
                if ('판매처' in c and '갑' in c) or '병원명' in c: FIXED['name'] = i
            break
        if any(c in ('실납처', '실 납처') for c in cells):
            fmt = 'raw'; data_start = ridx + 1
            for i, c in enumerate(cells):
                if c in ('실납처', '실 납처'):    FIXED['code'] = i
                if c in ('실납처명', '실 납처명'): FIXED['name'] = i
                if c == '유형':                    FIXED['type'] = i
            break
        # 거래처코드 없고 판매처(갑)+계약유형만 있는 요약 형식
        has_name = any('판매처' in c and '갑' in c for c in cells)
        has_type = any(c == '계약유형' for c in cells)
        if has_name and has_type and str(hrow[0] or '').strip() in ('No.', 'No', '번호', '순번'):
            fmt = 'summary'; data_start = ridx + 1
            for i, c in enumerate(cells):
                if '판매처' in c and '갑' in c: FIXED['name'] = i
                if c == '계약유형':              FIXED['type'] = i
                if '사업자' in c:               FIXED['biz'] = i
            FIXED['code'] = -1  # 코드 컬럼 없음 → 이름 매칭 전용
            # jiche 컬럼: 첫 데이터행에서 '지연' 텍스트 있는 컬럼 탐지
            for drow in ws.iter_rows(min_row=data_start, max_row=min(data_start+5, ws.max_row), values_only=True):
                if drow and str(drow[0] or '').strip().isdigit():
                    for i, v in enumerate(drow):
                        if v and ('지연' in str(v) or '1일당' in str(v) or '보상금' in str(v)):
                            FIXED['jiche'] = i; break
                    break
            break
    if fmt is None: fmt = 'full'
    FIXED.setdefault('code', 2); FIXED.setdefault('name', 1)
    FIXED.setdefault('type', 11); FIXED.setdefault('biz', 12); FIXED.setdefault('jiche', 24)

    result = {}
    for row in ws.iter_rows(min_row=data_start, max_row=ws.max_row, values_only=True):
        if not row or len(row) < 3: continue
        no_val = row[0]
        if not str(no_val or '').strip().isdigit(): continue

        def _cell(v):
            if v is None: return ''
            if isinstance(v, float): return str(int(v)) if v == int(v) else str(v)
            if isinstance(v, int):   return str(v)
            return str(v).strip()
        _ci = FIXED.get("code", 2)
        code  = (_cell(row[_ci]) if _ci >= 0 and len(row) > _ci else '')
        name  = _cell(row[FIXED["name"]]) if len(row) > FIXED["name"] else ''
        ctype = _cell(row[FIXED["type"]]) if len(row) > FIXED["type"] else ''
        _bi   = FIXED.get("biz", 12)
        biz   = _cell(row[_bi]) if _bi >= 0 and len(row) > _bi else ''
        _ji   = FIXED.get("jiche", 24)
        y_val = _cell(row[_ji]) if _ji >= 0 and len(row) > _ji else ''

        if fmt == 'raw':
            is_yes = True   # raw 파일에 있으면 계약 보유
            rate   = 1/1000
        else:
            is_yes = bool(y_val)
            rate_m = re.search(r'\[(\d+)/(\d+)\]', y_val)
            rate   = float(rate_m.group(1)) / float(rate_m.group(2)) if rate_m else 1/1000

        # 코드가 없거나 비어있으면 이름 매칭으로 보완
        if not code or code == 'nan':
            norm_n = _norm(name)
            code = sm.get(norm_n, '')
            if not code:
                for sn, sc in zip(sn_list, sc_list):
                    if len(sn) >= 4 and len(norm_n) >= 4 and (norm_n in sn or sn in norm_n):
                        code = sc; break

        if not code or code == 'nan': continue
        if code not in result or not result[code]["지체여부"]:
            result[code] = {
                "지체여부": is_yes, "요율": rate,
                "계약유형":  ctype if ctype not in ('nan','') else '',
                "사업자번호": biz   if biz   not in ('nan','') else '',
            }

    try: wb.close()
    except: pass
    n_yes = sum(1 for v in result.values() if v["지체여부"])
    n_no  = sum(1 for v in result.values() if not v["지체여부"])
    return result, n_yes, n_no

def load_base(file_bytes, cmap=None, think_rotations=None):
    df = pd.read_excel(io.BytesIO(file_bytes), dtype=str).dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]
    def _norm_code(v):
        s = str(v).strip()
        if '.' in s:
            try: return str(int(float(s)))
            except: pass
        return s
    out = []
    for _, r in df.iterrows():
        code = _norm_code(r.get("판매처","")); name = str(r.get("판매처명","")).strip()
        base = _num(r.get("기준회전일",0)); yak = _num(r.get("약정회전일",0))
        exc  = _num(r.get("예외회전일",0))
        if not code or code == "nan" or pd.isna(base) or base <= 0: continue
        eff = (int(exc) if not pd.isna(exc) and exc > 0
               else (int(yak) if not pd.isna(yak) and yak > 0 else int(base)))
        if eff == 40: eff = 30  # 씽크 계약 표준 30일 적용
        info = (cmap or {}).get(code, {})
        # cmap 제공 시 미등록 거래처는 지체 대상 아님 (계약 없음)
        use = info.get("지체여부", cmap is None); rate = info.get("요율", 1/1000)
        out.append({"코드":code,"명칭":name,"기준회전일":int(base),
                    "유효기준회전일":eff,"지체여부":use,"요율":rate})
    existing = {r["코드"] for r in out}
    if think_rotations:
        for code, rot in think_rotations.items():
            if code not in existing:
                info = (cmap or {}).get(code, {})
                use = info.get("지체여부", cmap is None); rate = info.get("요율", 1/1000)
                out.append({"코드":code,"명칭":"","기준회전일":rot,
                            "유효기준회전일":rot,"지체여부":use,"요율":rate})
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
    row2 = df.iloc[2].tolist()
    row3 = df.iloc[3].tolist()
    mo_cols = {}
    cur = None
    for i, (r2, r3) in enumerate(zip(row2, row3)):
        s2 = str(r2).strip()
        if '.' in s2 and s2 not in ('nan','None'):
            cur = s2; mo_cols[cur] = {}
        if cur and str(r3).strip() in ('매출','수금') and str(r3).strip() not in mo_cols[cur]:
            mo_cols[cur][str(r3).strip()] = i
    valid = {k: v for k, v in mo_cols.items()
             if '합계' not in k and '계약' not in k and len(k) == 7}
    months_dict = {}
    for _, row in df.iloc[4:].iterrows():
        code = str(row.iloc[2]).strip()
        name = str(row.iloc[3]).strip()
        if not code or code in ('nan','None','판매처'): continue
        기준 = _parse_rot(row.iloc[8])
        cum_sale = cum_coll = 0.0
        prev_현회전일 = 0
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
            cum_sale += sale; cum_coll += coll
            잔고 = cum_sale - cum_coll
            days = calendar.monthrange(yr, mo)[1]
            if 잔고 <= 0:
                prev_현회전일 = 0
                continue
            if sale > 0:
                현회전일 = 1
            else:
                현회전일 = prev_현회전일 + days
            prev_현회전일 = 현회전일
            key = (yr, mo)
            if key not in months_dict: months_dict[key] = []
            months_dict[key].append({
                "연": yr, "월": mo, "코드": code, "명칭": name,
                "채널": "씽크", "잔고": int(잔고), "현회전일": 현회전일,
                "_기준회전일_override": 기준,
                "사무소": "", "사업부": "", "담당자": "",
            })
    months = [{"yr": yr, "mo": mo, "data": pd.DataFrame(rows)}
              for (yr, mo), rows in sorted(months_dict.items())]
    return months

def load_monthly(file_bytes, think_codes=None):
    df = pd.read_excel(io.BytesIO(file_bytes), dtype=str).dropna(how="all")
    df.columns = [str(c).strip() for c in df.columns]
    def _ym(v):
        s = str(v).strip()
        return (int(s.split(".")[0]), int(s.split(".")[1])) if "." in s else (None, None)
    df["_yr"], df["_mo"] = zip(*df["년월"].map(_ym))
    df = df[df["_yr"].notna()].copy()
    df["_yr"] = df["_yr"].astype(int); df["_mo"] = df["_mo"].astype(int)
    df["_잔고"] = df["순실잔고"].apply(_num); df["_일수"] = df["순실회전일"].apply(_num)
    df = df.dropna(subset=["_잔고","_일수"]).copy(); df["_일수"] = df["_일수"].astype(int)
    # 당월 발생 미수금(납기 미도달): 잔고(1∼30) 버킷 — 지체보상금 기준 제외 대상
    _col_1_30 = next((c for c in df.columns if '잔고(1' in c and '30' in c), None)
    df["_이번달잔고"] = df[_col_1_30].apply(_num).fillna(0) if _col_1_30 else 0
    df = df[df["판매처"].notna() & (df["판매처"].str.strip() != "") & (df["판매처"] != "nan")]
    df = df[df["채널"].isin(ALLOW_CH)]; df = df[df["_잔고"] != 0]
    # 씽크 채널: 2025년 이전 데이터 제외 (12월 매출 → 납기 미도달, 지체보상금 구조적 불가)
    df = df[df["_yr"] >= 2025]
    # DW07은 2026년부터 씽크 신규 매출 없음 → 이번달잔고=0 강제 (잔고(1~30) 버킷에 비씽크 매출 혼입 방지)
    dw07_2026 = (df["채널"] == "DW07") & (df["_yr"] >= 2026)
    df.loc[dw07_2026, "_이번달잔고"] = 0.0
    if think_codes and "채널" in df.columns:
        dw07_mask = df["채널"] == "DW07"
        df = df[~dw07_mask | df["판매처"].isin(think_codes)]
    if "담당자"   in df.columns: df = df[~df["담당자"].isin(EXCLUDE_MGR)]
    if "사업부명" in df.columns: df = df[~df["사업부명"].apply(lambda x: any(k in str(x) for k in EXCLUDE_BU))]
    df = df[~df["판매처"].isin(EXCLUDE_CODES)]
    if "사무소명" in df.columns: df = df[~df["사무소명"].isin(EXCLUDE_OFFICE)]
    # 여러 파일 합산 시 같은 거래처·채널·년월 중복 제거 (마지막 파일 값 우선)
    df = df.drop_duplicates(subset=["년월", "판매처", "채널"], keep="last")
    months = []
    for (yr, mo), grp in df.groupby(["_yr","_mo"]):
        def _nc(v):
            s = str(v).strip()
            if '.' in s:
                try: return str(int(float(s)))
                except: pass
            return s
        rows = [{"연":int(yr),"월":int(mo),"코드":_nc(r["판매처"]),
                 "명칭":str(r.get("판매처명","")).strip(),"채널":str(r.get("채널","")).strip(),
                 "잔고":r["_잔고"],"현회전일":r["_일수"],
                 "이번달잔고": float(r.get("_이번달잔고", 0) or 0),
                 "사무소":str(r.get("사무소명","")).strip() if "사무소명" in r.index else "",
                 "사업부":str(r.get("사업부명","")).strip() if "사업부명" in r.index else "",
                 "담당자":str(r.get("담당자명","")).strip() if "담당자명" in r.index else "",
                 } for _, r in grp.iterrows()]
        months.append({"yr":int(yr),"mo":int(mo),"data":pd.DataFrame(rows)})
    months.sort(key=lambda x: (x["yr"], x["mo"]))
    return months

def calculate(months, base_df):
    bmap = {r["코드"]: r for _, r in base_df.iterrows()}
    results = []
    for m in months:
        yr, mo = m["yr"], m["mo"]
        for _, row in m["data"].iterrows():
            code = row["코드"]
            if code not in bmap: continue
            b = bmap[code]
            if not b["지체여부"]: continue
            기준 = int(row["_기준회전일_override"]) if "_기준회전일_override" in row and pd.notna(row["_기준회전일_override"]) else b["유효기준회전일"]
            요율 = b["요율"]
            현 = row["현회전일"]; 잔 = row["잔고"]
            이번달 = float(row.get("이번달잔고", 0) or 0)
            연체기준잔 = max(0.0, 잔 - 이번달)  # 당월 납기 미도달분 제외
            지연 = max(0, 현 - 기준)
            보상 = round(연체기준잔 * 지연 * 요율) if 지연 > 0 else 0
            rate_str = f"1/{int(round(1/요율))}" if 요율 > 0 else "-"
            results.append({
                "기준월":f"{yr}-{mo:02d}","판매처코드":code,"판매처명":row["명칭"],
                "채널":row["채널"],
                "사무소":row.get("사무소",""),"사업부":row.get("사업부",""),"담당자":row.get("담당자",""),
                "현 미수금":int(잔),"연체기준잔고":int(연체기준잔),"기준회전일(SAP)":b["기준회전일"],
                "유효기준회전일":기준,"현 회전일":현,"지연일수(누적)":지연,
                "지체보상금":보상,"요율":rate_str,
                "비고":"청구" if 보상 > 0 else "",
            })
    return pd.DataFrame(results)

def make_excel(df, title="전체내역"):
    wb = openpyxl.Workbook(); HDR = PatternFill("solid", fgColor="18293f")
    def _h(ws):
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF", size=10)
            c.fill = HDR; c.alignment = Alignment(horizontal="center", vertical="center")
    w = {"기준월":10,"판매처코드":12,"판매처명":22,"채널":10,"현 미수금":14,
         "기준회전일(SAP)":12,"유효기준회전일":12,"현 회전일":10,"지연일수(누적)":12,
         "지체보상금":14,"요율":8,"비고":8}
    ws = wb.active; ws.title = title
    cols = list(df.columns); ws.append(cols)
    for r in df.itertuples(index=False): ws.append(list(r))
    _h(ws)
    for ci, cn in enumerate(cols, 1):
        for ri in range(2, ws.max_row+1):
            c = ws.cell(ri, ci)
            if cn in {"현 미수금","지체보상금"}: c.number_format = "#,##0"
            c.alignment = Alignment(horizontal="center")
        ws.column_dimensions[gcl(ci)].width = w.get(cn, 12)
    ws.freeze_panes = "A2"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

def make_full_excel(df):
    wb = openpyxl.Workbook(); HDR = PatternFill("solid", fgColor="18293f")
    def _h(ws):
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF", size=10)
            c.fill = HDR; c.alignment = Alignment(horizontal="center", vertical="center")
    w = {"기준월":10,"판매처코드":12,"판매처명":22,"채널":10,"현 미수금":14,
         "기준회전일(SAP)":12,"유효기준회전일":12,"현 회전일":10,"지연일수(누적)":12,
         "지체보상금":14,"요율":8,"비고":8}
    ws1 = wb.active; ws1.title = "전체내역"
    cols = list(df.columns); ws1.append(cols)
    for r in df.itertuples(index=False): ws1.append(list(r))
    _h(ws1)
    for ci, cn in enumerate(cols, 1):
        for ri in range(2, ws1.max_row+1):
            c = ws1.cell(ri, ci)
            if cn in {"현 미수금","지체보상금"}: c.number_format = "#,##0"
            c.alignment = Alignment(horizontal="center")
        ws1.column_dimensions[gcl(ci)].width = w.get(cn, 12)
    ws1.freeze_panes = "A2"
    ws2 = wb.create_sheet("거래처별누계")
    pv = df.pivot_table(index=["판매처코드","판매처명"], columns="기준월",
                        values="지체보상금", aggfunc="sum", fill_value=0).reset_index()
    pc = list(pv.columns); ws2.append(pc)
    for r in pv.itertuples(index=False): ws2.append(list(r))
    _h(ws2)
    for row in ws2.iter_rows(min_row=2):
        for c in row:
            if isinstance(c.value, (int, float)) and c.value != 0: c.number_format = "#,##0"
            c.alignment = Alignment(horizontal="center")
    ws2.column_dimensions["A"].width = 12; ws2.column_dimensions["B"].width = 22
    for i in range(3, len(pc)+1): ws2.column_dimensions[gcl(i)].width = 14
    ws2.freeze_panes = "C2"
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

def _공문_doc_setup(doc):
    sec0 = doc.sections[0]
    sec0.top_margin = _Cm(1.3); sec0.bottom_margin = _Cm(1.0)
    sec0.left_margin = _Cm(1.8); sec0.right_margin = _Cm(1.8)
    def _rf(run, name='맑은 고딕'):
        rPr = run._r.get_or_add_rPr()
        for t in list(rPr):
            if t.tag == _dqn('w:rFonts'): rPr.remove(t)
        rf = _dxml('w:rFonts')
        rf.set(_dqn('w:ascii'), name); rf.set(_dqn('w:eastAsia'), name); rf.set(_dqn('w:hAnsi'), name)
        rPr.insert(0, rf)
    def _r(para, text, bold=False, size=10, color=None, name='맑은 고딕'):
        run = para.add_run(text)
        run.bold = bold; run.font.size = _Pt(size); run.font.name = name
        if color: run.font.color.rgb = _RGB.from_string(color)
        _rf(run, name); return run
    def _p(text='', bold=False, size=10, align=_WDA.LEFT, sa=2, color=None):
        p = doc.add_paragraph(); p.alignment = align
        p.paragraph_format.space_after = _Pt(sa); p.paragraph_format.space_before = _Pt(0)
        if text: _r(p, text, bold=bold, size=size, color=color)
        return p
    def _bline(side='bottom', sa=6):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = _Pt(sa); p.paragraph_format.space_before = _Pt(0)
        pPr = p._p.get_or_add_pPr()
        pBdr = _dxml('w:pBdr')
        bdr = _dxml(f'w:{side}')
        bdr.set(_dqn('w:val'), 'single'); bdr.set(_dqn('w:sz'), '18')
        bdr.set(_dqn('w:space'), '1'); bdr.set(_dqn('w:color'), '1A2A6C')
        pBdr.append(bdr); pPr.append(pBdr)
    def _tc(cell, text, bold=False, size=9.5, color=None, fill=None, align=_WDA.LEFT):
        para = cell.paragraphs[0]
        run = para.add_run(str(text))
        run.bold = bold; run.font.size = _Pt(size); run.font.name = '맑은 고딕'
        if color: run.font.color.rgb = _RGB.from_string(color)
        _rf(run); para.alignment = align
        if fill:
            tcPr = cell._tc.get_or_add_tcPr()
            for tag in list(tcPr):
                if tag.tag == _dqn('w:shd'): tcPr.remove(tag)
            shd = _dxml('w:shd')
            shd.set(_dqn('w:val'), 'clear'); shd.set(_dqn('w:color'), 'auto'); shd.set(_dqn('w:fill'), fill)
            tcPr.append(shd)
    def _sig():
        p1 = _p(sa=0, align=_WDA.RIGHT); _r(p1, '서울시 강남구 봉은사로 114길 12', size=9.5)
        p2 = _p(sa=0, align=_WDA.RIGHT); _r(p2, '주 식 회 사  대 웅 제 약', size=10)
        p3 = _p(sa=3, align=_WDA.RIGHT); _r(p3, '대 표 이 사   이 창 재', size=10)
    def _footer():
        _bline('top', sa=4)
        _p('DAEWOONG PHARMACEUTICAL CO., LTD.', bold=True, size=10, align=_WDA.CENTER, sa=0)
        _p('대웅제약 주식회사', size=8.5, align=_WDA.CENTER, sa=0)
    def _letterhead(판매처명, 제목, doc_no_suffix):
        from datetime import date as _d
        today = _d.today()
        p_lh = _p(sa=1); _r(p_lh, '(주)  대 웅 제 약', bold=True, size=20)
        _bline('bottom', sa=3)
        th = doc.add_table(rows=1, cols=2)
        for cell in th.rows[0].cells: cell._tc.get_or_add_tcPr()
        _tc(th.rows[0].cells[0], f'문서번호 : {doc_no_suffix}')
        _tc(th.rows[0].cells[1], today.strftime("%Y년 %m월 %d일"), align=_WDA.RIGHT)
        _p(sa=2)
        p = _p(sa=1); _r(p, '수  신', bold=True, size=10.5); _r(p, f'  :  {판매처명} 귀중', size=10.5)
        p = _p(sa=3); _r(p, '제  목', bold=True, size=10.5); _r(p, f'  :  {제목}', bold=True, size=10.5)
        _bline('top', sa=5)
        return today
    return _r, _p, _bline, _tc, _sig, _footer, _letterhead

def make_공문_word(row):
    if not _HAS_DOCX: raise ImportError("pip install python-docx 후 재시작 필요")
    doc = _DocxDoc()
    _r, _p, _bline, _tc, _sig, _footer, _letterhead = _공문_doc_setup(doc)
    판매처명 = str(row.get("판매처명", ""))
    기준월   = str(row.get("기준월", ""))
    _letterhead(판매처명, "지체보상금 청구의 건", f"지체-{기준월.replace('-','')}-{판매처명[:4]}")
    _p('1.', bold=True, size=10, sa=2)
    p = _p(sa=3)
    _r(p, '귀원의 무궁한 발전을 기원합니다.\n', size=10)
    _r(p, '당사 대웅제약과 귀 병원 간에 체결된 thync™ 상품공급 계약서에 따라 이루어진 거래와 관련하여 ', size=10)
    _r(p, 기준월, bold=True, size=10)
    _r(p, ' 기준 지체보상금이 아래와 같이 발생하였기에 청구드립니다.', size=10)
    _p('─ 청구 내역', bold=True, size=10, sa=2)
    details = [
        ('· 기준월', 기준월),
        ('· 현 미수금 잔액', f'{int(row["현 미수금"]):,} 원'),
        ('· 유효 기준회전일', f'{int(row["유효기준회전일"])}일'),
        ('· 현 회전일', f'{int(row["현 회전일"])}일'),
        ('· 지연일수', f'{int(row["지연일수(누적)"])}일'),
        ('· 요율', str(row["요율"])),
        ('· 지체보상금', f'{int(row["지체보상금"]):,} 원'),
        ('· 계산 근거', str(row.get("계산 근거", ""))),
    ]
    tbl = doc.add_table(rows=len(details), cols=2); tbl.style = 'Table Grid'
    for i, (lbl, val) in enumerate(details):
        _tc(tbl.rows[i].cells[0], lbl, bold=True, fill='EEEEF5')
        _tc(tbl.rows[i].cells[1], val, color='C0392B' if '지체보상금' in lbl else None)
    for r_obj in tbl.rows:
        r_obj.cells[0].width = _Cm(5.5); r_obj.cells[1].width = _Cm(9.5)
    사업자번호 = str(row.get("사업자번호", "")).strip()
    계약유형   = str(row.get("계약유형",   "")).strip()
    parts = [x for x in [
        f'사업자번호: {사업자번호}' if 사업자번호 not in ("", "nan") else "",
        f'계약유형: {계약유형}'     if 계약유형   not in ("", "nan") else "",
    ] if x]
    if parts: _p(f'※ {" | ".join(parts)}', size=8.5, sa=3, color='666666')
    else:     _p(sa=3)
    _p('2.', bold=True, size=10, sa=2)
    p = _p(sa=5)
    _r(p, '상기 내용에 대해 확인하시어, 가능하신 ', size=10)
    _r(p, '______년  ______월  ______일', size=10)
    _r(p, ' 일정에 맞추어 회신을 부탁드립니다.', size=10)
    _p('거래처명 :                                                                                   (인)', size=10, sa=6)
    _sig(); _footer()
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()

def make_공문_word_cumulative(vendor_df):
    if not _HAS_DOCX: raise ImportError("pip install python-docx 후 재시작 필요")
    doc = _DocxDoc()
    _r, _p, _bline, _tc, _sig, _footer, _letterhead = _공문_doc_setup(doc)
    first = vendor_df.sort_values("기준월").iloc[0]
    months = sorted(vendor_df["기준월"].unique())
    period = f"{months[0]} ~ {months[-1]}"
    판매처명 = str(first.get("판매처명", ""))
    total = int(vendor_df["지체보상금"].sum())
    from datetime import date as _d
    today = _d.today()
    _letterhead(판매처명, "지체보상금 청구의 건 (전체 기간)",
                f"지체-{today.strftime('%Y')}-전체-{판매처명[:4]}")
    _p('1.', bold=True, size=10, sa=2)
    p = _p(sa=3)
    _r(p, '귀원의 무궁한 발전을 기원합니다.\n', size=10)
    _r(p, '당사 대웅제약과 귀 병원 간에 체결된 thync™ 상품공급 계약서에 따라 이루어진 거래와 관련하여 ', size=10)
    _r(p, period, bold=True, size=10)
    _r(p, ' 기간 동안 발생한 지체보상금 합계를 아래와 같이 청구드립니다.', size=10)
    _p('─ 월별 청구 내역', bold=True, size=10, sa=2)
    hdrs = ["기준월", "연체기준잔고", "기준회전일", "현 회전일", "지연일수", "요율", "지체보상금"]
    sorted_df = vendor_df.sort_values("기준월")
    n = len(sorted_df)
    tbl = doc.add_table(rows=n + 2, cols=len(hdrs)); tbl.style = 'Table Grid'
    for ci, h in enumerate(hdrs):
        _tc(tbl.rows[0].cells[ci], h, bold=True, color='FFFFFF', fill='1A2A6C', align=_WDA.CENTER)
    for ri, (_, r) in enumerate(sorted_df.iterrows(), start=1):
        bg = 'F0F0F8' if ri % 2 == 0 else None
        vals = [r["기준월"], f'{int(r.get("연체기준잔고", r["현 미수금"])):,}', f'{int(r["유효기준회전일"])}',
                f'{int(r["현 회전일"])}', f'{int(r["지연일수(누적)"])}', r["요율"], f'{int(r["지체보상금"]):,}']
        for ci, v in enumerate(vals):
            _tc(tbl.rows[ri].cells[ci], v, size=9.5, fill=bg, align=_WDA.CENTER)
    last = tbl.rows[n + 1]
    _tc(last.cells[0], "합  계", bold=True, fill='E8EDF5', align=_WDA.CENTER)
    for ci in range(1, len(hdrs) - 1):
        tcPr = last.cells[ci]._tc.get_or_add_tcPr()
        for tag in list(tcPr):
            if tag.tag == _dqn('w:shd'): tcPr.remove(tag)
        shd = _dxml('w:shd')
        shd.set(_dqn('w:val'), 'clear'); shd.set(_dqn('w:color'), 'auto'); shd.set(_dqn('w:fill'), 'E8EDF5')
        tcPr.append(shd)
    _tc(last.cells[-1], f'{total:,}', bold=True, color='C0392B', fill='E8EDF5', align=_WDA.CENTER)
    widths = [_Cm(2.2), _Cm(3.0), _Cm(1.8), _Cm(1.8), _Cm(2.2), _Cm(1.5), _Cm(2.5)]
    for r_obj in tbl.rows:
        for ci, w in enumerate(widths): r_obj.cells[ci].width = w
    _p(sa=2)
    p_tot = _p(sa=5, align=_WDA.RIGHT)
    _r(p_tot, '※ 합계 청구금액 : ', bold=True, size=11)
    _r(p_tot, f'{total:,} 원', bold=True, size=13, color='C0392B')
    _p('2.', bold=True, size=10, sa=2)
    p = _p(sa=5)
    _r(p, '상기 내용에 대해 확인하시어, 가능하신 ', size=10)
    _r(p, '______년  ______월  ______일', size=10)
    _r(p, ' 일정에 맞추어 회신을 부탁드립니다.', size=10)
    _p('거래처명 :                                                                                   (인)', size=10, sa=6)
    _sig(); _footer()
    buf = io.BytesIO(); doc.save(buf); return buf.getvalue()

def make_공문_zip(df, month_label):
    buf = io.BytesIO()
    def _calc_str(r):
        if r["지체보상금"] <= 0: return "—"
        base = int(r.get("연체기준잔고", r["현 미수금"]))
        지연 = int(r.get("지연일수(누적)", 0))
        return f'{base:,} × {지연}일 × {r["요율"]} = {int(r["지체보상금"]):,}원'
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for _, row in df[df["지체보상금"] > 0].sort_values("지체보상금", ascending=False).iterrows():
            row = row.copy(); row["계산 근거"] = _calc_str(row)
            safe = re.sub(r'[\\/:*?"<>|]', '_', str(row["판매처명"]).strip())
            fname = f"지체보상금청구_{safe}_{row['기준월']}.docx"
            zf.writestr(fname, make_공문_word(row))
    return buf.getvalue()

def make_공문_zip_cumulative(result_df):
    charge_df = result_df[result_df["지체보상금"] > 0]
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for code, grp in charge_df.groupby("판매처코드"):
            name = grp["판매처명"].iloc[0]
            safe = re.sub(r'[\\/:*?"<>|]', '_', str(name).strip())
            fname = f"지체보상금청구_{safe}_전체기간.docx"
            zf.writestr(fname, make_공문_word_cumulative(grp))
    return buf.getvalue()

def load_raw_monthly_sales(file_bytes):
    """raw 파일 → {(코드, yr, mo): 매출금액} — 당월 매출 제외 계산용"""
    import openpyxl as _xl
    try:
        wb = _xl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
        ws = wb.active
    except Exception:
        return {}
    all_rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))
    if len(all_rows) < 4:
        try: wb.close()
        except: pass
        return {}
    date_row = all_rows[1]  # Excel 2행: 날짜 헤더 (YYYY.MM 또는 N월)
    date_cols = {}
    last_yr = None
    for i, v in enumerate(date_row):
        s = str(v or '').strip()
        if len(s) == 7 and s[4] == '.':  # YYYY.MM 형식
            try:
                yr, mo = int(s[:4]), int(s[5:])
                date_cols[(yr, mo)] = i; last_yr = yr
            except: pass
        elif s.endswith('월') and s[:-1].isdigit():  # N월 형식 → 다음 연도
            try:
                mo = int(s[:-1])
                yr = (last_yr + 1) if last_yr else 2026
                date_cols[(yr, mo)] = i
            except: pass
    result = {}
    for row in all_rows[3:]:  # Excel 4행부터 데이터
        if not row or not str(row[0] or '').strip().isdigit(): continue
        def _code(v):
            if v is None: return ''
            if isinstance(v, float): return str(int(v)) if v == int(v) else str(v)
            if isinstance(v, int): return str(v)
            return str(v).strip()
        def _amt(v):
            if v is None: return 0.0
            if isinstance(v, (int, float)): return float(v)
            try: return float(str(v).replace(',', '').strip() or 0)
            except: return 0.0
        code = _code(row[1])
        if not code: continue
        for (yr, mo), ci in date_cols.items():
            if ci < len(row):
                amt = _amt(row[ci])
                if amt: result[(code, yr, mo)] = amt
    try: wb.close()
    except: pass
    return result

def _load_biz_map(file_bytes):
    """계약서 파일 → {거래처코드: 사업자번호}
    거래처코드 컬럼 있으면 코드 직접 매칭, 없으면 판매처명 fallback.
    _계약파기/대시 정규화, 파기 아닌 계약 우선."""
    import openpyxl as _xl
    try:
        wb = _xl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
        ws = wb['계약서'] if '계약서' in wb.sheetnames else wb.active
    except Exception:
        return {}
    code_col = biz_col = name_col = None; data_start = 12
    for ridx, row in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), 1):
        cells = [str(v or '').strip() for v in row]
        _cc = next((i for i, c in enumerate(cells) if '거래처코드' in c), None)
        _bc = next((i for i, c in enumerate(cells) if '사업자' in c), None)
        _nc = next((i for i, c in enumerate(cells) if '판매처' in c and '갑' in c), None)
        if _bc is not None and (_cc is not None or _nc is not None):
            code_col = _cc; biz_col = _bc; name_col = _nc; data_start = ridx + 1; break
    if biz_col is None:
        try: wb.close()
        except: pass
        return {}
    def _c(v):
        if v is None: return ''
        if isinstance(v, float): return str(int(v)) if v == int(v) else str(v)
        return str(v).strip()
    temp = {}  # key -> (biz, is_파기)
    for row in ws.iter_rows(min_row=data_start, max_row=ws.max_row, values_only=True):
        if not row or not str(row[0] or '').strip().isdigit(): continue
        # 키: 거래처코드 우선, 없으면 정규화 병원명
        if code_col is not None and len(row) > code_col:
            key = _c(row[code_col])
        elif name_col is not None and len(row) > name_col:
            key = _norm(_c(row[name_col]))
        else:
            continue
        if not key: continue
        biz_raw = _c(row[biz_col]) if len(row) > biz_col else ''
        is_파기 = '_계약파기' in biz_raw
        biz = biz_raw.split('_')[0].replace('-', '').strip()
        if key not in temp or (not is_파기 and temp[key][1]):
            temp[key] = (biz, is_파기)
    try: wb.close()
    except: pass
    return {k: v[0] for k, v in temp.items()}

def load_raw_vendors(file_bytes):
    """raw 씽크 파일 → {코드: {명칭, 기준회전일, 계약유형}} — vendor code 직접 포함"""
    import openpyxl as _xl
    wb = _xl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb.active
    result = {}
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, values_only=True):
        if not row or len(row) < 9: continue
        no_val = str(row[0] or '').strip()
        if not no_val.isdigit(): continue
        def _c(v):
            if v is None: return ''
            if isinstance(v, float): return str(int(v)) if v == int(v) else str(v)
            if isinstance(v, int): return str(v)
            return str(v).strip()
        code = _c(row[1])
        if not code or code == 'nan': continue
        name = _c(row[2])
        유형 = _c(row[5])
        rot_str = _c(row[8])
        try: rot = int(''.join(c for c in rot_str if c.isdigit()) or '30')
        except: rot = 30
        if rot <= 0: rot = 30
        result[code] = {'명칭': name, '기준회전일': rot, '계약유형': 유형}
    try: wb.close()
    except: pass
    return result

# ──────────────────────────────────────────────────────────────
BASE_PATH = Path("input/기준회전일/SAP 거래처별 기준회전일_DW11_0722.xlsx")
MO_FOLDER = Path("input/월별잔고회전일")
RAW_PATH  = Path("input/계약정보/raw_thync.xlsx")
THINK_PATH = next(iter(sorted(Path("input").glob("씽크*.xlsx")) + sorted(Path("input").glob("씽크*.XLSX"))), None)
# raw_thync.xlsx 제외: CT_LIST에 섞이면 알파벳 순으로 계약서 파일보다 앞서 CT_PATH가 raw 파일을 잡음
CT_LIST   = [f for f in sorted(Path("input/계약정보").glob("*.xlsx")) + sorted(Path("input/계약정보").glob("*.xls"))
             if f.name not in ("raw_thync.xlsx", "ct_biz.xlsx")]
CT_PATH   = CT_LIST[0] if CT_LIST else None
CT_BIZ_PATH   = Path("input/계약정보/ct_biz.xlsx")    # 보조 사업자번호 소스 (이름 매칭)
THINK25_PATH  = Path("input/계약정보/think25.xlsx")   # 씽크 25년 현황 — 코드 직접 매핑 (187/187)
CT_CSV    = Path("input/계약정보/계약_매핑결과.csv")

@st.cache_data(show_spinner=False)
def _read_mo_folder(folder_str: str, _mtime: float = 0):
    folder = Path(folder_str)
    files = sorted(folder.glob("*.xlsx")) + sorted(folder.glob("*.xls"))
    # 기준·가공 안내 파일 제외 (월별 잔고 데이터가 아님)
    files = [f for f in files if not any(k in f.name for k in ["기준", "가공", "~$"])]
    if not files: return None
    dfs = [pd.read_excel(f, dtype=str).dropna(how="all") for f in files]
    combined = pd.concat(dfs, ignore_index=True)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        combined.to_excel(w, index=False)
    return buf.getvalue(), [f.name for f in files]

@st.cache_data(show_spinner=False)
def _read_bytes(path_str: str):
    p = Path(path_str)
    return p.read_bytes() if p.exists() else None

# ── 파일 로드
_mo_files = (sorted(MO_FOLDER.glob("*.xlsx")) + sorted(MO_FOLDER.glob("*.xls"))) if MO_FOLDER.exists() else []
_mo_result = _read_mo_folder(str(MO_FOLDER), _mtime=max((f.stat().st_mtime for f in _mo_files if f.exists()), default=0)) if MO_FOLDER.exists() else None
m_b, _mo_names = _mo_result if _mo_result else (None, [])
b_b = _read_bytes(str(BASE_PATH))
c_b = _read_bytes(str(CT_PATH)) if CT_PATH else None
t_b = _read_bytes(str(THINK_PATH)) if THINK_PATH else None
csv_b = CT_CSV.read_bytes() if CT_CSV.exists() else None
raw_b  = _read_bytes(str(RAW_PATH))
biz2_b     = _read_bytes(str(CT_BIZ_PATH))   # 보조 사업자번호 파일
think25_b  = _read_bytes(str(THINK25_PATH))  # 씽크 25년 현황 (코드 직접 매핑)

def _build_org_map(m_b):
    if not m_b: return {}
    try:
        df = pd.read_excel(io.BytesIO(m_b), dtype=str).dropna(how="all")
        df.columns = [str(c).strip() for c in df.columns]
        org_map = {}
        for _, r in df.iterrows():
            code = str(r.get("판매처","")).strip()
            if not code or code == "nan": continue
            org_map[code] = {
                "사무소": str(r.get("사무소명","")).strip() if "사무소명" in r.index else "",
                "사업부": str(r.get("사업부명","")).strip() if "사업부명" in r.index else "",
                "담당자": str(r.get("담당자명","")).strip() if "담당자명" in r.index else "",
            }
        return org_map
    except Exception: return {}

_CALC_VER = "v28"  # 계산 로직 변경 시 증가 → 캐시 자동 무효화

@st.cache_data(show_spinner=False)
def run_calc(m_b, b_b, c_b, t_b, csv_b=None, raw_b=None, biz2_b=None, think25_b=None, ver=_CALC_VER):
    months = load_monthly(m_b)
    org_map = _build_org_map(m_b)

    # ── raw 파일 우선: 거래처코드 직접 포함, 기준회전일·계약유형 정확
    raw_vendors = {}; sales_map = {}
    # raw 파일 월별 매출 → 이번달잔고 override (잔고(1∼30) 대신 실제 매출액 사용)
    if raw_b:
        try:
            sales_map = load_raw_monthly_sales(raw_b)
            if sales_map:
                for m_data in months:
                    yr, mo = m_data["yr"], m_data["mo"]
                    df_m = m_data["data"]
                    df_m["이번달잔고"] = df_m.apply(
                        lambda row: sales_map.get((row["코드"], yr, mo), row.get("이번달잔고", 0)),
                        axis=1
                    )
        except Exception:
            pass
    if raw_b:
        try: raw_vendors = load_raw_vendors(raw_b)
        except Exception: pass

    if raw_vendors:
        # raw 파일에 있는 거래처만 계산 대상
        bdf = pd.DataFrame([
            {"코드": code, "명칭": info["명칭"],
             "기준회전일": info["기준회전일"], "유효기준회전일": info["기준회전일"],
             "지체여부": True, "요율": 1/1000}
            for code, info in raw_vendors.items()
        ])
        # 사업자번호: think25.xlsx (코드 직접) → 계약서 파일 → 보조파일 (이름 매칭)
        # 1) think25.xlsx — 판매처(col4) + 사업자번호(col3), 187/187 커버
        think25_code_map = {}
        if think25_b:
            try:
                import openpyxl as _xl25
                _wb25 = _xl25.load_workbook(io.BytesIO(think25_b), data_only=True, read_only=True)
                _ws25 = _wb25['씽크현황'] if '씽크현황' in _wb25.sheetnames else _wb25.active
                def _c25(v):
                    if v is None: return ''
                    if isinstance(v, float): return str(int(v)) if v == int(v) else str(v)
                    return str(v).strip()
                for _row in _ws25.iter_rows(min_row=4, values_only=True):
                    if not _row: continue
                    _code = _c25(_row[4]); _biz = _c25(_row[3]).replace('-','').split('_')[0]
                    if _code and _biz and _code.isdigit(): think25_code_map[_code] = _biz
                try: _wb25.close()
                except: pass
            except Exception: pass
        biz_map = {}
        if c_b:
            try: biz_map = _load_biz_map(c_b)
            except Exception: pass
        # 보조 사업자번호 파일 (ct_biz.xlsx) — 이름 기반
        biz2_name = {}
        if biz2_b:
            try:
                import openpyxl as _xl2
                _wb2 = _xl2.load_workbook(io.BytesIO(biz2_b), data_only=True, read_only=True)
                _ws2 = _wb2.active
                _data_start = 7
                for _ridx, _row in enumerate(_ws2.iter_rows(min_row=1, max_row=15, values_only=True), 1):
                    _cells = [str(v or '').strip() for v in _row]
                    if any('판매처' in c for c in _cells) and any('사업자' in c for c in _cells):
                        _data_start = _ridx + 1; break
                def _cb(v):
                    if v is None: return ''
                    if isinstance(v, float): return str(int(v)) if v == int(v) else str(v)
                    return str(v).strip()
                for _row in _ws2.iter_rows(min_row=_data_start, values_only=True):
                    if not _row or not str(_row[0] or '').strip().isdigit(): continue
                    _name = _cb(_row[1]); _biz = _cb(_row[3]).replace('-','').split('_')[0]
                    if _name and _biz: biz2_name[_norm(_name)] = _biz
                try: _wb2.close()
                except: pass
            except Exception: pass
        cmap = {}
        for code, info in raw_vendors.items():
            biz = think25_code_map.get(code, '')  # 1순위: think25 코드 직접
            if not biz: biz = biz_map.get(code, '')  # 2순위: 계약서 코드
            if not biz: biz = biz_map.get(_norm(info["명칭"]), '')  # 3순위: 계약서 이름
            if not biz:  # 보조 파일 이름 매칭
                nk = _norm(info["명칭"])
                biz = biz2_name.get(nk, '')
                if not biz:
                    for k, b in biz2_name.items():
                        if len(k) >= 4 and len(nk) >= 4 and (nk in k or k in nk):
                            biz = b; break
            cmap[code] = {"지체여부": True, "요율": 1/1000,
                          "계약유형": info["계약유형"], "사업자번호": biz}
        n_yes = len(raw_vendors); n_no = 0
    else:
        # fallback: 기존 SAP + 계약서 로직
        sap_raw = pd.read_excel(io.BytesIO(b_b), dtype=str).dropna(how="all")
        sap_raw.columns = [str(c).strip() for c in sap_raw.columns]
        cmap = {}; n_yes = n_no = 0
        if c_b:
            try: cmap, n_yes, n_no = load_contract(c_b, sap_raw)
            except Exception: cmap = {}
        think_rot = load_think(t_b) if t_b else None
        _eff_cmap = cmap if cmap else None
        bdf = load_base(b_b, _eff_cmap, think_rot)
        if _eff_cmap is not None and not bdf.empty and not bdf["지체여부"].any():
            bdf = load_base(b_b, None, think_rot)

    result = calculate(months, bdf)
    if org_map and not result.empty:
        result["사무소"] = result["판매처코드"].map(lambda c: org_map.get(c, {}).get("사무소", ""))
        result["사업부"] = result["판매처코드"].map(lambda c: org_map.get(c, {}).get("사업부", ""))
        result["담당자"] = result["판매처코드"].map(lambda c: org_map.get(c, {}).get("담당자", ""))
    if cmap and not result.empty:
        result["계약유형"]  = result["판매처코드"].map(lambda c: cmap.get(c, {}).get("계약유형", ""))
        result["사업자번호"] = result["판매처코드"].map(lambda c: cmap.get(c, {}).get("사업자번호", ""))
    else:
        if "계약유형"  not in result.columns: result["계약유형"]  = ""
        if "사업자번호" not in result.columns: result["사업자번호"] = ""
    return result, cmap, n_yes, n_no

# ──────────────────────────────────────────────────────────────
# 헤더
# ──────────────────────────────────────────────────────────────
today_str = date.today().strftime("%Y. %m. %d")
st.markdown(f"""
<div class="top-header">
  <div class="hd-left">
    <div class="hd-title">대웅제약 지체보상금 대시보드</div>
    <div class="hd-sub">데이터 기준일: {today_str} &nbsp;|&nbsp; DW07 · DW11 채널 · 계약 지체보상금 조항 보유 거래처</div>
  </div>
  <div class="hd-badge">thync™ 채널</div>
</div>""", unsafe_allow_html=True)

# ── 데이터 계산
if not (m_b and b_b):
    st.warning("input 폴더에 월별잔고 파일 또는 기준회전일 파일이 없습니다.")
    st.stop()

with st.spinner("데이터 계산 중..."):
    result_df, contract_map, n_yes, n_no = run_calc(m_b, b_b, c_b, t_b, csv_b, raw_b=raw_b, biz2_b=biz2_b, think25_b=think25_b, ver=_CALC_VER)

if result_df is None or result_df.empty:
    st.warning("계산 결과가 없습니다."); st.stop()

months_sorted = sorted(result_df["기준월"].unique())
mo_sum = result_df.groupby("기준월")["지체보상금"].sum()

# ──────────────────────────────────────────────────────────────
# 메인 탭: 집계 | 상세
# ──────────────────────────────────────────────────────────────
tab_agg, tab_detail, tab_guide = st.tabs(["📊  집계표", "🔍  월별 상세", "📖  계산 가이드"])

# ══════════════════════════════════════════════════════════════
# 탭 1: 집계표
# ══════════════════════════════════════════════════════════════
with tab_agg:
    # ── 전체 기간 KPI ─────────────────────────────────────────
    total_all   = int(result_df["지체보상금"].sum())
    total_vend  = result_df[result_df["지체보상금"] > 0]["판매처코드"].nunique()
    total_months = len([m for m in months_sorted if mo_sum.get(m, 0) > 0])
    best_mo      = mo_sum.idxmax() if len(mo_sum) > 0 else "-"
    best_mo_amt  = int(mo_sum.max()) if len(mo_sum) > 0 else 0
    avg_mo_amt   = int(total_all / total_months) if total_months > 0 else 0
    cur_yr = str(date.today().year)
    ann_total = int(result_df[result_df["기준월"].str.startswith(cur_yr)]["지체보상금"].sum())

    st.markdown(f"""
    <div class="kpi-row">
      <div class="kpi-card orange">
        <div class="kpi-lbl">전체 기간 지체보상금</div>
        <div class="kpi-val orange">{fmt_won(total_all, short=True)}</div>
        <div class="kpi-sub">{months_sorted[0]} ~ {months_sorted[-1]}</div>
      </div>
      <div class="kpi-card blue">
        <div class="kpi-lbl">{cur_yr}년 누계</div>
        <div class="kpi-val blue">{fmt_won(ann_total, short=True)}</div>
        <div class="kpi-sub">{cur_yr}년 발생 합계</div>
      </div>
      <div class="kpi-card red">
        <div class="kpi-lbl">청구 거래처 수 (전 기간)</div>
        <div class="kpi-val red">{total_vend}<span class="kpi-unit">개</span></div>
        <div class="kpi-sub">지체보상금 발생 고유 거래처</div>
      </div>
      <div class="kpi-card teal">
        <div class="kpi-lbl">월 평균 지체보상금</div>
        <div class="kpi-val teal">{fmt_won(avg_mo_amt, short=True)}</div>
        <div class="kpi-sub">발생 {total_months}개월 평균</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── 인사이트: 누적 TOP + 사업부별 ──────────────────────────
    ins1, ins2 = st.columns(2, gap="medium")

    with ins1:
        st.markdown('<div class="sec-hd"><span class="sec-bar"></span>전 기간 누적 합계 리스트</div>', unsafe_allow_html=True)
        _top_n = 8
        _top5 = (result_df.groupby(["판매처코드","판매처명"])["지체보상금"]
                 .sum().reset_index()
                 .sort_values("지체보상금", ascending=False).head(_top_n).reset_index(drop=True))
        st.markdown(
            f'<div class="chg-title" style="margin-bottom:.4rem">전 기간 누적 합계 '
            f'<span style="font-size:.71rem;font-weight:400;color:#9E9E9E;margin-left:.35rem">상위 {_top_n}개사 · 클릭하면 월별 내역 표시</span></div>',
            unsafe_allow_html=True)
        st.markdown("""<style>
.top-btn button {
    background:#fff; border:1px solid #F0F0F0; border-radius:8px;
    padding:.35rem .7rem; text-align:left; width:100%;
    font-size:.84rem; color:#212121; margin-bottom:3px;
    transition:background .15s;
}
.top-btn button:hover { background:#FFF8F1 !important; border-color:#F57C00 !important; }
.top-btn-sel button { background:#FFF3E0 !important; border-color:#F57C00 !important; font-weight:700 !important; }
</style>""", unsafe_allow_html=True)
        if "top_sel_code" not in st.session_state:
            st.session_state["top_sel_code"] = None
        for _i, _row in _top5.iterrows():
            _v = int(_row["지체보상금"])
            _code = _row["판매처코드"]
            _is_sel = st.session_state["top_sel_code"] == _code
            _div_cls = "top-btn-sel" if _is_sel else "top-btn"
            with st.container():
                st.markdown(f'<div class="{_div_cls}">', unsafe_allow_html=True)
                if st.button(
                    f"{_i+1}.  {_row['판매처명']}   {fmt_won(_v, short=True)}",
                    key=f"top_btn_{_code}",
                    use_container_width=True,
                ):
                    st.session_state["top_sel_code"] = _code if not _is_sel else None
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

    with ins2:
        _bu_yr = months_sorted[-1][:4] if months_sorted else cur_yr
        _bu_last_mo = mo_label(months_sorted[-1]) if months_sorted else ""
        st.markdown('<div class="sec-hd"><span class="sec-bar"></span>사업부별 지체보상금 현황</div>', unsafe_allow_html=True)
        if "사업부" in result_df.columns:
            _bu_all = (result_df[result_df["사업부"].str.strip().ne("") & result_df["사업부"].notna()]
                       .groupby("사업부")["지체보상금"].sum().reset_index())
            _bu_all = _bu_all.sort_values("지체보상금", ascending=True)  # 하단부터 큰 값
            if not _bu_all.empty:
                _bu_palette = ["#F57C00","#EF5350","#42A5F5","#66BB6A","#AB47BC",
                               "#26C6DA","#FFA726","#EC407A","#8D6E63","#78909C",
                               "#5C6BC0","#26A69A"]
                _n = len(_bu_all)
                _bu_colors = [_bu_palette[i % len(_bu_palette)] for i in range(_n)]
                _x_max = _bu_all["지체보상금"].max()

                _fig_bu = go.Figure(go.Bar(
                    y=_bu_all["사업부"],
                    x=_bu_all["지체보상금"] / 1e8,
                    orientation="h",
                    marker=dict(color=_bu_colors, line_width=0, cornerradius=6),
                    text=[f'{v/1e8:.2f}억' if v >= 1e6 else (f'{v/1e4:.0f}만' if v > 0 else "")
                          for v in _bu_all["지체보상금"]],
                    textposition="outside",
                    cliponaxis=False,
                    textfont=dict(size=10, color="#546e7a", family=FONT),
                ))
                _fig_bu.update_layout(
                    title=dict(
                        text=f'<b>사업부별 지체보상금</b>  <span style="font-size:12px;color:#9E9E9E">{_bu_yr}년 누계</span>',
                        x=0, xanchor="left",
                        font=dict(size=14, family=FONT, color="#212121"),
                        pad=dict(l=4),
                    ),
                    margin=dict(l=4, r=80, t=44, b=8),
                    height=max(280, _n * 42 + 70),
                    paper_bgcolor="#FFFFFF", plot_bgcolor="#F9FAFB",
                    xaxis=dict(
                        range=[0, _x_max / 1e8 * 1.40],
                        tickformat=".1f", ticksuffix="억",
                        showgrid=True, gridcolor="#EEEEEE",
                        zeroline=True, zerolinecolor="#E0E0E0", zerolinewidth=1,
                        tickfont=dict(size=9, color="#BDBDBD"),
                        showline=False,
                    ),
                    yaxis=dict(
                        tickfont=dict(size=10.5, family=FONT, color="#424242"),
                        showgrid=False, automargin=True,
                    ),
                    font=dict(family=FONT),
                )
                # 단위: 오른쪽 위
                _fig_bu.add_annotation(
                    x=1, y=1.03, xref="paper", yref="paper",
                    text="단위: 억원",
                    showarrow=False, xanchor="right",
                    font=dict(size=10, color="#9E9E9E", family=FONT),
                    yanchor="bottom",
                )
                st.plotly_chart(_fig_bu, use_container_width=True, config={"displayModeBar": False})
            else:
                st.caption("사업부 정보 없음")
        else:
            st.caption("사업부 정보 없음")

    # ── 누적 TOP 클릭 → 월별 내역 패널 ──────────────────────────
    _tsel = st.session_state.get("top_sel_code")
    if _tsel:
        _tsel_name = result_df[result_df["판매처코드"] == _tsel]["판매처명"].iloc[0] \
                     if not result_df[result_df["판매처코드"] == _tsel].empty else _tsel
        _tsel_df = (result_df[result_df["판매처코드"] == _tsel]
                    .sort_values("기준월")[["기준월","지연일수(누적)","연체기준잔고","요율","지체보상금"]])
        _total_amt = int(_tsel_df["지체보상금"].sum())
        st.markdown(
            f'<div style="background:#FFF8F1;border:1px solid #F57C00;border-radius:12px;'
            f'padding:.9rem 1.2rem .5rem;margin:.6rem 0 .8rem">'
            f'<span style="font-size:.9rem;font-weight:700;color:#E65100">{_tsel_name}</span>'
            f'<span style="font-size:.8rem;color:#64748b;margin-left:.6rem">월별 지체보상금 내역</span>'
            f'<span style="font-size:.85rem;font-weight:700;color:#E65100;float:right">합계 {_total_amt:,}원</span>'
            f'</div>', unsafe_allow_html=True)
        _mo_cols = st.columns(min(len(_tsel_df), 6))
        for _ci, (_, _mr) in enumerate(_tsel_df.iterrows()):
            with _mo_cols[_ci % len(_mo_cols)]:
                _amt = int(_mr["지체보상금"])
                _days = int(_mr["지연일수(누적)"])
                _잔 = int(_mr.get("연체기준잔고", 0))
                _color = "#E65100" if _amt > 0 else "#94a3b8"
                _bg = "#FFF8F1" if _amt > 0 else "#f8fafc"
                _border = "#F57C00" if _amt > 0 else "#e2e8f0"
                st.markdown(
                    f'<div style="background:{_bg};border:1px solid {_border};border-radius:9px;'
                    f'padding:.65rem .8rem;text-align:center;margin-bottom:6px">'
                    f'<div style="font-size:.75rem;color:#64748b;font-weight:600">{_mr["기준월"]}</div>'
                    f'<div style="font-size:1rem;font-weight:800;color:{_color};margin:.25rem 0">{_amt:,}원</div>'
                    f'<div style="font-size:.7rem;color:#94a3b8">지연 <b style="color:#475569">{_days}일</b>'
                    f' · {_잔/1e4:.0f}만</div>'
                    f'</div>', unsafe_allow_html=True)

    col_tbl, col_chart = st.columns([1, 1], gap="large")

    with col_tbl:
        st.markdown('<div class="sec-hd"><span class="sec-bar"></span>월별 집계표</div>', unsafe_allow_html=True)

        # 월별 집계 테이블 구성
        agg_rows = []
        for mo in months_sorted:
            mdf = result_df[result_df["기준월"] == mo]
            agg_rows.append({
                "기준월": mo,
                "청구 거래처": int((mdf["지체보상금"] > 0).sum()),
                "청구 건수": int((mdf["지체보상금"] > 0).sum()),
                "지체보상금 (원)": int(mdf["지체보상금"].sum()),
            })
        agg_df = pd.DataFrame(agg_rows)
        # 합계 행 추가
        sum_row = pd.DataFrame([{
            "기준월": "▶ 합계",
            "청구 거래처": "-",
            "청구 건수": int(agg_df["청구 건수"].sum()),
            "지체보상금 (원)": int(agg_df["지체보상금 (원)"].sum()),
        }])
        agg_display = pd.concat([sum_row, agg_df], ignore_index=True)

        def _agg_style(row):
            if row["기준월"] == "▶ 합계":
                return ["background:#FFF3E0;font-weight:bold;color:#E65100"] * len(row)
            if row["지체보상금 (원)"] and row["지체보상금 (원)"] != "-" and int(row["지체보상금 (원)"]) > 0:
                return ["background:#fff7ed"] * len(row)
            return [""] * len(row)

        st.dataframe(
            agg_display.style
                .apply(_agg_style, axis=1)
                .format({"지체보상금 (원)": lambda v: f"{int(v):,}" if str(v).lstrip("-").isdigit() else v}),
            use_container_width=True,
            height=min(55 * len(agg_display) + 55, 620),
            hide_index=True,
            column_config={
                "기준월":         st.column_config.TextColumn("기준월",       width=110),
                "청구 거래처":    st.column_config.TextColumn("청구 거래처",  width=100),
                "청구 건수":      st.column_config.TextColumn("청구 건수",    width=90),
                "지체보상금 (원)": st.column_config.TextColumn("지체보상금",   width=160),
            },
        )

    with col_chart:
        _WINDOW = 18  # 최근 N개월만 표시 (데이터 늘어도 자동 슬라이딩)
        _mo_window = months_sorted[-_WINDOW:]
        _note = f"최근 {len(_mo_window)}개월" if len(months_sorted) > _WINDOW else f"전체 {len(_mo_window)}개월"
        st.markdown(f'<div class="sec-hd"><span class="sec-bar"></span>월별 지체보상금 추이 <span class="sec-sub">{_note}</span></div>', unsafe_allow_html=True)

        # 가로 막대: 최신 월이 위에 오도록 역순 정렬
        _mo_rev  = list(reversed(_mo_window))
        y_labels = [mo_label(m) for m in _mo_rev]
        x_vals   = [int(mo_sum.get(m, 0)) / 1e8 for m in _mo_rev]

        _bar_palette = ["#F57C00","#AB47BC","#42A5F5","#66BB6A","#EF5350","#26C6DA",
                        "#FFA726","#5C6BC0","#26A69A","#EC407A","#8D6E63","#78909C"]
        _n_w = len(_mo_rev)
        bar_colors_w = [_bar_palette[i % len(_bar_palette)] if v > 0 else "#E8E8E8"
                        for i, v in enumerate(x_vals)]
        _x_max = max(x_vals) if any(v > 0 for v in x_vals) else 1

        fig = go.Figure(go.Bar(
            y=y_labels, x=x_vals,
            orientation="h",
            marker=dict(color=bar_colors_w, line_width=0, cornerradius=6),
            text=[f"{v:.2f}억" if v >= 0.01 else "" for v in x_vals],
            textposition="outside", cliponaxis=False,
            textfont=dict(size=10, color="#546e7a", family=FONT),
            hovertemplate="<b>%{y}</b><br>%{x:.2f}억원<extra></extra>",
        ))
        fig.update_layout(
            title=dict(
                text=f'<b>월별 지체보상금 추이</b>  <span style="font-size:12px;color:#9E9E9E">최근 {_n_w}개월</span>',
                x=0, xanchor="left",
                font=dict(size=14, family=FONT, color="#212121"),
                pad=dict(l=4),
            ),
            plot_bgcolor="#F9FAFB", paper_bgcolor="#FFFFFF",
            margin=dict(l=4, r=80, t=44, b=8),
            height=max(300, _n_w * 34 + 80),
            xaxis=dict(
                range=[0, _x_max * 1.38],
                tickformat=".1f", ticksuffix="억",
                showgrid=True, gridcolor="#EEEEEE",
                zeroline=True, zerolinecolor="#E0E0E0", zerolinewidth=1,
                tickfont=dict(size=9, color="#BDBDBD"),
                showline=False,
            ),
            yaxis=dict(
                tickfont=dict(size=10, family=FONT, color="#424242"),
                showgrid=False, automargin=True,
            ),
            font=dict(family=FONT),
        )
        fig.add_annotation(
            x=1, y=1.03, xref="paper", yref="paper",
            text="단위: 억원",
            showarrow=False, xanchor="right",
            font=dict(size=10, color="#9E9E9E", family=FONT),
            yanchor="bottom",
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    # ── 전체 기간 거래처별 누계표 ──────────────────────────────
    st.markdown('<div class="sec-hd"><span class="sec-bar"></span>거래처별 누계 (전체 기간)</div>', unsafe_allow_html=True)

    pv_all = result_df.pivot_table(
        index=["판매처코드","판매처명"], columns="기준월",
        values="지체보상금", aggfunc="sum", fill_value=0,
    ).reset_index()
    pv_all.columns = [mo_label(c) if c not in ["판매처코드","판매처명"] else c for c in pv_all.columns]
    pv_all["합계"] = pv_all.iloc[:, 2:].sum(axis=1)
    pv_all = pv_all[pv_all["합계"] > 0].sort_values("합계", ascending=False).reset_index(drop=True)
    pv_all.index = pv_all.index + 1

    mo_cols_pv = [c for c in pv_all.columns if c not in ["판매처코드","판매처명"]]
    def _heat(s):
        mx = s.max() if s.max() > 0 else 1
        return [f"background-color: rgba(59,130,246,{v/mx*0.5:.2f})" for v in s]

    pv_event = st.dataframe(
        pv_all.style.format({c: "{:,.0f}" for c in mo_cols_pv}).apply(_heat, subset=["합계"]),
        use_container_width=True, height=420,
        on_select="rerun", selection_mode="multi-row", key="pv_all_df",
    )
    # 선택된 행 → 공문 다운로드
    _pv_pos = pv_event.selection.rows if pv_event.selection.rows else []
    if _pv_pos:
        _pv_codes = [pv_all.iloc[i]["판매처코드"] for i in _pv_pos if i < len(pv_all)]
        _pv_df = result_df[result_df["판매처코드"].isin(_pv_codes) & (result_df["지체보상금"] > 0)]
        n_pv = _pv_df["판매처코드"].nunique()
        ca, cb, _ = st.columns([2, 2, 4])
        with ca:
            st.download_button(
                f"📁 선택 전체기간 공문 ({n_pv}개사)",
                data=make_공문_zip_cumulative(_pv_df) if n_pv else b"",
                file_name=f"지체보상금_공문_전체기간_{date.today():%Y%m%d}.zip",
                mime="application/zip", disabled=(n_pv == 0),
                use_container_width=True, key="btn_pv_docx",
            )
        with cb:
            st.download_button(
                "⬇ 전체 기간 엑셀",
                data=make_full_excel(result_df),
                file_name=f"지체보상금_전체_{date.today():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key="btn_agg_full_excel",
            )
    else:
        st.caption("행을 클릭하면 선택 → 공문 다운로드 버튼이 표시됩니다")


# ══════════════════════════════════════════════════════════════
# 탭 2: 월별 상세
# ══════════════════════════════════════════════════════════════
with tab_detail:

    # ── 연도 / 월 선택기 ─────────────────────────────────────
    years_avail = sorted(set(m.split("-")[0] for m in months_sorted))

    _today_tmp = date.today()
    _unclosed_tmp = {f"{_today_tmp.year}-{mo:02d}" for mo in range(_today_tmp.month, 13)}
    _avail_sorted = [m for m in months_sorted if m not in _unclosed_tmp] or months_sorted
    if "sel_mo" not in st.session_state or st.session_state.sel_mo not in _avail_sorted:
        st.session_state.sel_mo = _avail_sorted[-1]
    if "sel_yr" not in st.session_state or st.session_state.sel_yr not in years_avail:
        st.session_state.sel_yr = st.session_state.sel_mo.split("-")[0]

    # 연도 버튼
    st.markdown('<div class="yr-label">연도 선택</div>', unsafe_allow_html=True)
    yr_cols = st.columns(len(years_avail) + 6)
    for i, yr in enumerate(years_avail):
        with yr_cols[i]:
            if st.button(f"{yr}년", key=f"yr_{yr}",
                         type="primary" if yr == st.session_state.sel_yr else "secondary",
                         use_container_width=True):
                st.session_state.sel_yr = yr
                yr_months = [m for m in months_sorted if m.startswith(yr)]
                if st.session_state.sel_mo not in yr_months:
                    st.session_state.sel_mo = yr_months[-1]
                st.rerun()

    # 미마감 월: 당월 포함 이후는 클릭 불가
    _today = date.today()
    _unclosed = {f"{_today.year}-{mo:02d}" for mo in range(_today.month, 13)}

    # 월 버튼 (12개 그리드)
    st.markdown('<div style="margin-top:.3rem"><div class="yr-label">월 선택</div></div>', unsafe_allow_html=True)
    yr_months = [m for m in months_sorted if m.startswith(st.session_state.sel_yr)]
    yr_months_set = set(yr_months)
    mo_cols_row1 = st.columns(6)
    mo_cols_row2 = st.columns(6)
    for i in range(6):
        mo = i + 1
        m = f"{st.session_state.sel_yr}-{mo:02d}"
        with mo_cols_row1[i]:
            if m in yr_months_set and m not in _unclosed:
                if st.button(f"{mo}월", key=f"mo_{m}",
                             type="primary" if m == st.session_state.sel_mo else "secondary",
                             use_container_width=True):
                    st.session_state.sel_mo = m; st.rerun()
            else:
                lbl = f"{mo}월 🔒" if m in _unclosed else f"{mo}월"
                st.button(lbl, key=f"mo_dis_{m}", disabled=True, use_container_width=True)
    for i in range(6):
        mo = i + 7
        m = f"{st.session_state.sel_yr}-{mo:02d}"
        with mo_cols_row2[i]:
            if m in yr_months_set and m not in _unclosed:
                if st.button(f"{mo}월", key=f"mo_{m}",
                             type="primary" if m == st.session_state.sel_mo else "secondary",
                             use_container_width=True):
                    st.session_state.sel_mo = m; st.rerun()
            else:
                lbl = f"{mo}월 🔒" if m in _unclosed else f"{mo}월"
                st.button(lbl, key=f"mo_dis_{m}", disabled=True, use_container_width=True)

    st.markdown('<div style="margin-top:.6rem"></div>', unsafe_allow_html=True)

    sel = st.session_state.sel_mo
    sel_idx = months_sorted.index(sel)
    prev = months_sorted[sel_idx-1] if sel_idx > 0 else None
    sel_df  = result_df[result_df["기준월"] == sel]
    prev_df = result_df[result_df["기준월"] == prev] if prev else pd.DataFrame()

    sel_total  = int(sel_df["지체보상금"].sum())
    sel_n      = int((sel_df["지체보상금"] > 0).sum())
    sel_cu     = sel_df[sel_df["지체보상금"] > 0]["판매처코드"].nunique()
    prev_total = int(prev_df["지체보상금"].sum()) if not prev_df.empty else 0
    prev_n     = int((prev_df["지체보상금"] > 0).sum()) if not prev_df.empty else 0
    prev_cu    = prev_df[prev_df["지체보상금"] > 0]["판매처코드"].nunique() if not prev_df.empty else 0
    _sel_yr    = sel.split("-")[0]
    ann_sel    = int(result_df[result_df["기준월"].str.startswith(_sel_yr)]["지체보상금"].sum())

    def _delta_html(cur, prv, unit="원"):
        if prv == 0: return '<span style="color:#94a3b8">— 전월 없음</span>'
        d = cur - prv
        if d == 0: return '<span style="color:#BDBDBD">전월 동일</span>'
        arr = "▲" if d > 0 else "▼"
        cls = "kpi-up" if d > 0 else "kpi-dn"
        val = fmt_won(abs(d)) if unit == "원" else f"{abs(d)}{unit}"
        return f'전월 <span class="{cls}">{arr} {val}</span>'

    d_amt = sel_total - prev_total
    if d_amt != 0 and prev_total > 0:
        _arr = "▲" if d_amt > 0 else "▼"
        _cls = "kpi-up" if d_amt > 0 else "kpi-dn"
        d_html = f'전월 <span class="{_cls}">{_arr} {fmt_won(abs(d_amt))}</span>'
    else:
        d_html = '<span style="color:#BDBDBD">전월 데이터 없음</span>'

    # ── 월 KPI ──────────────────────────────────────────────
    st.markdown(f"""
    <div class="kpi-row">
      <div class="kpi-card orange">
        <div class="kpi-lbl">{mo_label(sel)} 지체보상금</div>
        <div class="kpi-val orange">{fmt_won(sel_total, short=True)}</div>
        <div class="kpi-sub">{d_html}</div>
      </div>
      <div class="kpi-card red">
        <div class="kpi-lbl">청구 건수</div>
        <div class="kpi-val red">{sel_n}<span class="kpi-unit">건</span></div>
        <div class="kpi-sub">{_delta_html(sel_n, prev_n, "건")}</div>
      </div>
      <div class="kpi-card blue">
        <div class="kpi-lbl">청구 거래처</div>
        <div class="kpi-val blue">{sel_cu}<span class="kpi-unit">개</span></div>
        <div class="kpi-sub">{_delta_html(sel_cu, prev_cu, "개")}</div>
      </div>
      <div class="kpi-card purple">
        <div class="kpi-lbl">{_sel_yr}년 누계</div>
        <div class="kpi-val purple">{fmt_won(ann_sel, short=True)}</div>
        <div class="kpi-sub">{_sel_yr}년 1월 ~ {mo_label(sel)}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── 신규/추가/수금 카드 ───────────────────────────────────
    _prev_map = {r["판매처코드"]: r for _, r in prev_df.iterrows()}
    _curr_map = {r["판매처코드"]: r for _, r in sel_df.iterrows()}
    _신규, _추가, _수금 = [], [], []
    for _code in set(_prev_map) | set(_curr_map):
        _c = _curr_map.get(_code); _p = _prev_map.get(_code)
        _cj = int(_c["지체보상금"]) if _c is not None else 0
        _pj = int(_p["지체보상금"]) if _p is not None else 0
        _cp = int(_c["현 미수금"])   if _c is not None else 0
        _pp = int(_p["현 미수금"])   if _p is not None else 0
        _src = _c if _c is not None else _p
        _nm = _src["판매처명"]
        _org = str(_src.get("사무소", "") or _src.get("사업부", "") or "").strip()
        _org = "" if _org in ("nan", "None") else _org
        if _cj > 0 and _pj == 0:   _신규.append({"명칭": _nm, "금액": _cj, "org": _org})
        elif _cj > 0 and _pj > 0:  _추가.append({"명칭": _nm, "금액": _cj, "org": _org})
        if _pj > 0 and _cj == 0:   _수금.append({"명칭": _nm, "금액": max(_pp - _cp, 0) or _pp, "org": _org})
    for _lst in (_신규, _추가, _수금): _lst.sort(key=lambda x: -x["금액"])

    def _chg_html(lst, cls, title, amt_color):
        rows = ""
        for i, x in enumerate(lst[:5]):
            org_badge = (f'<span style="font-size:.65rem;color:#9E9E9E;margin-left:.3rem;'
                         f'background:#F5F5F5;border-radius:3px;padding:1px 5px">{x["org"]}</span>'
                         if x.get("org") else "")
            rows += (f'<div class="chg-row">'
                     f'<span class="chg-num">{i+1}</span>'
                     f'<span class="chg-name">{x["명칭"]}{org_badge}</span>'
                     f'<span class="chg-amt" style="color:{amt_color}">{fmt_won(x["금액"], short=True)}</span>'
                     f'</div>')
        if len(lst) > 5:
            etc_amt = sum(x["금액"] for x in lst[5:])
            rows += f'<div style="font-size:.71rem;color:#9E9E9E;padding:.35rem .9rem">그 외 {len(lst)-5}처 ({fmt_won(etc_amt,short=True)})</div>'
        if not lst:
            rows = '<div style="color:#BDBDBD;font-size:.81rem;padding:1.2rem .9rem;text-align:center">해당 없음</div>'
        total = sum(x["금액"] for x in lst)
        badge = f'<span style="font-size:.71rem;font-weight:400;color:#9E9E9E;margin-left:.35rem">{len(lst)}건 &nbsp;·&nbsp; 합계 {fmt_won(total,short=True)}</span>'
        return (f'<div class="chg-card {cls}">'
                f'<div class="chg-title">{title}{badge}</div>'
                f'{rows}</div>')

    cc1, cc2, cc3 = st.columns(3, gap="medium")
    with cc1: st.markdown(_chg_html(_신규, "red",   f"{mo_label(sel)} 신규 연체", "#ef4444"), unsafe_allow_html=True)
    with cc2: st.markdown(_chg_html(_추가, "amber", f"{mo_label(sel)} 추가 연체", "#d97706"), unsafe_allow_html=True)
    with cc3: st.markdown(_chg_html(_수금, "green", f"{mo_label(sel)} 수금 완료", "#059669"), unsafe_allow_html=True)

    st.markdown('<div style="margin:.8rem 0"></div>', unsafe_allow_html=True)

    # ── 청구 거래처 랭킹 + 차트 ──────────────────────────────
    left_col, right_col = st.columns([5, 5], gap="large")

    with left_col:
        st.markdown('<div class="sec-hd"><span class="sec-bar"></span>이달 청구 거래처 현황</div>', unsafe_allow_html=True)
        charge = sel_df[sel_df["지체보상금"] > 0].sort_values("지체보상금", ascending=False).reset_index(drop=True)
        all_max = int(charge["지체보상금"].max()) if not charge.empty else 1
        rank_colors = ["#ef4444","#f97316","#f59e0b","#84cc16","#22c55e",
                       "#14b8a6","#3b82f6","#6366f1","#8b5cf6","#ec4899"]
        if charge.empty:
            st.markdown('<div class="rank-card"><div class="rank-empty">해당 월 청구 발생 없음</div></div>', unsafe_allow_html=True)
        else:
            items_html = ""
            for i, row in charge.iterrows():
                v = int(row["지체보상금"])
                pct = min(v / all_max * 100, 100)
                color = rank_colors[i] if i < len(rank_colors) else "#9E9E9E"
                org = str(row.get("사무소","")).strip()
                if not org or org in ("nan",""):
                    org = str(row.get("사업부","")).strip()
                org_txt = f'<span style="font-size:.68rem;color:#BDBDBD;margin-left:.3rem">{org}</span>' if org and org not in ("nan","") else ""
                items_html += f"""<div class="rank-item">
                  <span class="rank-badge" style="background:{color}">{i+1}</span>
                  <div style="min-width:0">
                    <div style="font-size:.83rem;font-weight:600;color:#212121;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{row["판매처명"]}{org_txt}</div>
                    <div class="rank-bar-bg"><div class="rank-bar-fg" style="background:{color};width:{pct:.1f}%"></div></div>
                  </div>
                  <div style="text-align:right;font-size:.88rem;font-weight:700;color:{color};flex-shrink:0">{fmt_won(v, short=True)}</div>
                </div>"""
            st.markdown(f'<div class="rank-card">{items_html}</div>', unsafe_allow_html=True)

    with right_col:
        st.markdown(f'<div class="sec-hd"><span class="sec-bar"></span>{_sel_yr}년 월별 추이</div>', unsafe_allow_html=True)
        all_12 = [f"{_sel_yr}-{mo:02d}" for mo in range(1, 13)]
        x_12 = [f"{int(m.split('-')[1])}월" for m in all_12]
        y_12 = [int(mo_sum.get(m, 0)) / 10000 for m in all_12]
        bar_col_12 = ["#F57C00" if m == sel else ("#FFCC80" if mo_sum.get(m, 0) > 0 else "#E0E0E0") for m in all_12]
        cum_12 = []
        run = 0
        for v in y_12:
            run += v
            cum_12.append(round(run, 1))
        last_data_idx = max((i for i, v in enumerate(y_12) if v > 0), default=-1)
        cum_line_12 = [cum_12[i] if i <= last_data_idx else None for i in range(12)]

        fig2 = go.Figure()
        fig2.add_trace(go.Scatter(
            x=x_12, y=cum_line_12, name="_fill",
            mode="lines", line=dict(color="rgba(0,0,0,0)", width=0),
            fill="tozeroy", fillcolor="rgba(245,158,11,0.08)",
            hoverinfo="skip", yaxis="y2", showlegend=False,
        ))
        fig2.add_trace(go.Bar(
            x=x_12, y=y_12, name="월별",
            marker_color=bar_col_12, marker_line_width=0,
            text=[f"{v:,.0f}만" if v > 0 else "" for v in y_12],
            textposition="outside", cliponaxis=False,
            textfont=dict(size=10, color="#64748b"),
            hovertemplate="<b>%{x}</b><br>%{y:,.0f}만원<extra></extra>",
            yaxis="y1",
        ))
        fig2.add_trace(go.Scatter(
            x=x_12, y=cum_line_12, name="누적",
            mode="lines+markers",
            line=dict(color="#F57C00", width=2.5),
            marker=dict(size=6, color="#F57C00", line=dict(color="#fff", width=2)),
            hovertemplate="<b>%{x}</b><br>누적 %{y:,.0f}만원<extra></extra>",
            yaxis="y2",
        ))
        mo_max2 = max(y_12) if max(y_12) > 0 else 1
        cu_max2 = max(v for v in cum_12 if v > 0) if any(v > 0 for v in cum_12) else 1
        fig2.update_layout(
            plot_bgcolor="#F9FAFB", paper_bgcolor="#FFFFFF",
            margin=dict(t=44, b=8, l=4, r=60), height=360,
            bargap=0.38,
            legend=dict(orientation="h", x=1, xanchor="right", y=1.06,
                        font=dict(size=10, family=FONT), bgcolor="rgba(0,0,0,0)"),
            yaxis=dict(gridcolor="#EEEEEE", zeroline=False, showticklabels=False,
                       range=[0, mo_max2 * 1.65], title=None, showline=False),
            yaxis2=dict(overlaying="y", side="right", range=[0, cu_max2 * 1.3],
                        tickformat=",.0f", ticksuffix="만",
                        tickfont=dict(size=9, color="#E65100"),
                        showgrid=False, title=None, showline=False),
            xaxis=dict(showgrid=False, tickfont=dict(size=10, family=FONT), showline=False),
            font=dict(family=FONT),
        )
        fig2.add_annotation(
            x=1, y=1.03, xref="paper", yref="paper",
            text="단위: 만원",
            showarrow=False, xanchor="right",
            font=dict(size=10, color="#9E9E9E", family=FONT),
            yanchor="bottom",
        )
        st.plotly_chart(fig2, use_container_width=True, config={"displayModeBar": False})

    # ── 다운로드 버튼 ────────────────────────────────────────
    charge_only = sel_df[sel_df["지체보상금"] > 0].sort_values("지체보상금", ascending=False)
    n_charge = len(charge_only)
    all_charge_df = result_df[result_df["지체보상금"] > 0]
    n_vendors = all_charge_df["판매처코드"].nunique()

    st.markdown('<div class="sec-hd"><span class="sec-bar"></span>파일 다운로드</div>', unsafe_allow_html=True)
    st.markdown('<div class="dl-card">'
                '<div class="dl-card-title">엑셀 데이터</div>', unsafe_allow_html=True)
    dc1, dc2 = st.columns(2)
    with dc1:
        st.download_button(
            f"📊 {mo_label(sel)} 발생현황",
            data=make_excel(charge_only, f"{mo_label(sel)}"),
            file_name=f"지체보상금_{sel.replace('-','')}_발생현황.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, key="btn_dl_mo_excel",
            help=f"{mo_label(sel)} 청구 발생 거래처 내역",
        )
    with dc2:
        st.download_button(
            "📋 전체 기간 통합본",
            data=make_full_excel(result_df),
            file_name=f"지체보상금_전체_{date.today():%Y%m%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, key="btn_dl_full_excel",
            help="전체 기간 내역 + 거래처별 피벗",
        )
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('<div class="dl-card" style="margin-top:.5rem">'
                '<div class="dl-card-title">공문 (Word .docx)</div>', unsafe_allow_html=True)
    dc3, dc4 = st.columns(2)
    with dc3:
        st.download_button(
            f"📄 {mo_label(sel)} 청구 공문 ({n_charge}건)" if n_charge else f"📄 {mo_label(sel)} 청구 없음",
            data=make_공문_zip(charge_only, mo_label(sel)) if n_charge else b"",
            file_name=f"지체보상금_공문_{sel.replace('-','')}.zip",
            mime="application/zip",
            disabled=(n_charge == 0),
            use_container_width=True, key="btn_dl_mo_docx",
            help=f"{mo_label(sel)} 발생 거래처 개별 공문 ZIP",
        )
    with dc4:
        st.download_button(
            f"📁 전체기간 누적 공문 ({n_vendors}개사)",
            data=make_공문_zip_cumulative(result_df) if n_vendors else b"",
            file_name=f"지체보상금_공문_전체기간_{date.today():%Y%m%d}.zip",
            mime="application/zip",
            disabled=(n_vendors == 0),
            use_container_width=True, key="btn_dl_all_docx",
            help="전체 기간 거래처별 누적 공문 ZIP",
        )
    st.markdown('</div>', unsafe_allow_html=True)

    # ── 상세 테이블 ──────────────────────────────────────────
    st.markdown(f'<div class="sec-hd"><span class="sec-bar"></span>{mo_label(sel)} 상세 내역</div>', unsafe_allow_html=True)

    fc1, fc2 = st.columns([5, 1])
    kw     = fc1.text_input("거래처명 검색", placeholder="🔍  거래처명으로 검색",
                            label_visibility="collapsed", key="kw1")
    only_c = fc2.checkbox("청구만", key="oc1")
    view   = sel_df.copy()
    if kw:     view = view[view["판매처명"].str.contains(kw, na=False)]
    if only_c: view = view[view["지체보상금"] > 0]
    view = view.sort_values("지체보상금", ascending=False).reset_index(drop=True)

    def _calc_str(r):
        if r["지체보상금"] <= 0: return "—"
        base = int(r.get("연체기준잔고", r["현 미수금"]))
        지연 = int(r.get("지연일수(누적)", 0))
        return f'{base:,} × {지연}일 × {r["요율"]} = {int(r["지체보상금"]):,}원'
    view["계산 근거"] = view.apply(_calc_str, axis=1)

    _all_cols = ["판매처코드","판매처명","사업자번호","계약유형","채널",
                 "사업부","사무소","담당자","현 미수금",
                 "유효기준회전일","현 회전일","지연일수(누적)",
                 "요율","지체보상금","계산 근거","비고"]
    disp = view[[c for c in _all_cols if c in view.columns]]
    disp = disp.rename(columns={"판매처코드": "거래처코드", "채널": "거래채널"})
    disp.index = disp.index + 1

    def _sty(r):
        if r["지체보상금"] > 0: return ["background:#fffbeb;font-weight:bold"] * len(r)
        return ["color:#cbd5e1"] * len(r)

    t1_event = st.dataframe(
        disp.style.apply(_sty, axis=1)
            .format({"현 미수금": "{:,.0f}", "지체보상금": "{:,.0f}"}),
        use_container_width=True, height=440,
        on_select="rerun", selection_mode="multi-row", key="tab1_df",
        column_config={
            "거래처코드":     st.column_config.TextColumn("거래처코드",   width=115),
            "판매처명":       st.column_config.TextColumn("거래처명",     width=210),
            "사업자번호":     st.column_config.TextColumn("사업자번호",   width=115),
            "계약유형":       st.column_config.TextColumn("계약유형",     width=110),
            "거래채널":       st.column_config.TextColumn("채널",         width=75),
            "사업부":         st.column_config.TextColumn("사업부",       width=145),
            "사무소":         st.column_config.TextColumn("사무소",       width=130),
            "담당자":         st.column_config.TextColumn("담당자",       width=80),
            "현 미수금":      st.column_config.NumberColumn("현 미수금",  width=120, format="%,.0f"),
            "유효기준회전일": st.column_config.NumberColumn("기준회전일", width=80),
            "현 회전일":      st.column_config.NumberColumn("현회전일",   width=75),
            "지연일수(누적)": st.column_config.NumberColumn("지연일수",   width=75),
            "요율":           st.column_config.TextColumn("요율",         width=65),
            "지체보상금":     st.column_config.NumberColumn("지체보상금", width=125, format="%,.0f"),
            "계산 근거":      st.column_config.TextColumn("계산 근거",    width=280),
            "비고":           st.column_config.TextColumn("비고",         width=75),
        },
    )

    _sel_pos = t1_event.selection.rows if t1_event.selection.rows else []
    if _sel_pos:
        _sel_codes = [view.iloc[i]["판매처코드"] for i in _sel_pos if i < len(view)]
        _sel_mo_df  = view[view["판매처코드"].isin(_sel_codes) & (view["지체보상금"] > 0)]
        _sel_all_df = result_df[result_df["판매처코드"].isin(_sel_codes) & (result_df["지체보상금"] > 0)]
        n_mo = len(_sel_mo_df); n_all = _sel_all_df["판매처코드"].nunique()
        st.markdown(f'<div style="font-size:.75rem;color:#64748b;margin:.4rem 0 .3rem">'
                    f'선택: <b style="color:#0f172a">{len(_sel_pos)}개사</b> · 이번달 청구 {n_mo}건 · 전체기간 {n_all}개사</div>',
                    unsafe_allow_html=True)
        _ba, _bb, _bc = st.columns([2, 2, 4])
        with _ba:
            st.download_button(f"📄 선택 이번달 공문 ({n_mo}건)",
                data=make_공문_zip(_sel_mo_df, mo_label(sel)) if n_mo else b"",
                file_name=f"지체보상금_선택_{sel.replace('-','')}.zip",
                mime="application/zip", disabled=(n_mo == 0),
                use_container_width=True, key="btn_sel_mo_docx")
        with _bb:
            st.download_button(f"📁 선택 전체기간 공문 ({n_all}개사)",
                data=make_공문_zip_cumulative(_sel_all_df) if n_all else b"",
                file_name=f"지체보상금_선택_전체기간_{date.today():%Y%m%d}.zip",
                mime="application/zip", disabled=(n_all == 0),
                use_container_width=True, key="btn_sel_cum_t1")

        # ── 직전 3개월 상세 패널 (단일 거래처 선택 시) ────────────
        if len(_sel_codes) == 1:
            _code = _sel_codes[0]
            _nm_rows = view[view["판매처코드"] == _code]
            _name = _nm_rows["판매처명"].iloc[0] if not _nm_rows.empty else _code
            _sel_yr, _sel_mo_n = int(sel.split("-")[0]), int(sel.split("-")[1])
            _prev_months = []
            for i in range(3, 0, -1):
                mo = _sel_mo_n - i
                yr = _sel_yr
                if mo <= 0:
                    mo += 12; yr -= 1
                _prev_months.append(f"{yr}-{mo:02d}")
            _hist_df = result_df[result_df["판매처코드"] == _code].copy()
            _hist_df = _hist_df[_hist_df["기준월"].isin(_prev_months)].sort_values("기준월")

            st.markdown("---")
            st.markdown(
                f'<div style="font-size:.85rem;font-weight:700;color:#1e293b;margin:.2rem 0 .5rem">'
                f'📊 {_name} · 직전 3개월 지체보상금 내역'
                f'<span style="font-size:.75rem;color:#64748b;font-weight:400;margin-left:.5rem">'
                f'청구할 월 선택 후 공문 다운로드</span></div>',
                unsafe_allow_html=True)

            _hcols = st.columns(3)
            _checked_months = []
            for ci, pm in enumerate(_prev_months):
                with _hcols[ci]:
                    _row = _hist_df[_hist_df["기준월"] == pm]
                    if _row.empty:
                        st.markdown(
                            f'<div style="background:#f8fafc;border-radius:10px;padding:.8rem 1rem;'
                            f'text-align:center;border:1px solid #e2e8f0">'
                            f'<div style="font-size:.8rem;color:#94a3b8;font-weight:600">{pm}</div>'
                            f'<div style="font-size:.75rem;color:#cbd5e1;margin-top:.4rem">데이터 없음</div></div>',
                            unsafe_allow_html=True)
                    else:
                        _r = _row.iloc[0]
                        _amt = int(_r["지체보상금"])
                        _days = int(_r["지연일수(누적)"])
                        _color = "#E65100" if _amt > 0 else "#64748b"
                        _bg = "#FFF8F1" if _amt > 0 else "#f8fafc"
                        _border = "#F57C00" if _amt > 0 else "#e2e8f0"
                        st.markdown(
                            f'<div style="background:{_bg};border-radius:10px;padding:.8rem 1rem;'
                            f'text-align:center;border:1px solid {_border}">'
                            f'<div style="font-size:.8rem;color:#64748b;font-weight:600">{pm}</div>'
                            f'<div style="font-size:1.1rem;font-weight:800;color:{_color};margin:.3rem 0">{_amt:,}원</div>'
                            f'<div style="font-size:.75rem;color:#94a3b8">지연 <b style="color:#475569">{_days}일</b></div>'
                            f'</div>',
                            unsafe_allow_html=True)
                        if _amt > 0:
                            _cb = st.checkbox(pm, key=f"chk_{_code}_{pm}", value=True)
                            if _cb:
                                _checked_months.append(pm)

            _sel_hist = _hist_df[_hist_df["기준월"].isin(_checked_months)]
            if not _sel_hist.empty:
                _sel_label = ", ".join(_sel_hist["기준월"].tolist())
                st.markdown('<div style="margin-top:.4rem"></div>', unsafe_allow_html=True)
                st.download_button(
                    f"📋 선택 월 공문 다운로드 ({_sel_label})",
                    data=make_공문_zip_cumulative(_sel_hist),
                    file_name=f"지체보상금_{_code}_선택_{date.today():%Y%m%d}.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key="btn_hist_sel")
            elif _checked_months == [] and any(not _hist_df[_hist_df["기준월"] == pm].empty
                                               and int(_hist_df[_hist_df["기준월"] == pm].iloc[0]["지체보상금"]) > 0
                                               for pm in _prev_months):
                st.caption("월을 선택하면 공문 다운로드 버튼이 표시됩니다")
    else:
        st.caption("행을 클릭하면 선택 → 공문 다운로드 버튼이 표시됩니다")

# ══════════════════════════════════════════════════════════════
# 탭 3: 계산 가이드
# ══════════════════════════════════════════════════════════════
with tab_guide:
    st.markdown("""
<style>
.gd-wrap { max-width:860px; margin:0 auto; font-family:'Noto Sans KR','Malgun Gothic',sans-serif; }
.gd-hero {
  background: linear-gradient(135deg,#E65100 0%,#F57C00 60%,#FFB300 100%);
  border-radius:14px; padding:2rem 2.4rem 1.8rem; margin-bottom:1.6rem;
  color:#fff;
}
.gd-hero-title { font-size:1.45rem; font-weight:800; margin:0 0 .4rem; letter-spacing:-.5px; }
.gd-hero-sub { font-size:.9rem; opacity:.88; margin:0; font-weight:400; }
.gd-formula {
  background:#FFF8F1; border-left:4px solid #F57C00; border-radius:0 10px 10px 0;
  padding:1.1rem 1.4rem; margin:1rem 0 1.4rem; font-size:1.05rem; font-weight:700;
  color:#BF360C; letter-spacing:-.2px;
}
.gd-section {
  background:#fff; border-radius:12px; padding:1.4rem 1.6rem; margin-bottom:1.1rem;
  box-shadow:0 1px 6px rgba(0,0,0,.07); border:1px solid #F0F0F0;
}
.gd-section-hd {
  display:flex; align-items:center; gap:.6rem; margin-bottom:1rem;
  font-size:1rem; font-weight:700; color:#212121;
}
.gd-badge {
  background:#F57C00; color:#fff; border-radius:20px; padding:.15rem .6rem;
  font-size:.75rem; font-weight:700; flex-shrink:0;
}
.gd-badge.blue { background:#3B82F6; }
.gd-badge.green { background:#22C55E; }
.gd-badge.purple { background:#8B5CF6; }
.gd-badge.gray { background:#78909C; }
.gd-items { display:flex; flex-direction:column; gap:.55rem; }
.gd-item { display:flex; gap:.8rem; align-items:flex-start; font-size:.88rem; color:#374151; }
.gd-dot { width:7px; height:7px; border-radius:50%; background:#F57C00; flex-shrink:0; margin-top:.45rem; }
.gd-dot.blue { background:#3B82F6; }
.gd-dot.green { background:#22C55E; }
.gd-tbl { width:100%; border-collapse:collapse; font-size:.84rem; margin-top:.6rem; }
.gd-tbl th {
  background:#37474F; color:#fff; padding:.55rem .85rem; text-align:left;
  font-weight:600; font-size:.8rem;
}
.gd-tbl td { padding:.5rem .85rem; border-bottom:1px solid #F0F0F0; color:#374151; }
.gd-tbl tr:last-child td { border-bottom:none; }
.gd-tbl tr:nth-child(even) td { background:#FAFAFA; }
.gd-tbl td:first-child { font-weight:600; color:#1E293B; white-space:nowrap; }
.gd-ex-tbl { width:100%; border-collapse:collapse; font-size:.82rem; margin-top:.8rem; }
.gd-ex-tbl th {
  background:#1A2A6C; color:#fff; padding:.5rem .7rem; text-align:center;
  font-weight:600; font-size:.79rem;
}
.gd-ex-tbl td { padding:.45rem .7rem; border-bottom:1px solid #EEF0F5; text-align:right; color:#374151; }
.gd-ex-tbl td:first-child { text-align:center; font-weight:600; color:#1E293B; }
.gd-ex-tbl tr:nth-child(3) td { background:#FFF3E0; font-weight:600; color:#E65100; }
.gd-ex-tbl tr:last-child td { border-bottom:none; }
.gd-highlight {
  background:#FFF3E0; border:1px solid #FFB74D; border-radius:8px;
  padding:.75rem 1rem; margin-top:.9rem; font-size:.84rem; color:#BF360C;
  display:flex; gap:.5rem; align-items:flex-start;
}
.gd-code {
  background:#F1F5F9; border-radius:8px; padding:.9rem 1.2rem;
  font-family:monospace; font-size:.85rem; color:#1E293B;
  margin:.6rem 0; border-left:3px solid #3B82F6; white-space:pre-wrap;
}
.gd-flow {
  display:flex; align-items:center; gap:.5rem; flex-wrap:wrap;
  margin:.8rem 0; font-size:.87rem;
}
.gd-flow-box {
  background:#EFF6FF; border:1px solid #BFDBFE; border-radius:8px;
  padding:.4rem .8rem; color:#1D4ED8; font-weight:600; white-space:nowrap;
}
.gd-flow-arrow { color:#94A3B8; font-size:1.1rem; }
.gd-flow-box.result {
  background:#FEF3C7; border-color:#FCD34D; color:#92400E; font-size:.95rem;
}
.gd-rank { display:flex; flex-direction:column; gap:.5rem; margin-top:.5rem; }
.gd-rank-item {
  display:flex; align-items:center; gap:.7rem; background:#FAFAFA;
  border-radius:8px; padding:.5rem .9rem; font-size:.84rem;
}
.gd-rank-num {
  width:22px; height:22px; border-radius:50%; display:flex; align-items:center;
  justify-content:center; font-size:.72rem; font-weight:700; flex-shrink:0;
  color:#fff;
}
.gd-tag {
  display:inline-block; background:#F1F5F9; border-radius:4px;
  padding:.1rem .4rem; font-family:monospace; font-size:.8rem; color:#475569;
}
</style>

<div class="gd-wrap">

<div class="gd-hero">
  <div class="gd-hero-title">📖 지체보상금 계산 가이드</div>
  <div class="gd-hero-sub">thync™ 상품공급 계약 기준 · 대웅제약 채권관리팀</div>
</div>

<!-- 1. 기본 공식 -->
<div class="gd-section">
  <div class="gd-section-hd"><span class="gd-badge">1</span> 기본 공식</div>
  <div class="gd-formula">지체보상금 = 연체기준잔고 × 지연일수 × 요율 (1/1,000)</div>
  <table class="gd-tbl">
    <tr><th>항목</th><th>정의</th></tr>
    <tr><td>연체기준잔고</td><td>현 미수금 − 이번달 납기 미도달 매출액 (당월 아직 납기 안 된 금액 제외)</td></tr>
    <tr><td>지연일수 (누적)</td><td>현 회전일 − 기준회전일 (음수면 0으로 처리)</td></tr>
    <tr><td>요율</td><td>계약서 기준, 기본 1/1,000</td></tr>
  </table>
</div>

<!-- 2. 월별 청구 방식 -->
<div class="gd-section">
  <div class="gd-section-hd"><span class="gd-badge blue">2</span> 월별 청구 방식 <span style="font-size:.8rem;color:#64748b;font-weight:400">(핵심)</span></div>
  <div class="gd-items">
    <div class="gd-item"><div class="gd-dot blue"></div><span>각 월을 <b>독립적으로</b> 계산 — 전월 청구 여부와 무관</span></div>
    <div class="gd-item"><div class="gd-dot blue"></div><span>해당 월의 연체기준잔고 × 지연일수 × 요율 = 이번달 청구액</span></div>
    <div class="gd-item"><div class="gd-dot blue"></div><span>지연일수가 0 이하이면 → 이번달 청구 0</span></div>
  </div>
  <div class="gd-flow" style="margin-top:1rem">
    <div class="gd-flow-box">연체기준잔고</div>
    <div class="gd-flow-arrow">×</div>
    <div class="gd-flow-box">지연일수</div>
    <div class="gd-flow-arrow">×</div>
    <div class="gd-flow-box">요율 (1/1000)</div>
    <div class="gd-flow-arrow">=</div>
    <div class="gd-flow-box result">지체보상금</div>
  </div>
  <div style="font-size:.78rem;color:#94A3B8;margin-top:.4rem">* 지연일수 = 현 회전일 − 기준회전일 (음수는 0으로 처리)</div>
</div>

<!-- 3. 예시 -->
<div class="gd-section">
  <div class="gd-section-hd"><span class="gd-badge green">3</span> 월별 예시 <span style="font-size:.8rem;color:#64748b;font-weight:400">(기준회전일 30일, 요율 1/1,000)</span></div>
  <table class="gd-ex-tbl">
    <tr>
      <th>월</th><th>현회전일</th><th>지연일수</th><th>연체기준잔고</th><th>지체보상금</th>
    </tr>
    <tr>
      <td>6월</td><td>59일</td><td>29일</td><td>5,000,000</td>
      <td style="color:#E65100;font-weight:700">145,000원</td>
    </tr>
    <tr>
      <td>7월</td><td>62일</td><td>32일</td><td>5,000,000</td>
      <td style="color:#E65100;font-weight:700">160,000원</td>
    </tr>
    <tr>
      <td>8월</td><td>62일</td><td>32일</td><td>8,000,000</td>
      <td style="color:#E65100;font-weight:700">256,000원</td>
    </tr>
    <tr>
      <td>9월</td><td>65일</td><td>35일</td><td>8,000,000</td>
      <td style="color:#E65100;font-weight:700">280,000원</td>
    </tr>
  </table>
  <div class="gd-highlight">
    <span>💡</span>
    <span>각 월을 <b>독립적으로</b> 계산합니다. 전월 청구 여부와 무관하게 해당 월의 연체기준잔고 × 지연일수 × 요율로 청구액을 산출합니다.</span>
  </div>
</div>

<!-- 4. 컬럼 설명 -->
<div class="gd-section">
  <div class="gd-section-hd"><span class="gd-badge purple">4</span> 테이블 컬럼 설명</div>
  <table class="gd-tbl">
    <tr><th>컬럼명</th><th>설명</th></tr>
    <tr><td>현 미수금</td><td>해당 월 말 기준 총 미수금 잔고</td></tr>
    <tr><td>연체기준잔고</td><td>현 미수금 − 당월 납기 미달 매출 (이번달잔고) 차감한 금액</td></tr>
    <tr><td>기준회전일</td><td>SAP 기준 계약 목표 회전일 (거래처별 상이)</td></tr>
    <tr><td>현 회전일</td><td>실제 채권 회전일 (월별 잔고회전일 파일 기준)</td></tr>
    <tr><td>지연일수</td><td>현 회전일 − 기준회전일 (음수=0)</td></tr>
    <tr><td>지체보상금</td><td>연체기준잔고 × 지연일수 × 요율</td></tr>
    <tr><td>계산 근거</td><td>연체기준잔고 × 지연일수 × 요율 = 지체보상금 (상세 수식)</td></tr>
    <tr><td>비고</td><td>청구 (지체보상금 > 0)</td></tr>
  </table>
</div>


</div>
""", unsafe_allow_html=True)
