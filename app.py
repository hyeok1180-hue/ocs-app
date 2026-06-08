import streamlit as st
import pandas as pd
import re
import os
import json
from io import BytesIO
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.graph_objects as go
import plotly.express as px

# ─────────────────────────────────────────────
# 설정
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="OCS 처방 분석기",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Streamlit Cloud는 실행 위치가 저장소 루트이므로 파일명만 사용
import pathlib
_ROOT = pathlib.Path(__file__).parent if "__file__" in dir() else pathlib.Path.cwd()
PRICE_FILE = str(_ROOT / "약제급여목록및급여상한금액표_2026_4_1__공개용_1부.xlsx")
DB_FILE    = str(_ROOT / "ocs_history.json")
DAEWUNG_KW = ["대웅", "대웅제약", "대웅바이오"]

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&family=IBM+Plex+Mono:wght@400;600&display=swap');
html, body, [class*="css"] { font-family: 'Noto Sans KR', sans-serif; }
.main { background: #f0f4f8; }
.kpi-card {
    background: white; border-radius: 12px; padding: 20px 24px;
    box-shadow: 0 1px 4px rgba(0,0,0,.08); border-left: 4px solid #1F4E79; margin-bottom: 8px;
}
.kpi-label { font-size: 11px; color: #888; font-weight: 500; letter-spacing:.05em; text-transform:uppercase; margin-bottom:4px; }
.kpi-value { font-size: 22px; font-weight: 700; color: #1F4E79; font-family:'IBM Plex Mono',monospace; }
.kpi-delta { font-size: 12px; margin-top: 4px; }
.kpi-card.orange { border-left-color: #C55A11; }
.kpi-card.orange .kpi-value { color: #C55A11; }
.kpi-card.green  { border-left-color: #375623; }
.kpi-card.green .kpi-value  { color: #375623; }
.kpi-card.red    { border-left-color: #843c3c; }
.kpi-card.red .kpi-value    { color: #843c3c; }
.section-title {
    font-size: 14px; font-weight: 700; color: #1F4E79;
    border-bottom: 2px solid #1F4E79; padding-bottom: 6px; margin: 20px 0 12px 0;
}
.upload-hint {
    background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
    color: white; border-radius: 12px; padding: 20px 24px; margin-bottom: 16px;
}
.upload-hint h3 { margin: 0 0 6px 0; font-size: 15px; }
.upload-hint p  { margin: 0; font-size: 12px; opacity: .85; }
.stButton > button {
    background: #1F4E79 !important; color: white !important;
    border: none !important; border-radius: 8px !important;
    font-family: 'Noto Sans KR', sans-serif !important;
    font-weight: 600 !important; padding: 10px 28px !important;
}
.stButton > button:hover { background: #16375a !important; }
div[data-testid="stSidebarContent"] { background: #1a2e45; }
div[data-testid="stSidebarContent"] * { color: #e8edf3 !important; }
div[data-testid="stSidebarContent"] .stMarkdown h2 { color: #7ab3e0 !important; font-size: 13px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# 보험약가 로드
# ─────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_price_table():
    if not os.path.exists(PRICE_FILE):
        # 디버그: 실제 경로와 존재하는 파일 목록 표시
        import glob
        found = glob.glob("*.xlsx") + glob.glob("**/*.xlsx", recursive=True)
        st.error(f"⚠️ 약가 파일을 찾을 수 없습니다.\n\n찾는 경로: `{PRICE_FILE}`\n\n현재 디렉토리 xlsx 파일: {found}")
        return pd.DataFrame(columns=["제품명","업체명","상한금액","정규화명"])
    df = pd.read_excel(PRICE_FILE, sheet_name=0, header=0)
    df = df[["제품명","업체명","상한금액"]].copy()
    df["상한금액"] = pd.to_numeric(df["상한금액"], errors="coerce")
    df["정규화명"] = df["제품명"].apply(lambda x: re.sub(r"\s+","",str(x)).lower() if isinstance(x,str) else "")
    return df

def get_price_info(drug_name, price_df):
    if not isinstance(drug_name, str): return None, None, False
    if "비급여" in drug_name: return "비급여", None, True
    norm = re.sub(r"\s+","",drug_name).lower()
    match = price_df[price_df["정규화명"] == norm]
    if len(match) == 0:
        for l in [12,10,8,6]:
            prefix = norm[:l]
            if not prefix: continue
            match = price_df[price_df["정규화명"].str.startswith(prefix, na=False)]
            if len(match) > 0: break
    if len(match) > 0:
        row = match.iloc[0]
        return row["업체명"], row["상한금액"], False
    return None, None, False

# ─────────────────────────────────────────────
# 파일 파싱
# ─────────────────────────────────────────────
def parse_ocs_file(uploaded_file, price_df):
    fname = uploaded_file.name
    m = re.match(r"(.+?)(\d{6})", fname)
    hospital   = m.group(1) if m else fname
    period_str = m.group(2) if m else ""
    period     = f"{period_str[:4]}-{period_str[4:6]}" if len(period_str)==6 else period_str
    df   = pd.read_excel(uploaded_file, sheet_name="Sheet1", header=None)
    data = df.iloc[2:].copy()[[1,2,57]]
    data.columns = ["수가코드","약품명","합계수량"]
    data = data.dropna(subset=["약품명"])
    data["합계수량"] = pd.to_numeric(data["합계수량"], errors="coerce").fillna(0).astype(int)
    data = data[data["합계수량"] > 0].reset_index(drop=True)
    rows = []
    for _, row in data.iterrows():
        drug = row["약품명"]
        mfg, price, is_nongov = get_price_info(drug, price_df)
        if is_nongov:
            단가 = "비급여"; 매출액 = None
        elif price is not None:
            단가 = int(price); 매출액 = int(row["합계수량"] * price)
        else:
            단가 = None; 매출액 = None; mfg = None
        mfg_b = re.search(r"\(([^()]+)\)\s*$", drug)
        제약사 = mfg if mfg else (mfg_b.group(1) if mfg_b else "")
        rows.append({"병원명":hospital,"조회기간":period,"수가코드":row["수가코드"],
                     "약품명":drug,"제약사":제약사,"약가(단가)":단가,
                     "판매수량":int(row["합계수량"]),"매출액":매출액,
                     "신뢰도":"정상" if (is_nongov or price is not None) else "미매핑"})
    return pd.DataFrame(rows), hospital, period

# ─────────────────────────────────────────────
# DB
# ─────────────────────────────────────────────
def load_db():
    if os.path.exists(DB_FILE):
        with open(DB_FILE,"r",encoding="utf-8") as f: return json.load(f)
    return {}

def save_db(db):
    with open(DB_FILE,"w",encoding="utf-8") as f: json.dump(db,f,ensure_ascii=False,default=str)

def upsert_db(db, df):
    hospital = df["병원명"].iloc[0]; period = df["조회기간"].iloc[0]
    if hospital not in db: db[hospital] = {}
    db[hospital][period] = df.to_dict("records")
    return db

def get_periods(db, hospital):
    if hospital not in db: return []
    return sorted(db[hospital].keys())

def db_to_df(db, hospital, period):
    if hospital not in db or period not in db[hospital]: return pd.DataFrame()
    return pd.DataFrame(db[hospital][period])

def is_daewung(s): return isinstance(s,str) and any(k in s for k in DAEWUNG_KW)

# ─────────────────────────────────────────────
# 차트 헬퍼
# ─────────────────────────────────────────────
COLORS_UP  = ["#1F4E79","#2E75B6","#4A90D9","#6AAFE6","#9DCEF0"]
COLORS_DN  = ["#843c3c","#C0504D","#D9726F","#EE9A98","#F5BFBE"]
COLOR_DW   = "#C55A11"
COLOR_PREV = "#BDD7EE"
COLOR_CURR = "#1F4E79"

def short_name(name, max_len=14):
    """약품명을 그래프용으로 짧게 자르기"""
    if not isinstance(name, str): return str(name)
    # 괄호 앞까지만
    m = re.match(r"([^(]+)", name)
    s = m.group(1).strip() if m else name
    return s[:max_len] + "…" if len(s) > max_len else s

def chart_bar_top5(df_top, df_bot, title_up, title_dn, color_up, color_dn):
    """증가/감소 Top5 가로 바 차트 2개를 나란히"""
    col1, col2 = st.columns(2)

    # 증가 Top5
    with col1:
        if len(df_top) > 0:
            names = [short_name(n) for n in df_top["약품명"].tolist()][::-1]
            vals  = df_top["증감수량"].tolist()[::-1]
            fig = go.Figure(go.Bar(
                x=vals, y=names, orientation="h",
                marker_color=color_up,
                text=[f"+{v:,}" for v in vals],
                textposition="outside", textfont=dict(size=11),
            ))
            fig.update_layout(
                title=dict(text=title_up, font=dict(size=13, color="#1F4E79"), x=0),
                xaxis=dict(showgrid=True, gridcolor="#eee", zeroline=False),
                yaxis=dict(tickfont=dict(size=11)),
                plot_bgcolor="white", paper_bgcolor="white",
                margin=dict(l=10, r=60, t=40, b=10),
                height=280,
            )
            st.plotly_chart(fig, use_container_width=True)

    # 감소 Bottom5
    with col2:
        if len(df_bot) > 0:
            names = [short_name(n) for n in df_bot["약품명"].tolist()]
            vals  = df_bot["증감수량"].tolist()
            fig = go.Figure(go.Bar(
                x=vals, y=names, orientation="h",
                marker_color=color_dn,
                text=[f"{v:,}" for v in vals],
                textposition="outside", textfont=dict(size=11),
            ))
            fig.update_layout(
                title=dict(text=title_dn, font=dict(size=13, color="#843c3c"), x=0),
                xaxis=dict(showgrid=True, gridcolor="#eee", zeroline=False),
                yaxis=dict(tickfont=dict(size=11)),
                plot_bgcolor="white", paper_bgcolor="white",
                margin=dict(l=10, r=60, t=40, b=10),
                height=280,
            )
            st.plotly_chart(fig, use_container_width=True)

def chart_grouped_bar(df_top, df_bot, prev_period, curr_period, title):
    """전월/당월 수량 비교 묶음 막대그래프"""
    combined = pd.concat([df_top, df_bot]).drop_duplicates("약품명")
    if len(combined) == 0: return
    names = [short_name(n) for n in combined["약품명"].tolist()]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name=f"전월 ({prev_period})", x=names, y=combined["전월수량"].tolist(),
        marker_color=COLOR_PREV, text=combined["전월수량"].apply(lambda v: f"{v:,}"),
        textposition="outside", textfont=dict(size=10),
    ))
    fig.add_trace(go.Bar(
        name=f"당월 ({curr_period})", x=names, y=combined["당월수량"].tolist(),
        marker_color=COLOR_CURR, text=combined["당월수량"].apply(lambda v: f"{v:,}"),
        textposition="outside", textfont=dict(size=10),
    ))
    fig.update_layout(
        title=dict(text=title, font=dict(size=13, color="#1F4E79"), x=0),
        barmode="group",
        xaxis=dict(tickfont=dict(size=10)),
        yaxis=dict(showgrid=True, gridcolor="#eee"),
        plot_bgcolor="white", paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=10, r=10, t=50, b=10),
        height=320,
    )
    st.plotly_chart(fig, use_container_width=True)

def chart_dw_sales_gauge(dw_curr, total_curr, dw_prev, total_prev):
    """대웅제약 매출 비중 게이지 + 전월 비교"""
    col1, col2 = st.columns(2)
    with col1:
        ratio = (dw_curr / total_curr * 100) if total_curr > 0 else 0
        fig = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=ratio,
            delta={"reference": (dw_prev/total_prev*100) if total_prev>0 else 0,
                   "valueformat":".1f", "suffix":"%"},
            number={"suffix":"%", "valueformat":".1f",
                    "font":{"size":32,"color":"#C55A11"}},
            title={"text":"대웅제약 매출 비중", "font":{"size":13,"color":"#555"}},
            gauge={
                "axis":{"range":[0,30],"tickfont":{"size":10}},
                "bar":{"color":"#C55A11"},
                "bgcolor":"white",
                "steps":[{"range":[0,10],"color":"#FEF0E7"},
                         {"range":[10,20],"color":"#FDDCBD"},
                         {"range":[20,30],"color":"#F9C49A"}],
                "threshold":{"line":{"color":"#843c3c","width":3},"thickness":0.75,"value":20},
            }
        ))
        fig.update_layout(height=240, margin=dict(l=20,r=20,t=40,b=10),
                          paper_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        categories = ["전월", "당월"]
        fig = go.Figure()
        fig.add_trace(go.Bar(
            name="전체 매출", x=categories,
            y=[total_prev/1e8, total_curr/1e8],
            marker_color=COLOR_PREV,
            text=[f"{v/1e8:.1f}억" for v in [total_prev, total_curr]],
            textposition="outside",
        ))
        fig.add_trace(go.Bar(
            name="대웅제약", x=categories,
            y=[dw_prev/1e8, dw_curr/1e8],
            marker_color=COLOR_DW,
            text=[f"{v/1e8:.1f}억" for v in [dw_prev, dw_curr]],
            textposition="outside",
        ))
        fig.update_layout(
            title=dict(text="전월 대비 매출 비교 (억원)", font=dict(size=13,color="#1F4E79"), x=0),
            barmode="group",
            yaxis=dict(showgrid=True, gridcolor="#eee", title="억원"),
            plot_bgcolor="white", paper_bgcolor="white",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
            margin=dict(l=10,r=10,t=50,b=10), height=240,
        )
        st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────
# 엑셀 생성 (기존과 동일)
# ─────────────────────────────────────────────
def build_excel(curr_df, prev_df, hospital, curr_period, prev_period):
    thin  = Side(style="thin",   color="BFBFBF")
    BT    = Border(left=thin, right=thin, top=thin, bottom=thin)
    H1    = PatternFill("solid", start_color="1F4E79")
    H2    = PatternFill("solid", start_color="2E75B6")
    DW    = PatternFill("solid", start_color="FFF2CC")
    ALT   = PatternFill("solid", start_color="F5F9FF")
    WHT   = PatternFill("solid", start_color="FFFFFF")
    SUM   = PatternFill("solid", start_color="DEEAF1")
    OR1   = PatternFill("solid", start_color="833C00")
    OR2   = PatternFill("solid", start_color="C55A11")
    OR3   = PatternFill("solid", start_color="FFF2CC")
    TOP   = PatternFill("solid", start_color="E2EFDA")
    BOT   = PatternFill("solid", start_color="FFE7E7")

    def hc(ws,r,c,v,fill=H1,sz=10,bold=True,fc="FFFFFF",wrap=False):
        cell=ws.cell(r,c,v); cell.font=Font(name="맑은 고딕",size=sz,bold=bold,color=fc)
        cell.fill=fill; cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=wrap)
        cell.border=BT; return cell

    def dc(ws,r,c,v,fill=WHT,align="left",fmt=None,bold=False):
        cell=ws.cell(r,c,v); cell.font=Font(name="맑은 고딕",size=9,bold=bold)
        cell.fill=fill; cell.alignment=Alignment(horizontal=align,vertical="center")
        cell.border=BT
        if fmt: cell.number_format=fmt
        return cell

    prev_qty  = prev_df.groupby("수가코드")["판매수량"].sum().rename("전월수량") if len(prev_df) else pd.Series(dtype=int)
    curr_qty  = curr_df.groupby("수가코드")["판매수량"].sum().rename("당월수량")
    all_codes = pd.concat([prev_df[["수가코드","약품명","제약사","약가(단가)","매출액"]],
                           curr_df[["수가코드","약품명","제약사","약가(단가)","매출액"]]
                           ]).drop_duplicates("수가코드").set_index("수가코드")
    mom = all_codes.join(curr_qty).join(prev_qty)
    mom["전월수량"] = mom["전월수량"].fillna(0).astype(int)
    mom["당월수량"] = mom["당월수량"].fillna(0).astype(int)
    mom["증감수량"] = mom["당월수량"] - mom["전월수량"]
    mom = mom.reset_index()

    def build_rank(src,n=5,largest=True):
        fn=src.nlargest if largest else src.nsmallest
        sub=fn(n,"증감수량")[["약품명","제약사","약가(단가)","당월수량","매출액","증감수량"]].copy()
        sub.insert(0,"병원명",hospital); sub.insert(1,"순위",range(1,len(sub)+1))
        return sub

    top5_all=build_rank(mom,5,True);  bot5_all=build_rank(mom,5,False)
    mom_dw=mom[mom["제약사"].apply(is_daewung)]
    top5_dw=build_rank(mom_dw,5,True); bot5_dw=build_rank(mom_dw,5,False)

    total_prev=int(prev_df["매출액"].sum(skipna=True)) if len(prev_df) else 0
    total_curr=int(curr_df["매출액"].sum(skipna=True))
    dw_prev=int(prev_df[prev_df["제약사"].apply(is_daewung)]["매출액"].sum(skipna=True)) if len(prev_df) else 0
    dw_curr=int(curr_df[curr_df["제약사"].apply(is_daewung)]["매출액"].sum(skipna=True))

    wb=Workbook()

    # Sheet1
    ws1=wb.active; ws1.title="전체매출집계"
    ws1.sheet_view.showGridLines=False; ws1.freeze_panes="A5"
    ws1.merge_cells("A1:I2")
    c=ws1["A1"]; c.value=f"🏥 {hospital} OCS 처방 분석  |  {prev_period} → {curr_period}"
    c.font=Font(name="맑은 고딕",size=14,bold=True,color="FFFFFF")
    c.fill=H1; c.alignment=Alignment(horizontal="center",vertical="center")
    summary=[("전체 총 매출 (전월)",f"{total_prev:,}원"),("전체 총 매출 (당월)",f"{total_curr:,}원"),
             ("대웅제약 매출 (전월)",f"{dw_prev:,}원"),("대웅제약 매출 (당월)",f"{dw_curr:,}원"),
             ("MoM 대웅 증감",f"{dw_curr-dw_prev:+,}원")]
    for i,(lbl,val) in enumerate(summary):
        col=i*2+1
        c1=ws1.cell(3,col,lbl); c1.font=Font(name="맑은 고딕",size=8,bold=True,color="1F4E79")
        c1.fill=SUM; c1.alignment=Alignment(horizontal="center"); c1.border=BT
        c2=ws1.cell(3,col+1,val); c2.font=Font(name="맑은 고딕",size=9,bold=True,color="1F4E79")
        c2.fill=SUM; c2.alignment=Alignment(horizontal="right"); c2.border=BT
    hdrs1=["병원명","약품명","제약사","약가(단가)(원)",f"판매수량 {prev_period}",
           f"판매수량 {curr_period}","증감수량",f"매출액 {curr_period}(원)","신뢰도"]
    for ci,h in enumerate(hdrs1,1): hc(ws1,4,ci,h,wrap=True)
    ws1.row_dimensions[4].height=30
    merged=mom.sort_values("당월수량",ascending=False)
    for ri,(_,row) in enumerate(merged.iterrows()):
        fill=DW if is_daewung(row["제약사"]) else (ALT if ri%2 else WHT)
        vals=[hospital,row["약품명"],row["제약사"],row["약가(단가)"],
              int(row["전월수량"]),int(row["당월수량"]),int(row["증감수량"]),
              int(row["매출액"]) if pd.notna(row.get("매출액")) else None,"정상"]
        fmts=[None,None,None,"#,##0","#,##0","#,##0","#,##0","#,##0",None]
        for ci,(v,f) in enumerate(zip(vals,fmts),1):
            align="right" if ci in [4,5,6,7,8] else ("center" if ci in [1,9] else "left")
            dc(ws1,5+ri,ci,v,fill=fill,align=align,fmt=f)
    for i,w in enumerate([12,40,18,14,14,14,12,18,8],1):
        ws1.column_dimensions[get_column_letter(i)].width=w

    def write_rank_section(ws,sr,title,df_in,hdr_fill,row_fill):
        ws.merge_cells(start_row=sr,start_column=1,end_row=sr,end_column=8)
        c=ws.cell(sr,1,title); c.font=Font(name="맑은 고딕",size=10,bold=True,color="FFFFFF")
        c.fill=hdr_fill; c.alignment=Alignment(horizontal="center"); c.border=BT
        for ci,h in enumerate(["병원명","순위","약품명","제약사","약가(단가)","당월수량","매출액(원)","증감수량"],1):
            hc(ws,sr+1,ci,h,fill=hdr_fill)
        r=sr+2
        for _,(_,row) in enumerate(df_in.iterrows()):
            vals=[row["병원명"],row["순위"],row["약품명"],row["제약사"],row["약가(단가)"],
                  row["당월수량"],int(row["매출액"]) if pd.notna(row.get("매출액")) else None,row["증감수량"]]
            fmts=[None,None,None,None,"#,##0","#,##0","#,##0","#,##0"]
            for ci,(v,f) in enumerate(zip(vals,fmts),1):
                align="right" if ci in [5,6,7,8] else ("center" if ci in [1,2] else "left")
                dc(ws,r,ci,v,fill=row_fill,align=align,fmt=f)
            r+=1
        return r+1

    ws2=wb.create_sheet("전체_증감Top5"); ws2.sheet_view.showGridLines=False
    ws2.merge_cells("A1:H1"); c=ws2["A1"]
    c.value=f"전체 처방수량 증감 Top5/Bottom5  ({prev_period}→{curr_period})"
    c.font=Font(name="맑은 고딕",size=12,bold=True,color="FFFFFF")
    c.fill=H1; c.alignment=Alignment(horizontal="center",vertical="center"); ws2.row_dimensions[1].height=25
    r=write_rank_section(ws2,2,"▲ 증가 Top5 (전체)",top5_all,H2,TOP)
    r=write_rank_section(ws2,r,"▼ 감소 Bottom5 (전체)",bot5_all,H2,BOT)
    for i,w in enumerate([12,6,40,18,14,12,16,12],1):
        ws2.column_dimensions[get_column_letter(i)].width=w

    ws3=wb.create_sheet("대웅제약_증감Top5"); ws3.sheet_view.showGridLines=False
    ws3.merge_cells("A1:H1"); c=ws3["A1"]
    c.value=f"대웅제약 처방수량 증감 Top5/Bottom5  ({prev_period}→{curr_period})"
    c.font=Font(name="맑은 고딕",size=12,bold=True,color="FFFFFF")
    c.fill=OR1; c.alignment=Alignment(horizontal="center",vertical="center"); ws3.row_dimensions[1].height=25
    r=write_rank_section(ws3,2,"▲ 증가 Top5 (대웅제약)",top5_dw,OR2,OR3)
    r=write_rank_section(ws3,r,"▼ 감소 Bottom5 (대웅제약)",bot5_dw,OR2,PatternFill("solid",start_color="FCE4D6"))
    for i,w in enumerate([12,6,40,18,14,12,16,12],1):
        ws3.column_dimensions[get_column_letter(i)].width=w

    ws4=wb.create_sheet("대웅제약_전체"); ws4.sheet_view.showGridLines=False; ws4.freeze_panes="A4"
    ws4.merge_cells("A1:I2"); c=ws4["A1"]
    c.value=f"대웅제약 품목 전체  |  {prev_period} vs {curr_period}"
    c.font=Font(name="맑은 고딕",size=13,bold=True,color="FFFFFF")
    c.fill=OR1; c.alignment=Alignment(horizontal="center",vertical="center")
    for ci,h in enumerate(["병원명","약품명","제약사","약가(단가)(원)",
                            f"판매수량 {prev_period}",f"판매수량 {curr_period}",
                            "증감수량",f"매출액 {curr_period}(원)","신뢰도"],1):
        hc(ws4,3,ci,h,fill=OR2,wrap=True)
    ws4.row_dimensions[3].height=30
    dw_all=mom_dw.sort_values("당월수량",ascending=False).reset_index(drop=True)
    for ri,(_,row) in enumerate(dw_all.iterrows()):
        fill=DW if ri%2==0 else PatternFill("solid",start_color="FFF9EC")
        vals=[hospital,row["약품명"],row["제약사"],row["약가(단가)"],
              int(row["전월수량"]),int(row["당월수량"]),int(row["증감수량"]),
              int(row["매출액"]) if pd.notna(row.get("매출액")) else None,"정상"]
        fmts=[None,None,None,"#,##0","#,##0","#,##0","#,##0","#,##0",None]
        for ci,(v,f) in enumerate(zip(vals,fmts),1):
            align="right" if ci in [4,5,6,7,8] else ("center" if ci in [1,9] else "left")
            dc(ws4,4+ri,ci,v,fill=fill,align=align,fmt=f)
    tr=4+len(dw_all)
    ws4.merge_cells(start_row=tr,start_column=1,end_row=tr,end_column=3)
    c=ws4.cell(tr,1,"합  계"); c.font=Font(name="맑은 고딕",size=10,bold=True,color="FFFFFF")
    c.fill=OR2; c.alignment=Alignment(horizontal="center"); c.border=BT
    for ci in [2,3]: ws4.cell(tr,ci).fill=OR2; ws4.cell(tr,ci).border=BT
    for ci,val in {5:dw_all["전월수량"].sum(),6:dw_all["당월수량"].sum(),
                   7:int(dw_all["당월수량"].sum())-int(dw_all["전월수량"].sum()),
                   8:dw_all["매출액"].sum()}.items():
        c2=ws4.cell(tr,ci,int(val) if pd.notna(val) else "")
        c2.font=Font(name="맑은 고딕",size=9,bold=True,color="FFFFFF")
        c2.fill=OR2; c2.alignment=Alignment(horizontal="right"); c2.border=BT; c2.number_format="#,##0"
    for col in [4,9]: ws4.cell(tr,col).fill=OR2; ws4.cell(tr,col).border=BT
    for i,w in enumerate([12,40,18,14,14,14,12,18,8],1):
        ws4.column_dimensions[get_column_letter(i)].width=w

    buf=BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ─────────────────────────────────────────────
# 사이드바
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 💊 OCS 분석기")
    st.markdown("---")
    st.markdown("**사용 방법**")
    st.markdown("1. xlsx 파일 업로드\n2. 분석 실행 클릭\n3. 엑셀 다운로드")
    st.markdown("---")
    st.markdown("**파일명 규칙**")
    st.code("병원명YYYYMM.xlsx\n예) 김해복음병원202604po.xlsx")
    st.markdown("---")
    st.markdown("**범례**")
    st.markdown("🟡 대웅제약 품목\n\n🟢 수량 증가\n\n🔴 수량 감소")

# ─────────────────────────────────────────────
# 메인 UI
# ─────────────────────────────────────────────
st.markdown("""
<div class="upload-hint">
  <h3>📂 처방 통계 파일 업로드</h3>
  <p>병원 OCS에서 받은 xlsx 파일을 올려주세요. 여러 달치를 한번에 올려도 됩니다.</p>
</div>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "파일 선택 (xlsx, 여러 개 가능)",
    type=["xlsx"], accept_multiple_files=True, label_visibility="collapsed",
)

if uploaded_files:
    price_df = load_price_table()
    if price_df.empty:
        st.warning("⚠️ 보험약가 파일(xlsx)이 없습니다. 앱 폴더에 약가 파일을 넣어주세요.")

    db = load_db()
    with st.spinner("파일 분석 중..."):
        parsed = []
        for f in uploaded_files:
            try:
                df, hosp, period = parse_ocs_file(f, price_df)
                parsed.append((df, hosp, period))
                db = upsert_db(db, df)
            except Exception as e:
                st.error(f"❌ {f.name} 파싱 오류: {e}")
        save_db(db)

    if not parsed: st.stop()

    st.markdown(f"✅ **{len(parsed)}개 파일** 업로드 완료")
    for df, hosp, period in parsed:
        total_qty  = df["판매수량"].sum()
        total_sale = int(df["매출액"].sum(skipna=True))
        dw_sale    = int(df[df["제약사"].apply(is_daewung)]["매출액"].sum(skipna=True))
        st.markdown(f"&nbsp;&nbsp;• `{hosp}` **{period}** — 총 {total_qty:,}개 / 매출 {total_sale:,}원 / 대웅 {dw_sale:,}원")

    st.markdown("---")

    hospitals   = list({hosp for _,hosp,_ in parsed})
    sel_hosp    = st.selectbox("📍 분석할 병원", hospitals) if len(hospitals)>1 else hospitals[0]
    all_periods = get_periods(db, sel_hosp)

    if len(all_periods) < 2:
        st.info("ℹ️ 전월 비교를 위해 2개월치 이상 업로드해 주세요.")
        if len(all_periods)==1: st.markdown(f"현재 저장된 기간: `{all_periods[0]}`")
        st.stop()

    col1, col2 = st.columns(2)
    with col1: prev_period = st.selectbox("📅 전월 선택", all_periods[:-1], index=len(all_periods)-2)
    with col2: curr_period = st.selectbox("📅 당월 선택", all_periods[1:],  index=len(all_periods)-2)

    if st.button("🔍 분석 실행"):
        prev_df = db_to_df(db, sel_hosp, prev_period)
        curr_df = db_to_df(db, sel_hosp, curr_period)

        # ── KPI 카드 ──────────────────────────────────────
        total_curr  = int(curr_df["매출액"].sum(skipna=True))
        total_prev  = int(prev_df["매출액"].sum(skipna=True))
        dw_curr     = int(curr_df[curr_df["제약사"].apply(is_daewung)]["매출액"].sum(skipna=True))
        dw_prev     = int(prev_df[prev_df["제약사"].apply(is_daewung)]["매출액"].sum(skipna=True))
        delta_total = total_curr - total_prev
        delta_dw    = dw_curr   - dw_prev

        k1,k2,k3,k4 = st.columns(4)
        for col, label, val, delta, cls in [
            (k1, "전체 총 매출 (당월)",  total_curr,  delta_total, ""),
            (k2, "전월 대비 증감",        delta_total, None,        "green" if delta_total>=0 else "red"),
            (k3, "대웅제약 매출 (당월)",  dw_curr,     delta_dw,    "orange"),
            (k4, "대웅제약 전월 대비",    delta_dw,    None,        "green" if delta_dw>=0 else "red"),
        ]:
            arrow = ("▲" if val>=0 else "▼") if any(k in label for k in ["증감","대비"]) else ""
            color = "#375623" if val>=0 else "#843c3c"
            is_delta_card = any(k in label for k in ["증감","대비"])
            val_str = f"{val:+,}원" if is_delta_card else f"{val:,}원"
            delta_html = "" if delta is None else f'<div class="kpi-delta" style="color:{"#375623" if delta>=0 else "#843c3c"}">{"▲" if delta>=0 else "▼"} {abs(delta):,}원</div>'
            col.markdown(f"""
<div class="kpi-card {cls}">
  <div class="kpi-label">{label}</div>
  <div class="kpi-value">{val_str}</div>
  {delta_html}
</div>""", unsafe_allow_html=True)

        # ── MoM 계산 ──────────────────────────────────────
        prev_qty_map = prev_df.groupby("수가코드")["판매수량"].sum()
        curr_qty_map = curr_df.groupby("수가코드")["판매수량"].sum()
        all_codes    = pd.concat([
            prev_df[["수가코드","약품명","제약사"]],
            curr_df[["수가코드","약품명","제약사"]]
        ]).drop_duplicates("수가코드").set_index("수가코드")
        mom = all_codes.copy()
        mom["전월수량"] = mom.index.map(prev_qty_map).fillna(0).astype(int)
        mom["당월수량"] = mom.index.map(curr_qty_map).fillna(0).astype(int)
        mom["증감수량"] = mom["당월수량"] - mom["전월수량"]
        mom = mom.reset_index()
        mom_dw = mom[mom["제약사"].apply(is_daewung)]

        top5    = mom.nlargest(5,"증감수량")
        bot5    = mom.nsmallest(5,"증감수량")
        top5_dw = mom_dw.nlargest(5,"증감수량")
        bot5_dw = mom_dw.nsmallest(5,"증감수량")

        # ══════════════════════════════════════════════════
        # 대시보드 탭
        # ══════════════════════════════════════════════════
        tab_dash, tab_all, tab_dw, tab_table = st.tabs([
            "📊 대시보드", "🔵 전체 증감 Top5", "🟡 대웅제약 Top5", "📋 대웅제약 전체"
        ])

        # ── 대시보드 탭 ────────────────────────────────────
        with tab_dash:
            st.markdown('<div class="section-title">💰 대웅제약 매출 현황</div>', unsafe_allow_html=True)
            chart_dw_sales_gauge(dw_curr, total_curr, dw_prev, total_prev)

            st.markdown('<div class="section-title">🔵 전체 품목 처방수량 증감</div>', unsafe_allow_html=True)
            chart_bar_top5(top5, bot5,
                           f"▲ 증가 Top5  ({prev_period}→{curr_period})",
                           f"▼ 감소 Bottom5  ({prev_period}→{curr_period})",
                           COLORS_UP[0], COLORS_DN[0])
            chart_grouped_bar(top5, bot5, prev_period, curr_period,
                              f"전월/당월 수량 비교 — 전체 Top/Bottom5")

            st.markdown('<div class="section-title">🟡 대웅제약 품목 처방수량 증감</div>', unsafe_allow_html=True)
            chart_bar_top5(top5_dw, bot5_dw,
                           f"▲ 대웅 증가 Top5  ({prev_period}→{curr_period})",
                           f"▼ 대웅 감소 Bottom5  ({prev_period}→{curr_period})",
                           COLOR_DW, COLORS_DN[0])
            chart_grouped_bar(top5_dw, bot5_dw, prev_period, curr_period,
                              f"전월/당월 수량 비교 — 대웅제약 Top/Bottom5")

        # ── 전체 증감 탭 ──────────────────────────────────
        with tab_all:
            chart_bar_top5(top5, bot5,
                           f"▲ 증가 Top5", f"▼ 감소 Bottom5",
                           COLORS_UP[0], COLORS_DN[0])
            st.markdown('<div class="section-title">▲ 증가 Top 5</div>', unsafe_allow_html=True)
            st.dataframe(top5[["약품명","제약사","전월수량","당월수량","증감수량"]].reset_index(drop=True),
                         use_container_width=True, hide_index=True)
            st.markdown('<div class="section-title">▼ 감소 Bottom 5</div>', unsafe_allow_html=True)
            st.dataframe(bot5[["약품명","제약사","전월수량","당월수량","증감수량"]].reset_index(drop=True),
                         use_container_width=True, hide_index=True)

        # ── 대웅제약 탭 ────────────────────────────────────
        with tab_dw:
            chart_bar_top5(top5_dw, bot5_dw,
                           f"▲ 대웅 증가 Top5", f"▼ 대웅 감소 Bottom5",
                           COLOR_DW, COLORS_DN[0])
            st.markdown('<div class="section-title">▲ 대웅제약 증가 Top 5</div>', unsafe_allow_html=True)
            st.dataframe(top5_dw[["약품명","제약사","전월수량","당월수량","증감수량"]].reset_index(drop=True),
                         use_container_width=True, hide_index=True)
            st.markdown('<div class="section-title">▼ 대웅제약 감소 Bottom 5</div>', unsafe_allow_html=True)
            st.dataframe(bot5_dw[["약품명","제약사","전월수량","당월수량","증감수량"]].reset_index(drop=True),
                         use_container_width=True, hide_index=True)

        # ── 대웅제약 전체 탭 ──────────────────────────────
        with tab_table:
            dw_full = curr_df[curr_df["제약사"].apply(is_daewung)][
                ["약품명","제약사","약가(단가)","판매수량","매출액"]
            ].sort_values("판매수량", ascending=False)
            st.dataframe(dw_full.reset_index(drop=True), use_container_width=True, hide_index=True)

        # ── 엑셀 다운로드 ─────────────────────────────────
        st.markdown("---")
        st.markdown("### 📥 엑셀 다운로드")
        excel_buf = build_excel(
            pd.DataFrame(db[sel_hosp][curr_period]),
            pd.DataFrame(db[sel_hosp][prev_period]),
            sel_hosp, curr_period, prev_period
        )
        fname = f"{sel_hosp}_{curr_period}_OCS분석.xlsx"
        st.download_button(
            label=f"⬇️ {fname} 다운로드",
            data=excel_buf, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

else:
    st.markdown("""
    <div style="text-align:center; padding: 60px 0; color: #aaa;">
        <div style="font-size:48px; margin-bottom:12px;">📂</div>
        <div style="font-size:16px; font-weight:600;">위에서 파일을 업로드해 주세요</div>
        <div style="font-size:13px; margin-top:6px;">예) 김해복음병원202604po.xlsx</div>
    </div>
    """, unsafe_allow_html=True)
