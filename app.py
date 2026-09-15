import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path

# ------------------------------------------------------------------
# 기본 설정
# ------------------------------------------------------------------
st.set_page_config(page_title="연구실적 대시보드", layout="wide", initial_sidebar_state="expanded")

DATA_PATH = Path(__file__).parent / "rawdata_2608.xlsx"

COLOR_HIGHLIGHT = "#E8392A"   # 강조(빨간색)
COLOR_BASE = "#93B8E0"        # 나머지(파란색 계열)
COLOR_NAVY = "#0D2B5E"        # 사이드바 네이비
COLOR_BG = "#F5F7FA"          # 배경
GRADIENT_START = "#DCEAF7"    # 파란 그라데이션 시작(밝음)
GRADIENT_END = "#1D5FA8"      # 파란 그라데이션 끝(진함)

# 학사연도 기준 월 순서 (3월 ~ 다음해 2월)
FISCAL_MONTHS = ["3월", "4월", "5월", "6월", "7월", "8월",
                  "9월", "10월", "11월", "12월", "1월", "2월"]

DEPT_ORDER = [
    # 학부(대학)
    "인문대학", "사회과학대학", "경영대학",
    "화학·생명과학대학", "반도체·ICT대학", "스마트시스템공과대학",
    "인공지능·소프트웨어융합대학", "건축대학", "미디어·휴먼라이프대학",
    "스포츠·예술대학", "스포츠학부(체육학전공, 스포츠산업학전공)",
    "미래융합대학", "방목기초교육대학",
    # 대학원
    "(일반)대학원", "교육대학원", "통합치료대학원", "기록정보과학전문대학원",
]

# ------------------------------------------------------------------
# 스타일 (CSS)
# ------------------------------------------------------------------
def inject_css():
    st.markdown(f"""
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700;900&display=swap" rel="stylesheet">
    <style>
        html, body, [class*="css"] {{
            font-family: 'Noto Sans KR', sans-serif !important;
        }}
        .stApp {{
            background-color: {COLOR_BG};
        }}
        section[data-testid="stSidebar"] {{
            background-color: {COLOR_NAVY};
        }}
        section[data-testid="stSidebar"] * {{
            color: #FFFFFF !important;
        }}
        section[data-testid="stSidebar"] .stRadio > label {{
            color: #FFFFFF !important;
        }}
        div[role="radiogroup"] label span {{
            color: #FFFFFF !important;
        }}
        .kpi-card {{
            background-color: #FFFFFF;
            border-radius: 14px;
            padding: 22px 20px;
            box-shadow: 0 2px 8px rgba(13,43,94,0.08);
            border-left: 6px solid {COLOR_NAVY};
            height: 100%;
        }}
        .kpi-card.highlight {{
            border-left: 6px solid {COLOR_HIGHLIGHT};
        }}
        .kpi-label {{
            font-size: 15px;
            font-weight: 500;
            color: #5A6B85;
            margin-bottom: 6px;
        }}
        .kpi-value {{
            font-size: 34px;
            font-weight: 700;
            color: {COLOR_NAVY};
        }}
        .kpi-value.highlight {{
            color: {COLOR_HIGHLIGHT};
        }}
        .page-title {{
            font-size: 28px;
            font-weight: 700;
            color: {COLOR_NAVY};
            margin-bottom: 4px;
        }}
        .page-subtitle {{
            font-size: 15px;
            color: #5A6B85;
            margin-bottom: 24px;
        }}
        .section-title {{
            font-size: 18px;
            font-weight: 700;
            color: {COLOR_NAVY};
            margin-top: 10px;
            margin-bottom: 10px;
        }}
        .footnote {{
            font-size: 12.5px;
            color: #8A93A3;
            border-top: 1px solid #E1E6ED;
            margin-top: 36px;
            padding-top: 10px;
            line-height: 1.6;
        }}
    </style>
    """, unsafe_allow_html=True)


def kpi_card(label, value, highlight=False, suffix=""):
    cls = "kpi-card highlight" if highlight else "kpi-card"
    vcls = "kpi-value highlight" if highlight else "kpi-value"
    st.markdown(f"""
    <div class="{cls}">
        <div class="kpi-label">{label}</div>
        <div class="{vcls}">{value}{suffix}</div>
    </div>
    """, unsafe_allow_html=True)


def footnote(text):
    st.markdown(f'<div class="footnote">{text}</div>', unsafe_allow_html=True)


def page_header(title, subtitle=None):
    st.markdown(f'<div class="page-title">{title}</div>', unsafe_allow_html=True)
    if subtitle:
        st.markdown(f'<div class="page-subtitle">{subtitle}</div>', unsafe_allow_html=True)


# ------------------------------------------------------------------
# 유틸 함수
# ------------------------------------------------------------------
def month_to_int(m):
    """'3월' 또는 3 형태를 모두 정수로 변환"""
    if pd.isna(m):
        return np.nan
    if isinstance(m, str):
        return int(m.replace("월", "").strip())
    return int(m)


def blue_gradient(n):
    """n개의 색상을 GRADIENT_START -> GRADIENT_END 로 선형 보간 (밝은 -> 진한)"""
    def hex_to_rgb(h):
        h = h.lstrip("#")
        return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

    def rgb_to_hex(rgb):
        return "#{:02X}{:02X}{:02X}".format(*[int(round(c)) for c in rgb])

    start = hex_to_rgb(GRADIENT_START)
    end = hex_to_rgb(GRADIENT_END)
    if n <= 1:
        return [GRADIENT_END]
    colors = []
    for i in range(n):
        t = i / (n - 1)
        rgb = tuple(start[j] + (end[j] - start[j]) * t for j in range(3))
        colors.append(rgb_to_hex(rgb))
    return colors


def apply_chart_style(fig, dual_axis=False, legend=True):
    """차트 폰트 사이즈 등 공통 스타일 적용"""
    fig.update_layout(
        font=dict(family="Noto Sans KR, sans-serif", size=13, color="#2B3A55"),
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        margin=dict(l=10, r=10, t=40, b=10),
        legend=dict(
            font=dict(size=15),
            orientation="h",
            yanchor="bottom", y=1.02,
            xanchor="right", x=1
        ) if legend else dict(),
        showlegend=legend,
        hoverlabel=dict(font_size=13, font_family="Noto Sans KR, sans-serif"),
    )
    fig.update_xaxes(
        title_font=dict(size=15),
        tickfont=dict(size=13),
        showgrid=False,
        linecolor="#D8DEE8",
    )
    fig.update_yaxes(
        title_font=dict(size=15),
        tickfont=dict(size=13),
        gridcolor="#EEF1F6",
        secondary_y=False if dual_axis else None,
    )
    if dual_axis:
        fig.update_yaxes(title_font=dict(size=15), tickfont=dict(size=13), secondary_y=True, showgrid=False)
    return fig


# ------------------------------------------------------------------
# 데이터 로드
# ------------------------------------------------------------------
def _prep_sheet1(df):
    df = df.copy()
    df["일자"] = pd.to_datetime(df["일자"], errors="coerce")
    df["승인일자"] = pd.to_datetime(df["승인일자"], errors="coerce")
    df["월"] = df["일자"].dt.month
    df["승인월"] = df["승인일자"].dt.month
    return df


@st.cache_data
def load_sheet1():
    df = pd.read_excel(DATA_PATH, sheet_name="Sheet1")
    return _prep_sheet1(df)


@st.cache_data
def load_sheet1_1():
    df = pd.read_excel(DATA_PATH, sheet_name="Sheet1-1")
    return _prep_sheet1(df)


@st.cache_data
def load_sheet2():
    df = pd.read_excel(DATA_PATH, sheet_name="Sheet2")
    df["월_num"] = df["월"].apply(month_to_int)
    return df


@st.cache_data
def load_year_sheets():
    years = ["2023", "2024", "2025", "2026"]
    data = {}
    for y in years:
        df = pd.read_excel(DATA_PATH, sheet_name=y)
        data[y] = df
    return data


# ------------------------------------------------------------------
# 페이지 1: 이달 연구실적 개요
# ------------------------------------------------------------------
def page_overview():
    df = load_sheet1()        # 이번달까지 최신 스냅샷
    df_prev = load_sheet1_1()  # 지난달까지 스냅샷 (비교 기준)

    current_month = int(df["월"].max())

    page_header("이달 연구실적 개요", f"기준월: {current_month}월  ·  데이터: rawdata_2608.xlsx (Sheet1, Sheet1-1)")

    # KPI 계산
    count_this_month = int((df["월"] == current_month).sum())
    count_approved_this_month = len(df) - len(df_prev)  # Sheet1(최신) - Sheet1-1(전월 스냅샷)

    col1, col2 = st.columns(2)
    with col1:
        kpi_card("이달 실적 건수", f"{count_this_month:,}", highlight=True, suffix=" 건")
    with col2:
        kpi_card("이달 승인 건수", f"{count_approved_this_month:,}", suffix=" 건")

    st.markdown("<br>", unsafe_allow_html=True)

    col_a, col_b = st.columns(2)

    # 월별 실적 건수 추이 (Sheet1 vs Sheet1-1 비교)
    with col_a:
        st.markdown('<div class="section-title">월별 실적 건수 추이</div>', unsafe_allow_html=True)
        all_months = sorted(set(df["월"].dropna().astype(int)) | set(df_prev["월"].dropna().astype(int)))
        monthly_new = df.groupby("월").size().reindex(all_months)
        monthly_prev = df_prev.groupby("월").size().reindex(all_months)

        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=[f"{m}월" for m in all_months], y=monthly_prev.values,
            mode="lines+markers", name="Sheet1-1 (전월 기준)",
            line=dict(color=COLOR_BASE, width=2.5), marker=dict(size=6),
        ))
        fig.add_trace(go.Scatter(
            x=[f"{m}월" for m in all_months], y=monthly_new.values,
            mode="lines+markers", name="Sheet1 (이번달 기준)",
            line=dict(color=COLOR_HIGHLIGHT, width=4), marker=dict(size=8),
        ))
        fig.update_layout(yaxis_title="실적 건수 (건)")
        apply_chart_style(fig)
        st.plotly_chart(fig, use_container_width=True)

    # 소속별 실적 건수 차트 (Sheet1 vs Sheet1-1 비교)
    with col_b:
        st.markdown('<div class="section-title">소속별 실적 건수</div>', unsafe_allow_html=True)
        dept_new = df.groupby("소속(대)").size()
        dept_prev = df_prev.groupby("소속(대)").size()
        all_depts = set(dept_new.index) | set(dept_prev.index)
        ordered_depts = [d for d in DEPT_ORDER if d in all_depts]
        ordered_depts += [d for d in all_depts if d not in DEPT_ORDER]

        vals_new = [dept_new.get(d, 0) for d in ordered_depts]
        vals_prev = [dept_prev.get(d, 0) for d in ordered_depts]

        fig2 = go.Figure()
        fig2.add_trace(go.Bar(
            y=ordered_depts, x=vals_prev, name="Sheet1-1 (전월 기준)",
            orientation="h", marker_color=COLOR_BASE,
        ))
        fig2.add_trace(go.Bar(
            y=ordered_depts, x=vals_new, name="Sheet1 (이번달 기준)",
            orientation="h", marker_color=COLOR_HIGHLIGHT,
        ))
        fig2.update_layout(
            barmode="group",
            yaxis=dict(autorange="reversed"),
            xaxis_title="실적 건수 (건)",
            height=560,
        )
        apply_chart_style(fig2)
        st.plotly_chart(fig2, use_container_width=True)

    footnote(
        "※ 데이터 출처: rawdata_2608.xlsx (Sheet1, Sheet1-1) · 기준월은 Sheet1 내 최신 월(일자 기준)로 자동 산정됩니다.<br>"
        "※ 이달 승인 건수 = Sheet1(이번달 최신 스냅샷) 행 수 − Sheet1-1(전월 스냅샷) 행 수로, 스냅샷 사이에 새로 등록·승인된 건수를 의미합니다."
    )


# ------------------------------------------------------------------
# 페이지 2: 연도별 월별 실적 추이
# ------------------------------------------------------------------
def page_yearly_trend():
    page_header("연도별 월별 실적 추이", "데이터: rawdata_2608.xlsx (2023 ~ 2026 시트)")

    year_data = load_year_sheets()
    years = list(year_data.keys())
    highlight_year = "2026"
    other_years = [y for y in years if y != highlight_year]
    grad_colors = blue_gradient(len(other_years))
    color_map = dict(zip(other_years, grad_colors))
    color_map[highlight_year] = COLOR_HIGHLIGHT

    fig = go.Figure()
    for y in years:
        df = year_data[y]
        month_cols = [m for m in FISCAL_MONTHS if m in df.columns]
        totals = df[month_cols].sum(min_count=1)
        is_highlight = (y == highlight_year)
        fig.add_trace(go.Scatter(
            x=month_cols,
            y=[totals[m] for m in month_cols],
            mode="lines+markers",
            name=y,
            line=dict(color=color_map[y], width=4 if is_highlight else 2.5),
            marker=dict(size=8 if is_highlight else 6),
        ))

    fig.update_layout(
        xaxis_title="월",
        yaxis_title="실적 건수 (건)",
        height=520,
    )
    apply_chart_style(fig)
    st.plotly_chart(fig, use_container_width=True)

    footnote(
        "※ 데이터 출처: rawdata_2608.xlsx (2023~2026 연도별 시트) · 구분(논문/저서 등) 전체 합계 기준입니다.<br>"
        "※ 2026년은 실적이 집계되지 않은 미래 월을 NaN으로 처리하여 그래프에 표시하지 않습니다."
    )


# ------------------------------------------------------------------
# 페이지 3: 산학협력 현황
# ------------------------------------------------------------------
def page_collaboration():
    df = load_sheet2()
    month_order = [m for m in FISCAL_MONTHS if m in df["월"].unique()]
    current_month = month_order[-1] if month_order else None

    page_header("산학협력 현황", f"기준월: {current_month}  ·  데이터: rawdata_2608.xlsx (Sheet2)")

    col_a, col_b = st.columns([1, 1.4])

    # 이달 제안 수 중 선정 수
    with col_a:
        st.markdown('<div class="section-title">이달 제안 수 중 선정 수</div>', unsafe_allow_html=True)
        this_month_df = df[df["월"] == current_month]
        categories = this_month_df["분류"].tolist()
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=categories, y=this_month_df["제안 수"],
            name="제안 수", marker_color=COLOR_BASE,
            text=this_month_df["제안 수"], textposition="outside",
        ))
        fig.add_trace(go.Bar(
            x=categories, y=this_month_df["선정 수"],
            name="선정 수", marker_color=COLOR_HIGHLIGHT,
            text=this_month_df["선정 수"], textposition="outside",
        ))
        fig.update_layout(barmode="group", yaxis_title="건수", height=460)
        apply_chart_style(fig)
        st.plotly_chart(fig, use_container_width=True)

    # 월별 선정 수 및 금액 추이 (이중축)
    with col_b:
        st.markdown('<div class="section-title">월별 선정 수 및 금액(억원) 추이</div>', unsafe_allow_html=True)
        monthly = df.groupby("월").agg({"선정 수": "sum", "금액(억원)": "sum"}).reindex(month_order)

        fig2 = make_subplots(specs=[[{"secondary_y": True}]])
        fig2.add_trace(go.Bar(
            x=month_order, y=monthly["선정 수"],
            name="선정 수 (건)", marker_color=COLOR_BASE,
        ), secondary_y=False)
        fig2.add_trace(go.Scatter(
            x=month_order, y=monthly["금액(억원)"],
            name="금액 (억원)", mode="lines+markers",
            line=dict(color=COLOR_HIGHLIGHT, width=3),
            marker=dict(size=8),
        ), secondary_y=True)
        fig2.update_yaxes(title_text="선정 수 (건)", secondary_y=False)
        fig2.update_yaxes(title_text="금액 (억원)", secondary_y=True)
        fig2.update_layout(height=460)
        apply_chart_style(fig2, dual_axis=True)
        st.plotly_chart(fig2, use_container_width=True)

    footnote(
        "※ 데이터 출처: rawdata_2608.xlsx (Sheet2) · 분류: 국가R&D, 용역·사기업 포함 전체 합계 기준입니다.<br>"
        "※ 기준월은 데이터 내 최신 월로 자동 산정됩니다."
    )


# ------------------------------------------------------------------
# 메인 (사이드바 네비게이션)
# ------------------------------------------------------------------
def main():
    inject_css()

    st.sidebar.markdown(
        f"""
        <div style="padding: 10px 4px 24px 4px;">
            <div style="font-size:20px; font-weight:700; color:white;">연구전략기획팀</div>
            <div style="font-size:13px; color:#B7C4DE;">연구실적 대시보드</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    pages = {
        "이달 연구실적 개요": page_overview,
        "연도별 월별 실적 추이": page_yearly_trend,
        "산학협력 현황": page_collaboration,
    }

    selection = st.sidebar.radio("페이지 선택", list(pages.keys()), label_visibility="collapsed")
    pages[selection]()


if __name__ == "__main__":
    main()
