import streamlit as st
import pandas as pd
import plotly.express as px
from plotly.subplots import make_subplots
import plotly.graph_objects as go
from pathlib import Path

st.set_page_config(page_title="연구실적 대시보드", layout="wide")

DEFAULT_FILE = Path(__file__).parent / "rawdata_2608.xlsx"
MONTH_ORDER = ["3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월", "1월", "2월"]

BG_COLOR = "#F5F7FA"
SIDEBAR_COLOR = "#0D2B5E"
HIGHLIGHT_YEAR = "2026"
HIGHLIGHT_COLOR = "#E8392A"

LIGHT_BLUE = "#DCEAF7"
DARK_BLUE = "#1D5FA8"


def _hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i : i + 2], 16) for i in (0, 2, 4))


def _rgb_to_hex(rgb):
    return "#{:02X}{:02X}{:02X}".format(*[round(c) for c in rgb])


def blue_gradient(n):
    """가장 오래된 연도가 가장 옅고, 최신일수록 짙어지는 파란색 n개를 반환"""
    if n <= 1:
        return [DARK_BLUE]
    c1, c2 = _hex_to_rgb(LIGHT_BLUE), _hex_to_rgb(DARK_BLUE)
    return [
        _rgb_to_hex(tuple(c1[i] + (c2[i] - c1[i]) * (step / (n - 1)) for i in range(3)))
        for step in range(n)
    ]


DEPT_ORDER = [
    "인문대학",
    "사회과학대학",
    "경영대학",
    "미디어·휴먼라이프대학",
    "인공지능·소프트웨어융합대학",
    "미래융합대학",
    "화학·생명과학대학",
    "스마트시스템공과대학",
    "반도체·ICT대학",
    "스포츠·예술대학",
    "건축대학",
    "방목기초교육대학",
    "(일반)대학원",
    "기록정보과학전문대학원",
    "통합치료대학원",
    "교육대학원",
    "스포츠학부(체육학전공, 스포츠산업학전공)",
]

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&display=swap');

    html, body, [class*="css"] {{
        font-family: 'Noto Sans KR', sans-serif;
    }}

    .stApp {{
        background-color: {BG_COLOR};
    }}

    section[data-testid="stSidebar"] {{
        background-color: {SIDEBAR_COLOR};
    }}
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] .stMarkdown {{
        color: #FFFFFF !important;
    }}
    section[data-testid="stSidebar"] div[data-baseweb="select"] * {{
        color: {SIDEBAR_COLOR} !important;
    }}

    .footnote {{
        margin-top: 28px;
        padding-top: 12px;
        border-top: 1px solid #D0D5DD;
        font-size: 12px;
        color: #8A8F98;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


def render_footnote(text):
    st.markdown(f"<div class='footnote'>{text}</div>", unsafe_allow_html=True)


def apply_chart_style(fig, legend=True):
    """차트의 범례·축 제목·눈금 레이블 글자 크기를 키워 가독성을 높인다."""
    fig.update_layout(
        font=dict(size=15),
        legend=dict(font=dict(size=15), title_font=dict(size=15)) if legend else fig.layout.legend,
        xaxis=dict(title_font=dict(size=15), tickfont=dict(size=13)),
        yaxis=dict(title_font=dict(size=15), tickfont=dict(size=13)),
    )
    return fig


def month_to_int(value):
    """'3월' 같은 문자열이나 3, 3.0 같은 숫자를 모두 정수 월로 변환"""
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return int(value)
    digits = "".join(ch for ch in str(value) if ch.isdigit())
    return int(digits) if digits else None


@st.cache_data
def load_data(file):
    xls = pd.ExcelFile(file)
    sheet_names = xls.sheet_names

    perf_sheet = "Sheet1" if "Sheet1" in sheet_names else sheet_names[0]
    df_perf = pd.read_excel(xls, sheet_name=perf_sheet)
    df_perf.columns = [str(c).strip() for c in df_perf.columns]
    df_perf["월"] = df_perf["월"].apply(month_to_int)
    df_perf["승인월"] = pd.to_datetime(df_perf["승인일자"], errors="coerce").dt.month

    collab_sheet = "Sheet2" if "Sheet2" in sheet_names else sheet_names[1]
    df_collab = pd.read_excel(xls, sheet_name=collab_sheet)
    df_collab.columns = [str(c).strip() for c in df_collab.columns]
    df_collab["월_num"] = df_collab["월"].apply(month_to_int)

    year_sheets = {}
    for name in sheet_names:
        if str(name).strip().isdigit():
            df_year = pd.read_excel(xls, sheet_name=name)
            df_year.columns = [str(c).strip() for c in df_year.columns]
            first_col = df_year.columns[0]
            df_year = df_year.set_index(first_col)
            year_sheets[str(name).strip()] = df_year

    return df_perf, df_collab, year_sheets


def month_delta(current, previous):
    if previous is None or previous == 0:
        return None
    return f"{(current - previous) / previous * 100:+.1f}%"


st.sidebar.title("연구실적 대시보드")
uploaded = st.sidebar.file_uploader("데이터 파일 업로드 (.xlsx)", type=["xlsx"])
file_to_use = uploaded if uploaded is not None else DEFAULT_FILE

if uploaded is None and not DEFAULT_FILE.exists():
    st.error("데이터 파일을 찾을 수 없습니다. 사이드바에서 엑셀 파일을 업로드해주세요.")
    st.stop()

df_perf, df_collab, year_sheets = load_data(file_to_use)

all_months = sorted(
    set(df_perf["월"].dropna().astype(int)) | set(df_collab["월_num"].dropna().astype(int))
)
if not all_months:
    st.error("월 정보를 읽을 수 없습니다. 데이터 형식을 확인해주세요.")
    st.stop()

default_month = max(all_months)
selected_month = st.sidebar.selectbox(
    "기준월 선택", all_months, index=all_months.index(default_month), format_func=lambda m: f"{m}월"
)

page = st.sidebar.radio(
    "페이지 선택",
    ["이달 연구실적 개요", "연도별 월별 실적 추이", "산학협력 실적 개요"],
)

file_label = file_to_use.name if hasattr(file_to_use, "name") else Path(file_to_use).name
st.sidebar.markdown("---")
st.sidebar.caption(f"기준월: {selected_month}월 · 데이터 파일: {file_label}")


if page == "이달 연구실적 개요":
    st.title("이달 연구실적 개요")

    prev_month_candidates = [m for m in all_months if m < selected_month]
    prev_month = max(prev_month_candidates) if prev_month_candidates else None

    this_month_perf = df_perf[df_perf["월"] == selected_month]
    prev_month_perf = df_perf[df_perf["월"] == prev_month] if prev_month else df_perf.iloc[0:0]

    this_month_count = len(this_month_perf)
    prev_month_count = len(prev_month_perf)

    this_month_approved = df_perf[df_perf["승인월"] == selected_month].shape[0]
    prev_month_approved = df_perf[df_perf["승인월"] == prev_month].shape[0] if prev_month else 0

    col1, col2 = st.columns(2)
    with col1:
        st.metric(
            f"{selected_month}월 실적 건수",
            f"{this_month_count}건",
            month_delta(this_month_count, prev_month_count) if prev_month else None,
        )
    with col2:
        st.metric(
            f"{selected_month}월 승인 건수",
            f"{this_month_approved}건",
            month_delta(this_month_approved, prev_month_approved) if prev_month else None,
        )

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("월별 실적 건수 추이")
        trend = df_perf.dropna(subset=["월"]).groupby("월").size().reindex(all_months, fill_value=0)
        trend_df = pd.DataFrame({"월": [f"{m}월" for m in trend.index], "실적 건수": trend.values})
        fig = px.line(trend_df, x="월", y="실적 건수", markers=True)
        fig.update_traces(line_color="#378ADD")
        fig.update_layout(margin=dict(t=20, l=10, r=10, b=10), plot_bgcolor="white", paper_bgcolor="white")
        apply_chart_style(fig, legend=False)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.subheader(f"소속별 실적 건수 ({selected_month}월)")
        by_dept = this_month_perf.groupby("소속(대)").size()
        ordered_depts = [d for d in DEPT_ORDER if d in by_dept.index] + [
            d for d in by_dept.index if d not in DEPT_ORDER
        ]
        by_dept = by_dept.reindex(ordered_depts)
        fig = px.bar(
            x=by_dept.index,
            y=by_dept.values,
            labels={"x": "소속(대)", "y": "실적 건수"},
        )
        fig.update_traces(marker_color="#1D9E75")
        fig.update_layout(
            margin=dict(t=20, l=10, r=10, b=10), xaxis_tickangle=-30, plot_bgcolor="white", paper_bgcolor="white"
        )
        apply_chart_style(fig, legend=False)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader(f"{selected_month}월 실적 상세")
    display_cols = [c for c in ["성명", "소속(대)", "구분명칭", "연구분류", "참여구분", "일자", "승인일자", "평가점수"] if c in this_month_perf.columns]
    st.dataframe(this_month_perf[display_cols], use_container_width=True, hide_index=True)

    render_footnote("※ Sheet1 데이터 기준이며, 승인 건수는 승인일자의 월을 기준으로 집계됩니다.")

elif page == "연도별 월별 실적 추이":
    st.title("연도별 월별 실적 추이")

    if not year_sheets:
        st.warning("연도별 실적 시트(2023, 2024 ... 형태)를 찾을 수 없습니다.")
    else:
        categories = sorted(set().union(*[set(df.index.dropna()) for df in year_sheets.values()]))
        category = st.selectbox("구분 선택", ["전체"] + categories)

        available_months = [m for m in MONTH_ORDER if any(m in df.columns for df in year_sheets.values())]

        trend_data = {}
        for year in sorted(year_sheets.keys()):
            df_year = year_sheets[year]
            cols = [m for m in available_months if m in df_year.columns]
            if category == "전체":
                series = df_year[cols].sum(axis=0, skipna=True, min_count=1)
            elif category in df_year.index:
                series = df_year.loc[category, cols]
            else:
                series = pd.Series([None] * len(cols), index=cols)
            trend_data[year] = series.reindex(available_months)

        trend_df = pd.DataFrame(trend_data).reindex(available_months)
        trend_df.index.name = "월"
        plot_df = trend_df.reset_index().melt(id_vars="월", var_name="연도", value_name="실적 건수")
        plot_df["월"] = pd.Categorical(plot_df["월"], categories=available_months, ordered=True)
        plot_df = plot_df.sort_values("월")

        non_highlight_years = sorted(y for y in trend_data.keys() if y != HIGHLIGHT_YEAR)
        gradient = blue_gradient(len(non_highlight_years))
        color_map = dict(zip(non_highlight_years, gradient))
        color_map[HIGHLIGHT_YEAR] = HIGHLIGHT_COLOR

        fig = px.line(plot_df, x="월", y="실적 건수", color="연도", markers=True, color_discrete_map=color_map)
        for trace in fig.data:
            if trace.name == HIGHLIGHT_YEAR:
                trace.line.width = 4
                trace.marker.size = 9
                trace.opacity = 1
            else:
                trace.line.width = 2
                trace.marker.size = 6
                trace.opacity = 0.55
        fig.update_layout(
            margin=dict(t=20, l=10, r=10, b=10), legend_title_text="연도", plot_bgcolor="white", paper_bgcolor="white"
        )
        apply_chart_style(fig)
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("---")
        st.subheader("연도별 월별 실적 표")
        st.dataframe(trend_df, use_container_width=True)

        render_footnote(
            f"※ 2023~{max(trend_data.keys())} 시트 기준이며, {HIGHLIGHT_YEAR}년은 빨간색으로 강조 표시됩니다. "
            "아직 데이터가 없는 미래 월은 빈 값으로 처리되어 선이 끊깁니다."
        )

else:
    st.title("산학협력 실적 개요")

    this_month_collab = df_collab[df_collab["월_num"] == selected_month]

    col1, col2 = st.columns(2)
    with col1:
        st.metric(f"{selected_month}월 제안 수", f"{int(this_month_collab['제안 수'].sum())}건")
    with col2:
        st.metric(f"{selected_month}월 선정 수", f"{int(this_month_collab['선정 수'].sum())}건")

    st.markdown("---")

    st.subheader(f"{selected_month}월 분류별 제안 수 대비 선정 수")
    by_type = this_month_collab.groupby("분류")[["제안 수", "선정 수"]].sum().reset_index()
    fig = px.bar(by_type, x="분류", y=["제안 수", "선정 수"], barmode="group")
    fig.update_layout(margin=dict(t=20, l=10, r=10, b=10), legend_title_text="", plot_bgcolor="white", paper_bgcolor="white")
    apply_chart_style(fig)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    st.subheader("월별 선정 수 및 금액(억원) 추이")
    monthly = (
        df_collab.dropna(subset=["월_num"])
        .groupby("월_num")[["선정 수", "금액(억원)"]]
        .sum()
        .reindex(all_months, fill_value=0)
    )
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_trace(
        go.Bar(x=[f"{m}월" for m in monthly.index], y=monthly["선정 수"], name="선정 수", marker_color="#7F77DD"),
        secondary_y=False,
    )
    fig.add_trace(
        go.Scatter(
            x=[f"{m}월" for m in monthly.index],
            y=monthly["금액(억원)"],
            name="금액(억원)",
            mode="lines+markers",
            line=dict(color="#D85A30"),
        ),
        secondary_y=True,
    )
    fig.update_yaxes(title_text="선정 수", secondary_y=False, title_font=dict(size=15), tickfont=dict(size=13))
    fig.update_yaxes(title_text="금액(억원)", secondary_y=True, title_font=dict(size=15), tickfont=dict(size=13))
    fig.update_xaxes(tickfont=dict(size=13))
    fig.update_layout(
        margin=dict(t=20, l=10, r=10, b=10),
        legend_title_text="",
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(size=15),
        legend=dict(font=dict(size=15)),
    )
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader(f"{selected_month}월 산학협력 상세")
    st.dataframe(this_month_collab, use_container_width=True, hide_index=True)

    render_footnote("※ Sheet2 데이터 기준이며, 금액 단위는 억원입니다.")
