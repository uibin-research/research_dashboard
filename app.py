import re
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from pathlib import Path
from plotly.subplots import make_subplots

st.set_page_config(page_title="연구실적 대시보드", layout="wide")

DEFAULT_FILE = Path(__file__).parent / "rawdata_2608.xlsx"
MONTH_ORDER = ["3월", "4월", "5월", "6월", "7월", "8월", "9월", "10월", "11월", "12월", "1월", "2월"]
RESEARCH_MONTHS = [7, 8]

BG_COLOR = "#F5F7FA"
SIDEBAR_COLOR = "#0D2B5E"
HIGHLIGHT_COLOR = "#0057FF"
BASE_BLUE = "#93B8E0"
HIGHLIGHT_YEAR = "2026"

LIGHT_BLUE = "#DCEAF7"
DARK_BLUE = "#1D5FA8"
LIGHT_ORANGE = "#FDE0C2"
DARK_ORANGE = "#C2560A"

PERF_COLS = ["소속(대)", "구분명칭", "연구분류", "일자"]


def _hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i : i + 2], 16) for i in (0, 2, 4))


def _rgb_to_hex(rgb):
    return "#{:02X}{:02X}{:02X}".format(*[round(c) for c in rgb])


def color_gradient(n, light, dark):
    """light색에서 dark색으로 균등하게 보간한 색상 n개를 반환 (가장 오래된 순서가 light)"""
    if n <= 1:
        return [dark]
    c1, c2 = _hex_to_rgb(light), _hex_to_rgb(dark)
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


def find_month_sheets(sheet_names, suffix):
    """'8월승인', '7월미승인' 처럼 'N월<suffix>' 형태의 시트를 {월번호: 시트명}으로 반환"""
    pattern = re.compile(rf"^(\d{{1,2}})월\s*{suffix}$")
    found = {}
    for name in sheet_names:
        m = pattern.match(str(name).strip())
        if m:
            found[int(m.group(1))] = name
    return found


@st.cache_data
def load_data(file):
    xls = pd.ExcelFile(file)
    sheet_names = xls.sheet_names

    approved_sheets = find_month_sheets(sheet_names, "승인")
    unapproved_sheets = find_month_sheets(sheet_names, "미승인")
    collab_sheets = find_month_sheets(sheet_names, "산단")

    def read_perf(sheet_name):
        if sheet_name is None:
            return pd.DataFrame(columns=PERF_COLS + ["월"])
        df = pd.read_excel(xls, sheet_name=sheet_name)
        df.columns = [str(c).strip() for c in df.columns]
        df = df[[c for c in PERF_COLS if c in df.columns]].copy()
        df["월"] = pd.to_datetime(df["일자"], errors="coerce").dt.month
        return df

    approved_data = {m: read_perf(name) for m, name in approved_sheets.items()}
    unapproved_data = {m: read_perf(name) for m, name in unapproved_sheets.items()}

    def read_collab(sheet_name):
        if sheet_name is None:
            return None
        df = pd.read_excel(xls, sheet_name=sheet_name)
        df.columns = [str(c).strip() for c in df.columns]
        df["월_num"] = df["월"].apply(month_to_int)
        return df

    collab_data = {m: read_collab(name) for m, name in collab_sheets.items()}

    year_sheets = {}
    for name in sheet_names:
        if str(name).strip().isdigit():
            df_year = pd.read_excel(xls, sheet_name=name)
            df_year.columns = [str(c).strip() for c in df_year.columns]
            df_year = df_year.set_index(df_year.columns[0])
            year_sheets[str(name).strip()] = df_year

    return {
        "approved_data": approved_data,
        "unapproved_data": unapproved_data,
        "collab_data": collab_data,
        "year_sheets": year_sheets,
    }


def empty_perf_df():
    return pd.DataFrame(columns=PERF_COLS + ["월"])


def render_research_page(month, current_approved, prev_month, prev_approved, current_unapproved):
    st.title(f"{month}월 연구실적 개요")

    combined_current = pd.concat([current_approved, current_unapproved], ignore_index=True)
    this_month_combined = combined_current[combined_current["월"] == month]

    perf_count = len(current_approved[current_approved["월"] == month])
    approved_delta = len(current_approved) - len(prev_approved)

    col1, col2 = st.columns(2)
    with col1:
        st.metric(
            f"{month}월 실적 건수",
            f"{perf_count}건",
            help=f"{month}월승인 시트에서 일자가 {month}월인 건만 집계",
        )
    with col2:
        st.metric(
            f"{month}월 승인 건수 증감",
            f"{approved_delta:+d}건",
            help=f"{month}월승인 시트 행 수({len(current_approved)}) − {prev_month}월승인 시트 행 수({len(prev_approved)})",
        )

    st.markdown("---")

    st.subheader(f"{month}월 단과대학별 실적 건수")
    by_dept = this_month_combined.groupby("소속(대)").size()
    ordered_depts = [d for d in DEPT_ORDER if d in by_dept.index] + [
        d for d in by_dept.index if d not in DEPT_ORDER
    ]
    by_dept = by_dept.reindex(ordered_depts)
    fig = px.bar(x=by_dept.index, y=by_dept.values, labels={"x": "소속(대)", "y": "실적 건수"})
    fig.update_traces(
        marker_color="#1D9E75", texttemplate="%{y}", textposition="outside", textfont=dict(size=13)
    )
    fig.update_layout(
        margin=dict(t=20, l=10, r=10, b=10), xaxis_tickangle=-30, plot_bgcolor="white", paper_bgcolor="white"
    )
    apply_chart_style(fig, legend=False)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader(f"월별 실적 건수 추이 ({prev_month}월승인 vs {month}월승인)")
        fig = go.Figure()
        months_present = sorted(
            set(prev_approved["월"].dropna().astype(int)) | set(current_approved["월"].dropna().astype(int))
        )
        for label, df_snapshot, color, width in [
            (f"{prev_month}월승인", prev_approved, BASE_BLUE, 2),
            (f"{month}월승인", current_approved, HIGHLIGHT_COLOR, 4),
        ]:
            trend = df_snapshot.dropna(subset=["월"]).groupby("월").size().reindex(months_present, fill_value=0)
            fig.add_trace(
                go.Scatter(
                    x=[f"{m}월" for m in trend.index],
                    y=trend.values,
                    name=label,
                    mode="lines+markers+text",
                    line=dict(color=color, width=width),
                    marker=dict(size=9 if color == HIGHLIGHT_COLOR else 6),
                    text=[str(v) for v in trend.values],
                    textposition="top center",
                    textfont=dict(size=12, color=color),
                )
            )
        fig.update_layout(margin=dict(t=20, l=10, r=10, b=10), plot_bgcolor="white", paper_bgcolor="white")
        apply_chart_style(fig)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.subheader("월별 실적 건수 (승인 + 미승인)")
        combined_current2 = combined_current.copy()
        combined_current2["소스"] = "승인"
        combined_current2.loc[len(current_approved) :, "소스"] = "미승인"
        by_month_source = combined_current2.groupby(["월", "소스"]).size().unstack(fill_value=0)
        for col in ["승인", "미승인"]:
            if col not in by_month_source.columns:
                by_month_source[col] = 0
        by_month_source = by_month_source.reindex(sorted(by_month_source.index))

        fig = go.Figure()
        fig.add_trace(
            go.Bar(
                x=[f"{m}월" for m in by_month_source.index],
                y=by_month_source["승인"],
                name="승인",
                marker_color=HIGHLIGHT_COLOR,
                texttemplate="%{y}",
                textposition="inside",
                textfont=dict(size=12, color="white"),
            )
        )
        fig.add_trace(
            go.Bar(
                x=[f"{m}월" for m in by_month_source.index],
                y=by_month_source["미승인"],
                name="미승인",
                marker_color=BASE_BLUE,
                texttemplate="%{y}",
                textposition="inside",
                textfont=dict(size=12, color="white"),
            )
        )
        fig.update_layout(
            barmode="stack",
            margin=dict(t=20, l=10, r=10, b=10),
            plot_bgcolor="white",
            paper_bgcolor="white",
            legend_title_text="",
        )
        apply_chart_style(fig)
        st.plotly_chart(fig, use_container_width=True)

    render_footnote(
        f"※ {month}월 실적 건수는 {month}월승인 시트에서 일자가 {month}월인 건만 집계한 값입니다(아직 승인되지 않은 건은 제외). "
        f"승인 건수 증감은 {month}월승인 시트와 {prev_month}월승인 시트의 전체 행 수 차이입니다. "
        f"월별 실적 건수 추이는 {prev_month}월승인(연한 파랑)과 {month}월승인(진한 파랑) 두 시점의 승인 데이터를 비교하며, "
        "하단 막대는 승인(진한 파랑)·미승인(연한 파랑)을 함께 표시합니다."
    )


def render_collab_page(month, df_sheet):
    st.title(f"{month}월 산학협력 개요")

    if df_sheet is None or df_sheet.empty:
        st.warning(f"{month}월산단 시트를 찾을 수 없습니다.")
        return

    all_collab_months = sorted(df_sheet["월_num"].dropna().astype(int).unique())
    this_month_collab = df_sheet[df_sheet["월_num"] == month]

    col1, col2 = st.columns(2)
    with col1:
        st.metric(f"{month}월 제안 수", f"{int(this_month_collab['제안 수'].sum())}건")
    with col2:
        st.metric(f"{month}월 선정 수", f"{int(this_month_collab['선정 수'].sum())}건")

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader(f"{month}월 분류별 제안 수 대비 선정 수")
        by_type = this_month_collab.groupby("분류")[["제안 수", "선정 수"]].sum().reset_index()
        fig = px.bar(by_type, x="분류", y=["제안 수", "선정 수"], barmode="group")
        fig.update_traces(texttemplate="%{y}", textposition="outside", textfont=dict(size=12))
        fig.update_layout(
            margin=dict(t=20, l=10, r=10, b=10), legend_title_text="", plot_bgcolor="white", paper_bgcolor="white"
        )
        apply_chart_style(fig)
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.subheader("월별 선정 수 및 금액(억원) 추이")
        monthly = (
            df_sheet.dropna(subset=["월_num"])
            .groupby("월_num")[["선정 수", "금액(억원)"]]
            .sum()
            .reindex(all_collab_months, fill_value=0)
        )
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(
            go.Bar(
                x=[f"{m}월" for m in monthly.index],
                y=monthly["선정 수"],
                name="선정 수",
                marker_color="#7F77DD",
                texttemplate="%{y}",
                textposition="outside",
                textfont=dict(size=12),
            ),
            secondary_y=False,
        )
        fig.add_trace(
            go.Scatter(
                x=[f"{m}월" for m in monthly.index],
                y=monthly["금액(억원)"],
                name="금액(억원)",
                mode="lines+markers+text",
                line=dict(color="#D85A30"),
                text=[f"{v:.1f}" for v in monthly["금액(억원)"]],
                textposition="top center",
                textfont=dict(size=12, color="#D85A30"),
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
    st.subheader(f"{month}월 산학협력 상세")
    st.dataframe(this_month_collab, use_container_width=True, hide_index=True)

    render_footnote(f"※ {month}월산단 시트 데이터 기준이며, 금액 단위는 억원입니다.")


def render_year_trend_page(year_sheets):
    st.title("연도별 월별 연구실적 추이")

    if not year_sheets:
        st.warning("연도별 실적 시트(2023, 2024 ... 형태)를 찾을 수 없습니다.")
        return

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
    gradient = color_gradient(len(non_highlight_years), LIGHT_ORANGE, DARK_ORANGE)
    color_map = dict(zip(non_highlight_years, gradient))
    color_map[HIGHLIGHT_YEAR] = HIGHLIGHT_COLOR

    fig = px.line(plot_df, x="월", y="실적 건수", color="연도", markers=True, color_discrete_map=color_map)
    for trace in fig.data:
        values = [None if v is None or pd.isna(v) else v for v in trace.y]
        text = [("" if v is None else f"{v:.0f}") for v in values]
        if trace.name == HIGHLIGHT_YEAR:
            trace.line.width = 4
            trace.marker.size = 9
            trace.opacity = 1
            trace.mode = "lines+markers+text"
            trace.text = text
            trace.textposition = "top center"
            trace.textfont = dict(size=12, color=HIGHLIGHT_COLOR)
        else:
            trace.line.width = 2
            trace.marker.size = 6
            trace.opacity = 0.55
            trace.mode = "lines+markers+text"
            trace.text = text
            trace.textposition = "bottom center"
            trace.textfont = dict(size=10, color=color_map.get(trace.name, DARK_ORANGE))
    fig.update_layout(
        margin=dict(t=20, l=10, r=10, b=10), legend_title_text="연도", plot_bgcolor="white", paper_bgcolor="white"
    )
    apply_chart_style(fig)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader("연도별 월별 실적 표")
    st.dataframe(trend_df, use_container_width=True)

    render_footnote(
        f"※ {', '.join(sorted(year_sheets.keys()))} 시트 기준이며, {HIGHLIGHT_YEAR}년은 원색 파란색으로, "
        "나머지 연도는 오래될수록 옅고 최근일수록 짙은 주황색으로 표시됩니다. "
        "아직 데이터가 없는 미래 월은 빈 값으로 처리되어 선이 끊깁니다."
    )


st.sidebar.title("연구실적 대시보드")
uploaded = st.sidebar.file_uploader("데이터 파일 업로드 (.xlsx)", type=["xlsx"])
file_to_use = uploaded if uploaded is not None else DEFAULT_FILE

if uploaded is None and not DEFAULT_FILE.exists():
    st.error("데이터 파일을 찾을 수 없습니다. 사이드바에서 엑셀 파일을 업로드해주세요.")
    st.stop()

data = load_data(file_to_use)
approved_data = data["approved_data"]
unapproved_data = data["unapproved_data"]
collab_data = data["collab_data"]
year_sheets = data["year_sheets"]

PAGES = []
for m in RESEARCH_MONTHS:
    PAGES.append(f"{m}월 연구실적 개요")
    PAGES.append(f"{m}월 산학협력 개요")
PAGES.append("연도별 월별 연구실적 추이")

page = st.sidebar.radio("페이지 선택", PAGES)

file_label = file_to_use.name if hasattr(file_to_use, "name") else Path(file_to_use).name
st.sidebar.markdown("---")
st.sidebar.caption(f"데이터 파일: {file_label}")

matched_research = re.match(r"^(\d{1,2})월 연구실적 개요$", page)
matched_collab = re.match(r"^(\d{1,2})월 산학협력 개요$", page)

if matched_research:
    month = int(matched_research.group(1))
    render_research_page(
        month=month,
        current_approved=approved_data.get(month, empty_perf_df()),
        prev_month=month - 1,
        prev_approved=approved_data.get(month - 1, empty_perf_df()),
        current_unapproved=unapproved_data.get(month, empty_perf_df()),
    )
elif matched_collab:
    month = int(matched_collab.group(1))
    render_collab_page(month, collab_data.get(month))
else:
    render_year_trend_page(year_sheets)
