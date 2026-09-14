import streamlit as st
import pandas as pd
import plotly.express as px
from plotly.subplots import make_subplots
import plotly.graph_objects as go
from pathlib import Path

st.set_page_config(page_title="연구실적 대시보드", layout="wide")

DEFAULT_FILE = Path(__file__).parent / "rawdata_2608.xlsx"


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

    df_perf = pd.read_excel(xls, sheet_name=sheet_names[0])
    df_perf.columns = [str(c).strip() for c in df_perf.columns]
    df_perf["월"] = df_perf["월"].apply(month_to_int)
    df_perf["승인월"] = pd.to_datetime(df_perf["승인일자"], errors="coerce").dt.month

    collab_sheet = "Sheet2" if "Sheet2" in sheet_names else sheet_names[-1]
    df_collab = pd.read_excel(xls, sheet_name=collab_sheet)
    df_collab.columns = [str(c).strip() for c in df_collab.columns]
    df_collab["월_num"] = df_collab["월"].apply(month_to_int)

    return df_perf, df_collab


st.sidebar.title("연구실적 대시보드")
uploaded = st.sidebar.file_uploader("데이터 파일 업로드 (.xlsx)", type=["xlsx"])
file_to_use = uploaded if uploaded is not None else DEFAULT_FILE

if uploaded is None and not DEFAULT_FILE.exists():
    st.error("데이터 파일을 찾을 수 없습니다. 사이드바에서 엑셀 파일을 업로드해주세요.")
    st.stop()

df_perf, df_collab = load_data(file_to_use)

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

page = st.sidebar.radio("페이지 선택", ["이달 연구실적 개요", "산학협력 실적 개요"])

file_label = file_to_use.name if hasattr(file_to_use, "name") else Path(file_to_use).name
st.sidebar.markdown("---")
st.sidebar.caption(f"기준월: {selected_month}월 · 데이터 파일: {file_label}")


def month_delta(current, previous):
    if previous is None or previous == 0:
        return None
    return f"{(current - previous) / previous * 100:+.1f}%"


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
        fig.update_layout(margin=dict(t=20, l=10, r=10, b=10))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.subheader(f"소속별 실적 건수 ({selected_month}월)")
        by_dept = this_month_perf.groupby("소속(대)").size().sort_values(ascending=True)
        fig = px.bar(
            x=by_dept.values,
            y=by_dept.index,
            orientation="h",
            labels={"x": "실적 건수", "y": "소속(대)"},
        )
        fig.update_traces(marker_color="#1D9E75")
        fig.update_layout(margin=dict(t=20, l=10, r=10, b=10))
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader(f"{selected_month}월 실적 상세")
    display_cols = [c for c in ["성명", "소속(대)", "구분명칭", "연구분류", "참여구분", "일자", "승인일자", "평가점수"] if c in this_month_perf.columns]
    st.dataframe(this_month_perf[display_cols], use_container_width=True, hide_index=True)

else:
    st.title("산학협력 실적 개요")

    this_month_collab = df_collab[df_collab["월_num"] == selected_month]

    col1, col2 = st.columns(2)
    with col1:
        st.metric(f"{selected_month}월 제안 수", f"{int(this_month_collab['제안 수'].sum())}건")
    with col2:
        st.metric(f"{selected_month}월 선정 수", f"{int(this_month_collab['선정 수'].sum())}건")

    st.markdown("---")

    st.subheader(f"{selected_month}월 분류별 제안 수 / 선정 수")
    by_type = this_month_collab.groupby("분류")[["제안 수", "선정 수"]].sum().reset_index()
    fig = px.bar(by_type, x="분류", y=["제안 수", "선정 수"], barmode="group")
    fig.update_layout(margin=dict(t=20, l=10, r=10, b=10), legend_title_text="")
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
    fig.update_yaxes(title_text="선정 수", secondary_y=False)
    fig.update_yaxes(title_text="금액(억원)", secondary_y=True)
    fig.update_layout(margin=dict(t=20, l=10, r=10, b=10), legend_title_text="")
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader(f"{selected_month}월 산학협력 상세")
    st.dataframe(this_month_collab, use_container_width=True, hide_index=True)
