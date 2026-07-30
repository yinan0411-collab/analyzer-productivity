from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from analysis_engine import (
    AnalysisParams,
    build_person_productivity,
    build_task_table,
    demand_summary,
    order_level_actual,
    prepare_exam,
    prepare_pick,
    read_excel,
)

st.set_page_config(page_title="LAX2 出库产能分析", page_icon="📦", layout="wide")

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.4rem; padding-bottom: 2rem;}
    [data-testid="stMetricValue"] {font-size: 1.65rem;}
    .small-note {color:#5f6368; font-size:0.9rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📦 LAX2 出库产能与订单结构分析（初版）")
st.caption("当前先跑通：考核单 → 拣货结果 → 2B/2C与高低层结构 → 实际操作人效。出勤人数、排班工时和其他出库环节将在下一版接入。")

with st.sidebar:
    st.header("上传数据")
    exam_files = st.file_uploader(
        "① 实际需要生产的单（考核单，可一次多选一周文件）",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
    )
    pick_files = st.file_uploader(
        "② 拣货实际完成结果（可一次多选一周文件）",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
    )

    st.header("分析规则")
    two_b_keyword = st.text_input("2B识别关键词", value="2B")
    oversized_threshold = st.number_input("非2B超大异常单阈值（件/订单）", min_value=1, value=500, step=50)
    high_area_start = st.number_input("高层区域起始 A", min_value=0, value=1, step=1)
    high_area_end = st.number_input("高层区域结束 A", min_value=0, value=36, step=1)
    high_level_start = st.number_input("高层起始层级 L", min_value=1, value=3, step=1)
    max_task_minutes = st.number_input("有效任务最长时长（分钟）", min_value=5, value=240, step=15)

    st.info(
        f"高层默认规则：A{int(high_area_start):02d}–A{int(high_area_end):02d} 且 L{int(high_level_start)}及以上。其余全部计为地面拣选。"
    )

params = AnalysisParams(
    two_b_keyword=two_b_keyword,
    oversized_order_threshold=int(oversized_threshold),
    high_area_start=int(high_area_start),
    high_area_end=int(high_area_end),
    high_level_start=int(high_level_start),
    max_valid_task_minutes=int(max_task_minutes),
)

if not exam_files or not pick_files:
    st.subheader("这一版可以回答什么？")
    c1, c2, c3 = st.columns(3)
    c1.markdown("**当天需求**\n\n到期订单/件数、开班前已完成、开班时剩余、是否超过生产结束时间。")
    c2.markdown("**订单难度结构**\n\n2B、2C、超大异常单，以及高层/地面件数的交叉占比。")
    c3.markdown("**拣货操作人效**\n\n按员工合并重叠任务时间，计算件数、人效、高层占比和2B占比。")
    st.warning("请在左侧上传两张Excel表。")
    st.stop()

@st.cache_data(show_spinner=False)
def load_and_prepare(exam_payloads: tuple[bytes, ...], pick_payloads: tuple[bytes, ...], params_dict: dict):
    import io

    p = AnalysisParams(**params_dict)
    exam_raw = pd.concat([read_excel(io.BytesIO(data)) for data in exam_payloads], ignore_index=True)
    pick_raw = pd.concat([read_excel(io.BytesIO(data)) for data in pick_payloads], ignore_index=True)
    exam = prepare_exam(exam_raw, p)
    pick = prepare_pick(pick_raw, exam, p)
    return exam, pick

with st.spinner("正在读取并关联订单与拣货数据……"):
    try:
        exam, pick = load_and_prepare(
            tuple(file.getvalue() for file in exam_files),
            tuple(file.getvalue() for file in pick_files),
            params.__dict__,
        )
    except Exception as exc:
        st.error(f"数据读取失败：{exc}")
        st.stop()

actual_dates = sorted(d for d in pick["实际完成日期"].dropna().unique())
due_dates = sorted(d for d in exam["到期日期"].dropna().unique())
if not actual_dates:
    st.error("拣货结果中没有可识别的‘拣货完成时间’。")
    st.stop()

filter_col1, filter_col2 = st.columns(2)
with filter_col1:
    analysis_date = st.selectbox("实际拣货分析日期", actual_dates, index=len(actual_dates) - 1, format_func=lambda x: str(x))
with filter_col2:
    default_due_index = due_dates.index(analysis_date) if analysis_date in due_dates else len(due_dates) - 1
    demand_date = st.selectbox("应生产需求日期", due_dates, index=max(default_due_index, 0), format_func=lambda x: str(x))

pick_day = pick[pick["实际完成日期"].eq(analysis_date)].copy()
orders_day = order_level_actual(pick_day)
tasks_day = build_task_table(pick_day, params)
people_day = build_person_productivity(pick_day, tasks_day)
demand = demand_summary(exam, demand_date)

actual_units = float(pick_day["实际拣货量"].sum())
actual_orders = int(pick_day["订单号"].nunique())
actual_tasks = int(pick_day["任务单号"].nunique())
operators = int(pick_day["人员键"].nunique())
valid_active_hours = float(people_day["有效操作工时"].sum()) if not people_day.empty else 0
operation_uph = actual_units / valid_active_hours if valid_active_hours > 0 else 0
high_units = float(pick_day["高层件数"].sum())
high_share = high_units / actual_units if actual_units else 0
b2_units = float(pick_day.loc[pick_day["订单类型"].eq("2B"), "实际拣货量"].sum())
b2_share = b2_units / actual_units if actual_units else 0

st.divider()
k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("实际拣货件数", f"{actual_units:,.0f}")
k2.metric("实际订单数", f"{actual_orders:,}")
k3.metric("操作人数", f"{operators:,}")
k4.metric("有效操作工时", f"{valid_active_hours:,.1f}")
k5.metric("操作人效", f"{operation_uph:,.1f} 件/时")
k6.metric("高层 / 2B件数占比", f"{high_share:.1%} / {b2_share:.1%}")

if operators:
    st.caption("注意：目前显示的是系统任务时间口径的‘操作人效’，不是包含等待、培训和辅助工作的完整出勤人效。")

tab_overview, tab_structure, tab_people, tab_exceptions, tab_export = st.tabs(
    ["当日概览", "订单与储位结构", "员工操作人效", "异常与数据质量", "导出结果"]
)

with tab_overview:
    st.subheader(f"{demand_date} 应生产需求")
    d1, d2, d3, d4, d5 = st.columns(5)
    d1.metric("到期订单", f"{demand['到期订单数']:,.0f}")
    d2.metric("到期件数", f"{demand['到期件数']:,.0f}")
    d3.metric("开班前已拣", f"{demand['开班前已拣件数']:,.0f}")
    d4.metric("开班时剩余", f"{demand['开班时剩余件数']:,.0f}")
    d5.metric("超过生产结束时间/未完成", f"{demand['超过生产结束时间件数']:,.0f}")

    left, right = st.columns(2)
    with left:
        st.markdown("#### 当天实际在做哪一天的订单")
        timing = (
            pick_day.groupby("生产日期关系", dropna=False)["实际拣货量"]
            .sum()
            .reset_index(name="件数")
            .sort_values("件数", ascending=False)
        )
        fig = px.bar(timing, x="生产日期关系", y="件数", text_auto=",.0f")
        fig.update_layout(xaxis_title="", yaxis_title="实际拣货件数", showlegend=False)
        st.plotly_chart(fig, use_container_width=True)
        timing["件数占比"] = timing["件数"] / timing["件数"].sum() if timing["件数"].sum() else 0
        st.dataframe(timing, use_container_width=True, hide_index=True, column_config={"件数占比": st.column_config.NumberColumn(format="percent")})

    with right:
        st.markdown("#### 到期需求完成结构")
        due_status = pd.DataFrame(
            {
                "状态": ["开班前已完成", "当日仍需处理", "超过生产结束时间/未完成"],
                "件数": [
                    demand["开班前已拣件数"],
                    demand["开班时剩余件数"],
                    demand["超过生产结束时间件数"],
                ],
            }
        )
        fig = px.bar(due_status, x="状态", y="件数", text_auto=",.0f")
        fig.update_layout(xaxis_title="", yaxis_title="件数", showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

with tab_structure:
    st.subheader("2B / 2C 与高层 / 地面交叉结构")
    cross = pd.pivot_table(
        pick_day,
        index="订单类型",
        columns="拣选层级",
        values="实际拣货量",
        aggfunc="sum",
        fill_value=0,
        margins=True,
        margins_name="合计",
    )
    st.dataframe(cross, use_container_width=True)

    chart_data = (
        pick_day.groupby(["订单类型", "拣选层级"], dropna=False)["实际拣货量"]
        .sum()
        .reset_index(name="件数")
    )
    fig = px.bar(chart_data, x="订单类型", y="件数", color="拣选层级", barmode="stack", text_auto=",.0f")
    fig.update_layout(xaxis_title="", yaxis_title="实际拣货件数", legend_title="")
    st.plotly_chart(fig, use_container_width=True)

    left, right = st.columns(2)
    with left:
        st.markdown("#### 订单储位结构")
        order_loc = (
            orders_day.groupby("订单储位结构", dropna=False)
            .agg(订单数=("订单号", "nunique"), 件数=("实际件数", "sum"))
            .reset_index()
        )
        st.dataframe(order_loc, use_container_width=True, hide_index=True)
    with right:
        st.markdown("#### 订单类型结构")
        order_type = (
            orders_day.groupby("订单类型", dropna=False)
            .agg(订单数=("订单号", "nunique"), 件数=("实际件数", "sum"))
            .reset_index()
        )
        order_type["件数占比"] = order_type["件数"] / order_type["件数"].sum() if order_type["件数"].sum() else 0
        st.dataframe(order_type, use_container_width=True, hide_index=True, column_config={"件数占比": st.column_config.NumberColumn(format="percent")})

    st.markdown("#### 高层占比对员工操作人效的影响（当天横截面）")
    scatter_data = people_day.dropna(subset=["件数人效"]).copy()
    if len(scatter_data) >= 2:
        fig = px.scatter(
            scatter_data,
            x="高层件数占比",
            y="件数人效",
            size="实际件数",
            color="2B件数占比",
            hover_name="姓名",
            hover_data=["工号", "实际件数", "有效操作工时", "订单数"],
        )
        fig.update_layout(xaxis_tickformat=".0%", coloraxis_colorbar_title="2B占比", xaxis_title="高层件数占比", yaxis_title="件数/有效操作小时")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("当前日期有效人员数据不足，暂时不能绘制关系图。")

with tab_people:
    st.subheader("员工实际操作产出")
    st.caption("有效操作工时会先按任务单计算开始与结束时间，再合并同一员工的重叠区间；超过左侧任务时长阈值的任务不计入工时。")
    display_cols = [
        "工号", "姓名", "实际件数", "订单数", "任务数", "有效操作工时", "件数人效",
        "高层件数", "高层件数占比", "二B件数", "2B件数占比", "二C件数", "2C件数占比", "超大异常件数",
    ]
    st.dataframe(
        people_day[display_cols],
        use_container_width=True,
        hide_index=True,
        column_config={
            "有效操作工时": st.column_config.NumberColumn(format="%.2f"),
            "件数人效": st.column_config.NumberColumn(format="%.1f"),
            "高层件数占比": st.column_config.NumberColumn(format="percent"),
            "2B件数占比": st.column_config.NumberColumn(format="percent"),
            "2C件数占比": st.column_config.NumberColumn(format="percent"),
        },
    )

    top_n = st.slider("图表显示人数", min_value=5, max_value=max(5, min(50, len(people_day))), value=min(15, max(5, len(people_day)))) if len(people_day) else 5
    chart_people = people_day.head(top_n).sort_values("件数人效")
    if not chart_people.empty:
        fig = px.bar(chart_people, x="件数人效", y="姓名", orientation="h", hover_data=["实际件数", "有效操作工时", "高层件数占比", "2B件数占比"])
        fig.update_layout(xaxis_title="件数/有效操作小时", yaxis_title="")
        st.plotly_chart(fig, use_container_width=True)

with tab_exceptions:
    st.subheader("需要人工确认的异常")
    e1, e2, e3, e4 = st.columns(4)
    oversized = exam[(~exam["是否取消订单"]) & exam["订单类型"].eq("超大异常单")]
    unmatched = pick_day[pick_day["订单匹配状态"].eq("未匹配考核单")]
    unparsed = pick_day[~pick_day["储位可解析"]]
    abnormal_tasks = tasks_day[~tasks_day["任务时长有效"]]
    e1.metric("非2B超大异常订单", f"{oversized['京东订单号'].nunique():,}")
    e2.metric("未匹配考核单的拣货订单", f"{unmatched['订单号'].nunique():,}")
    e3.metric("无法解析储位明细", f"{len(unparsed):,}")
    e4.metric("异常任务时长", f"{len(abnormal_tasks):,}")

    if not oversized.empty:
        st.markdown("#### 非2B但超过阈值的超大异常单")
        st.dataframe(oversized[["京东订单号", "件数", "SPB名称", "生产结束时间", "货主名称"] if "货主名称" in oversized.columns else ["京东订单号", "件数", "SPB名称", "生产结束时间"]], use_container_width=True, hide_index=True)
    if not unmatched.empty:
        st.markdown("#### 拣货结果未匹配到考核单")
        st.dataframe(unmatched[["订单号", "实际拣货量", "拣货完成时间", "姓名"]].drop_duplicates(), use_container_width=True, hide_index=True)
    if not unparsed.empty:
        st.markdown("#### 无法解析储位（仍按地面拣选计入）")
        st.dataframe(unparsed[["订单号", "储位", "实际拣货量", "姓名"]].drop_duplicates().head(500), use_container_width=True, hide_index=True)
    if not abnormal_tasks.empty:
        st.markdown("#### 异常任务时长")
        st.dataframe(abnormal_tasks[["工号", "姓名", "任务单号", "任务开始", "任务结束", "任务分钟", "异常原因"]], use_container_width=True, hide_index=True)

with tab_export:
    st.subheader("下载分析结果")
    st.download_button(
        "下载员工人效 CSV",
        people_day.to_csv(index=False, encoding="utf-8-sig"),
        file_name=f"拣货员工人效_{analysis_date}.csv",
        mime="text/csv",
    )
    st.download_button(
        "下载订单级分析 CSV",
        orders_day.to_csv(index=False, encoding="utf-8-sig"),
        file_name=f"拣货订单分析_{analysis_date}.csv",
        mime="text/csv",
    )
    st.download_button(
        "下载任务级分析 CSV",
        tasks_day.to_csv(index=False, encoding="utf-8-sig"),
        file_name=f"拣货任务分析_{analysis_date}.csv",
        mime="text/csv",
    )
    st.download_button(
        "下载明细关联结果 CSV",
        pick_day.to_csv(index=False, encoding="utf-8-sig"),
        file_name=f"拣货关联明细_{analysis_date}.csv",
        mime="text/csv",
    )

st.divider()
st.caption("下一步：加入人员名单与实际出勤工时后，可同时计算操作人效与排班人效；再接入复核、Babylist/Mix/2B打包、机区拣货、退供、道口和5S。")
