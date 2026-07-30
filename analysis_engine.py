from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from typing import Iterable

import numpy as np
import pandas as pd


@dataclass(frozen=True)
class AnalysisParams:
    two_b_keyword: str = "2B"
    oversized_order_threshold: int = 500
    high_area_start: int = 1
    high_area_end: int = 36
    high_level_start: int = 3
    max_valid_task_minutes: int = 240


EXAM_REQUIRED = [
    "京东订单号",
    "件数",
    "生产结束时间",
    "SPB名称",
]

PICK_REQUIRED = [
    "订单号",
    "任务单号",
    "储位",
    "实际拣货量",
    "任务领取时间",
    "拣货完成时间",
    "工号",
    "姓名",
]


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    return out


def require_columns(df: pd.DataFrame, required: Iterable[str], label: str) -> None:
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"{label}缺少字段：{', '.join(missing)}")


def read_excel(source) -> pd.DataFrame:
    """Read the first worksheet from a file path or Streamlit UploadedFile."""
    return normalize_columns(pd.read_excel(source, sheet_name=0, engine="openpyxl"))


def _contains_2b(value: object, keyword: str) -> bool:
    text = "" if pd.isna(value) else str(value)
    return keyword.casefold() in text.casefold()


def _cancelled_mask(series: pd.Series) -> pd.Series:
    normalized = series.fillna("").astype(str).str.strip().str.casefold()
    return normalized.isin({"是", "yes", "y", "true", "1", "已取消", "cancelled", "canceled"})


def prepare_exam(exam_raw: pd.DataFrame, params: AnalysisParams) -> pd.DataFrame:
    exam = normalize_columns(exam_raw)
    require_columns(exam, EXAM_REQUIRED, "考核单")

    exam = exam.copy()
    exam["京东订单号"] = exam["京东订单号"].astype(str).str.strip()
    exam["件数"] = pd.to_numeric(exam["件数"], errors="coerce").fillna(0)
    exam["生产结束时间"] = pd.to_datetime(exam["生产结束时间"], errors="coerce")

    if "拣货完成时间" in exam.columns:
        exam["考核单拣货完成时间"] = pd.to_datetime(exam["拣货完成时间"], errors="coerce")
    else:
        exam["考核单拣货完成时间"] = pd.NaT

    if "是否取消" in exam.columns:
        exam["是否取消订单"] = _cancelled_mask(exam["是否取消"])
    else:
        exam["是否取消订单"] = False

    is_2b = exam["SPB名称"].map(lambda x: _contains_2b(x, params.two_b_keyword))
    is_oversized = (~is_2b) & (exam["件数"] > params.oversized_order_threshold)
    exam["订单类型"] = np.select(
        [is_2b, is_oversized],
        ["2B", "超大异常单"],
        default="2C",
    )
    exam["到期日期"] = exam["生产结束时间"].dt.date

    # The assessment export should be one row per order. Keep a single record if duplicates appear.
    exam = exam.sort_values(["京东订单号", "生产结束时间"]).drop_duplicates("京东订单号", keep="last")
    return exam


def parse_location(location: object, params: AnalysisParams) -> tuple[float, float, str, bool]:
    text = "" if pd.isna(location) else str(location).strip().upper()
    area_match = re.search(r"(?:^|-)A(\d{1,3})(?:-|$)", text)
    if area_match is None:
        # Most records begin directly with Axx, so also support that format.
        area_match = re.search(r"^A(\d{1,3})(?:-|$)", text)
    level_match = re.search(r"(?:^|-)L(\d+)(?:-|$)", text)

    area = float(area_match.group(1)) if area_match else np.nan
    level = float(level_match.group(1)) if level_match else np.nan
    parsed = bool(area_match and level_match)
    is_high = bool(
        parsed
        and params.high_area_start <= area <= params.high_area_end
        and level >= params.high_level_start
    )
    # Per user's rule: all records not matching the high-level range are ground picking.
    pick_level = "高层拣选" if is_high else "地面拣选"
    return area, level, pick_level, parsed


def prepare_pick(pick_raw: pd.DataFrame, exam: pd.DataFrame, params: AnalysisParams) -> pd.DataFrame:
    pick = normalize_columns(pick_raw)
    require_columns(pick, PICK_REQUIRED, "拣货结果")

    pick = pick.copy()
    pick["订单号"] = pick["订单号"].astype(str).str.strip()
    pick["任务单号"] = pick["任务单号"].astype(str).str.strip()
    pick["实际拣货量"] = pd.to_numeric(pick["实际拣货量"], errors="coerce").fillna(0)
    pick["任务领取时间"] = pd.to_datetime(pick["任务领取时间"], errors="coerce")
    pick["拣货完成时间"] = pd.to_datetime(pick["拣货完成时间"], errors="coerce")
    pick["实际完成日期"] = pick["拣货完成时间"].dt.date

    loc = pick["储位"].map(lambda x: parse_location(x, params))
    pick[["A区编号", "L层级", "拣选层级", "储位可解析"]] = pd.DataFrame(loc.tolist(), index=pick.index)
    pick["高层件数"] = np.where(pick["拣选层级"].eq("高层拣选"), pick["实际拣货量"], 0)
    pick["地面件数"] = np.where(pick["拣选层级"].eq("地面拣选"), pick["实际拣货量"], 0)

    # The picking export also contains a production-end column. Preserve it for comparison,
    # but use the assessment export as the standard demand deadline after the join.
    if "生产结束时间" in pick.columns:
        pick = pick.rename(columns={"生产结束时间": "拣货结果生产结束时间"})
        pick["拣货结果生产结束时间"] = pd.to_datetime(pick["拣货结果生产结束时间"], errors="coerce")

    order_lookup_cols = [
        "京东订单号",
        "件数",
        "生产结束时间",
        "到期日期",
        "SPB名称",
        "订单类型",
        "是否取消订单",
    ]
    lookup = exam[order_lookup_cols].rename(
        columns={
            "京东订单号": "订单号",
            "件数": "订单总件数",
        }
    )
    pick = pick.merge(lookup, on="订单号", how="left", validate="many_to_one")
    pick["订单匹配状态"] = np.where(pick["生产结束时间"].notna(), "已匹配", "未匹配考核单")
    # Future orders may not exist in a single-day assessment file. Their deadline is still
    # available in the picking result, so use it as a fallback for timing analysis.
    if "拣货结果生产结束时间" in pick.columns:
        pick["生产结束时间"] = pick["生产结束时间"].combine_first(pick["拣货结果生产结束时间"])
    pick["到期日期"] = pd.to_datetime(pick["生产结束时间"], errors="coerce").dt.date
    pick["订单类型"] = pick["订单类型"].fillna("未匹配")

    actual_day = pd.to_datetime(pick["拣货完成时间"]).dt.normalize()
    due_day = pd.to_datetime(pick["生产结束时间"]).dt.normalize()
    pick["生产日期关系"] = np.select(
        [due_day.eq(actual_day), due_day.gt(actual_day), due_day.lt(actual_day)],
        ["当天做当天", "当天做未来", "逾期补做"],
        default="未知到期时间",
    )
    pick["是否按生产结束时间完成"] = np.where(
        pick["生产结束时间"].isna() | pick["拣货完成时间"].isna(),
        np.nan,
        pick["拣货完成时间"] <= pick["生产结束时间"],
    )

    pick["人员键"] = pick["工号"].fillna("").astype(str).str.strip()
    blank_person = pick["人员键"].eq("") | pick["人员键"].eq("nan")
    pick.loc[blank_person, "人员键"] = pick.loc[blank_person, "姓名"].fillna("未知人员").astype(str)
    return pick


def merge_intervals(intervals: list[tuple[pd.Timestamp, pd.Timestamp]]) -> float:
    valid = sorted((s, e) for s, e in intervals if pd.notna(s) and pd.notna(e) and e > s)
    if not valid:
        return 0.0
    merged: list[list[pd.Timestamp]] = [[valid[0][0], valid[0][1]]]
    for start, end in valid[1:]:
        if start <= merged[-1][1]:
            if end > merged[-1][1]:
                merged[-1][1] = end
        else:
            merged.append([start, end])
    seconds = sum((end - start).total_seconds() for start, end in merged)
    return seconds / 3600


def build_task_table(pick_filtered: pd.DataFrame, params: AnalysisParams) -> pd.DataFrame:
    tasks = (
        pick_filtered.groupby(["人员键", "工号", "姓名", "任务单号"], dropna=False)
        .agg(
            任务开始=("任务领取时间", "min"),
            任务结束=("拣货完成时间", "max"),
            任务件数=("实际拣货量", "sum"),
            订单数=("订单号", "nunique"),
            高层件数=("高层件数", "sum"),
        )
        .reset_index()
    )
    tasks["任务分钟"] = (tasks["任务结束"] - tasks["任务开始"]).dt.total_seconds() / 60
    tasks["任务时长有效"] = tasks["任务分钟"].gt(0) & tasks["任务分钟"].le(params.max_valid_task_minutes)
    tasks["异常原因"] = np.select(
        [tasks["任务分钟"].isna(), tasks["任务分钟"].le(0), tasks["任务分钟"].gt(params.max_valid_task_minutes)],
        ["缺少开始或结束时间", "结束时间不晚于开始时间", "超过任务时长阈值"],
        default="",
    )
    return tasks


def build_person_productivity(pick_filtered: pd.DataFrame, tasks: pd.DataFrame) -> pd.DataFrame:
    if pick_filtered.empty:
        return pd.DataFrame()

    people = (
        pick_filtered.groupby("人员键", dropna=False)
        .agg(
            工号=("工号", "first"),
            姓名=("姓名", "first"),
            实际件数=("实际拣货量", "sum"),
            订单数=("订单号", "nunique"),
            任务数=("任务单号", "nunique"),
            高层件数=("高层件数", "sum"),
            地面件数=("地面件数", "sum"),
            二B件数=("实际拣货量", lambda s: s[pick_filtered.loc[s.index, "订单类型"].eq("2B")].sum()),
            二C件数=("实际拣货量", lambda s: s[pick_filtered.loc[s.index, "订单类型"].eq("2C")].sum()),
            超大异常件数=("实际拣货量", lambda s: s[pick_filtered.loc[s.index, "订单类型"].eq("超大异常单")].sum()),
        )
        .reset_index()
    )

    valid_tasks = tasks[tasks["任务时长有效"]].copy()
    hours_by_person = {}
    for person, group in valid_tasks.groupby("人员键"):
        intervals = list(zip(group["任务开始"], group["任务结束"]))
        hours_by_person[person] = merge_intervals(intervals)

    people["有效操作工时"] = people["人员键"].map(hours_by_person).fillna(0)
    people["件数人效"] = np.where(
        people["有效操作工时"] > 0,
        people["实际件数"] / people["有效操作工时"],
        np.nan,
    )
    people["高层件数占比"] = np.where(people["实际件数"] > 0, people["高层件数"] / people["实际件数"], 0)
    people["2B件数占比"] = np.where(people["实际件数"] > 0, people["二B件数"] / people["实际件数"], 0)
    people["2C件数占比"] = np.where(people["实际件数"] > 0, people["二C件数"] / people["实际件数"], 0)
    return people.sort_values("件数人效", ascending=False, na_position="last")


def demand_summary(exam: pd.DataFrame, due_date: date) -> dict[str, float]:
    valid = exam[~exam["是否取消订单"]].copy()
    due = valid[valid["到期日期"].eq(due_date)].copy()
    if due.empty:
        return {
            "到期订单数": 0,
            "到期件数": 0,
            "开班前已拣件数": 0,
            "开班时剩余件数": 0,
            "到期日前完成件数": 0,
            "按时完成件数": 0,
            "超过生产结束时间件数": 0,
        }

    start = pd.Timestamp(due_date)
    completion = due["考核单拣货完成时间"]
    early_before_day = completion.notna() & completion.lt(start)
    by_due_time = completion.notna() & completion.le(due["生产结束时间"])
    late_or_open = completion.isna() | completion.gt(due["生产结束时间"])

    return {
        "到期订单数": float(due["京东订单号"].nunique()),
        "到期件数": float(due["件数"].sum()),
        "开班前已拣件数": float(due.loc[early_before_day, "件数"].sum()),
        "开班时剩余件数": float(due.loc[~early_before_day, "件数"].sum()),
        "到期日前完成件数": float(due.loc[early_before_day, "件数"].sum()),
        "按时完成件数": float(due.loc[by_due_time, "件数"].sum()),
        "超过生产结束时间件数": float(due.loc[late_or_open, "件数"].sum()),
    }


def order_level_actual(pick_filtered: pd.DataFrame) -> pd.DataFrame:
    if pick_filtered.empty:
        return pd.DataFrame()
    orders = (
        pick_filtered.groupby("订单号", dropna=False)
        .agg(
            订单类型=("订单类型", "first"),
            SPB名称=("SPB名称", "first"),
            生产结束时间=("生产结束时间", "first"),
            实际完成时间=("拣货完成时间", "max"),
            实际件数=("实际拣货量", "sum"),
            高层件数=("高层件数", "sum"),
            地面件数=("地面件数", "sum"),
            任务数=("任务单号", "nunique"),
        )
        .reset_index()
    )
    orders["高层件数占比"] = np.where(orders["实际件数"] > 0, orders["高层件数"] / orders["实际件数"], 0)
    orders["订单储位结构"] = np.select(
        [orders["高层件数"].eq(0), orders["地面件数"].eq(0)],
        ["纯地面订单", "纯高层订单"],
        default="高低层混合订单",
    )
    actual_day = orders["实际完成时间"].dt.normalize()
    due_day = orders["生产结束时间"].dt.normalize()
    orders["生产日期关系"] = np.select(
        [due_day.eq(actual_day), due_day.gt(actual_day), due_day.lt(actual_day)],
        ["当天做当天", "当天做未来", "逾期补做"],
        default="未知到期时间",
    )
    return orders
