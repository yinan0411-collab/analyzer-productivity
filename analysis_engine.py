from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable
import pandas as pd


@dataclass(frozen=True)
class Thresholds:
    early_start: int = 30
    late_start: int = 30
    early_leave: int = 30
    late_leave: int = 30


COLUMN_ALIASES = {
    "姓名": ["姓名", "员工姓名"],
    "用户编码": ["用户编码", "员工编码", "账号", "User ID"],
    "日期": ["日期", "考勤日期"],
    "考勤组": ["考勤组", "考情组", "考勤组名称"],
    "班休": ["班休", "是否上班"],
    "班次名称": ["班次名称", "排班班次"],
    "计划上班时间": ["计划上班时间", "排班上班时间"],
    "实际上班时间": ["实际上班时间", "实际签到时间", "上班打卡时间"],
    "计划下班时间": ["计划下班时间", "排班下班时间"],
    "实际下班时间": ["实际下班时间", "实际签退时间", "下班打卡时间"],
    "排班时长(时)": ["排班时长(时)", "排班时长"],
    "打卡时长(时)": ["打卡时长(时)", "打卡时长"],
    "核算工时(时)": ["核算工时(时)", "核算工时"],
    "一级供应商": ["一级供应商", "供应商", "劳务公司"],
    "部门": ["部门"],
    "岗位": ["岗位"],
    "异常状态": ["异常状态"],
    "修改时间": ["修改时间"],
}


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """清理列名，并把常见别名统一为系统使用的标准列名。"""
    result = df.copy()
    result.columns = [
        str(c).replace("\n", "").replace("\r", "").strip()
        for c in result.columns
    ]

    rename_map: dict[str, str] = {}
    existing = set(result.columns)
    for standard_name, aliases in COLUMN_ALIASES.items():
        if standard_name in existing:
            continue
        for alias in aliases:
            if alias in existing:
                rename_map[alias] = standard_name
                break

    if rename_map:
        result = result.rename(columns=rename_map)
    return result


def validate_columns(df: pd.DataFrame) -> list[str]:
    required = [
        "姓名",
        "用户编码",
        "日期",
        "考勤组",
        "计划上班时间",
        "实际上班时间",
        "计划下班时间",
        "实际下班时间",
    ]
    return [column for column in required if column not in df.columns]


def prepare_data(df: pd.DataFrame, deduplicate: bool = True) -> pd.DataFrame:
    result = normalize_columns(df)

    text_columns = [
        "姓名", "用户编码", "考勤组", "班休", "班次名称",
        "一级供应商", "部门", "岗位", "异常状态",
    ]
    for column in text_columns:
        if column not in result.columns:
            result[column] = ""
        result[column] = result[column].fillna("").astype(str).str.strip()

    datetime_columns = [
        "计划上班时间", "实际上班时间",
        "计划下班时间", "实际下班时间", "修改时间",
    ]
    for column in datetime_columns:
        if column not in result.columns:
            result[column] = pd.NaT
        result[column] = pd.to_datetime(result[column], errors="coerce")

    result["日期"] = pd.to_datetime(result["日期"], errors="coerce").dt.date

    numeric_columns = ["排班时长(时)", "打卡时长(时)", "核算工时(时)"]
    for column in numeric_columns:
        if column not in result.columns:
            result[column] = 0.0
        result[column] = pd.to_numeric(result[column], errors="coerce").fillna(0.0)

    if deduplicate:
        key_columns = [
            column for column in ["用户编码", "日期", "考勤组"]
            if column in result.columns
        ]
        if key_columns:
            if result["修改时间"].notna().any():
                result = result.sort_values("修改时间")
            result = result.drop_duplicates(subset=key_columns, keep="last")

    return result.reset_index(drop=True)


def _minutes(actual: pd.Timestamp, planned: pd.Timestamp):
    if pd.isna(actual) or pd.isna(planned):
        return None
    return round((actual - planned).total_seconds() / 60, 1)


def analyze_attendance(
    df: pd.DataFrame,
    thresholds: Thresholds,
    flag_missing_punches: bool = True,
    flag_rest_day_punches: bool = True,
    flag_punch_without_schedule: bool = True,
) -> pd.DataFrame:
    """返回仅包含问题员工/日期的异常明细。每个员工每天保留一行。"""
    data = prepare_data(df, deduplicate=False)
    output_rows: list[dict] = []

    for _, row in data.iterrows():
        planned_start = row["计划上班时间"]
        actual_start = row["实际上班时间"]
        planned_end = row["计划下班时间"]
        actual_end = row["实际下班时间"]

        start_diff = _minutes(actual_start, planned_start)
        end_diff = _minutes(actual_end, planned_end)

        rest_text = f'{row.get("班休", "")} {row.get("班次名称", "")}'
        is_rest_day = (
            str(row.get("班休", "")).strip() == "休"
            or "休息" in rest_text
        )
        has_actual_punch = (
            pd.notna(actual_start)
            or pd.notna(actual_end)
            or float(row.get("打卡时长(时)", 0) or 0) > 0
        )
        has_planned_shift = pd.notna(planned_start) or pd.notna(planned_end)

        issue_types: list[str] = []
        issue_details: list[str] = []

        if is_rest_day:
            if flag_rest_day_punches and has_actual_punch:
                issue_types.append("休息日打卡")
                issue_details.append("排班为休息，但存在实际打卡或打卡工时")
        else:
            if not has_planned_shift:
                if flag_punch_without_schedule and has_actual_punch:
                    issue_types.append("有打卡无排班")
                    issue_details.append("存在实际打卡，但未配置计划班次")
            else:
                no_start = pd.isna(actual_start)
                no_end = pd.isna(actual_end)

                if flag_missing_punches and no_start and no_end:
                    issue_types.append("整班无打卡")
                    issue_details.append("有计划班次，但上下班均无实际打卡")
                else:
                    if flag_missing_punches and pd.notna(planned_start) and no_start:
                        issue_types.append("缺上班卡")
                        issue_details.append("有计划上班时间，但无实际上班打卡")
                    if flag_missing_punches and pd.notna(planned_end) and no_end:
                        issue_types.append("缺下班卡")
                        issue_details.append("有计划下班时间，但无实际下班打卡")

                if start_diff is not None:
                    if start_diff > thresholds.late_start:
                        issue_types.append("迟到超阈值")
                        issue_details.append(f"实际上班晚于计划 {start_diff:g} 分钟")
                    elif start_diff < -thresholds.early_start:
                        issue_types.append("提前上班超阈值")
                        issue_details.append(f"实际上班早于计划 {abs(start_diff):g} 分钟")

                if end_diff is not None:
                    if end_diff < -thresholds.early_leave:
                        issue_types.append("早退超阈值")
                        issue_details.append(f"实际下班早于计划 {abs(end_diff):g} 分钟")
                    elif end_diff > thresholds.late_leave:
                        issue_types.append("晚下班超阈值")
                        issue_details.append(f"实际下班晚于计划 {end_diff:g} 分钟")

        if not issue_types:
            continue

        output_rows.append(
            {
                "日期": row["日期"],
                "考勤组": row.get("考勤组", ""),
                "用户编码": row.get("用户编码", ""),
                "姓名": row.get("姓名", ""),
                "一级供应商": row.get("一级供应商", ""),
                "部门": row.get("部门", ""),
                "岗位": row.get("岗位", ""),
                "班次名称": row.get("班次名称", ""),
                "班休": row.get("班休", ""),
                "计划上班时间": planned_start,
                "实际上班时间": actual_start,
                "上班偏差(分钟)": start_diff,
                "计划下班时间": planned_end,
                "实际下班时间": actual_end,
                "下班偏差(分钟)": end_diff,
                "排班时长(时)": row.get("排班时长(时)", 0),
                "打卡时长(时)": row.get("打卡时长(时)", 0),
                "核算工时(时)": row.get("核算工时(时)", 0),
                "原系统异常状态": row.get("异常状态", ""),
                "异常类型": "；".join(issue_types),
                "异常说明": "；".join(issue_details),
                "来源文件": row.get("来源文件", ""),
            }
        )

    columns = [
        "日期", "考勤组", "用户编码", "姓名", "一级供应商",
        "部门", "岗位", "班次名称", "班休",
        "计划上班时间", "实际上班时间", "上班偏差(分钟)",
        "计划下班时间", "实际下班时间", "下班偏差(分钟)",
        "排班时长(时)", "打卡时长(时)", "核算工时(时)",
        "原系统异常状态", "异常类型", "异常说明", "来源文件",
    ]
    result = pd.DataFrame(output_rows, columns=columns)

    if result.empty:
        return result

    return result.sort_values(
        ["日期", "考勤组", "姓名", "用户编码"],
        na_position="last",
    ).reset_index(drop=True)


def build_summary(result: pd.DataFrame) -> pd.DataFrame:
    if result.empty:
        return pd.DataFrame(
            columns=["日期", "考勤组", "异常人次", "异常员工数"]
        )

    return (
        result.groupby(["日期", "考勤组"], dropna=False)
        .agg(
            异常人次=("用户编码", "size"),
            异常员工数=("用户编码", "nunique"),
        )
        .reset_index()
        .sort_values(["日期", "异常人次", "考勤组"], ascending=[True, False, True])
    )
