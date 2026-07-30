from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from typing import Iterable

import numpy as np
import pandas as pd


@dataclass(frozen=True)
class AnalysisParams:
    oversized_order_threshold: int = 500
    high_area_start: int = 1
    high_area_end: int = 36
    high_level_start: int = 3
    max_valid_task_minutes: int = 240


EXAM_REQUIRED = ["京东订单号", "件数", "生产结束时间", "SPB名称"]
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

TRAFILEA_OWNER_PATTERNS: tuple[tuple[str, str], ...] = (
    ("Walmart", r"\bwalmart\b"),
    ("Belk", r"\bbelk\b"),
    ("Kohl's", r"\bkohl\s*['’`-]?\s*s\b"),
    ("Kohl's", r"\bkolh\s*['’`-]?\s*s\b"),
    ("Nordstrom", r"\bnordstrom\b"),
)
NORMAL_2B_SPB_KEYS = {"ww 2b self pick up", "ww pick first transport"}


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    if "货主名称" not in out.columns and "货主名" in out.columns:
        out = out.rename(columns={"货主名": "货主名称"})
    return out


def require_columns(df: pd.DataFrame, required: Iterable[str], label: str) -> None:
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"{label}缺少字段：{', '.join(missing)}")


def read_excel(source) -> pd.DataFrame:
    return normalize_columns(pd.read_excel(source, sheet_name=0, engine="openpyxl"))


def normalize_so(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text.upper()


def _normalized_key(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).casefold().replace("’", "'").replace("`", "'")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def _trafilea_owner(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).casefold().replace("’", "'").replace("`", "'")
    for canonical_name, pattern in TRAFILEA_OWNER_PATTERNS:
        if re.search(pattern, text, flags=re.IGNORECASE):
            return canonical_name
    return ""


def _is_normal_2b_spb(value: object) -> bool:
    return _normalized_key(value) in NORMAL_2B_SPB_KEYS


def _cancelled_mask(series: pd.Series) -> pd.Series:
    normalized = series.fillna("").astype(str).str.strip().str.casefold()
    return normalized.isin({"是", "yes", "y", "true", "1", "已取消", "cancelled", "canceled"})


def prepare_exam(exam_raw: pd.DataFrame, params: AnalysisParams) -> pd.DataFrame:
    exam = normalize_columns(exam_raw)
    require_columns(exam, EXAM_REQUIRED, "考核单")
    exam = exam.copy()

    if "货主名称" not in exam.columns:
        exam["货主名称"] = ""

    exam["SO"] = exam["京东订单号"].map(normalize_so)
    exam["京东订单号"] = exam["SO"]
    exam["件数"] = pd.to_numeric(exam["件数"], errors="coerce").fillna(0)
    exam["生产结束时间"] = pd.to_datetime(exam["生产结束时间"], errors="coerce")
    exam["考核单拣货完成时间"] = (
        pd.to_datetime(exam["拣货完成时间"], errors="coerce")
        if "拣货完成时间" in exam.columns
        else pd.NaT
    )
    exam["是否取消订单"] = _cancelled_mask(exam["是否取消"]) if "是否取消" in exam.columns else False

    exam["Trafilea货主"] = exam["货主名称"].map(_trafilea_owner)
    is_trafilea_2b = exam["Trafilea货主"].ne("")
    is_normal_2b = exam["SPB名称"].map(_is_normal_2b_spb)
    is_2b = is_trafilea_2b | is_normal_2b

    exam["2B子类型"] = np.select(
        [is_trafilea_2b, is_normal_2b],
        ["Trafilea-2B", "普通2B"],
        default="非2B",
    )
    exam["2B识别依据"] = np.select(
        [is_trafilea_2b & is_normal_2b, is_trafilea_2b, is_normal_2b],
        ["货主名称+SPB名称", "货主名称", "SPB名称"],
        default="非2B",
    )

    is_oversized = (~is_2b) & (exam["件数"] > params.oversized_order_threshold)
    exam["订单类型"] = np.select([is_2b, is_oversized], ["2B", "超大异常单"], default="2C")
    exam["订单分析类型"] = np.where(exam["订单类型"].eq("2B"), exam["2B子类型"], exam["订单类型"])
    exam["到期日期"] = exam["生产结束时间"].dt.date

    exam = exam[exam["SO"].ne("")].sort_values(["SO", "生产结束时间"]).drop_duplicates("SO", keep="last")
    return exam


def parse_location(location: object, params: AnalysisParams) -> tuple[float, float, str, bool]:
    text = "" if pd.isna(location) else str(location).strip().upper()
    area_match = re.search(r"(?:^|-)A(\d{1,3})(?:-|$)", text) or re.search(r"^A(\d{1,3})(?:-|$)", text)
    level_match = re.search(r"(?:^|-)L(\d+)(?:-|$)", text)
    area = float(area_match.group(1)) if area_match else np.nan
    level = float(level_match.group(1)) if level_match else np.nan
    parsed = bool(area_match and level_match)
    is_high = bool(
        parsed
        and params.high_area_start <= area <= params.high_area_end
        and level >= params.high_level_start
    )
    return area, level, "高层拣选" if is_high else "地面拣选", parsed


def prepare_pick(pick_raw: pd.DataFrame, exam: pd.DataFrame, params: AnalysisParams) -> pd.DataFrame:
    pick = normalize_columns(pick_raw)
    require_columns(pick, PICK_REQUIRED, "拣货结果")
    pick = pick.copy()

    pick["SO"] = pick["订单号"].map(normalize_so)
    pick["订单号"] = pick["SO"]
    pick["任务单号"] = pick["任务单号"].astype(str).str.strip()
    pick["实际拣货量"] = pd.to_numeric(pick["实际拣货量"], errors="coerce").fillna(0)
    pick["任务领取时间"] = pd.to_datetime(pick["任务领取时间"], errors="coerce")
    pick["拣货完成时间"] = pd.to_datetime(pick["拣货完成时间"], errors="coerce")
    pick["实际完成日期"] = pick["拣货完成时间"].dt.date

    if "货主名称" in pick.columns:
        pick = pick.rename(columns={"货主名称": "拣货结果货主名称"})
    else:
        pick["拣货结果货主名称"] = ""
    if "生产结束时间" in pick.columns:
        pick = pick.rename(columns={"生产结束时间": "拣货结果生产结束时间"})
        pick["拣货结果生产结束时间"] = pd.to_datetime(pick["拣货结果生产结束时间"], errors="coerce")
    else:
        pick["拣货结果生产结束时间"] = pd.NaT

    loc = pick["储位"].map(lambda x: parse_location(x, params))
    pick[["A区编号", "L层级", "拣选层级", "储位可解析"]] = pd.DataFrame(loc.tolist(), index=pick.index)
    pick["高层件数"] = np.where(pick["拣选层级"].eq("高层拣选"), pick["实际拣货量"], 0)
    pick["地面件数"] = np.where(pick["拣选层级"].eq("地面拣选"), pick["实际拣货量"], 0)

    lookup = exam[
        [
            "SO",
            "件数",
            "生产结束时间",
            "到期日期",
            "SPB名称",
            "货主名称",
            "Trafilea货主",
            "订单类型",
            "2B子类型",
            "2B识别依据",
            "订单分析类型",
            "是否取消订单",
        ]
    ].rename(
        columns={
            "件数": "考核单订单件数",
            "生产结束时间": "考核单生产结束时间",
            "到期日期": "考核单到期日期",
            "SPB名称": "考核单SPB名称",
            "货主名称": "考核单货主名称",
        }
    )
    pick = pick.merge(lookup, on="SO", how="left", validate="many_to_one")

    matched = pick["订单类型"].notna()
    pick["SO匹配状态"] = np.where(matched, "已匹配考核单", "未匹配考核单")
    pick["订单匹配状态"] = pick["SO匹配状态"]
    pick["SO分类可用"] = matched

    pick["生产结束时间"] = pick["考核单生产结束时间"].combine_first(pick["拣货结果生产结束时间"])
    pick["到期日期"] = pd.to_datetime(pick["生产结束时间"], errors="coerce").dt.date
    pick["SPB名称"] = pick["考核单SPB名称"].fillna("")
    pick["货主名称"] = pick["考核单货主名称"].fillna("")
    pick["订单类型"] = pick["订单类型"].fillna("未匹配SO")
    pick["2B子类型"] = pick["2B子类型"].fillna("未匹配SO")
    pick["2B识别依据"] = pick["2B识别依据"].fillna("未匹配SO")
    pick["订单分析类型"] = pick["订单分析类型"].fillna("未匹配SO")
    pick["Trafilea货主"] = pick["Trafilea货主"].fillna("")
    pick["是否取消订单"] = pick["是否取消订单"].fillna(False)

    pick["已匹配件数"] = np.where(matched, pick["实际拣货量"], 0)
    pick["未匹配件数"] = np.where(~matched, pick["实际拣货量"], 0)

    actual_day = pick["拣货完成时间"].dt.normalize()
    due_day = pick["生产结束时间"].dt.normalize()
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
            merged[-1][1] = max(merged[-1][1], end)
        else:
            merged.append([start, end])
    return sum((end - start).total_seconds() for start, end in merged) / 3600


def _single_or_mixed(series: pd.Series, mixed_label: str, empty_label: str = "未知") -> str:
    values = [str(v) for v in series.dropna().unique() if str(v).strip()]
    if not values:
        return empty_label
    if len(values) == 1:
        return values[0]
    return mixed_label


def build_task_table(pick_filtered: pd.DataFrame, params: AnalysisParams) -> pd.DataFrame:
    if pick_filtered.empty:
        return pd.DataFrame()
    work = pick_filtered.copy()
    work["Trafilea二B件数行"] = np.where(work["2B子类型"].eq("Trafilea-2B"), work["实际拣货量"], 0)
    work["普通二B件数行"] = np.where(work["2B子类型"].eq("普通2B"), work["实际拣货量"], 0)

    tasks = (
        work.groupby(["人员键", "工号", "姓名", "任务单号"], dropna=False)
        .agg(
            任务开始=("任务领取时间", "min"),
            任务结束=("拣货完成时间", "max"),
            任务件数=("实际拣货量", "sum"),
            订单数=("SO", "nunique"),
            高层件数=("高层件数", "sum"),
            地面件数=("地面件数", "sum"),
            已匹配件数=("已匹配件数", "sum"),
            未匹配件数=("未匹配件数", "sum"),
            Trafilea二B件数=("Trafilea二B件数行", "sum"),
            普通二B件数=("普通二B件数行", "sum"),
            任务订单类型=("订单分析类型", lambda s: _single_or_mixed(s, "混合订单类型")),
            任务SO匹配状态=("SO匹配状态", lambda s: _single_or_mixed(s, "部分SO未匹配")),
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
    tasks["任务储位结构"] = np.select(
        [tasks["高层件数"].eq(0), tasks["地面件数"].eq(0)],
        ["纯地面任务", "纯高层任务"],
        default="高低层混合任务",
    )
    tasks["有效任务工时"] = np.where(tasks["任务时长有效"], tasks["任务分钟"] / 60, 0)
    tasks["任务人效"] = np.where(tasks["有效任务工时"] > 0, tasks["任务件数"] / tasks["有效任务工时"], np.nan)
    tasks["高层件数占比"] = np.where(tasks["任务件数"] > 0, tasks["高层件数"] / tasks["任务件数"], 0)
    return tasks


def build_person_productivity(pick_filtered: pd.DataFrame, tasks: pd.DataFrame) -> pd.DataFrame:
    if pick_filtered.empty:
        return pd.DataFrame()

    work = pick_filtered.copy()
    work["二B件数行"] = np.where(work["订单类型"].eq("2B"), work["实际拣货量"], 0)
    work["Trafilea二B件数行"] = np.where(work["2B子类型"].eq("Trafilea-2B"), work["实际拣货量"], 0)
    work["普通二B件数行"] = np.where(work["2B子类型"].eq("普通2B"), work["实际拣货量"], 0)
    work["二C件数行"] = np.where(work["订单类型"].eq("2C"), work["实际拣货量"], 0)
    work["超大异常件数行"] = np.where(work["订单类型"].eq("超大异常单"), work["实际拣货量"], 0)

    people = (
        work.groupby("人员键", dropna=False)
        .agg(
            工号=("工号", "first"),
            姓名=("姓名", "first"),
            实际件数=("实际拣货量", "sum"),
            已匹配件数=("已匹配件数", "sum"),
            未匹配件数=("未匹配件数", "sum"),
            订单数=("SO", "nunique"),
            任务数=("任务单号", "nunique"),
            高层件数=("高层件数", "sum"),
            地面件数=("地面件数", "sum"),
            二B件数=("二B件数行", "sum"),
            Trafilea二B件数=("Trafilea二B件数行", "sum"),
            普通二B件数=("普通二B件数行", "sum"),
            二C件数=("二C件数行", "sum"),
            超大异常件数=("超大异常件数行", "sum"),
        )
        .reset_index()
    )

    valid_tasks = tasks[tasks["任务时长有效"]].copy() if not tasks.empty else pd.DataFrame()
    hours_by_person: dict[object, float] = {}
    if not valid_tasks.empty:
        for person, group in valid_tasks.groupby("人员键"):
            hours_by_person[person] = merge_intervals(list(zip(group["任务开始"], group["任务结束"])))

    people["有效操作工时"] = people["人员键"].map(hours_by_person).fillna(0)
    people["件数人效"] = np.where(people["有效操作工时"] > 0, people["实际件数"] / people["有效操作工时"], np.nan)
    people["SO匹配件数占比"] = np.where(people["实际件数"] > 0, people["已匹配件数"] / people["实际件数"], 0)
    people["高层件数占比"] = np.where(people["实际件数"] > 0, people["高层件数"] / people["实际件数"], 0)
    people["2B件数占已匹配比例"] = np.where(people["已匹配件数"] > 0, people["二B件数"] / people["已匹配件数"], 0)
    people["Trafilea-2B件数占已匹配比例"] = np.where(
        people["已匹配件数"] > 0, people["Trafilea二B件数"] / people["已匹配件数"], 0
    )
    people["普通2B件数占已匹配比例"] = np.where(
        people["已匹配件数"] > 0, people["普通二B件数"] / people["已匹配件数"], 0
    )
    people["2C件数占已匹配比例"] = np.where(people["已匹配件数"] > 0, people["二C件数"] / people["已匹配件数"], 0)
    return people.sort_values("件数人效", ascending=False, na_position="last")


def so_match_summary(pick_filtered: pd.DataFrame) -> dict[str, float]:
    if pick_filtered.empty:
        return {
            "总SO": 0, "已匹配SO": 0, "未匹配SO": 0, "SO匹配率": 0,
            "总件数": 0, "已匹配件数": 0, "未匹配件数": 0, "件数匹配率": 0,
        }
    order_match = pick_filtered.groupby("SO", dropna=False)["SO分类可用"].max()
    total_so = int(order_match.size)
    matched_so = int(order_match.sum())
    total_units = float(pick_filtered["实际拣货量"].sum())
    matched_units = float(pick_filtered["已匹配件数"].sum())
    return {
        "总SO": total_so,
        "已匹配SO": matched_so,
        "未匹配SO": total_so - matched_so,
        "SO匹配率": matched_so / total_so if total_so else 0,
        "总件数": total_units,
        "已匹配件数": matched_units,
        "未匹配件数": total_units - matched_units,
        "件数匹配率": matched_units / total_units if total_units else 0,
    }


def task_productivity_summary(tasks: pd.DataFrame) -> pd.DataFrame:
    if tasks.empty:
        return pd.DataFrame()
    valid = tasks[tasks["任务时长有效"]].copy()
    if valid.empty:
        return pd.DataFrame()
    summary = (
        valid.groupby(["任务订单类型", "任务储位结构"], dropna=False)
        .agg(
            任务数=("任务单号", "nunique"),
            操作人数=("人员键", "nunique"),
            件数=("任务件数", "sum"),
            有效任务工时=("有效任务工时", "sum"),
            中位任务分钟=("任务分钟", "median"),
        )
        .reset_index()
    )
    summary["任务口径人效"] = np.where(summary["有效任务工时"] > 0, summary["件数"] / summary["有效任务工时"], np.nan)
    return summary.sort_values(["任务订单类型", "任务储位结构"])


def demand_summary(exam: pd.DataFrame, due_date: date) -> dict[str, float]:
    valid = exam[~exam["是否取消订单"]].copy()
    due = valid[valid["到期日期"].eq(due_date)].copy()
    if due.empty:
        return {
            "到期订单数": 0, "到期件数": 0, "开班前已拣件数": 0, "开班时剩余件数": 0,
            "到期日前完成件数": 0, "按时完成件数": 0, "超过生产结束时间件数": 0,
        }
    start = pd.Timestamp(due_date)
    completion = due["考核单拣货完成时间"]
    early_before_day = completion.notna() & completion.lt(start)
    by_due_time = completion.notna() & completion.le(due["生产结束时间"])
    late_or_open = completion.isna() | completion.gt(due["生产结束时间"])
    return {
        "到期订单数": float(due["SO"].nunique()),
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
        pick_filtered.groupby("SO", dropna=False)
        .agg(
            SO匹配状态=("SO匹配状态", "first"),
            订单类型=("订单类型", "first"),
            二B子类型=("2B子类型", "first"),
            订单分析类型=("订单分析类型", "first"),
            二B识别依据=("2B识别依据", "first"),
            Trafilea货主=("Trafilea货主", "first"),
            考核单货主名称=("考核单货主名称", "first"),
            拣货结果货主名称=("拣货结果货主名称", "first"),
            SPB名称=("SPB名称", "first"),
            生产结束时间=("生产结束时间", "first"),
            实际完成时间=("拣货完成时间", "max"),
            实际件数=("实际拣货量", "sum"),
            高层件数=("高层件数", "sum"),
            地面件数=("地面件数", "sum"),
            任务数=("任务单号", "nunique"),
            操作人数=("人员键", "nunique"),
        )
        .reset_index()
        .rename(columns={"二B子类型": "2B子类型", "二B识别依据": "2B识别依据"})
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
    return orders.rename(columns={"SO": "订单号"})
