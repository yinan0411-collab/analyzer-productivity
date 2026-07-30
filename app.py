from __future__ import annotations

import hashlib
import io
import os
import sqlite3
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any, Iterable

import pandas as pd
import streamlit as st


APP_TITLE = "耗材使用与库存管理"
LOCAL_DB_PATH = "consumables_local.db"
USAGE_TABLE = "consumable_usage"
MASTER_TABLE = "consumable_master"
INVENTORY_TABLE = "consumable_inventory"

COLUMN_ALIASES = {
    "pack_time": ["打包时间", "Packing Time", "pack_time"],
    "so_number": ["订单号", "SO", "SO号", "销售订单号", "so_number"],
    "package_number": ["包裹编号", "包裹号", "Package No", "package_number"],
    "owner_name": ["货主名称", "客户名称", "Owner Name", "owner_name"],
    "material_code": ["实际耗材编码", "实际使用耗材编码", "耗材编码", "material_code"],
    "material_name": ["实际耗材名称", "实际使用耗材", "耗材名称", "material_name"],
}

MASTER_COLUMN_ALIASES = {
    "material_code": ["耗材编码", "实际耗材编码", "Material Code", "material_code"],
    "material_name": ["耗材名称", "实际耗材名称", "Material Name", "material_name"],
    "material_size": ["尺寸", "耗材尺寸", "规格", "Size", "material_size"],
    "pieces_per_pallet": [
        "每托数量",
        "每托Pieces",
        "每托件数",
        "Pieces per Pallet",
        "pieces_per_pallet",
    ],
}


@dataclass
class ProcessResult:
    package_records: pd.DataFrame
    material_summary: pd.DataFrame
    so_summary: pd.DataFrame
    daily_summary: pd.DataFrame
    invalid_rows: pd.DataFrame
    conflicts: pd.DataFrame
    warnings: list[str]
    raw_row_count: int
    file_count: int
    date_min: date | None
    date_max: date | None


@dataclass
class MasterProcessResult:
    records: pd.DataFrame
    invalid_rows: pd.DataFrame
    warnings: list[str]


# -----------------------------
# Utility functions
# -----------------------------
def normalize_text(series: pd.Series) -> pd.Series:
    out = series.astype("string").str.strip()
    return out.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA})


def find_column(columns: Iterable[str], aliases: list[str]) -> str | None:
    exact_map = {str(c).strip(): c for c in columns}
    lower_map = {str(c).strip().lower(): c for c in columns}
    for alias in aliases:
        if alias in exact_map:
            return exact_map[alias]
        if alias.lower() in lower_map:
            return lower_map[alias.lower()]
    return None


def file_sha256(file_bytes: bytes) -> str:
    return hashlib.sha256(file_bytes).hexdigest()


def build_usage_key(usage_date: Any, so_number: str, package_number: str) -> str:
    raw = f"{usage_date}|{so_number}|{package_number}".encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def build_inventory_key(inventory_date: Any, material_code: str) -> str:
    raw = f"{inventory_date}|{material_code}".encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def mode_or_first(series: pd.Series) -> Any:
    clean = series.dropna()
    if clean.empty:
        return None
    mode = clean.mode()
    return mode.iloc[0] if not mode.empty else clean.iloc[0]


def read_tabular_file(filename: str, file_bytes: bytes) -> pd.DataFrame:
    suffix = filename.lower().rsplit(".", 1)[-1]
    buffer = io.BytesIO(file_bytes)
    if suffix in {"xlsx", "xlsm", "xls"}:
        return pd.read_excel(buffer)
    if suffix == "csv":
        try:
            return pd.read_csv(buffer, encoding="utf-8-sig")
        except UnicodeDecodeError:
            buffer.seek(0)
            return pd.read_csv(buffer, encoding="gb18030")
    raise ValueError(f"暂不支持文件格式：{filename}")


def latest_tuesday(reference: date | None = None) -> date:
    reference = reference or date.today()
    days_since_tuesday = (reference.weekday() - 1) % 7
    return reference - timedelta(days=days_since_tuesday)


def process_uploaded_files(uploaded_files: list[Any]) -> ProcessResult:
    frames: list[pd.DataFrame] = []
    warnings: list[str] = []
    raw_row_count = 0

    for uploaded in uploaded_files:
        file_bytes = uploaded.getvalue()
        fingerprint = file_sha256(file_bytes)
        raw = read_tabular_file(uploaded.name, file_bytes)
        raw_row_count += len(raw)

        resolved: dict[str, str] = {}
        missing_columns: list[str] = []
        for canonical, aliases in COLUMN_ALIASES.items():
            matched = find_column(raw.columns, aliases)
            if matched is None:
                missing_columns.append(" / ".join(aliases[:2]))
            else:
                resolved[canonical] = matched

        if missing_columns:
            raise ValueError(
                f"文件 {uploaded.name} 缺少必要字段：{', '.join(missing_columns)}"
            )

        part = pd.DataFrame(
            {
                "pack_time": pd.to_datetime(raw[resolved["pack_time"]], errors="coerce"),
                "so_number": normalize_text(raw[resolved["so_number"]]),
                "package_number": normalize_text(raw[resolved["package_number"]]),
                "owner_name": normalize_text(raw[resolved["owner_name"]]),
                "material_code": normalize_text(raw[resolved["material_code"]]),
                "material_name": normalize_text(raw[resolved["material_name"]]),
            }
        )
        part["source_file"] = uploaded.name
        part["file_fingerprint"] = fingerprint
        part["source_row"] = range(2, len(part) + 2)

        # 业务口径：打包时间所在日期就是耗材使用日期。
        part["usage_date"] = part["pack_time"].dt.date
        frames.append(part)

    if not frames:
        return ProcessResult(
            package_records=pd.DataFrame(),
            material_summary=pd.DataFrame(),
            so_summary=pd.DataFrame(),
            daily_summary=pd.DataFrame(),
            invalid_rows=pd.DataFrame(),
            conflicts=pd.DataFrame(),
            warnings=[],
            raw_row_count=0,
            file_count=0,
            date_min=None,
            date_max=None,
        )

    combined = pd.concat(frames, ignore_index=True)
    required = [
        "pack_time",
        "usage_date",
        "so_number",
        "package_number",
        "owner_name",
        "material_code",
        "material_name",
    ]
    invalid_mask = combined[required].isna().any(axis=1)
    invalid_rows = combined.loc[invalid_mask].copy()
    valid = combined.loc[~invalid_mask].copy()

    if not invalid_rows.empty:
        warnings.append(f"有 {len(invalid_rows):,} 行缺少必要字段，未纳入统计。")

    if valid.empty:
        return ProcessResult(
            package_records=pd.DataFrame(),
            material_summary=pd.DataFrame(),
            so_summary=pd.DataFrame(),
            daily_summary=pd.DataFrame(),
            invalid_rows=invalid_rows,
            conflicts=pd.DataFrame(),
            warnings=warnings,
            raw_row_count=raw_row_count,
            file_count=len(uploaded_files),
            date_min=None,
            date_max=None,
        )

    package_group_cols = ["usage_date", "so_number", "package_number"]

    conflict_stats = (
        valid.groupby(package_group_cols, dropna=False)
        .agg(
            owner_names=("owner_name", lambda x: " | ".join(sorted(set(x.dropna())))),
            material_code_count=("material_code", "nunique"),
            material_name_count=("material_name", "nunique"),
            material_codes=("material_code", lambda x: " | ".join(sorted(set(x.dropna())))),
            material_names=("material_name", lambda x: " | ".join(sorted(set(x.dropna())))),
            raw_row_count=("so_number", "size"),
            source_files=("source_file", lambda x: " | ".join(sorted(set(x)))),
        )
        .reset_index()
    )
    conflicts = conflict_stats.loc[conflict_stats["material_code_count"] > 1].copy()

    if not conflicts.empty:
        warnings.append(
            f"发现 {len(conflicts):,} 个包裹对应多个实际耗材编码。这些包裹暂不保存。"
        )

    conflict_keys = set(
        zip(
            conflicts["usage_date"],
            conflicts["so_number"],
            conflicts["package_number"],
        )
    )
    if conflict_keys:
        valid["_key_tuple"] = list(
            zip(valid["usage_date"], valid["so_number"], valid["package_number"])
        )
        valid = valid.loc[~valid["_key_tuple"].isin(conflict_keys)].drop(columns="_key_tuple")

    name_variation_count = int(
        (
            (conflict_stats["material_code_count"] == 1)
            & (conflict_stats["material_name_count"] > 1)
        ).sum()
    )
    if name_variation_count:
        warnings.append(
            f"有 {name_variation_count:,} 个包裹耗材编码一致但名称不同，系统采用出现次数最多的名称。"
        )

    package_records = (
        valid.groupby(package_group_cols, dropna=False)
        .agg(
            pack_time=("pack_time", "max"),
            owner_name=("owner_name", mode_or_first),
            material_code=("material_code", mode_or_first),
            material_name=("material_name", mode_or_first),
            source_file=("source_file", lambda x: " | ".join(sorted(set(x)))),
            file_fingerprint=("file_fingerprint", lambda x: " | ".join(sorted(set(x)))),
            raw_row_count=("so_number", "size"),
        )
        .reset_index()
    )

    package_records["usage_key"] = package_records.apply(
        lambda r: build_usage_key(r["usage_date"], r["so_number"], r["package_number"]),
        axis=1,
    )

    # 拆单按“打包日期 + SO”判断。一个SO有多个包裹即为拆单。
    so_summary = (
        package_records.groupby(["usage_date", "so_number"], dropna=False)
        .agg(
            owner_name=("owner_name", mode_or_first),
            package_count=("package_number", "nunique"),
        )
        .reset_index()
    )
    so_summary["split_flag"] = (so_summary["package_count"] > 1).astype(int)
    so_summary["so_type"] = so_summary["split_flag"].map(
        {0: "单包裹SO", 1: "拆单SO"}
    )

    package_records = package_records.merge(
        so_summary[
            ["usage_date", "so_number", "package_count", "split_flag", "so_type"]
        ],
        on=["usage_date", "so_number"],
        how="left",
    )

    material_summary = (
        package_records.groupby(["material_code", "material_name"], dropna=False)
        .agg(
            usage_qty=("usage_key", "nunique"),
            so_count=("so_number", "nunique"),
        )
        .reset_index()
        .sort_values(["usage_qty", "material_code"], ascending=[False, True])
    )

    daily_summary = (
        package_records.groupby("usage_date")
        .agg(
            usage_qty=("usage_key", "nunique"),
            so_count=("so_number", "nunique"),
            split_so_count=(
                "so_number",
                lambda x: x[
                    package_records.loc[x.index, "split_flag"] == 1
                ].nunique(),
            ),
        )
        .reset_index()
        .sort_values("usage_date")
    )

    date_min = package_records["usage_date"].min() if not package_records.empty else None
    date_max = package_records["usage_date"].max() if not package_records.empty else None

    return ProcessResult(
        package_records=package_records,
        material_summary=material_summary,
        so_summary=so_summary,
        daily_summary=daily_summary,
        invalid_rows=invalid_rows,
        conflicts=conflicts,
        warnings=warnings,
        raw_row_count=raw_row_count,
        file_count=len(uploaded_files),
        date_min=date_min,
        date_max=date_max,
    )


def process_master_file(uploaded: Any) -> MasterProcessResult:
    raw = read_tabular_file(uploaded.name, uploaded.getvalue())
    resolved: dict[str, str] = {}
    missing: list[str] = []

    for canonical, aliases in MASTER_COLUMN_ALIASES.items():
        matched = find_column(raw.columns, aliases)
        if canonical == "material_size" and matched is None:
            continue
        if matched is None:
            missing.append(" / ".join(aliases[:2]))
        else:
            resolved[canonical] = matched

    if missing:
        raise ValueError(f"耗材主数据缺少字段：{', '.join(missing)}")

    records = pd.DataFrame(
        {
            "material_code": normalize_text(raw[resolved["material_code"]]),
            "material_name": normalize_text(raw[resolved["material_name"]]),
            "pieces_per_pallet": pd.to_numeric(
                raw[resolved["pieces_per_pallet"]], errors="coerce"
            ),
        }
    )
    if "material_size" in resolved:
        records["material_size"] = normalize_text(raw[resolved["material_size"]])
    else:
        records["material_size"] = ""

    invalid_mask = (
        records["material_code"].isna()
        | records["material_name"].isna()
        | records["pieces_per_pallet"].isna()
        | (records["pieces_per_pallet"] <= 0)
    )
    invalid_rows = records.loc[invalid_mask].copy()
    valid = records.loc[~invalid_mask].copy()
    valid["material_size"] = valid["material_size"].fillna("")
    valid["active"] = True

    duplicate_count = int(valid.duplicated("material_code", keep=False).sum())
    valid = valid.drop_duplicates("material_code", keep="last").sort_values("material_code")

    warnings: list[str] = []
    if not invalid_rows.empty:
        warnings.append(f"有 {len(invalid_rows):,} 行编码、名称或每托数量无效，未纳入。")
    if duplicate_count:
        warnings.append(f"发现重复耗材编码，系统保留每个编码最后一条记录。")

    return MasterProcessResult(
        records=valid.reset_index(drop=True),
        invalid_rows=invalid_rows,
        warnings=warnings,
    )


# -----------------------------
# Database layer
# -----------------------------
class StorageBase:
    mode_name = "Unknown"

    def fetch_usage(self, start_date: date, end_date: date) -> pd.DataFrame:
        raise NotImplementedError

    def upsert_usage(self, records: pd.DataFrame, uploader: str) -> dict[str, int]:
        raise NotImplementedError

    def fetch_master(self, active_only: bool = True) -> pd.DataFrame:
        raise NotImplementedError

    def upsert_master(self, records: pd.DataFrame, uploader: str) -> dict[str, int]:
        raise NotImplementedError

    def fetch_inventory(self, start_date: date, end_date: date) -> pd.DataFrame:
        raise NotImplementedError

    def upsert_inventory(self, records: pd.DataFrame, recorder: str) -> dict[str, int]:
        raise NotImplementedError


class SQLiteStorage(StorageBase):
    mode_name = "本地 SQLite（仅适合测试）"

    def __init__(self, db_path: str = LOCAL_DB_PATH):
        self.db_path = db_path
        self._init_db()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _columns(self, conn: sqlite3.Connection, table_name: str) -> set[str]:
        return {row[1] for row in conn.execute(f"PRAGMA table_info({table_name})")}

    def _init_db(self) -> None:
        with self._connect() as conn:
            conn.execute(
                f"""
                CREATE TABLE IF NOT EXISTS {USAGE_TABLE} (
                    usage_key TEXT PRIMARY KEY,
                    usage_date TEXT NOT NULL,
                    pack_time TEXT,
                    so_number TEXT NOT NULL,
                    package_number TEXT NOT NULL,
                    owner_name TEXT,
                    material_code TEXT NOT NULL,
                    material_name TEXT,
                    source_file TEXT,
                    file_fingerprint TEXT,
                    raw_row_count INTEGER DEFAULT 1,
                    uploader TEXT,
                    uploaded_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                )
                """
            )
            if "owner_name" not in self._columns(conn, USAGE_TABLE):
                conn.execute(f"ALTER TABLE {USAGE_TABLE} ADD COLUMN owner_name TEXT")

            conn.execute(
                f"""
                CREATE TABLE IF NOT EXISTS {MASTER_TABLE} (
                    material_code TEXT PRIMARY KEY,
                    material_name TEXT NOT NULL,
                    material_size TEXT,
                    pieces_per_pallet REAL NOT NULL,
                    active INTEGER NOT NULL DEFAULT 1,
                    uploader TEXT,
                    updated_at TEXT NOT NULL
                )
                """
            )
            conn.execute(
                f"""
                CREATE TABLE IF NOT EXISTS {INVENTORY_TABLE} (
                    inventory_key TEXT PRIMARY KEY,
                    inventory_date TEXT NOT NULL,
                    material_code TEXT NOT NULL,
                    material_name TEXT NOT NULL,
                    material_size TEXT,
                    pieces_per_pallet REAL NOT NULL,
                    pallet_qty REAL NOT NULL,
                    inventory_pieces INTEGER NOT NULL,
                    recorder TEXT NOT NULL,
                    notes TEXT,
                    recorded_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                )
                """
            )
            conn.execute(
                f"CREATE INDEX IF NOT EXISTS idx_usage_date ON {USAGE_TABLE}(usage_date)"
            )
            conn.execute(
                f"CREATE INDEX IF NOT EXISTS idx_material_code ON {USAGE_TABLE}(material_code)"
            )
            conn.execute(
                f"CREATE INDEX IF NOT EXISTS idx_inventory_date ON {INVENTORY_TABLE}(inventory_date)"
            )

    def fetch_usage(self, start_date: date, end_date: date) -> pd.DataFrame:
        with self._connect() as conn:
            return pd.read_sql_query(
                f"""
                SELECT * FROM {USAGE_TABLE}
                WHERE usage_date BETWEEN ? AND ?
                ORDER BY usage_date, pack_time, so_number, package_number
                """,
                conn,
                params=[start_date.isoformat(), end_date.isoformat()],
            )

    def upsert_usage(self, records: pd.DataFrame, uploader: str) -> dict[str, int]:
        if records.empty:
            return {"new": 0, "updated": 0, "unchanged": 0}

        existing = self.fetch_usage(records["usage_date"].min(), records["usage_date"].max())
        existing_map = {row["usage_key"]: row for _, row in existing.iterrows()}
        new_count = updated_count = unchanged_count = 0
        now = datetime.now().isoformat(timespec="seconds")

        sql = f"""
            INSERT INTO {USAGE_TABLE} (
                usage_key, usage_date, pack_time, so_number, package_number,
                owner_name, material_code, material_name, source_file, file_fingerprint,
                raw_row_count, uploader, uploaded_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(usage_key) DO UPDATE SET
                pack_time=excluded.pack_time,
                owner_name=excluded.owner_name,
                material_code=excluded.material_code,
                material_name=excluded.material_name,
                source_file=excluded.source_file,
                file_fingerprint=excluded.file_fingerprint,
                raw_row_count=excluded.raw_row_count,
                uploader=excluded.uploader,
                updated_at=excluded.updated_at
        """

        values = []
        for _, row in records.iterrows():
            old = existing_map.get(row["usage_key"])
            pack_time = pd.Timestamp(row["pack_time"]).isoformat()
            current_signature = (
                pack_time,
                str(row["owner_name"]),
                str(row["material_code"]),
                str(row["material_name"]),
            )
            if old is None:
                new_count += 1
            else:
                old_signature = (
                    pd.to_datetime(old.get("pack_time"), errors="coerce").isoformat(),
                    str(old.get("owner_name")),
                    str(old.get("material_code")),
                    str(old.get("material_name")),
                )
                if current_signature == old_signature:
                    unchanged_count += 1
                else:
                    updated_count += 1

            values.append(
                (
                    row["usage_key"],
                    row["usage_date"].isoformat(),
                    pack_time,
                    str(row["so_number"]),
                    str(row["package_number"]),
                    str(row["owner_name"]),
                    str(row["material_code"]),
                    str(row["material_name"]),
                    str(row["source_file"]),
                    str(row["file_fingerprint"]),
                    int(row["raw_row_count"]),
                    uploader,
                    now,
                    now,
                )
            )

        with self._connect() as conn:
            conn.executemany(sql, values)
        return {"new": new_count, "updated": updated_count, "unchanged": unchanged_count}

    def fetch_master(self, active_only: bool = True) -> pd.DataFrame:
        where = "WHERE active = 1" if active_only else ""
        with self._connect() as conn:
            return pd.read_sql_query(
                f"SELECT * FROM {MASTER_TABLE} {where} ORDER BY material_code",
                conn,
            )

    def upsert_master(self, records: pd.DataFrame, uploader: str) -> dict[str, int]:
        if records.empty:
            return {"new": 0, "updated": 0}
        existing = self.fetch_master(active_only=False)
        existing_codes = set(existing["material_code"]) if not existing.empty else set()
        new_count = sum(code not in existing_codes for code in records["material_code"])
        updated_count = len(records) - new_count
        now = datetime.now().isoformat(timespec="seconds")
        sql = f"""
            INSERT INTO {MASTER_TABLE} (
                material_code, material_name, material_size, pieces_per_pallet,
                active, uploader, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(material_code) DO UPDATE SET
                material_name=excluded.material_name,
                material_size=excluded.material_size,
                pieces_per_pallet=excluded.pieces_per_pallet,
                active=excluded.active,
                uploader=excluded.uploader,
                updated_at=excluded.updated_at
        """
        values = [
            (
                str(row["material_code"]),
                str(row["material_name"]),
                str(row["material_size"]),
                float(row["pieces_per_pallet"]),
                1 if bool(row["active"]) else 0,
                uploader,
                now,
            )
            for _, row in records.iterrows()
        ]
        with self._connect() as conn:
            conn.executemany(sql, values)
        return {"new": int(new_count), "updated": int(updated_count)}

    def fetch_inventory(self, start_date: date, end_date: date) -> pd.DataFrame:
        with self._connect() as conn:
            return pd.read_sql_query(
                f"""
                SELECT * FROM {INVENTORY_TABLE}
                WHERE inventory_date BETWEEN ? AND ?
                ORDER BY inventory_date, material_code
                """,
                conn,
                params=[start_date.isoformat(), end_date.isoformat()],
            )

    def upsert_inventory(self, records: pd.DataFrame, recorder: str) -> dict[str, int]:
        if records.empty:
            return {"new": 0, "updated": 0}
        existing = self.fetch_inventory(
            records["inventory_date"].min(), records["inventory_date"].max()
        )
        existing_keys = set(existing["inventory_key"]) if not existing.empty else set()
        new_count = sum(key not in existing_keys for key in records["inventory_key"])
        updated_count = len(records) - new_count
        now = datetime.now().isoformat(timespec="seconds")
        sql = f"""
            INSERT INTO {INVENTORY_TABLE} (
                inventory_key, inventory_date, material_code, material_name,
                material_size, pieces_per_pallet, pallet_qty, inventory_pieces,
                recorder, notes, recorded_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(inventory_key) DO UPDATE SET
                material_name=excluded.material_name,
                material_size=excluded.material_size,
                pieces_per_pallet=excluded.pieces_per_pallet,
                pallet_qty=excluded.pallet_qty,
                inventory_pieces=excluded.inventory_pieces,
                recorder=excluded.recorder,
                notes=excluded.notes,
                updated_at=excluded.updated_at
        """
        values = [
            (
                row["inventory_key"],
                row["inventory_date"].isoformat(),
                str(row["material_code"]),
                str(row["material_name"]),
                str(row["material_size"]),
                float(row["pieces_per_pallet"]),
                float(row["pallet_qty"]),
                int(row["inventory_pieces"]),
                recorder,
                str(row.get("notes", "") or ""),
                now,
                now,
            )
            for _, row in records.iterrows()
        ]
        with self._connect() as conn:
            conn.executemany(sql, values)
        return {"new": int(new_count), "updated": int(updated_count)}


class SupabaseStorage(StorageBase):
    mode_name = "Supabase 云端数据库"

    def __init__(self, url: str, key: str):
        from supabase import create_client
        self.client = create_client(url, key)

    def _fetch_paged(self, table: str, query_builder: Any) -> pd.DataFrame:
        rows: list[dict[str, Any]] = []
        offset = 0
        page_size = 1000
        while True:
            response = query_builder(offset, page_size).execute()
            batch = response.data or []
            rows.extend(batch)
            if len(batch) < page_size:
                break
            offset += page_size
        return pd.DataFrame(rows)

    def fetch_usage(self, start_date: date, end_date: date) -> pd.DataFrame:
        return self._fetch_paged(
            USAGE_TABLE,
            lambda offset, size: (
                self.client.table(USAGE_TABLE)
                .select("*")
                .gte("usage_date", start_date.isoformat())
                .lte("usage_date", end_date.isoformat())
                .order("usage_date")
                .order("pack_time")
                .range(offset, offset + size - 1)
            ),
        )

    def upsert_usage(self, records: pd.DataFrame, uploader: str) -> dict[str, int]:
        if records.empty:
            return {"new": 0, "updated": 0, "unchanged": 0}

        existing = self.fetch_usage(records["usage_date"].min(), records["usage_date"].max())
        existing_map = (
            {row["usage_key"]: row for _, row in existing.iterrows()}
            if not existing.empty else {}
        )
        new_count = updated_count = unchanged_count = 0
        now = datetime.now().isoformat(timespec="seconds")
        payload: list[dict[str, Any]] = []

        for _, row in records.iterrows():
            pack_time = pd.Timestamp(row["pack_time"]).isoformat()
            old = existing_map.get(row["usage_key"])
            current_signature = (
                pack_time,
                str(row["owner_name"]),
                str(row["material_code"]),
                str(row["material_name"]),
            )
            if old is None:
                new_count += 1
            else:
                old_signature = (
                    pd.to_datetime(old.get("pack_time"), errors="coerce").isoformat(),
                    str(old.get("owner_name")),
                    str(old.get("material_code")),
                    str(old.get("material_name")),
                )
                if current_signature == old_signature:
                    unchanged_count += 1
                else:
                    updated_count += 1

            payload.append(
                {
                    "usage_key": row["usage_key"],
                    "usage_date": row["usage_date"].isoformat(),
                    "pack_time": pack_time,
                    "so_number": str(row["so_number"]),
                    "package_number": str(row["package_number"]),
                    "owner_name": str(row["owner_name"]),
                    "material_code": str(row["material_code"]),
                    "material_name": str(row["material_name"]),
                    "source_file": str(row["source_file"]),
                    "file_fingerprint": str(row["file_fingerprint"]),
                    "raw_row_count": int(row["raw_row_count"]),
                    "uploader": uploader,
                    "uploaded_at": now,
                    "updated_at": now,
                }
            )

        for start in range(0, len(payload), 500):
            self.client.table(USAGE_TABLE).upsert(
                payload[start:start + 500], on_conflict="usage_key"
            ).execute()

        return {"new": new_count, "updated": updated_count, "unchanged": unchanged_count}

    def fetch_master(self, active_only: bool = True) -> pd.DataFrame:
        query = self.client.table(MASTER_TABLE).select("*").order("material_code")
        if active_only:
            query = query.eq("active", True)
        response = query.execute()
        return pd.DataFrame(response.data or [])

    def upsert_master(self, records: pd.DataFrame, uploader: str) -> dict[str, int]:
        if records.empty:
            return {"new": 0, "updated": 0}
        existing = self.fetch_master(active_only=False)
        existing_codes = set(existing["material_code"]) if not existing.empty else set()
        new_count = sum(code not in existing_codes for code in records["material_code"])
        now = datetime.now().isoformat(timespec="seconds")
        payload = [
            {
                "material_code": str(row["material_code"]),
                "material_name": str(row["material_name"]),
                "material_size": str(row["material_size"]),
                "pieces_per_pallet": float(row["pieces_per_pallet"]),
                "active": bool(row["active"]),
                "uploader": uploader,
                "updated_at": now,
            }
            for _, row in records.iterrows()
        ]
        for start in range(0, len(payload), 500):
            self.client.table(MASTER_TABLE).upsert(
                payload[start:start + 500], on_conflict="material_code"
            ).execute()
        return {"new": int(new_count), "updated": int(len(records) - new_count)}

    def fetch_inventory(self, start_date: date, end_date: date) -> pd.DataFrame:
        return self._fetch_paged(
            INVENTORY_TABLE,
            lambda offset, size: (
                self.client.table(INVENTORY_TABLE)
                .select("*")
                .gte("inventory_date", start_date.isoformat())
                .lte("inventory_date", end_date.isoformat())
                .order("inventory_date")
                .order("material_code")
                .range(offset, offset + size - 1)
            ),
        )

    def upsert_inventory(self, records: pd.DataFrame, recorder: str) -> dict[str, int]:
        if records.empty:
            return {"new": 0, "updated": 0}
        existing = self.fetch_inventory(
            records["inventory_date"].min(), records["inventory_date"].max()
        )
        existing_keys = set(existing["inventory_key"]) if not existing.empty else set()
        new_count = sum(key not in existing_keys for key in records["inventory_key"])
        now = datetime.now().isoformat(timespec="seconds")
        payload = [
            {
                "inventory_key": row["inventory_key"],
                "inventory_date": row["inventory_date"].isoformat(),
                "material_code": str(row["material_code"]),
                "material_name": str(row["material_name"]),
                "material_size": str(row["material_size"]),
                "pieces_per_pallet": float(row["pieces_per_pallet"]),
                "pallet_qty": float(row["pallet_qty"]),
                "inventory_pieces": int(row["inventory_pieces"]),
                "recorder": recorder,
                "notes": str(row.get("notes", "") or ""),
                "recorded_at": now,
                "updated_at": now,
            }
            for _, row in records.iterrows()
        ]
        for start in range(0, len(payload), 500):
            self.client.table(INVENTORY_TABLE).upsert(
                payload[start:start + 500], on_conflict="inventory_key"
            ).execute()
        return {"new": int(new_count), "updated": int(len(records) - new_count)}


@st.cache_resource
def get_storage() -> StorageBase:
    url = os.getenv("SUPABASE_URL")
    key = os.getenv("SUPABASE_KEY")
    try:
        url = st.secrets.get("SUPABASE_URL", url)
        key = st.secrets.get("SUPABASE_KEY", key)
    except Exception:
        pass
    if url and key:
        return SupabaseStorage(url, key)
    return SQLiteStorage()


# -----------------------------
# Historical calculations
# -----------------------------
def prepare_history(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    if df.empty:
        empty = pd.DataFrame()
        return {
            "detail": empty,
            "so_summary": empty,
            "material_summary": empty,
            "daily_summary": empty,
        }

    data = df.copy()
    data["usage_date"] = pd.to_datetime(data["usage_date"]).dt.date
    data["pack_time"] = pd.to_datetime(data["pack_time"], errors="coerce")
    if "owner_name" not in data.columns:
        data["owner_name"] = ""

    so_summary = (
        data.groupby(["usage_date", "so_number"], dropna=False)
        .agg(
            owner_name=("owner_name", mode_or_first),
            package_count=("package_number", "nunique"),
        )
        .reset_index()
    )
    so_summary["split_flag"] = (so_summary["package_count"] > 1).astype(int)
    so_summary["so_type"] = so_summary["split_flag"].map(
        {0: "单包裹SO", 1: "拆单SO"}
    )

    detail = data.merge(
        so_summary[
            ["usage_date", "so_number", "package_count", "split_flag", "so_type"]
        ],
        on=["usage_date", "so_number"],
        how="left",
    )

    material_summary = (
        detail.groupby(["material_code", "material_name"], dropna=False)
        .agg(
            usage_qty=("usage_key", "nunique"),
            so_count=("so_number", "nunique"),
            split_so_count=(
                "so_number",
                lambda x: x[detail.loc[x.index, "split_flag"] == 1].nunique(),
            ),
        )
        .reset_index()
        .sort_values(["usage_qty", "material_code"], ascending=[False, True])
    )

    daily_summary = (
        detail.groupby("usage_date")
        .agg(
            usage_qty=("usage_key", "nunique"),
            so_count=("so_number", "nunique"),
            split_so_count=(
                "so_number",
                lambda x: x[detail.loc[x.index, "split_flag"] == 1].nunique(),
            ),
        )
        .reset_index()
        .sort_values("usage_date")
    )

    return {
        "detail": detail,
        "so_summary": so_summary,
        "material_summary": material_summary,
        "daily_summary": daily_summary,
    }


def dataframe_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title=APP_TITLE, page_icon="📦", layout="wide")
st.title("📦 耗材使用与库存管理")
st.caption("打包时间所在日期就是耗材使用日期；SO判断拆单，包裹计算实际耗材数量。")

storage = get_storage()
with st.sidebar:
    st.subheader("系统状态")
    st.write(f"数据存储：**{storage.mode_name}**")
    if isinstance(storage, SQLiteStorage):
        st.warning("当前为本地测试模式。部署后请配置Supabase保证长期保存。")
    st.divider()
    st.markdown(
        "**核心规则**\n\n"
        "1. 同一包裹多商品行只算一次。\n"
        "2. 一个SO多个包裹属于拆单，每个包裹都计算耗材。\n"
        "3. 耗材使用日期取打包时间的日期。\n"
        "4. 重复上传不会重复累计。"
    )

upload_tab, history_tab, detail_tab, inventory_tab = st.tabs(
    ["① 上传并保存", "② 历史用量查询", "③ 明细与异常", "④ 每周二库存记录"]
)

with upload_tab:
    st.subheader("上传复核打包结果")
    st.write("可以一次上传一天、一周、多天或多个文件，系统直接读取“打包时间”作为耗材使用日期。")
    uploaded_files = st.file_uploader(
        "选择Excel或CSV文件",
        type=["xlsx", "xlsm", "xls", "csv"],
        accept_multiple_files=True,
        key="usage_upload",
    )

    if uploaded_files:
        try:
            result = process_uploaded_files(uploaded_files)
            st.session_state["last_process_result"] = result
        except Exception as exc:
            st.error(f"文件处理失败：{exc}")
            st.stop()

        unique_dates = (
            result.package_records["usage_date"].nunique()
            if not result.package_records.empty else 0
        )
        if result.date_min == result.date_max and result.date_min:
            date_text = str(result.date_min)
        elif result.date_min and result.date_max:
            date_text = f"{result.date_min} 至 {result.date_max}（{unique_dates}个打包日期）"
        else:
            date_text = "无法识别"

        so_count = len(result.so_summary)
        split_so_count = int((result.so_summary["split_flag"] == 1).sum()) if so_count else 0
        single_so_count = int((result.so_summary["split_flag"] == 0).sum()) if so_count else 0
        package_count = len(result.package_records)
        split_rate = split_so_count / so_count if so_count else 0

        st.info(f"耗材使用日期：**{date_text}**；共上传 **{result.file_count}** 个文件。")
        metric_cols = st.columns(6)
        metric_cols[0].metric("原始明细行", f"{result.raw_row_count:,}")
        metric_cols[1].metric("SO数量", f"{so_count:,}")
        metric_cols[2].metric("实际包裹/耗材", f"{package_count:,}")
        metric_cols[3].metric("单包裹SO", f"{single_so_count:,}")
        metric_cols[4].metric("拆单SO", f"{split_so_count:,}")
        metric_cols[5].metric("拆单率", f"{split_rate:.2%}")

        for warning in result.warnings:
            st.warning(warning)

        st.markdown("#### 本次上传耗材汇总")
        material_display = result.material_summary.rename(
            columns={
                "material_code": "实际耗材编码",
                "material_name": "实际耗材名称",
                "usage_qty": "使用数量",
                "so_count": "涉及SO数",
            }
        )
        st.dataframe(material_display, use_container_width=True, hide_index=True)

        with st.expander("查看每日汇总"):
            st.dataframe(
                result.daily_summary.rename(
                    columns={
                        "usage_date": "耗材使用日期",
                        "usage_qty": "耗材数量",
                        "so_count": "SO数量",
                        "split_so_count": "拆单SO数量",
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

        with st.expander("查看拆单SO（含货主名称）", expanded=split_so_count > 0):
            split_display = result.so_summary.loc[
                result.so_summary["split_flag"] == 1
            ].rename(
                columns={
                    "usage_date": "耗材使用日期",
                    "owner_name": "货主名称",
                    "so_number": "订单号",
                    "package_count": "包裹数",
                    "split_flag": "拆单标记",
                    "so_type": "SO类型",
                }
            )
            st.dataframe(split_display, use_container_width=True, hide_index=True)

        if not result.conflicts.empty:
            st.error("存在同一包裹对应多个耗材编码的冲突，冲突包裹不会保存。")
            st.dataframe(result.conflicts, use_container_width=True, hide_index=True)

        uploader = st.text_input("上传人姓名或账号", placeholder="例如：Yinan")
        can_save = not result.package_records.empty and bool(uploader.strip())
        if st.button("确认保存并累计", type="primary", disabled=not can_save):
            try:
                counts = storage.upsert_usage(result.package_records, uploader.strip())
                st.success(
                    "保存完成："
                    f"新增 {counts['new']:,} 个包裹；"
                    f"更新 {counts['updated']:,} 个包裹；"
                    f"重复且未变化 {counts['unchanged']:,} 个包裹。"
                )
                st.caption("重复文件或重复日期不会造成双倍累计。")
            except Exception as exc:
                st.error(f"保存失败：{exc}")
                st.info("如刚更新V2，请先在Supabase SQL Editor运行 supabase_update_v2.sql。")
        elif not uploader.strip():
            st.caption("填写上传人后即可保存。")
    else:
        st.info("请上传复核打包结果文件。")

with history_tab:
    st.subheader("按日期查询累计用量")
    today = date.today()
    default_start = today.replace(day=1)
    date_cols = st.columns([1, 1, 2])
    start_date = date_cols[0].date_input("开始日期", value=default_start)
    end_date = date_cols[1].date_input("结束日期", value=today)
    material_keyword = date_cols[2].text_input(
        "耗材编码或名称筛选", placeholder="可留空"
    )

    if start_date > end_date:
        st.error("开始日期不能晚于结束日期。")
    elif st.button("查询历史用量", type="primary"):
        try:
            history_raw = storage.fetch_usage(start_date, end_date)
            history = prepare_history(history_raw)
            st.session_state["history_result"] = history
            st.session_state["history_range"] = (start_date, end_date)
        except Exception as exc:
            st.error(f"查询失败：{exc}")

    history = st.session_state.get("history_result")
    if history is not None:
        detail = history["detail"].copy()
        if material_keyword.strip() and not detail.empty:
            keyword = material_keyword.strip().lower()
            mask = (
                detail["material_code"].astype(str).str.lower().str.contains(keyword, na=False)
                | detail["material_name"].astype(str).str.lower().str.contains(keyword, na=False)
            )
            detail = detail.loc[mask]
            history = prepare_history(detail)

        if detail.empty:
            st.warning("所选日期范围内没有数据。")
        else:
            so_summary = history["so_summary"]
            material_summary = history["material_summary"]
            daily_summary = history["daily_summary"]
            package_count = detail["usage_key"].nunique()
            so_count = len(so_summary)
            split_so_count = int((so_summary["split_flag"] == 1).sum())
            single_so_count = int((so_summary["split_flag"] == 0).sum())

            metric_cols = st.columns(5)
            metric_cols[0].metric("累计耗材数量", f"{package_count:,}")
            metric_cols[1].metric("SO数量", f"{so_count:,}")
            metric_cols[2].metric("单包裹SO", f"{single_so_count:,}")
            metric_cols[3].metric("拆单SO", f"{split_so_count:,}")
            metric_cols[4].metric(
                "拆单率",
                f"{split_so_count / so_count:.2%}" if so_count else "0.00%",
            )

            st.markdown("#### 日期用量趋势")
            chart_df = daily_summary.set_index("usage_date")[["usage_qty"]].rename(
                columns={"usage_qty": "耗材数量"}
            )
            st.bar_chart(chart_df)

            st.markdown("#### 耗材累计使用量")
            material_display = material_summary.rename(
                columns={
                    "material_code": "实际耗材编码",
                    "material_name": "实际耗材名称",
                    "usage_qty": "使用数量",
                    "so_count": "涉及SO数",
                    "split_so_count": "涉及拆单SO数",
                }
            )
            st.dataframe(material_display, use_container_width=True, hide_index=True)
            st.download_button(
                "下载用量汇总CSV",
                data=dataframe_to_csv_bytes(material_display),
                file_name=f"耗材用量_{start_date}_{end_date}.csv",
                mime="text/csv",
            )

            with st.expander("查看该日期范围内的拆单SO"):
                split_history = so_summary.loc[
                    so_summary["split_flag"] == 1
                ].rename(
                    columns={
                        "usage_date": "耗材使用日期",
                        "owner_name": "货主名称",
                        "so_number": "订单号",
                        "package_count": "包裹数",
                        "split_flag": "拆单标记",
                        "so_type": "SO类型",
                    }
                )
                st.dataframe(split_history, use_container_width=True, hide_index=True)

with detail_tab:
    st.subheader("包裹级明细")
    history = st.session_state.get("history_result")
    if history is None or history["detail"].empty:
        st.info("请先在“历史用量查询”中选择日期并查询。")
    else:
        detail = history["detail"].copy()
        detail_display = detail.rename(
            columns={
                "usage_date": "耗材使用日期",
                "pack_time": "打包时间",
                "owner_name": "货主名称",
                "so_number": "订单号",
                "package_number": "包裹编号",
                "material_code": "实际耗材编码",
                "material_name": "实际耗材名称",
                "source_file": "来源文件",
                "raw_row_count": "原始重复行数",
                "uploader": "上传人",
                "package_count": "SO包裹数",
                "split_flag": "拆单标记",
                "so_type": "SO类型",
            }
        )
        wanted = [
            "耗材使用日期",
            "打包时间",
            "货主名称",
            "订单号",
            "包裹编号",
            "实际耗材编码",
            "实际耗材名称",
            "SO包裹数",
            "拆单标记",
            "SO类型",
            "原始重复行数",
            "来源文件",
            "上传人",
        ]
        existing_cols = [c for c in wanted if c in detail_display.columns]
        st.dataframe(detail_display[existing_cols], use_container_width=True, hide_index=True)
        st.download_button(
            "下载包裹明细CSV",
            data=dataframe_to_csv_bytes(detail_display[existing_cols]),
            file_name="耗材包裹明细.csv",
            mime="text/csv",
        )

with inventory_tab:
    st.subheader("每周二库存记录")
    st.write("员工输入实际库存托数，系统根据“每托数量”自动换算为Pieces并永久保存。")
    employee_tab, master_tab, inventory_history_tab = st.tabs(
        ["员工库存盘点", "管理员导入耗材主数据", "库存历史"]
    )

    with employee_tab:
        input_cols = st.columns([1, 1, 2])
        inventory_date = input_cols[0].date_input(
            "盘点日期",
            value=latest_tuesday(),
            key="inventory_date",
            help="默认显示最近一个周二，也可以选择其他日期补录。",
        )
        recorder = input_cols[1].text_input(
            "记录人姓名或账号", placeholder="例如：Maria", key="inventory_recorder"
        )
        search_keyword = input_cols[2].text_input(
            "搜索耗材编码或名称", placeholder="可留空", key="inventory_search"
        )

        if inventory_date.weekday() != 1:
            st.warning("所选日期不是周二。系统允许补录，但请确认日期是否正确。")

        try:
            master = storage.fetch_master(active_only=True)
        except Exception as exc:
            st.error(f"读取耗材主数据失败：{exc}")
            master = pd.DataFrame()

        if master.empty:
            st.info("还没有耗材主数据。请管理员先到“管理员导入耗材主数据”上传标准耗材表。")
        else:
            try:
                existing_inventory = storage.fetch_inventory(inventory_date, inventory_date)
            except Exception as exc:
                st.error(f"读取当日库存记录失败：{exc}")
                existing_inventory = pd.DataFrame()

            editor_source = master[
                ["material_code", "material_name", "material_size", "pieces_per_pallet"]
            ].copy()
            if not existing_inventory.empty:
                merge_cols = ["material_code", "pallet_qty", "notes"]
                editor_source = editor_source.merge(
                    existing_inventory[merge_cols], on="material_code", how="left"
                )
            else:
                editor_source["pallet_qty"] = pd.NA
                editor_source["notes"] = ""

            if search_keyword.strip():
                kw = search_keyword.strip().lower()
                editor_source = editor_source.loc[
                    editor_source["material_code"].astype(str).str.lower().str.contains(kw, na=False)
                    | editor_source["material_name"].astype(str).str.lower().str.contains(kw, na=False)
                ]

            editor_source = editor_source.rename(
                columns={
                    "material_code": "耗材编码",
                    "material_name": "耗材名称",
                    "material_size": "尺寸",
                    "pieces_per_pallet": "每托数量",
                    "pallet_qty": "实际库存托数",
                    "notes": "备注",
                }
            )

            edited = st.data_editor(
                editor_source,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                disabled=["耗材编码", "耗材名称", "尺寸", "每托数量"],
                column_config={
                    "每托数量": st.column_config.NumberColumn(format="%.0f"),
                    "实际库存托数": st.column_config.NumberColumn(
                        min_value=0.0,
                        step=0.25,
                        format="%.2f",
                        help="可输入小数托，例如0.5托、1.25托。",
                    ),
                },
                key=f"inventory_editor_{inventory_date}",
            )

            edited["系统换算Pieces"] = (
                pd.to_numeric(edited["实际库存托数"], errors="coerce")
                * pd.to_numeric(edited["每托数量"], errors="coerce")
            ).round().astype("Int64")

            filled_count = int(edited["实际库存托数"].notna().sum())
            total_rows = len(edited)
            total_pieces = int(edited["系统换算Pieces"].fillna(0).sum())

            metric_cols = st.columns(3)
            metric_cols[0].metric("页面耗材数", f"{total_rows:,}")
            metric_cols[1].metric("已填写", f"{filled_count:,}")
            metric_cols[2].metric("换算库存Pieces", f"{total_pieces:,}")

            st.caption("系统换算结果：实际库存托数 × 每托数量，四舍五入为Pieces。")
            st.dataframe(
                edited[
                    ["耗材编码", "耗材名称", "实际库存托数", "每托数量", "系统换算Pieces"]
                ],
                use_container_width=True,
                hide_index=True,
            )

            save_inventory = st.button(
                "保存本周库存记录",
                type="primary",
                disabled=not recorder.strip() or filled_count == 0,
            )
            if save_inventory:
                filled = edited.loc[edited["实际库存托数"].notna()].copy()
                save_df = pd.DataFrame(
                    {
                        "inventory_date": inventory_date,
                        "material_code": filled["耗材编码"].astype(str),
                        "material_name": filled["耗材名称"].astype(str),
                        "material_size": filled["尺寸"].fillna("").astype(str),
                        "pieces_per_pallet": pd.to_numeric(
                            filled["每托数量"], errors="coerce"
                        ),
                        "pallet_qty": pd.to_numeric(
                            filled["实际库存托数"], errors="coerce"
                        ),
                        "inventory_pieces": filled["系统换算Pieces"].astype(int),
                        "notes": filled["备注"].fillna("").astype(str),
                    }
                )
                save_df["inventory_key"] = save_df.apply(
                    lambda r: build_inventory_key(
                        r["inventory_date"], r["material_code"]
                    ),
                    axis=1,
                )
                try:
                    counts = storage.upsert_inventory(save_df, recorder.strip())
                    st.success(
                        f"库存记录已保存：新增 {counts['new']:,} 项，更新 {counts['updated']:,} 项。"
                    )
                    st.caption("同一个盘点日期和耗材编码再次保存时，会更新原记录，不会重复累计。")
                except Exception as exc:
                    st.error(f"库存保存失败：{exc}")
                    st.info("请确认已在Supabase运行 supabase_update_v2.sql。")
            elif not recorder.strip():
                st.caption("填写记录人后即可保存。")

    with master_tab:
        st.write("标准耗材表需要包含：耗材编码、耗材名称、尺寸（可选）、每托数量。")
        master_file = st.file_uploader(
            "上传标准耗材Excel或CSV",
            type=["xlsx", "xlsm", "xls", "csv"],
            key="master_upload",
        )
        if master_file:
            try:
                master_result = process_master_file(master_file)
            except Exception as exc:
                st.error(f"主数据处理失败：{exc}")
            else:
                for warning in master_result.warnings:
                    st.warning(warning)
                master_preview = master_result.records.rename(
                    columns={
                        "material_code": "耗材编码",
                        "material_name": "耗材名称",
                        "material_size": "尺寸",
                        "pieces_per_pallet": "每托数量",
                        "active": "启用",
                    }
                )
                st.dataframe(master_preview, use_container_width=True, hide_index=True)
                master_uploader = st.text_input(
                    "管理员姓名或账号", placeholder="例如：Yinan", key="master_uploader"
                )
                if st.button(
                    "保存耗材主数据",
                    type="primary",
                    disabled=not master_uploader.strip() or master_result.records.empty,
                ):
                    try:
                        counts = storage.upsert_master(
                            master_result.records, master_uploader.strip()
                        )
                        st.success(
                            f"耗材主数据已保存：新增 {counts['new']:,} 项，更新 {counts['updated']:,} 项。"
                        )
                    except Exception as exc:
                        st.error(f"主数据保存失败：{exc}")
                        st.info("请确认已在Supabase运行 supabase_update_v2.sql。")

        st.markdown("#### 当前耗材主数据")
        try:
            current_master = storage.fetch_master(active_only=False)
            if current_master.empty:
                st.info("暂无耗材主数据。")
            else:
                st.dataframe(
                    current_master.rename(
                        columns={
                            "material_code": "耗材编码",
                            "material_name": "耗材名称",
                            "material_size": "尺寸",
                            "pieces_per_pallet": "每托数量",
                            "active": "启用",
                            "uploader": "更新人",
                            "updated_at": "更新时间",
                        }
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
        except Exception as exc:
            st.error(f"读取耗材主数据失败：{exc}")

    with inventory_history_tab:
        history_cols = st.columns([1, 1, 2])
        inv_start = history_cols[0].date_input(
            "库存开始日期",
            value=latest_tuesday() - timedelta(days=28),
            key="inv_history_start",
        )
        inv_end = history_cols[1].date_input(
            "库存结束日期",
            value=latest_tuesday(),
            key="inv_history_end",
        )
        inv_keyword = history_cols[2].text_input(
            "耗材编码或名称", placeholder="可留空", key="inv_history_keyword"
        )

        if st.button("查询库存历史"):
            try:
                inventory_history = storage.fetch_inventory(inv_start, inv_end)
                st.session_state["inventory_history"] = inventory_history
            except Exception as exc:
                st.error(f"库存历史查询失败：{exc}")

        inventory_history = st.session_state.get("inventory_history")
        if inventory_history is not None:
            if inventory_history.empty:
                st.info("所选日期范围内没有库存记录。")
            else:
                display = inventory_history.copy()
                if inv_keyword.strip():
                    kw = inv_keyword.strip().lower()
                    display = display.loc[
                        display["material_code"].astype(str).str.lower().str.contains(kw, na=False)
                        | display["material_name"].astype(str).str.lower().str.contains(kw, na=False)
                    ]
                display = display.rename(
                    columns={
                        "inventory_date": "盘点日期",
                        "material_code": "耗材编码",
                        "material_name": "耗材名称",
                        "material_size": "尺寸",
                        "pieces_per_pallet": "每托数量",
                        "pallet_qty": "实际库存托数",
                        "inventory_pieces": "库存Pieces",
                        "recorder": "记录人",
                        "notes": "备注",
                        "updated_at": "更新时间",
                    }
                )
                wanted = [
                    "盘点日期",
                    "耗材编码",
                    "耗材名称",
                    "尺寸",
                    "每托数量",
                    "实际库存托数",
                    "库存Pieces",
                    "记录人",
                    "备注",
                    "更新时间",
                ]
                st.dataframe(display[wanted], use_container_width=True, hide_index=True)
                st.download_button(
                    "下载库存历史CSV",
                    data=dataframe_to_csv_bytes(display[wanted]),
                    file_name=f"库存记录_{inv_start}_{inv_end}.csv",
                    mime="text/csv",
                )
