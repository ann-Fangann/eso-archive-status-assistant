from __future__ import annotations

import io
import re
from dataclasses import dataclass
from typing import Iterable, List

import pandas as pd


@dataclass
class ExcelTable:
    dataframe: pd.DataFrame
    sheet_name: str
    header_row_index: int
    original_columns: List[str]


def sanitize_header(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower().startswith("unnamed:"):
        return ""
    text = text.replace("\n", " ").replace("\r", " ").strip()
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"[^\w\u4e00-\u9fff.%-]+", "_", text)
    text = text.strip("_")
    return text or "未命名列"


def deduplicate_columns(columns: Iterable[object]) -> List[str]:
    seen = {}
    result = []
    for raw in columns:
        base = sanitize_header(raw)
        if not base:
            base = "未命名列"
        count = seen.get(base, 0)
        seen[base] = count + 1
        result.append(base if count == 0 else f"{base}_{count}")
    return result


def drop_empty_edges(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(how="all")
    empty_cols = [col for col in df.columns if df[col].isna().all()]
    if empty_cols:
        df = df.drop(columns=empty_cols)
    return df.reset_index(drop=True)


def read_first_sheet_candidates(
    file_bytes: bytes,
    header_candidates: Iterable[int],
    max_rows: int | None = None,
) -> List[ExcelTable]:
    excel = pd.ExcelFile(io.BytesIO(file_bytes))
    if not excel.sheet_names:
        return []

    sheet_name = excel.sheet_names[0]
    candidates: List[ExcelTable] = []
    for header in header_candidates:
        try:
            df = pd.read_excel(
                io.BytesIO(file_bytes),
                sheet_name=sheet_name,
                header=header,
                nrows=max_rows,
            )
        except Exception:
            continue

        df = drop_empty_edges(df)
        if df.empty or len(df.columns) == 0:
            continue

        original_columns = ["" if col is None else str(col) for col in df.columns]
        df.columns = deduplicate_columns(df.columns)
        useful_columns = [col for col in df.columns if not col.startswith("未命名列")]
        if not useful_columns:
            continue

        candidates.append(
            ExcelTable(
                dataframe=df,
                sheet_name=sheet_name,
                header_row_index=header,
                original_columns=original_columns,
            )
        )
    return candidates


def dataframe_preview(df: pd.DataFrame, rows: int = 5) -> List[dict]:
    return df.head(rows).fillna("").astype(str).to_dict(orient="records")
