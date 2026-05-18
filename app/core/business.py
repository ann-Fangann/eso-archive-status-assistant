from __future__ import annotations

from collections import Counter
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd

from .field_mapper import is_effective_value, mapping_dict


def normalize_part_no(value: object) -> str:
    if not is_effective_value(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text.replace(" ", "")


def parse_date(value: object) -> date | None:
    if not is_effective_value(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.date()
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and 1 <= float(value) <= 60000:
        parsed = pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
        return None if pd.isna(parsed) else parsed.date()
    text = str(value).strip().split(" ")[0].replace("/", "-")
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y%m%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    parsed = pd.to_datetime(value, errors="coerce")
    return None if pd.isna(parsed) else parsed.date()


def format_date(value: date | None) -> str:
    return value.isoformat() if isinstance(value, date) else ""


def add_one_month(value: date) -> date:
    return (pd.Timestamp(value) + pd.DateOffset(months=1)).date()


def resolve_target_date(target_date: str | None) -> date:
    if not target_date:
        return date.today() - timedelta(days=1)
    return datetime.strptime(target_date, "%Y-%m-%d").date()


def get_cell(row: pd.Series, columns: Dict[str, str], field: str) -> Any:
    column = columns.get(field)
    if not column:
        return ""
    return row.get(column, "")


def build_latest_date_map(df: pd.DataFrame | None, mapping: List[Dict[str, Any]], date_field: str) -> Dict[str, date]:
    if df is None or df.empty:
        return {}
    columns = mapping_dict(mapping)
    part_col = columns.get("part_no")
    date_col = columns.get(date_field)
    if not part_col or not date_col:
        return {}

    result: Dict[str, date] = {}
    for _, row in df.iterrows():
        part_no = normalize_part_no(row.get(part_col))
        parsed = parse_date(row.get(date_col))
        if not part_no or not parsed:
            continue
        if part_no not in result or parsed >= result[part_no]:
            result[part_no] = parsed
    return result


def row_base(row: pd.Series, columns: Dict[str, str], source_row: int) -> Dict[str, Any]:
    return {
        "源行号": source_row,
        "零件号": normalize_part_no(get_cell(row, columns, "part_no")),
        "产品": str(get_cell(row, columns, "product") or ""),
        "项目": str(get_cell(row, columns, "project") or ""),
        "功能组": str(get_cell(row, columns, "group") or ""),
        "工程师": str(get_cell(row, columns, "engineer") or ""),
        "操作类型": str(get_cell(row, columns, "operation") or ""),
    }


def group_stats(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    counter = Counter(str(row.get("功能组") or "未知功能组") for row in rows)
    return [{"功能组": key, "未完成数量": value} for key, value in counter.most_common()]


def type_stats(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    counter = Counter(str(row.get("未完成类型") or "未知类型") for row in rows)
    return [{"未完成类型": key, "数量": value} for key, value in counter.most_common()]


def analyze_eso(
    primary_df: pd.DataFrame,
    primary_mapping: List[Dict[str, Any]],
    archive_df: pd.DataFrame | None,
    archive_mapping: List[Dict[str, Any]],
    target_date: date,
) -> Dict[str, Any]:
    columns = mapping_dict(primary_mapping)
    archive_dates = build_latest_date_map(archive_df, archive_mapping, "archive_date")

    rows: List[Dict[str, Any]] = []
    planned_count = 0
    completed_count = 0
    due_planned_count = 0
    due_completed_count = 0
    delete_row_count = 0
    delete_completed_count = 0
    delete_unfinished_excluded_count = 0
    backfilled_count = 0

    for index, row in primary_df.iterrows():
        operation = str(get_cell(row, columns, "operation") or "").strip().upper()
        is_delete = operation == "D"
        part_no = normalize_part_no(get_cell(row, columns, "part_no"))
        plan_date = parse_date(get_cell(row, columns, "eso_plan_date"))
        actual_date = parse_date(get_cell(row, columns, "eso_actual_date"))

        archive_date = archive_dates.get(part_no)
        if not is_delete and archive_date and (not actual_date or archive_date > actual_date):
            actual_date = archive_date
            backfilled_count += 1

        has_plan = plan_date is not None
        due = has_plan and plan_date <= target_date
        done = actual_date is not None

        if is_delete:
            delete_row_count += 1
            if done:
                delete_completed_count += 1
            if due and not done:
                delete_unfinished_excluded_count += 1
            continue

        if has_plan:
            planned_count += 1
            if done:
                completed_count += 1
        if due:
            due_planned_count += 1
            if done:
                due_completed_count += 1

        if due and not done:
            item = row_base(row, columns, int(index) + 2)
            item.update(
                {
                    "描述": str(get_cell(row, columns, "description") or ""),
                    "未完成类型": "ESO未完成",
                    "状态": "审批中",
                    "计划日期": format_date(plan_date),
                    "实际日期": "",
                    "到期日期": format_date(plan_date),
                    "未完成原因": "ESO计划日期已到期，但实际归档日期为空",
                }
            )
            rows.append(item)

    summary = {
        "统计日期": target_date.isoformat(),
        "计划数量": planned_count,
        "截至统计日期计划数量": due_planned_count,
        "已完成数量": completed_count,
        "截至统计日期已完成数量": due_completed_count,
        "未完成数量": len(rows),
        "本次按归档清单可回填数量": backfilled_count,
        "D行总数": delete_row_count,
        "D行已有实际日期数量": delete_completed_count,
        "D行排除未完成数量": delete_unfinished_excluded_count,
        "统计口径": "ESO计划日期 <= 统计日期，ESO实际归档日期为空，且操作类型不是 D",
    }
    return {
        "summary": summary,
        "rows": rows,
        "group_stats": group_stats(rows),
        "type_stats": type_stats(rows),
    }


def analyze_drawing(
    primary_df: pd.DataFrame,
    primary_mapping: List[Dict[str, Any]],
    archive_df: pd.DataFrame | None,
    archive_mapping: List[Dict[str, Any]],
    target_date: date,
) -> Dict[str, Any]:
    columns = mapping_dict(primary_mapping)
    archive_dates = build_latest_date_map(archive_df, archive_mapping, "drawing_actual_date")

    rows: List[Dict[str, Any]] = []
    delete_excluded_count = 0
    model_due_count = 0
    model_completed_count = 0
    drawing_due_count = 0
    drawing_completed_count = 0
    backfilled_count = 0

    for index, row in primary_df.iterrows():
        operation = str(get_cell(row, columns, "operation") or "").strip().upper()
        if operation == "D":
            delete_excluded_count += 1
            continue

        part_no = normalize_part_no(get_cell(row, columns, "part_no"))
        if not part_no:
            continue

        model_plan = parse_date(get_cell(row, columns, "model_plan_date"))
        model_actual = parse_date(get_cell(row, columns, "model_actual_date"))
        drawing_plan = parse_date(get_cell(row, columns, "drawing_plan_date"))
        drawing_actual = parse_date(get_cell(row, columns, "drawing_actual_date"))
        archive_actual = archive_dates.get(part_no)
        if archive_actual and not drawing_actual:
            drawing_actual = archive_actual
            backfilled_count += 1

        if model_actual:
            model_completed_count += 1
        if drawing_actual:
            drawing_completed_count += 1

        base = row_base(row, columns, int(index) + 2)
        base.update(
            {
                "图纸号": str(get_cell(row, columns, "drawing_no") or ""),
                "数模计划日期": format_date(model_plan),
                "数模实际日期": format_date(model_actual),
                "图纸实际日期": format_date(drawing_actual),
            }
        )

        if model_plan and model_plan <= target_date:
            model_due_count += 1

        if model_plan and model_plan <= target_date and not model_actual:
            rows.append(
                {
                    **base,
                    "未完成类型": "数模未完成",
                    "状态": "计划已到期未完成",
                    "到期日期": format_date(model_plan),
                    "未完成原因": "数模计划日期已到期，但数模实际日期为空",
                }
            )
            rows.append(
                {
                    **base,
                    "未完成类型": "图纸未发布",
                    "状态": "受数模未完成影响",
                    "到期日期": "待数模完成后一个月",
                    "未完成原因": "数模未完成，图纸发布无法闭环",
                }
            )
            continue

        drawing_due = add_one_month(model_actual) if model_actual else drawing_plan
        if drawing_due and drawing_due <= target_date:
            drawing_due_count += 1
        if drawing_due and drawing_due <= target_date and not drawing_actual:
            rows.append(
                {
                    **base,
                    "未完成类型": "图纸未发布",
                    "状态": "到期未发布",
                    "到期日期": format_date(drawing_due),
                    "未完成原因": "图纸要求完成日期已到期，但图纸实际日期为空",
                }
            )

    summary = {
        "统计日期": target_date.isoformat(),
        "未完成数量": len(rows),
        "数模到期数量": model_due_count,
        "数模已完成数量": model_completed_count,
        "图纸到期数量": drawing_due_count,
        "图纸已完成数量": drawing_completed_count,
        "本次按发布清单可回填数量": backfilled_count,
        "D行排除数量": delete_excluded_count,
        "统计口径": "数模计划日期 <= 统计日期且数模未完成，或数模完成后一个月内图纸未发布；操作类型 D 不进入未完成清单",
    }
    return {
        "summary": summary,
        "rows": rows,
        "group_stats": group_stats(rows),
        "type_stats": type_stats(rows),
    }


def export_result_to_excel(result: Dict[str, Any], output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    scene = result.get("scene", "analysis")
    target = str(result.get("target_date") or "").replace("-", "")
    path = output_dir / f"{scene}-unfinished-{target}-{result.get('session_id', 'latest')[:8]}.xlsx"

    summary_rows = [{"指标": key, "数值": value} for key, value in result.get("summary", {}).items()]
    mapping_rows = []
    for role, mappings in result.get("mappings", {}).items():
        for item in mappings:
            mapping_rows.append(
                {
                    "文件角色": "主清单" if role == "primary" else "归档/发布清单",
                    "标准字段": item.get("label"),
                    "匹配列": item.get("column") or "",
                    "置信度": item.get("confidence"),
                    "分数": item.get("score"),
                    "证据": item.get("evidence"),
                }
            )

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame(summary_rows).to_excel(writer, index=False, sheet_name="汇总")
        pd.DataFrame(result.get("rows", [])).to_excel(writer, index=False, sheet_name="未完成清单")
        pd.DataFrame(result.get("group_stats", [])).to_excel(writer, index=False, sheet_name="功能组统计")
        pd.DataFrame(result.get("type_stats", [])).to_excel(writer, index=False, sheet_name="类型统计")
        pd.DataFrame(mapping_rows).to_excel(writer, index=False, sheet_name="字段映射")
    return path
