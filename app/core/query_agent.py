from __future__ import annotations

from collections import Counter
from typing import Any, Dict, List

from .llm_client import call_llm_json, llm_enabled


FIELD_LABELS = {
    "part_no": ["零件号"],
    "engineer": ["工程师"],
    "group": ["功能组"],
    "product": ["产品"],
    "project": ["项目"],
    "type": ["未完成类型"],
    "status": ["状态"],
}


def make_table(records: List[Dict[str, Any]]) -> Dict[str, Any]:
    columns = []
    if records:
        columns = [{"prop": key, "label": key} for key in records[0].keys()]
    return {"records": records, "columns": columns, "total": len(records)}


def get_value(row: Dict[str, Any], field: str) -> str:
    for label in FIELD_LABELS.get(field, []):
        value = row.get(label)
        if value not in (None, ""):
            return str(value)
    return ""


def count_by(rows: List[Dict[str, Any]], field: str, label: str) -> List[Dict[str, Any]]:
    counter = Counter(get_value(row, field) or f"未知{label}" for row in rows)
    return [{label: key, "未完成数量": value} for key, value in counter.most_common()]


def filter_by(rows: List[Dict[str, Any]], field: str, value: str) -> List[Dict[str, Any]]:
    value_lower = value.lower()
    return [row for row in rows if value_lower in get_value(row, field).lower()]


def find_value_in_question(rows: List[Dict[str, Any]], field: str, question: str) -> str | None:
    values = sorted({get_value(row, field) for row in rows if get_value(row, field)}, key=len, reverse=True)
    lower_question = question.lower()
    for value in values:
        if value in question or value.lower() in lower_question:
            return value
    return None


def local_plan(question: str, rows: List[Dict[str, Any]], scene: str) -> Dict[str, Any]:
    text = question.strip()
    lower = text.lower()
    wants_top = any(token in lower for token in ["top", "最多", "排名", "前五", "前5", "前十", "前10"])
    wants_count = any(token in text for token in ["统计", "数量", "多少", "几个", "汇总", "总数"])
    wants_detail = any(token in text for token in ["清单", "明细", "列出", "有哪些", "告诉我", "全部"])

    if any(token in text for token in ["帮助", "怎么问", "示例", "样例"]):
        return {"tool": "help"}

    field_keywords = [
        ("engineer", "工程师", ["工程师", "负责人", "engineer"]),
        ("group", "功能组", ["功能组", "部门", "专业组", "group"]),
        ("type", "未完成类型", ["未完成类型", "类型", "数模", "图纸"]),
        ("product", "产品", ["产品", "车型", "平台"]),
        ("project", "项目", ["项目", "首次应用", "首次申请"]),
        ("status", "状态", ["状态"]),
    ]
    for field, label, keywords in field_keywords:
        matched_value = find_value_in_question(rows, field, text)
        if matched_value:
            return {"tool": "filter", "field": field, "label": label, "value": matched_value}
        if any(keyword.lower() in lower for keyword in keywords):
            return {"tool": "count_by" if wants_count or not wants_detail else "list_by", "field": field, "label": label, "limit": 10 if wants_top else 500}

    part_value = find_value_in_question(rows, "part_no", text)
    if part_value:
        return {"tool": "filter", "field": "part_no", "label": "零件号", "value": part_value}
    if "零件号" in text or "part" in lower:
        return {"tool": "list_parts", "limit": 500}

    if wants_count and not wants_detail:
        return {"tool": "summary"}
    return {"tool": "details", "limit": 500}


def llm_plan(question: str, rows: List[Dict[str, Any]], scene: str) -> Dict[str, Any] | None:
    if not llm_enabled("ENABLE_LLM_QUERY"):
        return None
    values = {
        field: sorted({get_value(row, field) for row in rows if get_value(row, field)})[:80]
        for field in FIELD_LABELS
    }
    system_prompt = (
        "你是受控数据问答规划器。只把用户问题转换为 JSON 查询计划，不直接回答。"
        "可用 tool: summary, details, list_parts, count_by, filter, help。"
        "field 只能是 part_no, engineer, group, product, project, type, status。"
        "如果要筛选，value 必须来自候选值或原问题明确给出。"
    )
    payload = {"scene": scene, "question": question, "available_values": values}
    response = call_llm_json(system_prompt, payload)
    if not response or response.get("tool") not in {"summary", "details", "list_parts", "count_by", "filter", "help"}:
        return None
    return response


def execute_plan(plan: Dict[str, Any], result: Dict[str, Any]) -> Dict[str, Any]:
    rows = result.get("rows", [])
    scene_label = "ESO" if result.get("scene") == "eso" else "图纸/数模"
    tool = plan.get("tool")

    if tool == "help":
        records = [
            {"可问问题": "当前未完成总数是多少？"},
            {"可问问题": "按工程师统计未完成数量"},
            {"可问问题": "按功能组统计未完成数量"},
            {"可问问题": "列出未完成零件号"},
            {"可问问题": "某个工程师的未完成明细"},
            {"可问问题": "按未完成类型统计"},
        ]
        return {"answer": "你可以这样查询当前批次。", "table_data": make_table(records), "plan": plan}

    if tool == "summary":
        records = [{"指标": key, "数值": value} for key, value in result.get("summary", {}).items()]
        return {"answer": f"{scene_label}当前未完成数量为 {len(rows)} 项。", "table_data": make_table(records), "plan": plan}

    if tool == "list_parts":
        records = [{"零件号": get_value(row, "part_no")} for row in rows if get_value(row, "part_no")]
        return {"answer": f"当前批次共有 {len(records)} 个未完成零件号。", "table_data": make_table(records), "plan": plan}

    if tool in {"count_by", "list_by"}:
        field = plan.get("field") or "group"
        label = plan.get("label") or "维度"
        records = count_by(rows, field, label)
        limit = int(plan.get("limit") or 500)
        return {
            "answer": f"已按{label}统计，合计 {len(rows)} 项。",
            "table_data": make_table(records[:limit]),
            "plan": plan,
        }

    if tool == "filter":
        field = plan.get("field") or "part_no"
        label = plan.get("label") or "字段"
        value = str(plan.get("value") or "")
        records = filter_by(rows, field, value)
        return {
            "answer": f"{label}包含“{value}”的未完成明细共 {len(records)} 条。",
            "table_data": make_table(records),
            "plan": plan,
        }

    limit = int(plan.get("limit") or 500)
    return {"answer": f"已返回当前批次未完成明细 {min(len(rows), limit)} 条。", "table_data": make_table(rows[:limit]), "plan": plan}


def answer_question(question: str, result: Dict[str, Any]) -> Dict[str, Any]:
    rows = result.get("rows", [])
    plan = llm_plan(question, rows, result.get("scene", "eso")) or local_plan(question, rows, result.get("scene", "eso"))
    response = execute_plan(plan, result)
    response["session_id"] = result.get("session_id")
    response["scene"] = result.get("scene")
    return response
