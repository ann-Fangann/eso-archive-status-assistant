from __future__ import annotations

import math
import re
from dataclasses import asdict
from difflib import SequenceMatcher
from typing import Any, Dict, Iterable, List

import pandas as pd

from .llm_client import call_llm_json, llm_enabled
from .schemas import FieldSpec


EMPTY_MARKERS = {"", "NA", "N/A", "NANA", "NONE", "NULL", "NAN", "NAT", "<NA>", "-", "--"}


def normalize_text(value: object) -> str:
    text = "" if value is None else str(value)
    text = text.strip().lower()
    text = text.replace("_", " ").replace("-", " ").replace(".", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def is_effective_value(value: object) -> bool:
    if value is None:
        return False
    try:
        if pd.isna(value):
            return False
    except (TypeError, ValueError):
        pass
    return str(value).strip().upper() not in EMPTY_MARKERS


def parseable_date_ratio(values: Iterable[object]) -> float:
    total = 0
    hits = 0
    for value in values:
        if not is_effective_value(value):
            continue
        total += 1
        if isinstance(value, (pd.Timestamp,)):
            hits += 1
            continue
        parsed = pd.to_datetime(value, errors="coerce")
        if not pd.isna(parsed):
            hits += 1
    return hits / total if total else 0.0


def part_no_ratio(values: Iterable[object]) -> float:
    total = 0
    hits = 0
    for value in values:
        if not is_effective_value(value):
            continue
        total += 1
        text = str(value).strip()
        compact = text.replace(" ", "")
        if re.fullmatch(r"[A-Za-z0-9_.-]{4,}", compact) or re.fullmatch(r"\d{5,}", compact):
            hits += 1
    return hits / total if total else 0.0


def operation_ratio(values: Iterable[object]) -> float:
    total = 0
    hits = 0
    allowed = {"A", "U", "D", "N", "ADD", "UPDATE", "DELETE"}
    for value in values:
        if not is_effective_value(value):
            continue
        total += 1
        if str(value).strip().upper() in allowed:
            hits += 1
    return hits / total if total else 0.0


def sample_values(series: pd.Series, limit: int = 6) -> List[str]:
    values = []
    for value in series.tolist():
        if is_effective_value(value):
            text = str(value).strip()
            if text not in values:
                values.append(text[:60])
            if len(values) >= limit:
                break
    return values


def header_similarity(column: str, spec: FieldSpec) -> tuple[float, str]:
    column_norm = normalize_text(column)
    best_score = 0.0
    best_reason = ""
    for alias in [spec.label, *spec.aliases]:
        alias_norm = normalize_text(alias)
        if not alias_norm:
            continue
        if column_norm == alias_norm:
            return 1.0, f"表头与别名“{alias}”完全一致"
        if alias_norm in column_norm or column_norm in alias_norm:
            score = 0.88 if len(alias_norm) > 2 else 0.72
            if score > best_score:
                best_score = score
                best_reason = f"表头包含别名“{alias}”"
        ratio = SequenceMatcher(None, column_norm, alias_norm).ratio()
        if ratio > best_score and ratio >= 0.62:
            best_score = ratio * 0.82
            best_reason = f"表头与别名“{alias}”相似"
    return best_score, best_reason


def value_hint_score(series: pd.Series, spec: FieldSpec) -> tuple[float, str]:
    values = series.dropna().head(80).tolist()
    if spec.data_type == "date":
        ratio = parseable_date_ratio(values)
        return min(ratio, 1.0), f"样本中 {round(ratio * 100)}% 可解析为日期"
    if spec.data_type == "part_no":
        ratio = part_no_ratio(values)
        return min(ratio, 1.0), f"样本中 {round(ratio * 100)}% 像零件号"
    if spec.data_type == "operation":
        ratio = operation_ratio(values)
        return min(ratio, 1.0), f"样本中 {round(ratio * 100)}% 像操作类型"
    non_empty = sum(1 for value in values if is_effective_value(value))
    ratio = non_empty / len(values) if values else 0.0
    return min(ratio * 0.35, 0.35), f"样本非空率 {round(ratio * 100)}%"


def confidence_label(score: float) -> str:
    if score >= 0.78:
        return "high"
    if score >= 0.58:
        return "medium"
    if score >= 0.42:
        return "low"
    return "missing"


def profile_columns(df: pd.DataFrame) -> List[Dict[str, Any]]:
    profiles = []
    for column in df.columns:
        series = df[column]
        profiles.append(
            {
                "column": str(column),
                "samples": sample_values(series),
                "non_empty": int(series.apply(is_effective_value).sum()),
            }
        )
    return profiles


def local_map_fields(df: pd.DataFrame, specs: List[FieldSpec]) -> List[Dict[str, Any]]:
    candidates_by_field: Dict[str, List[Dict[str, Any]]] = {}
    for spec in specs:
        candidates = []
        for column in df.columns:
            header_score, header_reason = header_similarity(str(column), spec)
            value_score, value_reason = value_hint_score(df[column], spec)
            score = min(1.0, header_score * 0.82 + value_score * 0.18)
            if header_score < 0.1 and value_score < 0.75:
                score *= 0.5
            evidence = header_reason or value_reason
            if header_reason and value_reason:
                evidence = f"{header_reason}；{value_reason}"
            candidates.append(
                {
                    "field_key": spec.key,
                    "label": spec.label,
                    "column": str(column),
                    "score": round(score, 4),
                    "evidence": evidence,
                    "samples": sample_values(df[column], limit=3),
                    "required": spec.required,
                }
            )
        candidates_by_field[spec.key] = sorted(candidates, key=lambda item: item["score"], reverse=True)

    assigned = set()
    output: List[Dict[str, Any]] = []
    for spec in sorted(specs, key=lambda item: item.required, reverse=True):
        chosen = None
        for candidate in candidates_by_field[spec.key]:
            if candidate["column"] in assigned:
                continue
            if candidate["score"] >= 0.42:
                chosen = candidate
                break
        if chosen:
            assigned.add(chosen["column"])
            score = chosen["score"]
            output.append(
                {
                    **chosen,
                    "confidence": confidence_label(score),
                    "source": "local_semantic_mapper",
                }
            )
        else:
            output.append(
                {
                    "field_key": spec.key,
                    "label": spec.label,
                    "column": None,
                    "score": 0.0,
                    "confidence": "missing",
                    "evidence": "没有找到足够可信的候选列",
                    "samples": [],
                    "required": spec.required,
                    "source": "local_semantic_mapper",
                }
            )
    order = {spec.key: idx for idx, spec in enumerate(specs)}
    return sorted(output, key=lambda item: order[item["field_key"]])


def merge_llm_mapping(
    local_mapping: List[Dict[str, Any]],
    df: pd.DataFrame,
    specs: List[FieldSpec],
    scene: str,
    role: str,
) -> List[Dict[str, Any]]:
    if not llm_enabled("ENABLE_LLM_MAPPING"):
        return local_mapping

    low_confidence = [item for item in local_mapping if item["confidence"] in {"low", "missing"}]
    if not low_confidence:
        return local_mapping

    system_prompt = (
        "你是企业数据治理助手。你的任务只是在 Excel 表头和标准业务字段之间做语义映射，"
        "不得做业务计算。只返回 JSON：{\"mappings\":[{\"field_key\":\"...\",\"column\":\"...\","
        "\"confidence\":0.0,\"reason\":\"...\"}]}。column 必须来自候选列，无法判断则填 null。"
    )
    payload = {
        "scene": scene,
        "role": role,
        "fields": [asdict(spec) for spec in specs],
        "columns": profile_columns(df),
        "local_mapping": local_mapping,
    }
    response = call_llm_json(system_prompt, payload)
    if not response or not isinstance(response.get("mappings"), list):
        return local_mapping

    valid_columns = {str(col) for col in df.columns}
    by_key = {item["field_key"]: dict(item) for item in local_mapping}
    for item in response["mappings"]:
        field_key = item.get("field_key")
        column = item.get("column")
        if field_key not in by_key:
            continue
        if column is None:
            continue
        column = str(column)
        if column not in valid_columns:
            continue
        confidence = float(item.get("confidence") or 0)
        if math.isnan(confidence) or confidence < 0.55:
            continue
        existing = by_key[field_key]
        if existing.get("confidence") in {"high", "medium"} and existing.get("score", 0) >= confidence:
            continue
        existing.update(
            {
                "column": column,
                "score": round(min(confidence, 0.92), 4),
                "confidence": confidence_label(confidence),
                "evidence": f"LLM语义映射：{item.get('reason') or '根据表头和样本判断'}",
                "source": "llm_semantic_mapper",
                "samples": sample_values(df[column], limit=3),
            }
        )
    return list(by_key.values())


def map_fields(df: pd.DataFrame, specs: List[FieldSpec], scene: str, role: str) -> List[Dict[str, Any]]:
    local = local_map_fields(df, specs)
    return merge_llm_mapping(local, df, specs, scene, role)


def mapping_dict(mapping: List[Dict[str, Any]]) -> Dict[str, str]:
    return {
        item["field_key"]: item["column"]
        for item in mapping
        if item.get("column")
    }


def mapping_score(mapping: List[Dict[str, Any]]) -> float:
    if not mapping:
        return 0.0
    required = [item for item in mapping if item.get("required")]
    required_score = sum(float(item.get("score") or 0) for item in required) / len(required) if required else 1.0
    all_score = sum(float(item.get("score") or 0) for item in mapping) / len(mapping)
    missing_required = sum(1 for item in required if not item.get("column"))
    penalty = missing_required * 0.45
    return max(0.0, required_score * 0.72 + all_score * 0.28 - penalty)


def missing_required_fields(mapping: List[Dict[str, Any]]) -> List[str]:
    return [item["label"] for item in mapping if item.get("required") and not item.get("column")]
