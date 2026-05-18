from __future__ import annotations

import io
import os
from pathlib import Path
from typing import Any, Dict, List

from dotenv import load_dotenv
from fastapi import Body, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from .core.business import analyze_drawing, analyze_eso, export_result_to_excel, resolve_target_date
from .core.excel_reader import ExcelTable, dataframe_preview, read_first_sheet_candidates
from .core.field_mapper import map_fields, mapping_score, missing_required_fields
from .core.query_agent import answer_question
from .core.schemas import get_scene_config
from .core.session_store import store


load_dotenv()

ROOT_DIR = Path(__file__).resolve().parents[1]
STATIC_DIR = ROOT_DIR / "static"
FRONTEND_DIST_DIR = ROOT_DIR / "eso" / "dist"
OUTPUT_DIR = ROOT_DIR / "outputs"

app = FastAPI(title="ESO v3.0 Intelligent Archive Assistant", version="3.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
app.mount(
    "/assets",
    StaticFiles(directory=FRONTEND_DIST_DIR / "assets", check_dir=False),
    name="frontend-assets",
)

COMPAT_CONTEXT: Dict[str, Dict[str, Any]] = {}


def validate_excel(file: UploadFile | None, required: bool = True) -> None:
    if file is None:
        if required:
            raise HTTPException(status_code=400, detail="请上传 Excel 文件")
        return
    filename = file.filename or ""
    if not filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail=f"{filename} 不是支持的 Excel 文件")


async def read_upload(file: UploadFile | None) -> bytes | None:
    if file is None:
        return None
    return await file.read()


def upload_from_bytes(filename: str, file_bytes: bytes) -> UploadFile:
    return UploadFile(filename=filename, file=io.BytesIO(file_bytes))


def pick_best_table(
    file_bytes: bytes,
    scene: str,
    role: str,
    header_candidates: List[int],
    fields: List[Any],
) -> tuple[ExcelTable, List[Dict[str, Any]], float]:
    candidates = read_first_sheet_candidates(file_bytes, header_candidates)
    if not candidates:
        raise HTTPException(status_code=400, detail="Excel 第一个工作表为空，或无法识别表头")

    scored = []
    for table in candidates:
        mapping = map_fields(table.dataframe, fields, scene, role)
        scored.append((mapping_score(mapping), table, mapping))
    scored.sort(key=lambda item: item[0], reverse=True)
    score, table, mapping = scored[0]
    return table, mapping, score


def mapping_warnings(role_label: str, mapping: List[Dict[str, Any]], score: float) -> List[str]:
    warnings = []
    missing = missing_required_fields(mapping)
    if missing:
        warnings.append(f"{role_label}缺少必要字段：{', '.join(missing)}")
    low_fields = [item["label"] for item in mapping if item.get("confidence") == "low"]
    if low_fields:
        warnings.append(f"{role_label}有低置信度字段，建议复核：{', '.join(low_fields)}")
    if score < 0.62:
        warnings.append(f"{role_label}整体映射置信度偏低，请检查表头行是否正确")
    return warnings


async def run_analysis(
    scene: str,
    primary_file: UploadFile,
    archive_file: UploadFile | None,
    target_date: str | None,
) -> Dict[str, Any]:
    validate_excel(primary_file, required=True)
    validate_excel(archive_file, required=False)

    config = get_scene_config(scene)
    primary_bytes = await read_upload(primary_file)
    archive_bytes = await read_upload(archive_file)
    if not primary_bytes:
        raise HTTPException(status_code=400, detail="主清单文件为空")

    target = resolve_target_date(target_date)
    primary_table, primary_mapping, primary_score = pick_best_table(
        primary_bytes,
        scene,
        "primary",
        config["primary_header_candidates"],
        config["primary_fields"],
    )
    primary_missing = missing_required_fields(primary_mapping)
    if primary_missing:
        raise HTTPException(
            status_code=422,
            detail={
                "message": f"主清单缺少必要字段：{', '.join(primary_missing)}",
                "mapping": primary_mapping,
                "columns": primary_table.dataframe.columns.tolist(),
                "preview": dataframe_preview(primary_table.dataframe),
            },
        )

    archive_table = None
    archive_mapping: List[Dict[str, Any]] = []
    archive_score = 0.0
    if archive_bytes:
        archive_table, archive_mapping, archive_score = pick_best_table(
            archive_bytes,
            scene,
            "archive",
            config["archive_header_candidates"],
            config["archive_fields"],
        )

    if scene == "eso":
        analysis = analyze_eso(
            primary_table.dataframe,
            primary_mapping,
            archive_table.dataframe if archive_table else None,
            archive_mapping,
            target,
        )
    else:
        analysis = analyze_drawing(
            primary_table.dataframe,
            primary_mapping,
            archive_table.dataframe if archive_table else None,
            archive_mapping,
            target,
        )

    warnings = []
    warnings.extend(mapping_warnings("主清单", primary_mapping, primary_score))
    if archive_table:
        warnings.extend(mapping_warnings("归档/发布清单", archive_mapping, archive_score))

    result = {
        "scene": scene,
        "scene_label": config["label"],
        "target_date": target.isoformat(),
        "source_files": {
            "primary": primary_file.filename,
            "archive": archive_file.filename if archive_file else None,
        },
        "detected_tables": {
            "primary": {
                "sheet_name": primary_table.sheet_name,
                "header_row": primary_table.header_row_index + 1,
                "mapping_score": round(primary_score, 4),
                "columns": primary_table.dataframe.columns.tolist(),
                "preview": dataframe_preview(primary_table.dataframe),
            },
            "archive": {
                "sheet_name": archive_table.sheet_name,
                "header_row": archive_table.header_row_index + 1,
                "mapping_score": round(archive_score, 4),
                "columns": archive_table.dataframe.columns.tolist(),
                "preview": dataframe_preview(archive_table.dataframe),
            }
            if archive_table
            else None,
        },
        "mappings": {
            "primary": primary_mapping,
            "archive": archive_mapping,
        },
        "warnings": warnings,
        **analysis,
    }
    saved = store.save(result)
    saved["export_url"] = f"/api/export/{saved['session_id']}"
    return saved


async def run_analysis_from_bytes(
    scene: str,
    primary_name: str,
    primary_bytes: bytes,
    archive_name: str | None,
    archive_bytes: bytes | None,
    target_date: str | None,
) -> Dict[str, Any]:
    primary_upload = upload_from_bytes(primary_name, primary_bytes)
    archive_upload = upload_from_bytes(archive_name, archive_bytes) if archive_name and archive_bytes else None
    return await run_analysis(scene, primary_upload, archive_upload, target_date)


def _number(value: Any, default: int = 0) -> int:
    try:
        if value is None or value == "":
            return default
        return int(value)
    except (TypeError, ValueError):
        return default


def _compat_group_stats(result: Dict[str, Any]) -> List[Dict[str, Any]]:
    rows = []
    for item in result.get("group_stats", []):
        group_name = item.get("功能组") or item.get("部门") or "未知功能组"
        count = item.get("count", item.get("未完成数量", item.get("数量", 0)))
        rows.append({"功能组": group_name, "count": _number(count)})
    return rows


def _compat_type_stats(result: Dict[str, Any]) -> List[Dict[str, Any]]:
    rows = []
    for item in result.get("type_stats", []):
        type_name = item.get("未完成类型") or "未知类型"
        count = item.get("count", item.get("数量", item.get("未完成数量", 0)))
        rows.append({"未完成类型": type_name, "count": _number(count)})
    return rows


def _compat_eso_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    output = []
    for row in rows:
        item = dict(row)
        item.setdefault("FFC中文描述", item.get("描述", ""))
        item.setdefault("首次申请项目", item.get("项目", ""))
        item.setdefault("首次应用项目", item.get("项目", ""))
        item.setdefault("ESO_Plan_Date", item.get("计划日期", ""))
        item.setdefault("ESO_Actual_Date", item.get("实际日期", ""))
        item.setdefault("ESO状态", item.get("状态", "审批中") or "审批中")
        output.append(item)
    return output


def _compat_summary(result: Dict[str, Any], scene: str) -> Dict[str, Any]:
    summary = result.get("summary", {})
    target_date = summary.get("统计日期") or result.get("target_date", "")
    if scene == "eso":
        return {
            "target_date": target_date,
            "planned_count": _number(summary.get("计划数量")),
            "due_planned_count": _number(summary.get("截至统计日期计划数量")),
            "completed_count": _number(summary.get("已完成数量")),
            "due_completed_count": _number(summary.get("截至统计日期已完成数量")),
            "unfinished_count": _number(summary.get("未完成数量"), len(result.get("rows", []))),
            "matched_completed_count": _number(summary.get("本次按归档清单可回填数量")),
            "delete_row_count": _number(summary.get("D行总数")),
            "delete_completed_count": _number(summary.get("D行已有实际日期数量")),
            "delete_unfinished_excluded_count": _number(summary.get("D行排除未完成数量")),
            "formula": summary.get("统计口径", ""),
            "group_field": "功能组",
        }
    return {
        "target_date": target_date,
        "unfinished_count": _number(summary.get("未完成数量"), len(result.get("rows", []))),
        "model_due_count": _number(summary.get("数模到期数量")),
        "model_completed_count": _number(summary.get("数模已完成数量")),
        "drawing_due_count": _number(summary.get("图纸到期数量")),
        "drawing_completed_count": _number(summary.get("图纸已完成数量")),
        "matched_completed_count": _number(summary.get("本次按发布清单可回填数量")),
        "delete_excluded_count": _number(summary.get("D行排除数量")),
        "formula": summary.get("统计口径", ""),
    }


def to_eso_frontend_select_result(result: Dict[str, Any], scene: str) -> Dict[str, Any]:
    rows = result.get("rows", [])
    if scene == "eso":
        data = _compat_eso_rows(rows)
    else:
        data = [dict(row) for row in rows]
    return {
        "data": data,
        "row_count": len(data),
        "total_empty_count": len(data),
        "summary": _compat_summary(result, scene),
        "group_stats": _compat_group_stats(result),
        "type_stats": _compat_type_stats(result),
        "mappings": result.get("mappings", {}),
        "warnings": result.get("warnings", []),
        "session_id": result.get("session_id"),
    }


def _compat_persistence(result: Dict[str, Any]) -> Dict[str, Any]:
    source = "LLM语义映射 + v3确定性计算"
    if result.get("warnings"):
        return {"saved": True, "message": f"{source}已完成；存在字段映射提醒，请复核。"}
    return {"saved": True, "message": f"{source}已完成。"}


def _compat_download(result: Dict[str, Any]) -> Dict[str, Any]:
    return {
        "download_url": f"/export/{result['session_id']}",
        "filled_count": _number(
            result.get("summary", {}).get("本次按归档清单可回填数量")
            or result.get("summary", {}).get("本次按发布清单可回填数量")
        ),
    }


def infer_scene_from_question(question: str, requested_scene: str | None = None) -> str:
    if requested_scene in {"eso", "drawing"}:
        return requested_scene
    lower = question.lower()
    if "图纸" in question or "数模" in question or "drawing" in lower:
        return "drawing"
    return "eso"


@app.get("/")
async def index() -> FileResponse:
    built_index = FRONTEND_DIST_DIR / "index.html"
    if built_index.exists():
        return FileResponse(built_index)
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/api/health")
async def health() -> Dict[str, Any]:
    return {
        "status": "ok",
        "llm_mapping": os.getenv("ENABLE_LLM_MAPPING", "false"),
        "llm_query": os.getenv("ENABLE_LLM_QUERY", "false"),
        "llm_provider": "minimax",
        "llm_model": os.getenv("MINIMAX_MODEL") or os.getenv("ANTHROPIC_MODEL") or "MiniMax-M2.7-highspeed",
    }


@app.post("/api/analyze/eso")
async def analyze_eso_endpoint(
    primary_file: UploadFile = File(...),
    archive_file: UploadFile | None = File(None),
    target_date: str | None = Form(None),
) -> Dict[str, Any]:
    return await run_analysis("eso", primary_file, archive_file, target_date)


@app.post("/api/analyze/drawing")
async def analyze_drawing_endpoint(
    primary_file: UploadFile = File(...),
    archive_file: UploadFile | None = File(None),
    target_date: str | None = Form(None),
) -> Dict[str, Any]:
    return await run_analysis("drawing", primary_file, archive_file, target_date)


@app.post("/api/upload-two-excel/", summary="ESO兼容入口：上传两个Excel并生成未完成清单")
async def upload_two_excel_compat(
    file1: UploadFile = File(...),
    file2: UploadFile = File(...),
    target_date: str | None = Form(None),
) -> Dict[str, Any]:
    validate_excel(file1, required=True)
    validate_excel(file2, required=True)
    primary_bytes = await read_upload(file1)
    archive_bytes = await read_upload(file2)
    if not primary_bytes or not archive_bytes:
        raise HTTPException(status_code=400, detail="上传文件为空")

    result = await run_analysis_from_bytes(
        "eso",
        file1.filename or "ESO主清单.xlsx",
        primary_bytes,
        file2.filename or "ESO归档清单.xlsx",
        archive_bytes,
        target_date,
    )
    COMPAT_CONTEXT["eso"] = {
        "primary_name": file1.filename or "ESO主清单.xlsx",
        "primary_bytes": primary_bytes,
        "archive_name": file2.filename or "ESO归档清单.xlsx",
        "archive_bytes": archive_bytes,
    }
    select_result = to_eso_frontend_select_result(result, "eso")
    return {
        "file1": file1.filename,
        "file2": file2.filename,
        "target_date": result.get("target_date"),
        "sheets": 2,
        "results": {"LLM字段映射": "已识别", "v3确定性计算": "已完成"},
        "modified_workbook": _compat_download(result),
        "persistence": _compat_persistence(result),
        "sql_results": {
            "update_result": f"成功按 v3 规则回填 {_compat_download(result).get('filled_count', 0)} 条记录",
            "select_result": select_result,
        },
    }


@app.post("/api/unfinished-list/", summary="ESO兼容入口：按统计日期重新生成未完成清单")
async def regenerate_unfinished_list_compat(target_date: str | None = Body(None, embed=True)) -> Dict[str, Any]:
    context = COMPAT_CONTEXT.get("eso")
    if not context:
        raise HTTPException(status_code=400, detail="请先上传并处理 ESO Excel 文件，再重新生成未完成清单")
    result = await run_analysis_from_bytes(
        "eso",
        context["primary_name"],
        context["primary_bytes"],
        context["archive_name"],
        context["archive_bytes"],
        target_date,
    )
    return {
        "target_date": result.get("target_date"),
        "select_result": to_eso_frontend_select_result(result, "eso"),
        "persistence": _compat_persistence(result),
    }


@app.post("/api/upload-drawing-excel/", summary="ESO兼容入口：上传图纸Excel并生成未完成清单")
async def upload_drawing_excel_compat(
    club_file: UploadFile = File(...),
    published_file: UploadFile | None = File(None),
    target_date: str | None = Form(None),
) -> Dict[str, Any]:
    validate_excel(club_file, required=True)
    validate_excel(published_file, required=False)
    primary_bytes = await read_upload(club_file)
    archive_bytes = await read_upload(published_file)
    if not primary_bytes:
        raise HTTPException(status_code=400, detail="上传文件为空")

    result = await run_analysis_from_bytes(
        "drawing",
        club_file.filename or "图纸主清单.xlsx",
        primary_bytes,
        published_file.filename if published_file else None,
        archive_bytes,
        target_date,
    )
    COMPAT_CONTEXT["drawing"] = {
        "primary_name": club_file.filename or "图纸主清单.xlsx",
        "primary_bytes": primary_bytes,
        "archive_name": published_file.filename if published_file else None,
        "archive_bytes": archive_bytes,
    }
    return {
        "club_file": club_file.filename,
        "published_file": published_file.filename if published_file else None,
        "target_date": result.get("target_date"),
        "select_result": to_eso_frontend_select_result(result, "drawing"),
        "modified_workbook": _compat_download(result),
        "persistence": _compat_persistence(result),
    }


@app.post("/api/drawing-unfinished-list/", summary="ESO兼容入口：按统计日期重新生成图纸未完成清单")
async def regenerate_drawing_unfinished_list_compat(target_date: str | None = Body(None, embed=True)) -> Dict[str, Any]:
    context = COMPAT_CONTEXT.get("drawing")
    if not context:
        raise HTTPException(status_code=400, detail="请先上传并处理图纸 Excel 文件，再重新生成未完成清单")
    result = await run_analysis_from_bytes(
        "drawing",
        context["primary_name"],
        context["primary_bytes"],
        context.get("archive_name"),
        context.get("archive_bytes"),
        target_date,
    )
    return {
        "target_date": result.get("target_date"),
        "select_result": to_eso_frontend_select_result(result, "drawing"),
        "modified_workbook": _compat_download(result),
        "persistence": _compat_persistence(result),
    }


@app.post("/api/smart-query/", summary="ESO兼容入口：v3受控智能问答")
async def smart_query_compat(
    question: str = Body(..., embed=True),
    scene: str | None = Body(None, embed=True),
) -> Dict[str, Any]:
    question = question.strip()
    if not question:
        raise HTTPException(status_code=400, detail="问题不能为空")
    resolved_scene = infer_scene_from_question(question, scene)
    result = store.latest(resolved_scene) or store.latest()
    if not result:
        raise HTTPException(status_code=400, detail="请先上传并生成一个分析批次")
    return answer_question(question, result)


@app.get("/api/session/{session_id}")
async def get_session(session_id: str) -> Dict[str, Any]:
    result = store.get(session_id)
    if not result:
        raise HTTPException(status_code=404, detail="未找到该分析会话")
    result["export_url"] = f"/api/export/{session_id}"
    return result


@app.post("/api/chat")
async def chat(payload: Dict[str, Any] = Body(...)) -> Dict[str, Any]:
    question = str(payload.get("question") or "").strip()
    if not question:
        raise HTTPException(status_code=400, detail="请输入问题")
    session_id = payload.get("session_id")
    scene = payload.get("scene")
    result = store.get(session_id) if session_id else store.latest(scene)
    if not result:
        raise HTTPException(status_code=400, detail="请先上传并生成一个分析批次")
    return answer_question(question, result)


@app.get("/api/export/{session_id}")
async def export_excel(session_id: str) -> FileResponse:
    result = store.get(session_id)
    if not result:
        raise HTTPException(status_code=404, detail="未找到该分析会话")
    result["session_id"] = session_id
    path = export_result_to_excel(result, OUTPUT_DIR)
    return FileResponse(
        path,
        filename=path.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
