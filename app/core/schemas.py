from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List


@dataclass(frozen=True)
class FieldSpec:
    key: str
    label: str
    description: str
    aliases: List[str]
    data_type: str = "text"
    required: bool = False


COMMON_FIELDS = {
    "part_no": FieldSpec(
        key="part_no",
        label="零件号",
        description="用于连接主清单和归档/发布清单的零件编号",
        aliases=[
            "零件号",
            "零件编号",
            "物料号",
            "part no",
            "part_no",
            "part number",
            "part_number",
            "pn",
            "part id",
        ],
        data_type="part_no",
        required=True,
    ),
    "operation": FieldSpec(
        key="operation",
        label="操作类型",
        description="A/U/D 等变更操作类型，D 行不进入未完成统计",
        aliases=["操作类型", "操作", "op type", "operation", "change type", "变更类型"],
        data_type="operation",
    ),
    "group": FieldSpec(
        key="group",
        label="功能组",
        description="功能组、部门或专业组，用于分组统计",
        aliases=["功能组", "部门", "专业组", "function group", "group", "dept", "department"],
    ),
    "engineer": FieldSpec(
        key="engineer",
        label="工程师",
        description="负责工程师或责任人",
        aliases=["工程师", "工程师姓名", "设计工程师", "负责人", "owner", "engineer"],
    ),
    "product": FieldSpec(
        key="product",
        label="产品",
        description="产品或车型平台",
        aliases=["产品", "车型", "平台", "product", "vehicle", "program"],
    ),
    "project": FieldSpec(
        key="project",
        label="项目",
        description="首次申请项目、首次应用项目或项目名称",
        aliases=["首次申请项目", "首次应用项目", "项目", "项目名称", "project", "application"],
    ),
}


ESO_MAIN_FIELDS = [
    COMMON_FIELDS["part_no"],
    COMMON_FIELDS["operation"],
    COMMON_FIELDS["group"],
    COMMON_FIELDS["engineer"],
    COMMON_FIELDS["product"],
    COMMON_FIELDS["project"],
    FieldSpec(
        key="description",
        label="FFC中文描述",
        description="零件描述或中文描述",
        aliases=["FFC中文描述", "中文描述", "零件描述", "描述", "description", "part description"],
    ),
    FieldSpec(
        key="eso_plan_date",
        label="ESO计划日期",
        description="ESO 计划完成或计划归档日期",
        aliases=[
            "ESO_Plan_Date",
            "ESO Plan Date",
            "ESO计划日期",
            "ESO计划完成日期",
            "计划归档日期",
            "ESO材料送检计划",
        ],
        data_type="date",
        required=True,
    ),
    FieldSpec(
        key="eso_actual_date",
        label="ESO实际归档日期",
        description="ESO 实际完成、实际归档或发布时间",
        aliases=[
            "ESO_Actual_Date",
            "ESO Actual Date",
            "ESO实际归档日期",
            "ESO实际日期",
            "实际归档日期",
            "归档日期",
            "发布日期",
        ],
        data_type="date",
        required=True,
    ),
    FieldSpec(
        key="eso_note",
        label="ESO备注",
        description="ESO 备注、类型或说明",
        aliases=["ESO备注", "ESO备注_类型", "备注", "note", "remark"],
    ),
]

ESO_ARCHIVE_FIELDS = [
    COMMON_FIELDS["part_no"],
    FieldSpec(
        key="archive_date",
        label="归档日期",
        description="已归档清单中的归档日期，用于回填 ESO 实际归档日期",
        aliases=["归档日期", "发布日期", "发布时间", "实际归档日期", "archive date", "release date"],
        data_type="date",
        required=True,
    ),
]

DRAWING_MAIN_FIELDS = [
    COMMON_FIELDS["part_no"],
    COMMON_FIELDS["operation"],
    COMMON_FIELDS["group"],
    COMMON_FIELDS["engineer"],
    COMMON_FIELDS["product"],
    COMMON_FIELDS["project"],
    FieldSpec(
        key="drawing_no",
        label="图纸号",
        description="2D 图纸编号",
        aliases=["图纸号", "2D_Drawing_No.", "2D_Drawing_No_", "Drawing No", "drawing number"],
    ),
    FieldSpec(
        key="model_plan_date",
        label="数模计划日期",
        description="数模或 TG2 计划完成日期",
        aliases=["数模计划日期", "数模_Plan_Date", "TG2_Plan_Date", "TG2 Plan Date", "model plan date"],
        data_type="date",
        required=True,
    ),
    FieldSpec(
        key="model_actual_date",
        label="数模实际日期",
        description="数模或 TG2 实际完成日期",
        aliases=["数模实际日期", "数模发布日期", "TG2_Actual_Date", "TG2 Actual Date", "model actual date"],
        data_type="date",
        required=True,
    ),
    FieldSpec(
        key="drawing_plan_date",
        label="图纸计划日期",
        description="图纸计划发布日期，可作为数模实际日期缺失时的兜底",
        aliases=["图纸计划日期", "图纸_Plan_Date", "2D_Drawing_Plan_Date", "drawing plan date"],
        data_type="date",
    ),
    FieldSpec(
        key="drawing_actual_date",
        label="图纸实际日期",
        description="图纸实际发布或归档日期",
        aliases=[
            "图纸实际日期",
            "图纸发布日期",
            "图纸发布实际日期",
            "2D_Drawing_Actual_Date",
            "Drawing Actual Date",
            "drawing release date",
        ],
        data_type="date",
    ),
]

DRAWING_ARCHIVE_FIELDS = [
    COMMON_FIELDS["part_no"],
    FieldSpec(
        key="drawing_actual_date",
        label="图纸实际日期",
        description="图纸发布/归档清单中的图纸实际日期",
        aliases=["图纸实际日期", "图纸发布日期", "归档日期", "发布日期", "发布时间", "release date"],
        data_type="date",
        required=True,
    ),
]


SCENES: Dict[str, Dict[str, object]] = {
    "eso": {
        "label": "ESO未完成清单",
        "primary_fields": ESO_MAIN_FIELDS,
        "archive_fields": ESO_ARCHIVE_FIELDS,
        "primary_header_candidates": [1, 0, 2, 3],
        "archive_header_candidates": [0, 1, 2],
    },
    "drawing": {
        "label": "图纸/数模未完成清单",
        "primary_fields": DRAWING_MAIN_FIELDS,
        "archive_fields": DRAWING_ARCHIVE_FIELDS,
        "primary_header_candidates": [1, 0, 2, 3],
        "archive_header_candidates": [0, 1, 2],
    },
}


def get_scene_config(scene: str) -> Dict[str, object]:
    if scene not in SCENES:
        raise ValueError(f"Unsupported scene: {scene}")
    return SCENES[scene]
