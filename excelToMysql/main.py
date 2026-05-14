from fastapi import FastAPI, UploadFile, HTTPException, File, Body, Form
from fastapi.responses import FileResponse
import uvicorn
from pydantic import BaseModel
import pandas as pd
import xlsxwriter
import os
import logging
from typing import List, Dict
import json
import io
import re
from datetime import datetime, date, timedelta
from pathlib import Path
from uuid import uuid4
from sqlalchemy import create_engine, MetaData, Table, Column, inspect, text
from sqlalchemy.engine import URL
from sqlalchemy.types import Integer, String, Text, Float
from dotenv import load_dotenv
from openpyxl import load_workbook
import sys
import traceback

# 设置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

# 定义允许的工作表名和字段
ALLOWED_SHEET_NAME = "工程车近14天数据"
REQUIRED_COLUMNS = [
    "vin", "统计日期", "日行里程", "累计里程", "日第一帧信号时间", 
    "日最后一帧信号时间", "日最后省份", "日最后城市", "日最后区域", 
    "日行程数", "日在线时长(小时)", "key_id"
]

app = FastAPI()

OUTPUT_DIR = Path(__file__).resolve().parent / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

EMPTY_MARKERS = {"", "NA", "N/A", "NANA", "NONE", "NULL", "NAN"}

@app.post("/upload-excel/")
async def upload_excel(file: UploadFile):
    # 读取Excel文件
    try:
        # 使用pandas读取Excel文件
        df = pd.read_excel(file.file)
        
        # 检查工作表名称
        if df.shape[0] == 0:
            raise HTTPException(status_code=400, detail="Excel文件为空")
            
        # 获取工作表名称（这里假设只有一个工作表）
        sheet_name = df.columns.name if df.columns.name else "Sheet1"
        if sheet_name != ALLOWED_SHEET_NAME:
            raise HTTPException(status_code=400, detail=f"工作表名称必须为'{ALLOWED_SHEET_NAME}'，当前为'{sheet_name}'")
            
        # 检查字段是否完整
        missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
        if missing_columns:
            raise HTTPException(status_code=400, detail=f"缺少必要字段：{', '.join(missing_columns)}")
            
        # 添加字段说明
        field_descriptions = {
            "vin": "车辆唯一标识符",
            "统计日期": "数据统计的日期",
            "日行里程": "当日行驶的里程数",
            "累计里程": "累计行驶的总里程",
            "日第一帧信号时间": "当日第一次收到信号的时间",
            "日最后一帧信号时间": "当日最后一次收到信号的时间",
            "日最后省份": "当日最后所在省份",
            "日最后城市": "当日最后所在城市",
            "日最后区域": "当日最后所在区域",
            "日行程数": "当日行程次数",
            "日在线时长(小时)": "当日在线时长（单位：小时）",
            "key_id": "唯一标识ID"
        }
        
        # 返回结果
        return {
            "success": True,
            "message": "文件上传成功",
            "field_descriptions": field_descriptions,
            "data_preview": df.head().to_dict(),
            "total_rows": len(df)
        }
        
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))

# 数据库配置。默认值保留公司内网配置，本地开发可在 .env 中覆盖。
load_dotenv()

DB_HOST = os.getenv("DB_HOST", "10.182.37.12")
DB_PORT = int(os.getenv("DB_PORT", "3306"))
DB_NAME = os.getenv("DB_NAME", "ESO")
DB_USER = os.getenv("DB_USER", "root")
DB_PASSWORD = os.getenv("DB_PASSWORD", "123456")
DB_CONNECT_TIMEOUT = int(os.getenv("DB_CONNECT_TIMEOUT", "10"))

# 使用utf8mb4字符集以支持中文
DATABASE_URL = URL.create(
    "mysql+pymysql",
    username=DB_USER,
    password=DB_PASSWORD,
    host=DB_HOST,
    port=DB_PORT,
    database=DB_NAME,
    query={"charset": "utf8mb4"},
)

logger.info(f"数据库配置: {DB_USER}@{DB_HOST}:{DB_PORT}/{DB_NAME}")

# 创建引擎时启用连接池预检查和支持utf8mb4
engine = create_engine(
    DATABASE_URL,
    echo=False,
    pool_pre_ping=True,
    pool_recycle=3600,
    connect_args={"connect_timeout": DB_CONNECT_TIMEOUT},
)

# 安全：处理表名和字段名，但保留中文字符
def sanitize_name(name: str, is_column=False) -> str:
    """
    清理名称以确保数据库兼容性
    对于表名：保留中文字符，但清理可能导致SQL问题的特殊字符
    对于列名：保留中文字符，但清理可能导致SQL问题的特殊字符
    """
    if not isinstance(name, str):
        name = str(name)
    
    # 对于表名和列名，都保留中文、英文字母、数字，移除可能导致问题的特殊字符
    name = re.sub(r'[^\w\u4e00-\u9fff]', '_', name)
    # 确保不以数字开头
    if name and name[0].isdigit():
        name = "t_" + name
    
    # 限制长度
    return name[:64] or "unnamed"

# 推断 SQL 类型 - 简化版本，全部使用字符串类型
def infer_sql_type(series):
    # 所有列都使用VARCHAR(255)类型以避免类型转换问题
    return String(255)

def create_table_with_chunked_columns(table_name, columns, type_mapping):
    """
    当列数过多时，将表按列分组创建多个子表
    """
    MAX_COLS_PER_TABLE = 50  # 每个表最多50列
    
    if len(columns) <= MAX_COLS_PER_TABLE:
        # 列数不多，直接创建一个表
        metadata = MetaData()
        table = Table(
            table_name, metadata, *columns,
            mysql_charset='utf8mb4',
            extend_existing=True,
            mysql_row_format='DYNAMIC',
            mysql_engine='InnoDB'
        )
        metadata.create_all(engine)
        return [table_name]  # 返回表名列表
    else:
        # 列数过多，需要拆分成多个表
        table_names = []
        chunk_num = 0
        
        # 主表包含前50列，包括关键标识列
        main_columns = columns[:MAX_COLS_PER_TABLE-1]  # 保留一列用于ID
        # 确保第一张表包含关键列（如果存在）
        essential_cols = [col for col in columns if '零件号' in col.name or 'ID' in col.name or 'id' in col.name]
        if essential_cols and main_columns[0] != essential_cols[0]:
            # 如果关键列不在前面，则调整顺序
            main_columns = [essential_cols[0]] + columns[1:MAX_COLS_PER_TABLE-1]
        
        chunk_table_name = f"{table_name}_part{chunk_num}"
        table_names.append(chunk_table_name)
        
        metadata = MetaData()
        main_table = Table(
            chunk_table_name, metadata, *main_columns,
            mysql_charset='utf8mb4',
            extend_existing=True,
            mysql_row_format='DYNAMIC',
            mysql_engine='InnoDB'
        )
        metadata.create_all(engine)
        chunk_num += 1
        
        # 处理其余的列
        remaining_columns = columns[MAX_COLS_PER_TABLE-1:]
        for i in range(0, len(remaining_columns), MAX_COLS_PER_TABLE-1):  # 保留一列用于ID关联
            chunk = remaining_columns[i:i+MAX_COLS_PER_TABLE-1]
            chunk_table_name = f"{table_name}_part{chunk_num}"
            table_names.append(chunk_table_name)
            
            # 为了关联，每个分表都需要一个ID列
            id_col = Column('common_id', Integer, primary_key=True)
            chunk_with_id = [id_col] + chunk
            
            metadata = MetaData()
            part_table = Table(
                chunk_table_name, metadata, *chunk_with_id,
                mysql_charset='utf8mb4',
                extend_existing=True,
                mysql_row_format='DYNAMIC',
                mysql_engine='InnoDB'
            )
            metadata.create_all(engine)
            chunk_num += 1
            
        return table_names

def convert_to_int(value):
    """安全地将值转换为整数"""
    if pd.isna(value) or value == '' or str(value).strip() == '':
        return None
    try:
        # 先转换为浮点数，再转换为整数
        float_val = float(value)
        if float_val.is_integer():
            return int(float_val)
        else:
            # 如果不是整数，取整
            return int(float_val)
    except (ValueError, TypeError):
        # 如果不能转换为整数，返回None
        return None

def convert_to_float(value):
    """安全地将值转换为浮点数"""
    if pd.isna(value) or value == '' or str(value).strip() == '':
        return None
    try:
        return float(value)
    except (ValueError, TypeError):
        # 如果不能转换为浮点数，返回None
        return None

def process_date_value(value):
    """
    处理日期值，只保留年月日部分
    """
    if pd.isna(value) or value == '' or value == 'None' or value == 'nan':
        return value

    # 尝试解析为pandas的Timestamp或datetime对象
    try:
        if isinstance(value, (pd.Timestamp, datetime, date)):
            # 只保留年月日部分
            if isinstance(value, (pd.Timestamp, datetime)):
                return value.strftime('%Y-%m-%d')
            else:  # date对象
                return value.strftime('%Y-%m-%d')
    except:
        pass

    # 尝试解析字符串格式的日期（可能包含时间部分）
    try:
        value_str = str(value)
        # 匹配常见的日期格式：YYYY-MM-DD HH:MM:SS 或 YYYY/MM/DD HH:MM:SS 等
        if ' ' in value_str and ('-' in value_str[:10] or '/' in value_str[:10]):
            # 分割日期和时间部分
            date_part = value_str.split(' ')[0]
            # 验证是否是有效的日期格式
            if '-' in date_part and len(date_part) == 10:
                return date_part
            elif '/' in date_part:
                # 将 YYYY/MM/DD 转换为 YYYY-MM-DD
                parts = date_part.split('/')
                if len(parts) == 3:
                    return f"{parts[0]}-{parts[1].zfill(2)}-{parts[2].zfill(2)}"
    except:
        pass

    # 如果不是日期格式，返回原值
    return value

def preprocess_data(df, type_mapping):
    """预处理数据，将所有数据转换为字符串并处理空值，日期只保留年月日"""
    for col in df.columns:
        # 先处理日期类型 - 如果列包含日期数据，只保留年月日部分
        if df[col].dtype == 'object':  # 对象类型可能包含日期字符串
            df[col] = df[col].apply(lambda x: process_date_value(x))

        # 将所有数据转换为字符串，处理各种空值情况
        df[col] = df[col].apply(lambda x: '' if pd.isna(x) or x == 'nan' or x == '<NA>' or x == 'None' else str(x))
        
        # 截断过长的字符串以避免行大小问题
        if isinstance(type_mapping.get(col), String) and hasattr(type_mapping.get(col), 'length'):
            max_length = type_mapping[col].length
            df[col] = df[col].apply(lambda x: x[:max_length] if len(x) > max_length else x)
        else:
            # 对于其他类型，限制为较小的长度
            df[col] = df[col].apply(lambda x: x[:255] if len(x) > 255 else x)
    
    return df

def is_effective_value(value) -> bool:
    """判断业务字段是否有有效值，排除空值和 NA/NANA 等无需 ESO 标记。"""
    if value is None:
        return False
    value_str = str(value).strip()
    return value_str.upper() not in EMPTY_MARKERS

def normalize_part_no(value) -> str:
    """零件号匹配用标准化，避免 Excel 数字/空格导致匹配失败。"""
    if value is None:
        return ""
    value_str = str(value).strip()
    if value_str.endswith(".0"):
        value_str = value_str[:-2]
    return value_str

def parse_date_like_value(value):
    """将常见 Excel/字符串日期转成 date/datetime，无法识别时返回原值。"""
    if value is None or not is_effective_value(value):
        return value
    if isinstance(value, (datetime, date)):
        return value

    value_str = str(value).strip()
    value_str = value_str.split(" ")[0].replace("/", "-")
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y%m%d"):
        try:
            return datetime.strptime(value_str, fmt)
        except ValueError:
            continue
    return value

def resolve_target_date(target_date: str | None) -> date:
    """统计日期默认昨天；用户传入时按 YYYY-MM-DD 校验。"""
    if not target_date or not target_date.strip():
        return date.today() - timedelta(days=1)
    try:
        return datetime.strptime(target_date.strip(), "%Y-%m-%d").date()
    except ValueError:
        raise HTTPException(status_code=400, detail="统计日期格式错误，应为 YYYY-MM-DD")

def build_unfinished_select_result(conn, resolved_target_date: date, matched_completed_count=None):
    """基于当前 sheet 表，按统计日期生成未完成清单和汇总。"""
    try:
        main_table_columns = conn.execute(text("SHOW COLUMNS FROM `sheet`")).fetchall()
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"请先上传并处理 Excel 文件，再生成未完成清单: {str(e)}")

    main_columns = [col[0] for col in main_table_columns]
    required_main_cols = ['零件号', 'ESO_Plan_Date', 'ESO_Actual_Date']
    missing_main_cols = [col for col in required_main_cols if col not in main_columns]
    if missing_main_cols:
        raise HTTPException(status_code=400, detail=f"主表(sheet)缺少必要列: {missing_main_cols}")

    plan_valid_condition = """
    ESO_Plan_Date IS NOT NULL
    AND TRIM(ESO_Plan_Date) != ''
    AND UPPER(TRIM(ESO_Plan_Date)) NOT IN ('NA', 'N/A', 'NANA', 'NONE', 'NULL', 'NAN')
    """
    plan_date_expr = """
    STR_TO_DATE(
        REPLACE(SUBSTRING_INDEX(TRIM(ESO_Plan_Date), ' ', 1), '/', '-'),
        '%Y-%m-%d'
    )
    """
    plan_due_condition = f"""
    {plan_valid_condition}
    AND {plan_date_expr} IS NOT NULL
    AND {plan_date_expr} <= :target_date
    """
    actual_filled_condition = """
    ESO_Actual_Date IS NOT NULL
    AND TRIM(ESO_Actual_Date) != ''
    """
    actual_empty_condition = f"NOT ({actual_filled_condition})"

    summary_sql = f"""
    SELECT
        SUM(CASE WHEN {plan_valid_condition} THEN 1 ELSE 0 END) AS planned_count,
        SUM(CASE WHEN {plan_valid_condition} AND {actual_filled_condition} THEN 1 ELSE 0 END) AS completed_count,
        SUM(CASE WHEN {plan_due_condition} AND {actual_empty_condition} THEN 1 ELSE 0 END) AS unfinished_count,
        SUM(CASE WHEN {plan_due_condition} THEN 1 ELSE 0 END) AS due_planned_count,
        SUM(CASE WHEN {plan_due_condition} AND {actual_filled_condition} THEN 1 ELSE 0 END) AS due_completed_count,
        SUM(CASE WHEN 操作类型 = 'D' AND {actual_filled_condition} THEN 1 ELSE 0 END) AS delete_completed_count
    FROM sheet;
    """

    group_sql = f"""
    SELECT
        COALESCE(NULLIF(TRIM(功能组), ''), '未知功能组') AS 功能组,
        COUNT(*) AS count
    FROM sheet
    WHERE {plan_due_condition}
      AND {actual_empty_condition}
    GROUP BY COALESCE(NULLIF(TRIM(功能组), ''), '未知功能组')
    ORDER BY count DESC;
    """

    select_sql = f"""
    SELECT *
    FROM sheet
    WHERE {plan_due_condition}
      AND {actual_empty_condition};
    """

    query_params = {"target_date": resolved_target_date.isoformat()}
    summary_result = conn.execute(text(summary_sql), query_params).mappings().first() or {}
    planned_count = int(summary_result.get("planned_count") or 0)
    completed_count = int(summary_result.get("completed_count") or 0)
    unfinished_count = int(summary_result.get("unfinished_count") or 0)
    due_planned_count = int(summary_result.get("due_planned_count") or 0)
    due_completed_count = int(summary_result.get("due_completed_count") or 0)
    delete_completed_count = int(summary_result.get("delete_completed_count") or 0)

    group_result = conn.execute(text(group_sql), query_params)
    group_rows = [dict(row) for row in group_result.mappings().all()]

    result = conn.execute(text(select_sql), query_params)
    columns = result.keys()
    rows = [dict(zip(columns, row)) for row in result.fetchall()]
    logger.info(f"未完成清单查询成功，统计日期 {resolved_target_date.isoformat()}，返回 {len(rows)} 条记录")

    return {
        "row_count": len(rows),
        "total_empty_count": unfinished_count,
        "data": rows,
        "summary": {
            "planned_count": planned_count,
            "completed_count": completed_count,
            "unfinished_count": unfinished_count,
            "due_planned_count": due_planned_count,
            "due_completed_count": due_completed_count,
            "delete_completed_count": delete_completed_count,
            "matched_completed_count": matched_completed_count,
            "unfinished_by_difference": max(due_planned_count - due_completed_count, 0),
            "target_date": resolved_target_date.isoformat(),
            "formula": "ESO Plan Date <= 统计日期，且 ESO Actual Date 为空 = 未完成数量",
        },
        "group_stats": group_rows,
    }

def find_column_by_sanitized_header(ws, header_row: int, expected_name: str):
    for cell in ws[header_row]:
        if sanitize_name(cell.value, is_column=True) == expected_name:
            return cell.column
    return None

def build_archive_date_map(file_bytes: bytes):
    """
    从 TDSP 已归档清单中读取零件号 -> 归档日期。
    返回原始 openpyxl 单元格值和格式，便于写回时尽量保留日期显示格式。
    """
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb[wb.sheetnames[0]]
    part_col = find_column_by_sanitized_header(ws, 1, "零件号")
    archive_col = find_column_by_sanitized_header(ws, 1, "归档日期")

    if not part_col or not archive_col:
        raise ValueError("ESO已归档清单缺少必要列：零件号或归档日期")

    archive_map = {}
    for row in range(2, ws.max_row + 1):
        part_no = normalize_part_no(ws.cell(row=row, column=part_col).value)
        archive_cell = ws.cell(row=row, column=archive_col)
        if part_no and is_effective_value(archive_cell.value):
            archive_map[part_no] = {
                "value": archive_cell.value,
                "number_format": archive_cell.number_format,
            }
    return archive_map

def create_modified_club_workbook(file_bytes: bytes, archive_map: Dict, original_filename: str):
    """
    在原始零件俱乐部 workbook 上回填 ESO Actual Date。
    只写实际日期列，不触碰 ESO Plan Date，因此计划日期格式会保留原样。
    """
    wb = load_workbook(io.BytesIO(file_bytes))
    ws = wb[wb.sheetnames[0]]
    header_row = 2
    part_col = find_column_by_sanitized_header(ws, header_row, "零件号")
    actual_col = find_column_by_sanitized_header(ws, header_row, "ESO_Actual_Date")
    operation_col = find_column_by_sanitized_header(ws, header_row, "操作类型")

    if not part_col or not actual_col:
        raise ValueError("零件俱乐部清单缺少必要列：零件号或 ESO Actual Date")

    filled_count = 0
    delete_skipped_count = 0
    for row in range(header_row + 1, ws.max_row + 1):
        part_no = normalize_part_no(ws.cell(row=row, column=part_col).value)
        archive_info = archive_map.get(part_no)
        if not archive_info:
            continue

        operation_value = ""
        if operation_col:
            operation_value = str(ws.cell(row=row, column=operation_col).value or "").strip().upper()
        if operation_value == "D":
            delete_skipped_count += 1
            continue

        actual_cell = ws.cell(row=row, column=actual_col)
        actual_cell.value = parse_date_like_value(archive_info["value"])
        actual_cell.number_format = "yyyy/mm/dd"
        filled_count += 1

    safe_stem = re.sub(r"[^\w\u4e00-\u9fff.-]+", "_", Path(original_filename).stem)
    output_name = f"{safe_stem}-已回填ESO实际归档日期-{datetime.now().strftime('%Y%m%d%H%M%S')}-{uuid4().hex[:8]}.xlsx"
    output_path = OUTPUT_DIR / output_name
    wb.save(output_path)

    return {
        "filename": output_name,
        "download_url": f"/download/{output_name}",
        "filled_count": filled_count,
        "delete_skipped_count": delete_skipped_count,
    }

def split_dataframe_for_tables(df, table_names, max_cols_per_table=49):  # 保留1列给ID
    """
    将DataFrame按列拆分以适应多个表
    """
    if len(table_names) == 1:
        # 只有一个表，不需要拆分
        return [(table_names[0], df)]
    
    dataframes_parts = []
    
    # 第一部分 - 主表
    main_df = df.iloc[:, :max_cols_per_table]
    dataframes_parts.append((table_names[0], main_df))
    
    # 其余部分 - 按照分表策略拆分
    remaining_df = df.iloc[:, max_cols_per_table:]
    
    part_num = 1
    for i in range(0, remaining_df.shape[1], max_cols_per_table):
        if part_num >= len(table_names):
            break
            
        chunk_df = remaining_df.iloc[:, i:i+max_cols_per_table].copy()
        # 添加一个公共ID列用于关联
        chunk_df['common_id'] = chunk_df.index + 1
        dataframes_parts.append((table_names[part_num], chunk_df))
        part_num += 1
    
    return dataframes_parts

def append_dataframe_to_table(table_name: str, df: pd.DataFrame):
    """使用 SQLAlchemy 原生批量插入，避免 pandas.to_sql 在不同服务器版本下不兼容。"""
    if df.empty:
        return

    metadata = MetaData()
    table = Table(table_name, metadata, autoload_with=engine)
    records = df.to_dict(orient="records")
    with engine.begin() as conn:
        conn.execute(table.insert(), records)

@app.post("/upload-excel/", summary="上传Excel文件")
async def upload_excel(file: UploadFile = File(...)):
    """
    上传Excel文件并将其数据导入MySQL数据库
    
    - Excel工作表名称将作为数据库表名
    - 第一行将作为数据库字段名
    - 会删除已存在的同名表
    """
    logger.info(f"收到上传请求，文件名: {file.filename}")
    
    if not file.filename.endswith(('.xlsx', '.xls')):
        logger.warning(f"不支持的文件类型: {file.filename}")
        raise HTTPException(status_code=400, detail="仅支持 .xlsx 或 .xls 文件")

    try:
        logger.info(f"开始处理文件: {file.filename}")
        contents = await file.read()
        logger.info(f"文件大小: {len(contents)} bytes")
        excel_file = io.BytesIO(contents)

        # 读取所有 sheet
        sheet_names = pd.ExcelFile(excel_file).sheet_names
        logger.info(f"发现工作表: {sheet_names}")
        results = {}

        for sheet_name in sheet_names:
            logger.info(f"处理工作表: {sheet_name}")
            # 重置指针
            excel_file.seek(0)
            # 读取数据，第一行作为列名
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            logger.info(f"工作表 {sheet_name} 数据形状: {df.shape}")

            if df.empty:
                results[sheet_name] = "跳过（空表）"
                continue

            # 使用原始的第一行作为列名，但需要清理以支持中文
            df.columns = [sanitize_name(col, is_column=True) for col in df.columns]
            logger.info(f"清理后的列名: {df.columns.tolist()}")
            
            # 去重列名（避免重复）
            seen = set()
            new_cols = []
            for col in df.columns:
                suffix = 1
                orig = col
                while col in seen:
                    col = f"{orig}_{suffix}"
                    suffix += 1
                seen.add(col)
                new_cols.append(col)
            df.columns = new_cols

            # 表名使用工作表名称
            table_name = sanitize_name(sheet_name, is_column=False)
            logger.info(f"表名: {table_name}")
            
            # 获取数据库元数据
            metadata = MetaData()
            inspector = inspect(engine)

            # 删除已存在的同名表
            if inspector.has_table(table_name):
                logger.info(f"删除已存在的表: {table_name}")
                old_table = Table(table_name, metadata, autoload_with=engine)
                old_table.drop(engine)
                logger.info(f"成功删除表: {table_name}")

            # 推断列类型并构建表结构
            columns = []
            type_mapping = {}
            
            for col in df.columns:
                # 推断合适的SQL类型而不是全部使用TEXT
                sql_type = infer_sql_type(df[col])
                columns.append(Column(col, sql_type))
                # 构建 SQLAlchemy 类型映射（用于 to_sql）
                type_mapping[col] = sql_type

            # 创建新表 - 使用分表策略
            created_table_names = create_table_with_chunked_columns(table_name, columns, type_mapping)
            logger.info(f"成功创建表: {created_table_names}")

            # 数据预处理 - 确保所有数据都符合对应的SQL类型
            df = preprocess_data(df, type_mapping)

            # 插入数据
            logger.info(f"准备插入 {len(df)} 行数据到表 {table_name}")
            
            # 按表拆分数据
            table_dataframes = split_dataframe_for_tables(df, created_table_names)
            
            for t_name, t_df in table_dataframes:
                logger.info(f"准备插入 {len(t_df)} 行数据到表 {t_name}")
                # 获取当前表的类型映射
                current_type_mapping = {}
                for col_name in t_df.columns:
                    if col_name in type_mapping:
                        current_type_mapping[col_name] = type_mapping[col_name]
                
                # 数据预处理 - 确保所有数据都符合对应的SQL类型
                t_df = preprocess_data(t_df, current_type_mapping)
                
                # 分批插入数据，避免一次性插入大量数据
                batch_size = 1000
                total_rows = len(t_df)
                
                for start_idx in range(0, total_rows, batch_size):
                    end_idx = min(start_idx + batch_size, total_rows)
                    batch_df = t_df.iloc[start_idx:end_idx]
                    
                    logger.info(f"正在插入批次 {start_idx//batch_size + 1}/{(total_rows-1)//batch_size + 1} 到表 {t_name}")
                    
                    append_dataframe_to_table(t_name, batch_df)
                
                logger.info(f"成功插入 {len(t_df)} 行数据到表 {t_name}")

            results[sheet_name] = f"成功导入 {len(df)} 行数据到表 `{', '.join(created_table_names)}`"
            logger.info(f"成功插入数据到表: {created_table_names}")

        logger.info(f"文件处理完成: {file.filename}")
        return {
            "filename": file.filename,
            "sheets": len(sheet_names),
            "results": results
        }

    except Exception as e:
        logger.error(f"处理失败: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"处理失败: {str(e)}")

@app.post("/upload-two-excel/", summary="上传两个Excel文件并执行SQL")
async def upload_two_excel(
    file1: UploadFile = File(...),
    file2: UploadFile = File(...),
    target_date: str = Form(None),
):
    """
    上传两个Excel文件并将其数据导入MySQL数据库，然后执行指定的SQL操作
    
    - 强制将第一个Excel文件导入为'sheet'表，第二个Excel文件导入为'sheet1'表
    - 第一个文件从第二行开始读取列名，只导入指定的列；第二个文件按原有逻辑处理
    - 第一行将作为数据库字段名
    - 会删除已存在的同名表
    - 导入完成后执行更新和查询操作
    """
    logger.info(f"收到两个上传请求，文件1: {file1.filename}, 文件2: {file2.filename}")
    
    if not file1.filename.endswith(('.xlsx', '.xls')):
        logger.warning(f"不支持的文件类型: {file1.filename}")
        raise HTTPException(status_code=400, detail="仅支持 .xlsx 或 .xls 文件 (file1)")

    if not file2.filename.endswith(('.xlsx', '.xls')):
        logger.warning(f"不支持的文件类型: {file2.filename}")
        raise HTTPException(status_code=400, detail="仅支持 .xlsx 或 .xls 文件 (file2)")

    try:
        resolved_target_date = resolve_target_date(target_date)
        logger.info(f"未完成清单统计日期: {resolved_target_date.isoformat()}")

        # 处理第一个文件 - 强制导入为 'sheet' 表，但只导入指定列，从第二行开始读取列名
        logger.info(f"开始处理文件1: {file1.filename}")
        contents1 = await file1.read()
        logger.info(f"文件1大小: {len(contents1)} bytes")
        excel_file1 = io.BytesIO(contents1)

        # 读取所有 sheet
        sheet_names1 = pd.ExcelFile(excel_file1).sheet_names
        logger.info(f"文件1发现工作表: {sheet_names1}")
        results1 = {}

        # 定义第一个文件要保留的列
        required_columns_sheet1 = [
            '操作类型', '层级', '数据类型', '产品', '车型年', '零件号', '短SVPPS', 'FFC',
            'FFC中文描述', '状态', '工程采购级别', '左右件', '功能组', '工程师代码',
            '工程师', '工厂编号', '节点组', '本色件标识', 'TG2_Plan_Date', 'TG2_Actual_Date',
            '2D_Drawing_Actual_Date', '2D_Drawing_No.', 'ESO_Plan_Date', 'ESO材料送检计划',
            'ESO备注_类型', '是否需要ESO_物流提供_', 'ESO_Actual_Date', '备注', '首次申请项目'
        ]

        # 强制将第一个文件的第一个工作表导入为 'sheet' 表
        for idx, sheet_name in enumerate(sheet_names1):
            logger.info(f"开始处理工作表: {sheet_name}")
            # 重置指针
            excel_file1.seek(0)
            # 读取数据，从第二行开始读取列名（header=1）
            logger.info("开始读取Excel文件...")
            df = pd.read_excel(excel_file1, sheet_name=sheet_name, header=1)
            logger.info(f"工作表 {sheet_name} 数据形状: {df.shape}")

            if df.empty:
                results1[sheet_name] = "跳过（空表）"
                logger.info(f"工作表 {sheet_name} 为空，跳过处理")
                continue

            logger.info("开始清理列名...")

            # 在sanitize之前，先记录原始列名并检测包含"ESO"的列名
            original_columns = list(df.columns)
            eso_column_names = []
            for col in original_columns:
                if isinstance(col, str) and 'ESO' in col.upper():
                    eso_column_names.append(col)
                    logger.info(f"发现包含ESO的原始列: {col}")

            # 使用第二行作为列名，但需要清理以支持中文
            df.columns = [sanitize_name(col, is_column=True) for col in df.columns]
            logger.info(f"清理后的列名: {df.columns.tolist()}")

            # 去重列名（避免重复）
            seen = set()
            new_cols = []
            old_to_new_col_map = {}  # 记录列名映射
            for col in df.columns:
                suffix = 1
                orig = col
                old_to_new_col_map[col] = col  # 先记录原始映射
                while col in seen:
                    col = f"{orig}_{suffix}"
                    suffix += 1
                seen.add(col)
                old_to_new_col_map[orig] = col  # 更新映射
                new_cols.append(col)
            df.columns = new_cols
            logger.info(f"去重后的列名: {df.columns.tolist()}")

            logger.info("开始筛选指定列...")
            # 只保留指定的列
            available_columns = [col for col in required_columns_sheet1 if col in df.columns]

            # 添加所有包含"ESO"的列（先sanitize原始列名，再在去重后的列中查找）
            # 跳过已经在required_columns_sheet1中定义的ESO列
            eso_columns_in_required = ['ESO_Plan_Date', 'ESO_Actual_Date', 'ESO备注_类型', '是否需要ESO_物流提供_']

            for original_col in eso_column_names:
                sanitized_name = sanitize_name(original_col, is_column=True)
                # 如果这个列已经在required_columns_sheet1中定义，跳过
                if sanitized_name in eso_columns_in_required:
                    continue
                # 先尝试直接找sanitized后的名称
                if sanitized_name in df.columns and sanitized_name not in available_columns:
                    available_columns.append(sanitized_name)
                    logger.info(f"添加包含ESO的列: {sanitized_name} (原名: {original_col})")
                # 如果直接找不到，可能在去重时被加了后缀，尝试查找
                elif sanitized_name not in df.columns:
                    for col in df.columns:
                        if col.startswith(sanitized_name + '_') or sanitized_name in col:
                            if col not in available_columns:
                                available_columns.append(col)
                                logger.info(f"添加包含ESO的列: {col} (原名: {original_col})")
                            break

            # 添加首次申请项目列（如果存在且未包含）
            if '首次申请项目' in df.columns and '首次申请项目' not in available_columns:
                available_columns.append('首次申请项目')
                logger.info(f"添加首次申请项目列")

            if not available_columns:
                logger.warning(f"没有找到任何指定的列，将使用所有可用列")
                available_columns = df.columns.tolist()
            else:
                logger.info(f"将保留的列: {available_columns}")

            df = df[available_columns]
            logger.info(f"筛选后的数据形状: {df.shape}")

            # 强制将第一个文件的第一个工作表导入为 'sheet' 表
            table_name = "sheet" if idx == 0 else f"sheet_{idx}"
            logger.info(f"表名: {table_name}")
            
            # 获取数据库元数据
            metadata = MetaData()
            inspector = inspect(engine)

            # 删除已存在的同名表
            if inspector.has_table(table_name):
                logger.info(f"删除已存在的表: {table_name}")
                old_table = Table(table_name, metadata, autoload_with=engine)
                old_table.drop(engine)
                logger.info(f"成功删除表: {table_name}")

            logger.info("开始推断列类型并构建表结构...")
            # 推断列类型并构建表结构
            columns = []
            type_mapping = {}
            
            for col in df.columns:
                # 推断合适的SQL类型而不是全部使用TEXT
                sql_type = infer_sql_type(df[col])
                columns.append(Column(col, sql_type))
                # 构建 SQLAlchemy 类型映射（用于 to_sql）
                type_mapping[col] = sql_type

            # 创建新表
            table = Table(
                table_name, metadata, *columns,
                mysql_charset='utf8mb4',
                extend_existing=True,
                mysql_row_format='DYNAMIC',
                mysql_engine='InnoDB'
            )
            logger.info(f"创建表 {table_name}")
            metadata.create_all(engine)
            logger.info(f"成功创建表: {table_name}")

            logger.info("开始数据预处理...")
            # 数据预处理 - 确保所有数据都符合对应的SQL类型
            df = preprocess_data(df, type_mapping)
            logger.info("数据预处理完成")

            logger.info("开始插入数据...")
            # 插入数据
            logger.info(f"准备插入 {len(df)} 行数据到表 {table_name}")
            
            # 按表拆分数据 - 由于我们已经筛选了列，不再需要拆分
            logger.info(f"插入 {len(df)} 行数据到表 {table_name}")
            
            # 获取当前表的类型映射
            current_type_mapping = {}
            for col_name in df.columns:
                if col_name in type_mapping:
                    current_type_mapping[col_name] = type_mapping[col_name]
            
            # 分批插入数据，避免一次性插入大量数据
            batch_size = 1000
            total_rows = len(df)
            
            for start_idx in range(0, total_rows, batch_size):
                end_idx = min(start_idx + batch_size, total_rows)
                batch_df = df.iloc[start_idx:end_idx]
                
                logger.info(f"正在插入批次 {start_idx//batch_size + 1}/{(total_rows-1)//batch_size + 1} 到表 {table_name}")
                
                append_dataframe_to_table(table_name, batch_df)
            
            results1[sheet_name] = f"成功导入 {len(df)} 行数据到表 `{table_name}`"
            logger.info(f"成功插入数据到表: {table_name}")
            
            # 只处理第一个工作表，其他工作表跳过
            if idx == 0:
                break

        # 处理第二个文件 - 强制导入为 'sheet1' 表，保持原有逻辑
        logger.info(f"开始处理文件2: {file2.filename}")
        contents2 = await file2.read()
        logger.info(f"文件2大小: {len(contents2)} bytes")
        excel_file2 = io.BytesIO(contents2)
        modified_workbook = None

        try:
            archive_map = build_archive_date_map(contents2)
            modified_workbook = create_modified_club_workbook(contents1, archive_map, file1.filename)
            logger.info(
                "已生成回填后的零件俱乐部文件: %s，回填 %s 行，跳过 Delete 行 %s 行",
                modified_workbook["filename"],
                modified_workbook["filled_count"],
                modified_workbook["delete_skipped_count"],
            )
        except Exception as workbook_error:
            logger.error(f"生成回填后的零件俱乐部文件失败: {str(workbook_error)}", exc_info=True)
            modified_workbook = {
                "error": f"生成回填后的零件俱乐部文件失败: {str(workbook_error)}"
            }

        # 读取所有 sheet
        sheet_names2 = pd.ExcelFile(excel_file2).sheet_names
        logger.info(f"文件2发现工作表: {sheet_names2}")
        results2 = {}

        # 强制将第二个文件导入为 'sheet1' 表
        for idx, sheet_name in enumerate(sheet_names2):
            logger.info(f"开始处理工作表: {sheet_name}")
            # 重置指针
            excel_file2.seek(0)
            # 读取数据，第一行作为列名（保持原有逻辑）
            logger.info("开始读取Excel文件...")
            df = pd.read_excel(excel_file2, sheet_name=sheet_name)
            logger.info(f"工作表 {sheet_name} 数据形状: {df.shape}")

            if df.empty:
                results2[sheet_name] = "跳过（空表）"
                logger.info(f"工作表 {sheet_name} 为空，跳过处理")
                continue

            logger.info("开始清理列名...")
            # 使用原始的第一行作为列名，但需要清理以支持中文
            df.columns = [sanitize_name(col, is_column=True) for col in df.columns]
            logger.info(f"清理后的列名: {df.columns.tolist()}")
            
            # 去重列名（避免重复）
            seen = set()
            new_cols = []
            for col in df.columns:
                suffix = 1
                orig = col
                while col in seen:
                    col = f"{orig}_{suffix}"
                    suffix += 1
                seen.add(col)
                new_cols.append(col)
            df.columns = new_cols

            # 强制将第二个文件的第一个工作表导入为 'sheet1' 表
            table_name = "sheet1" if idx == 0 else f"sheet1_{idx}"
            logger.info(f"表名: {table_name}")
            
            # 获取数据库元数据
            metadata = MetaData()
            inspector = inspect(engine)

            # 删除已存在的同名表
            if inspector.has_table(table_name):
                logger.info(f"删除已存在的表: {table_name}")
                old_table = Table(table_name, metadata, autoload_with=engine)
                old_table.drop(engine)
                logger.info(f"成功删除表: {table_name}")

            logger.info("开始推断列类型并构建表结构...")
            # 推断列类型并构建表结构
            columns = []
            type_mapping = {}
            
            for col in df.columns:
                # 推断合适的SQL类型而不是全部使用TEXT
                sql_type = infer_sql_type(df[col])
                columns.append(Column(col, sql_type))
                # 构建 SQLAlchemy 类型映射（用于 to_sql）
                type_mapping[col] = sql_type

            # 创建新表
            table = Table(
                table_name, metadata, *columns,
                mysql_charset='utf8mb4',
                extend_existing=True,
                mysql_row_format='DYNAMIC',
                mysql_engine='InnoDB'
            )
            logger.info(f"创建表 {table_name}")
            metadata.create_all(engine)
            logger.info(f"成功创建表: {table_name}")

            logger.info("开始数据预处理...")
            # 数据预处理 - 确保所有数据都符合对应的SQL类型
            df = preprocess_data(df, type_mapping)
            logger.info("数据预处理完成")

            logger.info("开始插入数据...")
            # 插入数据
            logger.info(f"准备插入 {len(df)} 行数据到表 {table_name}")
            
            # 按表拆分数据 - 由于我们已经简化了处理逻辑，不再需要拆分
            logger.info(f"插入 {len(df)} 行数据到表 {table_name}")
            
            # 获取当前表的类型映射
            current_type_mapping = {}
            for col_name in df.columns:
                if col_name in type_mapping:
                    current_type_mapping[col_name] = type_mapping[col_name]
            
            # 分批插入数据，避免一次性插入大量数据
            batch_size = 1000
            total_rows = len(df)
            
            for start_idx in range(0, total_rows, batch_size):
                end_idx = min(start_idx + batch_size, total_rows)
                batch_df = df.iloc[start_idx:end_idx]
                
                logger.info(f"正在插入批次 {start_idx//batch_size + 1}/{(total_rows-1)//batch_size + 1} 到表 {table_name}")
                
                append_dataframe_to_table(table_name, batch_df)
            
            results2[sheet_name] = f"成功导入 {len(df)} 行数据到表 `{table_name}`"
            logger.info(f"成功插入数据到表: {table_name}")
            
            # 只处理第一个工作表，其他工作表跳过
            if idx == 0:
                break

        # 检查是否成功导入了两个表
        if len(results1) == 0 or len(results2) == 0:
            logger.error("至少有一个文件没有成功导入任何表")
            raise HTTPException(status_code=500, detail="至少有一个文件没有成功导入任何表")

        logger.info("开始执行SQL语句")
        
        # 检查表中是否包含必要的列
        with engine.connect() as conn:
            # 获取固定表名的列信息
            main_table = "sheet"  # 固定主表名
            join_table = "sheet1" # 固定关联表名
            
            logger.info(f"使用固定表名: 主表={main_table}, 关联表={join_table}")
            
            # 获取主表的列信息
            try:
                main_table_columns = conn.execute(text(f"SHOW COLUMNS FROM `{main_table}`")).fetchall()
                main_columns = [col[0] for col in main_table_columns]
                
                # 获取关联表的列信息
                join_table_columns = conn.execute(text(f"SHOW COLUMNS FROM `{join_table}`")).fetchall()
                join_columns = [col[0] for col in join_table_columns]
                
                logger.info(f"主表 {main_table} 的列: {main_columns}")
                logger.info(f"关联表 {join_table} 的列: {join_columns}")
                
                # 检查必需的列是否存在
                # 根据错误信息，主表(sheet)应包含'零件号'和'ESO_Actual_Date'，关联表(sheet1)应包含'零件号'和'归档日期'
                required_main_cols = ['零件号', 'ESO_Plan_Date', 'ESO_Actual_Date']
                required_join_cols = ['零件号', '归档日期']
                
                missing_main_cols = [col for col in required_main_cols if col not in main_columns]
                missing_join_cols = [col for col in required_join_cols if col not in join_columns]
                
                if missing_main_cols or missing_join_cols:
                    error_msg = f"缺少必要的列: "
                    if missing_main_cols:
                        error_msg += f"主表({main_table})缺少: {missing_main_cols}; "
                    if missing_join_cols:
                        error_msg += f"关联表({join_table})缺少: {missing_join_cols}"
                    logger.error(error_msg)
                    
                    update_result = f"更新失败: {error_msg}"
                    select_result = {
                        "row_count": 0,
                        "total_empty_count": 0,
                        "data": [],
                        "error": f"缺少必要的列: {error_msg}"
                    }
                else:
                    updated_row_count = 0
                    # SQL 1: 更新非空归档日期到ESO Actual Date - 使用固定SQL
                    update_sql = """
                    UPDATE sheet
                    JOIN sheet1 ON sheet.零件号 = sheet1.零件号
                    SET sheet.ESO_Actual_Date = sheet1.归档日期
                    WHERE sheet1.归档日期 IS NOT NULL
                      AND TRIM(sheet1.归档日期) != ''
                      AND (sheet.操作类型 IS NULL OR UPPER(TRIM(sheet.操作类型)) != 'D');
                    """
                    
                    try:
                        result = conn.execute(text(update_sql))
                        conn.commit()
                        updated_row_count = result.rowcount
                        logger.info(f"更新SQL执行成功，受影响行数: {result.rowcount}")
                        update_result = f"成功更新 {result.rowcount} 条记录"
                    except Exception as e:
                        logger.error(f"更新SQL执行失败: {str(e)}")
                        update_result = f"更新失败: {str(e)}"
                    
                    try:
                        select_result = build_unfinished_select_result(conn, resolved_target_date, updated_row_count)
                    except Exception as e:
                        logger.error(f"查询SQL执行失败: {str(e)}")
                        select_result = {
                            "row_count": 0,
                            "total_empty_count": 0,
                            "data": [],
                            "error": f"查询失败: {str(e)}"
                        }
            except Exception as e:
                logger.error(f"检查表结构时出错: {str(e)}")
                update_result = f"检查表结构失败: {str(e)}"
                select_result = {
                    "row_count": 0,
                    "total_empty_count": 0,
                    "data": [],
                    "error": f"检查表结构失败: {str(e)}"
                }

        logger.info(f"两个文件处理完成，SQL执行完成")

        return {
            "file1": file1.filename,
            "file2": file2.filename,
            "target_date": resolved_target_date.isoformat(),
            "sheets": len(sheet_names1) + len(sheet_names2),
            "results": {**results1, **results2},
            "modified_workbook": modified_workbook,
            "sql_results": {
                "update_result": update_result,
                "select_result": select_result
            }
        }

    except Exception as e:
        logger.error(f"处理失败: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"处理失败: {str(e)}")

@app.post("/unfinished-list/", summary="按统计日期重新生成未完成清单")
async def regenerate_unfinished_list(target_date: str = Body(None, embed=True)):
    """
    基于已经导入并回填后的 sheet 表，重新计算未完成清单。
    不重新上传 Excel，不重新入库，不重新执行回填。
    """
    resolved_target_date = resolve_target_date(target_date)
    logger.info(f"重新生成未完成清单，统计日期: {resolved_target_date.isoformat()}")

    try:
        with engine.connect() as conn:
            select_result = build_unfinished_select_result(conn, resolved_target_date)

        return {
            "target_date": resolved_target_date.isoformat(),
            "select_result": select_result,
        }
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"重新生成未完成清单失败: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"重新生成未完成清单失败: {str(e)}")

@app.get("/download/{filename}", summary="下载生成的Excel文件")
async def download_generated_file(filename: str):
    safe_name = Path(filename).name
    file_path = OUTPUT_DIR / safe_name
    if not file_path.exists() or not file_path.is_file():
        raise HTTPException(status_code=404, detail="文件不存在或已被清理")

    return FileResponse(
        path=file_path,
        filename=safe_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

@app.post("/execute-sql/")
async def execute_sql(sql_query: str = Body(..., embed=True)):
    """
    执行SQL语句
    
    - **sql_query**: 要执行的SQL语句
    """
    logger.info(f"收到SQL执行请求: {sql_query}")
    
    # 检查SQL语句是否为空
    if not sql_query or not sql_query.strip():
        logger.warning("SQL语句为空")
        raise HTTPException(status_code=400, detail="SQL语句不能为空")
    
    # 检查SQL是否包含危险操作
    dangerous_keywords = ["drop", "delete", "truncate", "alter", "create"]
    sql_lower = sql_query.strip().lower()
    for keyword in dangerous_keywords:
        if keyword in sql_lower and keyword == sql_lower.split()[0]:  # 检查是否是语句开头的关键词
            logger.warning(f"检测到潜在危险SQL操作: {sql_query}")
            raise HTTPException(status_code=400, detail=f"不允许执行{keyword.upper()}操作")
    
    try:
        logger.info(f"执行SQL查询: {sql_query}")
        # 执行SQL查询
        with engine.connect() as conn:
            result = conn.execute(text(sql_query))
            
            # 判断是否是查询语句
            if sql_query.strip().upper().startswith("SELECT"):
                # 获取列名
                columns = result.keys()
                # 获取所有行数据
                rows = [dict(zip(columns, row)) for row in result.fetchall()]
                
                logger.info(f"查询成功，返回 {len(rows)} 行数据")
                return {
                    "success": True,
                    "row_count": len(rows),
                    "data": rows,
                    "sql_query": sql_query
                }
            else:
                # 对于非查询语句，提交事务并返回受影响的行数
                conn.commit()
                logger.info(f"执行成功，受影响行数: {result.rowcount}")
                return {
                    "success": True,
                    "row_count": result.rowcount,
                    "sql_query": sql_query
                }
    except Exception as e:
        logger.error(f"执行SQL失败: {str(e)}", exc_info=True)
        logger.error(f"SQL语句: {sql_query}")
        logger.error(f"错误详情: {traceback.format_exc()}")
        
        # 返回更具体的错误信息
        error_detail = str(e).lower()
        if "unknown table" in error_detail or "table" in error_detail and "doesn't exist" in error_detail:
            return {
                "success": False,
                "error": "表不存在，请检查表名是否正确",
                "sql_query": sql_query
            }
        elif "unknown column" in error_detail:
            return {
                "success": False,
                "error": f"字段不存在: {str(e)}",
                "sql_query": sql_query
            }
        elif "syntax error" in error_detail or "you have an error in your sql syntax" in error_detail:
            return {
                "success": False,
                "error": f"SQL语法错误: {str(e)}",
                "sql_query": sql_query
            }
        else:
            return {
                "success": False,
                "error": f"执行SQL失败: {str(e)}",
                "sql_query": sql_query
            }

@app.get("/", summary="根路径")
async def root():
    logger.info("访问根路径")
    return {"message": "Excel to MySQL 服务正在运行", "status": "ok"}

@app.get("/health", summary="健康检查")
async def health_check():
    logger.info("健康检查请求")
    try:
        # 测试数据库连接
        with engine.connect() as conn:
            result = conn.execute(text("SELECT 1"))
            logger.info("数据库连接测试成功")
        return {"status": "healthy", "database": "connected"}
    except Exception as e:
        logger.error(f"数据库连接测试失败: {str(e)}", exc_info=True)
        return {"status": "unhealthy", "database_error": str(e)}

@app.on_event("startup")
async def startup_event():
    logger.info("服务启动中...")
    try:
        # 测试数据库连接
        with engine.connect() as conn:
            result = conn.execute(text("SELECT 1"))
        logger.info("数据库连接初始化成功")
    except Exception as e:
        logger.error(f"数据库连接初始化失败: {str(e)}")
        
@app.on_event("shutdown")
async def shutdown_event():
    logger.info("服务关闭中...")
    try:
        engine.dispose()
        logger.info("数据库连接已关闭")
    except Exception as e:
        logger.error(f"关闭数据库连接时出错: {str(e)}")


if __name__ == "__main__":
    logger.info("启动Excel to MySQL服务...")
   
    try:
        # 启动服务
        uvicorn.run("main:app", host="0.0.0.0", port=8020, reload=False)
    except Exception as e:
        logger.error(f"启动服务时出错: {str(e)}", exc_info=True)
