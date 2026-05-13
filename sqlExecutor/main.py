from fastapi import FastAPI, HTTPException, Body
from fastapi.middleware.cors import CORSMiddleware
import logging
import uvicorn
import sys
import os
from sqlalchemy import create_engine, text, MetaData, Table, Column
from sqlalchemy.engine import URL
from sqlalchemy.types import String, Integer
from dotenv import load_dotenv
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

app = FastAPI(title="SQL Executor Service", description="专门用于执行SQL语句的服务")

# 允许跨域（开发用）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

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

@app.post("/execute-sql/")
async def execute_sql(sql_query: str = Body(..., embed=True), ss: str = Body(None, embed=True)):
    """
    执行SQL语句

    - **sql_query**: 要执行的第一条SQL语句（通常是UPDATE）
    - **ss**: 要执行的第二条SQL语句（SELECT），并在其结果中选出两列
    """
    logger.info(f"收到SQL执行请求: {sql_query}")
    if ss:
        logger.info(f"收到第二条SQL: {ss}")

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
    
    # 清理SQL查询，去除可能的注释或非法字符
    cleaned_sql = sql_query.strip()
    
    try:
        logger.info(f"执行SQL查询: {cleaned_sql}")
        # 执行SQL查询
        with engine.connect() as conn:
            result = conn.execute(text(cleaned_sql))
            
            # 判断是否是查询语句
            if cleaned_sql.strip().upper().startswith("SELECT"):
                # 获取列名
                columns = result.keys()
                # 获取所有行数据
                rows = [dict(zip(columns, row)) for row in result.fetchall()]

                logger.info(f"查询成功，返回 {len(rows)} 行数据")

                # 如果执行的是对sheet表的查询，确保包含ESO备注和首次申请项目列
                if "FROM sheet" in cleaned_sql.upper() or "FROM `sheet`" in cleaned_sql.upper():
                    # 检查第一行数据是否包含这些列
                    if len(rows) > 0:
                        first_row = rows[0]
                        # 添加ESO备注列（如果不存在）
                        if 'ESO备注' not in first_row:
                            for row in rows:
                                row['ESO备注'] = ''
                        # 添加首次申请项目列（如果不存在）
                        if '首次申请项目' not in first_row:
                            for row in rows:
                                row['首次申请项目'] = ''

                # 如果执行的是对sheet表的查询，也执行功能组统计查询
                if "FROM sheet" in cleaned_sql.upper() or "FROM `sheet`" in cleaned_sql.upper():
                    # 如果提供了ss参数，执行该SQL并选出两列结果
                    if ss:
                        try:
                            logger.info(f"执行第二条SQL查询: {ss}")
                            ss_result = conn.execute(text(ss))
                            ss_columns = ss_result.keys()
                            ss_rows = [dict(zip(ss_columns, row)) for row in ss_result.fetchall()]

                            logger.info(f"第二条SQL查询成功，返回 {len(ss_rows)} 行数据")

                            # 将第二条SQL的结果添加到原始查询结果中
                            rows.extend(ss_rows)

                        except Exception as ss_e:
                            logger.error(f"第二条SQL查询执行失败: {str(ss_e)}")
                
                return {
                    "success": True,
                    "row_count": len(rows),
                    "data": rows,
                    "sql_query": cleaned_sql
                }
            else:
                # 对于非查询语句，提交事务并返回受影响的行数
                conn.commit()
                logger.info(f"执行成功，受影响行数: {result.rowcount}")
                
                # 检查是否涉及sheet表
                is_about_sheet = "sheet" in cleaned_sql.lower()
                
                # 如果执行的是UPDATE、INSERT或DELETE语句，且涉及sheet表，先添加列
                if is_about_sheet:
                    # 为sheet表添加两个新列（如果不存在）- 使用MySQL兼容语法
                    try:
                        # 尝试添加功能组统计列
                        alter_table_sql1 = """
                        ALTER TABLE sheet 
                        ADD COLUMN `部门` VARCHAR(255)
                        """
                        conn.execute(text(alter_table_sql1))
                        logger.info("成功为 'sheet' 表添加新列 '部门'")
                    except Exception as e:
                        if "Duplicate column" in str(e) or "same name" in str(e):
                            logger.info("'部门' 列已存在")
                        else:
                            logger.error(f"添加 '部门' 列时出错: {str(e)}")
                    
                    try:
                        # 尝试添加统计数量列
                        alter_table_sql2 = """
                        ALTER TABLE sheet
                        ADD COLUMN `延期未发布数量` INT
                        """
                        conn.execute(text(alter_table_sql2))
                        logger.info("成功为 'sheet' 表添加新列 '延期未发布数量'")
                    except Exception as e:
                        if "Duplicate column" in str(e) or "same name" in str(e):
                            logger.info("'延期未发布数量' 列已存在")
                        else:
                            logger.error(f"添加 '延期未发布数量' 列时出错: {str(e)}")

                    try:
                        # 尝试添加ESO备注列
                        alter_table_sql3 = """
                        ALTER TABLE sheet
                        ADD COLUMN `ESO备注` VARCHAR(255)
                        """
                        conn.execute(text(alter_table_sql3))
                        logger.info("成功为 'sheet' 表添加新列 'ESO备注'")
                    except Exception as e:
                        if "Duplicate column" in str(e) or "same name" in str(e):
                            logger.info("'ESO备注' 列已存在")
                        else:
                            logger.error(f"添加 'ESO备注' 列时出错: {str(e)}")

                    try:
                        # 尝试添加首次申请项目列
                        alter_table_sql4 = """
                        ALTER TABLE sheet
                        ADD COLUMN `首次申请项目` VARCHAR(255)
                        """
                        conn.execute(text(alter_table_sql4))
                        logger.info("成功为 'sheet' 表添加新列 '首次申请项目'")
                    except Exception as e:
                        if "Duplicate column" in str(e) or "same name" in str(e):
                            logger.info("'首次申请项目' 列已存在")
                        else:
                            logger.error(f"添加 '首次申请项目' 列时出错: {str(e)}")

                    conn.commit()
                    logger.info("已为 'sheet' 表添加新列 '部门'、'延期未发布数量'、'ESO备注' 和 '首次申请项目' (如果它们之前不存在)")

                    # 检查是否是UPDATE语句且提供了ss参数
                    if "update" in cleaned_sql.lower() and "sheet" in cleaned_sql.lower() and ss:
                        try:
                            logger.info(f"UPDATE已提交，执行ss查询: {ss}")

                            # 使用新的连接来查询，确保获取最新数据
                            with engine.connect() as new_conn:
                                stats_result = new_conn.execute(text(ss))
                                stats_columns = stats_result.keys()
                                stats_rows = [dict(zip(stats_columns, row)) for row in stats_result.fetchall()]

                            logger.info(f"ss查询成功，返回 {len(stats_rows)} 行数据")

                            # 确保ss查询结果包含ESO备注和首次申请项目列
                            if len(stats_rows) > 0:
                                first_row = stats_rows[0]
                                if 'ESO备注' not in first_row:
                                    for row in stats_rows:
                                        row['ESO备注'] = ''
                                if '首次申请项目' not in first_row:
                                    for row in stats_rows:
                                        row['首次申请项目'] = ''

                            # 从ss查询结果中统计功能组（在内存中对895行数据进行统计）
                            group_stats = {}
                            for row in stats_rows:
                                功能组 = row.get('功能组', '')
                                if 功能组:
                                    group_stats[功能组] = group_stats.get(功能组, 0) + 1

                            # 转换为列表格式
                            group_rows = [{'部门': k, 'count': v} for k, v in group_stats.items()]

                            logger.info(f"部门完成，共 {len(group_rows)} 个功能组")
                            logger.info(f"部门数据: {group_rows}")

                            # 返回ss查询结果和功能组统计结果
                            return {
                                "success": True,
                                "row_count": result.rowcount,
                                "sql_query": cleaned_sql,
                                "ss_info": {
                                    "ss_query": ss,
                                    "ss_row_count": len(stats_rows),
                                    "ss_data": stats_rows
                                },
                                "group_stats": {
                                    "group_row_count": len(group_rows),
                                    "group_data": group_rows
                                }
                            }
                        except Exception as stats_e:
                            logger.error(f"部门统计查询执行失败: {str(stats_e)}")
                            logger.error(f"错误详情: {traceback.format_exc()}")
                
                return {
                    "success": True,
                    "row_count": result.rowcount,
                    "sql_query": cleaned_sql
                }
    except Exception as e:
        logger.error(f"执行SQL失败: {str(e)}", exc_info=True)
        logger.error(f"SQL语句: {cleaned_sql}")
        logger.error(f"错误详情: {traceback.format_exc()}")
        
        # 返回更具体的错误信息
        error_detail = str(e).lower()
        if "unknown table" in error_detail or "table" in error_detail and "doesn't exist" in error_detail:
            return {
                "success": False,
                "error": "表不存在，请检查表名是否正确",
                "sql_query": cleaned_sql
            }
        elif "unknown column" in error_detail:
            return {
                "success": False,
                "error": f"字段不存在: {str(e)}",
                "sql_query": cleaned_sql
            }
        elif "syntax error" in error_detail or "you have an error in your sql syntax" in error_detail:
            return {
                "success": False,
                "error": f"SQL语法错误: {str(e)}",
                "sql_query": cleaned_sql
            }
        else:
            return {
                "success": False,
                "error": f"执行SQL失败: {str(e)}",
                "sql_query": cleaned_sql
            }

@app.get("/health")
async def health_check():
    """健康检查端点"""
    try:
        # 测试数据库连接
        with engine.connect() as conn:
            result = conn.execute(text("SELECT 1"))
        return {"status": "healthy", "database": "connected"}
    except Exception as e:
        logger.error(f"数据库连接测试失败: {str(e)}")
        return {"status": "unhealthy", "database_error": str(e)}

@app.get("/")
async def root():
    return {"message": "SQL Executor Service正在运行", "status": "ok"}

if __name__ == "__main__":
    logger.info("启动SQL Executor服务...")
    try:
        # 启动服务
        uvicorn.run("main:app", host="0.0.0.0", port=8001, reload=False)
    except Exception as e:
        logger.error(f"启动服务时出错: {str(e)}", exc_info=True)
