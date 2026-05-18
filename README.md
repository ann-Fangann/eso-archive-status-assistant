# ESO 数据智能体 v3.0

ESO 数据智能体是一个面向 ESO 和图纸归档跟踪的智能数据处理工具，用于自动生成未完成清单、做统计分析，并通过 AI 降低人工整理 Excel 和查询数据的成本。

当前版本是接入 AI 的 `v3.0`，和旧版未接 AI 的 `ESO` 做了区分：旧版成熟前端体验已经合并进本目录，AI 能力由 `ESOv3.0` 后端提供。

## 主要能力

- 上传 ESO 或图纸相关 Excel，自动识别字段并生成未完成清单。
- 按零件号回填实际归档/发布日期。
- 支持统计日期、D 行排除、功能组统计、未完成类型统计。
- 支持前端表格查看、列选择、Excel 导出和图表可视化。
- 支持智能问答，例如“按工程师统计未完成数量”“某个功能组有哪些未完成”。

## AI 用在哪里

AI 只做辅助理解，不直接决定业务结果。

1. 字段语义识别  
   不同 Excel 模板里的字段名可能不一致，例如 `ESO_Plan_Date`、`ESO计划日期`、`计划归档日期`。系统先用本地规则识别，低置信度或缺失字段再交给 AI 辅助判断。

2. 自然语言问答理解  
   用户输入问题后，AI 把问题转换成受控查询计划，例如按工程师统计、按功能组筛选、列出零件号等。

业务计算仍由 Python 确定性执行，包括回填、逾期判断、D 行排除、未完成数量统计和导出结果生成。

## 运行方式

需要 Python 3.10+ 和 Node.js。

```bash
cd ESOv3.0
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

cd eso
npm install
npm run build
cd ..

uvicorn app.main:app --reload --port 8030
```

打开：

```text
http://127.0.0.1:8030
```

Windows 激活虚拟环境：

```bat
.venv\Scripts\activate
```

## AI 配置

复制环境变量模板：

```bash
cp .env.example .env
```

配置 MiniMax：

```text
MINIMAX_BASE_URL=https://api.minimaxi.com/anthropic
MINIMAX_API_KEY=你的MiniMax API Key
MINIMAX_MODEL=MiniMax-M2.7-highspeed
ENABLE_LLM_MAPPING=true
ENABLE_LLM_QUERY=true
```

`.env` 不要提交到 GitHub。

## 开发前端

如果需要单独调试前端：

```bash
cd ESOv3.0/eso
npm run dev
```

Vite 端口是 `3030`，会把 `/api` 代理到 `http://127.0.0.1:8030`。

## 目录说明

```text
app/main.py                  FastAPI 入口和旧前端兼容接口
app/core/field_mapper.py     字段识别，本地规则 + 可选 AI 辅助
app/core/query_agent.py      智能问答，AI 生成受控查询计划
app/core/llm_client.py       MiniMax/Anthropic 兼容接口调用
app/core/business.py         确定性业务计算
app/core/excel_reader.py     Excel 表头候选读取
app/core/schemas.py          标准业务字段定义
eso/                         Vue 前端，沿用旧版 ESO 体验
tests/                       轻量回归测试
```

## GitHub 分支

当前 AI 版本建议放在：

```text
esov3-ai
```

旧版未接 AI 的 `main` 分支可作为历史版本保留；确认 v3.0 在公司电脑测试稳定后，再考虑合并或替换主分支。
