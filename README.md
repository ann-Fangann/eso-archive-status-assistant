# ESO v3.0 智能归档状态助手

这是对现有 ESO/图纸未完成清单工具的 v3.0 初版重写。核心思路是：

> LLM 只做语义映射和自然语言理解；业务计算继续由确定性代码执行。

这样既能适配“不完全统一的 Excel 模板”，又不会把未完成数量、回填、逾期判断交给模型自由发挥。

## 这版解决什么

- 自动识别不同模板里的相近字段，例如 `ESO_Plan_Date`、`ESO计划日期`、`计划归档日期`。
- 给每个字段映射输出置信度、样本和证据，避免黑盒。
- ESO 场景：按零件号从归档清单回填 ESO 实际归档日期，再按统计日期生成未完成清单。
- 图纸/数模场景：识别数模计划/实际、图纸实际/发布字段，按规则生成未完成清单。
- 智能问答：把自然语言问题转成受控查询动作，例如按工程师统计、查某个功能组明细、列出零件号。
- 导出 Excel：包含汇总、未完成清单、功能组统计、类型统计和字段映射。

## 为什么不是“全智能体”

这个需求的关键风险不是“模型不够聪明”，而是“统计结果必须可解释、可复核、可重复”。所以 v3.0 把系统拆成四层：

1. Excel Profiling：读取表头候选行，采样字段值。
2. Semantic Mapping：用本地语义规则 + 可选 LLM，把表头映射到标准字段。
3. Deterministic Engine：所有回填、日期判断、D 行排除、统计都在 Python 里确定性执行。
4. Controlled Agent：自然语言问答只生成查询计划，实际查询由工具函数执行。

## 前沿 Agent 概念在这里怎么落地

- Tool：`summary`、`count_by`、`filter`、`list_parts` 是受控查询工具，不允许模型直接改数据或编 SQL。
- Memory：当前初版用内存保存分析会话；后续可以换成 SQLite/MySQL，保存历史批次、字段映射和用户确认记录。
- RAG：这版暂不需要向量库。真正适合 RAG 的是“业务规则文档/会议纪要/字段说明”，不是表格计算本身。
- MCP：后续可以把“读取最新批次、查询清单、导出 Excel”暴露为 MCP tools，让 Claude Desktop、Cursor、内部 Agent 平台调用。
- Harness/Evals：建议沉淀一批黄金 Excel 样例，测试字段映射、未完成数量和导出结果，作为回归评估集。

## 当前合并版说明

当前目录已经把旧版 `ESO` 的成熟 Vue 前端复制到了 `ESOv3.0/eso`，并在 `app/main.py`
里增加了旧前端需要的兼容接口：

- 页面、菜单、上传区、统计日期、列选择导出、邮件汇报文本、图表和智能问答交互沿用旧版 `ESO`。
- 上传和问答接口内部使用 `ESOv3.0` 的 LLM 字段语义映射、确定性业务计算和受控 Agent。
- 旧版 `ESO` 目录没有被修改，可继续作为回退版本。

## 运行方式

```bash
cd ~/Downloads/ESOv3.0
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cd eso
npm install
npm run build
cd ..
uvicorn app.main:app --reload --port 8030
```

需要 Python 3.10+。这台机器上 `python` 是 3.11；`/usr/local/bin/python3` 可能是旧版 3.6，不建议用它创建虚拟环境。

打开：

```text
http://127.0.0.1:8030
```

开发前端时也可以单独运行 Vite：

```bash
cd ~/Downloads/ESOv3.0/eso
npm run dev
```

Vite 端口是 `3030`，会把 `/api` 代理到 `http://127.0.0.1:8030`。

## 可选 MiniMax LLM 能力

默认不依赖任何 LLM，完全本地可跑。如果要启用 LLM 辅助字段映射或问答规划：

```bash
cp .env.example .env
```

然后设置：

```text
MINIMAX_BASE_URL=https://api.minimaxi.com/anthropic
MINIMAX_API_KEY=你的MiniMax API Key
MINIMAX_MODEL=MiniMax-M2.7-highspeed
ENABLE_LLM_MAPPING=true
ENABLE_LLM_QUERY=true
```

项目调用 MiniMax Anthropic 兼容接口：`https://api.minimaxi.com/anthropic/v1/messages`。
也兼容 Claude Code 风格环境变量：`ANTHROPIC_BASE_URL`、`ANTHROPIC_AUTH_TOKEN`、`ANTHROPIC_MODEL`。
即使启用 LLM，它也只返回字段映射或查询计划，业务结果仍由代码计算。

## 目录

```text
app/main.py                  FastAPI 入口
app/core/schemas.py          标准业务字段本体
app/core/excel_reader.py     Excel 表头候选读取
app/core/field_mapper.py     字段语义映射
app/core/business.py         确定性业务计算和导出
app/core/query_agent.py      受控问答规划和工具执行
eso/                         沿用 ESO 旧版体验的 Vue 前端
static/                      前端页面
tests/                       轻量回归测试
```

## 下一步建议

1. 增加“人工确认字段映射”页面，把确认结果保存成模板记忆。
2. 用 SQLite/MySQL 保存历史批次和映射记忆，形成轻量 memory。
3. 建立 `golden_cases/`，每个样例包含输入 Excel、期望映射、期望未完成数量。
4. 给字段映射加主动学习：低置信度字段由用户确认后进入别名字典。
5. 再考虑 MCP，把这个服务包装成可被其他 Agent 调用的工具层。
