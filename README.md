# OpsButler-LLM

基于 LLM 的运维实施方案自动生成工具。读取 Excel 上线清单，调用 LLM 分析并映射到自动化部署平台的变更步骤，生成标准格式的实施方案 Word 文档。

生成的 Word 文档既供运维人员操作参考，也可导入自动化部署平台生成在线变更单。

## 功能特性

- **动态 Excel 解析** — 不硬编码 Sheet 名和列名，新增 Sheet 页类型无需改代码
- **LLM 智能映射** — 依据人工维护的 `mapping_rules.md` 规则，将 Excel 内容精确映射到平台变更步骤
- **一对多映射** — 支持单个 Excel Sheet 映射到多个平台步骤（如"ROMA任务与事件"拆分为"ROMA任务"和"ROMA事件"）
- **程序化 Word 构建** — 动态生成多级标题、按步骤分组的操作表格，适配不同规模的上线内容
- **多 LLM 后端** — 支持 OpenAI 兼容 API 和 Ollama 本地模型
- **可扩展规则** — 平台新增变更步骤类型时，只需更新 `mapping_rules.md` 文件
- **Agent Skill** — 配套 Skill 描述文件，Agent 自动识别用户意图，通过本地 CLI 命令生成方案
- **MCP Server** — 可选，可作为本地 MCP Server 接入 Claude Code 等 AI Agent

## 前置准备

### 环境要求

- Python 3.10+

### 安装

在项目根目录下以开发模式安装：

```bash
pip install -e .
```

依赖列表：

| 包 | 用途 |
|---|---|
| openpyxl | Excel 文件解析 |
| python-docx | Word 文档生成 |
| pydantic | 数据模型验证 |
| requests | LLM API 调用 |
| pyyaml | 配置文件加载 |
| mcp | MCP Server SDK |

## 配置指导

### 1. LLM 配置

编辑 `config.yaml`：

```yaml
llm:
  provider: "openai_compatible"     # 可选: openai_compatible | ollama
  base_url: "https://api.openai.com/v1"
  api_key: "${OPENAI_API_KEY}"      # 支持环境变量引用
  model: "gpt-4o"
  temperature: 0.3
  max_tokens: 4096
  retry_count: 2
```

**使用 OpenAI 兼容 API：**

设置环境变量：

```bash
export OPENAI_API_KEY="your-api-key"
```

**使用 Ollama 本地模型：**

```yaml
llm:
  provider: "ollama"
  base_url: "http://localhost:11434"
  model: "qwen2.5:14b"
```

确保 Ollama 服务已启动：`ollama serve`

### 2. Excel 解析配置

`config.yaml` 中的 `excel` 部分控制列名自动检测：

```yaml
excel:
  action_column_candidates:    # 操作类型列的候选列名（按优先级匹配）
    - "操作类型"
    - "操作"
    - "变更事项"
    - "变更类型"
  app_column_candidates:       # 应用标识列的候选列名
    - "APPID"
    - "应用"
    - "应用名称"
  skip_sheets:                 # 跳过不处理的 Sheet
    - "变更安排"
```

### 3. 映射规则配置

`mapping_rules.md` 定义 Excel Sheet 到平台变更步骤的映射规则，LLM 做映射时参考此文件。

每个来源 Sheet 通过"步骤定义"划分为一个或多个平台步骤，支持按操作类型拆分为不同的平台步骤名：

```markdown
## 步骤名称
- 来源Sheet: 对应的 Excel Sheet 名
- 说明: 步骤用途简述
- 步骤定义:
  - 平台步骤名: 平台变更步骤名
    - 筛选规则: 操作类型为"X" | 操作类型为"X"或"Y" | 某列非空 | 无
    - 操作描述: 生成到 Word 文档中的操作说明段落
```

筛选规则限定为四种固定句式，LLM 按句式匹配分配行数据：
- `操作类型为"X"` — 精确匹配操作类型列值
- `操作类型为"X"或"Y"` — 匹配多个操作类型值
- `某列非空` — 按指定列是否有值筛选（如 ROMA 任务名列非空）
- `无` — 不筛选，匹配所有行

新增平台步骤时，只需在此文件中添加对应规则，无需改动代码。

## 命令介绍

### 基本用法

```bash
python -m opsbutler --excel <Excel文件路径> --output <输出Word路径>
```

### 参数说明

| 参数 | 缩写 | 必填 | 说明 |
|---|---|---|---|
| `--excel` | `-e` | 是 | 输入的 Excel 上线清单文件路径 |
| `--output` | `-o` | 是 | 输出的 Word 实施方案文件路径 |
| `--config` | `-c` | 否 | 配置文件路径，默认 `config.yaml` |
| `--log-level` | — | 否 | 日志级别：DEBUG / INFO / WARNING / ERROR，默认 INFO |

### 示例

**使用默认配置：**

```bash
python -m opsbutler --excel sample/上线checklist.xlsx --output output/实施方案.docx
```

**指定配置文件：**

```bash
python -m opsbutler -e input/checklist.xlsx -o output/plan.docx -c my_config.yaml
```

**开启调试日志：**

```bash
python -m opsbutler -e sample/上线checklist.xlsx -o output/test.docx --log-level DEBUG
```

## Agent Skill（推荐）

`.claude/skills/opsbutler-deployment-plan/SKILL.md` 提供了 Agent Skill 描述，Claude Code 会自动识别并在用户提到上线清单、部署方案等关键词时触发。

Agent 触发后会自动获取用户本地的 Excel 文件路径和输出路径，然后通过 Bash 执行 `python -m opsbutler` 命令完成生成。

## 作为 MCP Server 使用（可选）

OpsButler-LLM 也可作为 MCP Server 接入 Claude Code 等 AI Agent，Agent 通过工具调用直接生成部署方案。

项目根目录已包含 `.mcp.json` 配置文件，在 Claude Code 中进入项目目录即可自动加载。或手动添加：

```bash
claude mcp add opsbutler -- python3 -m opsbutler.mcp_server
```

MCP Server 以 stdio 传输模式运行，暴露 `generate_deployment_plan` 工具，参数为 `excel_path`（Excel 路径）和 `output_path`（输出路径）。

## 项目结构

```
OpsButler-LLM/
├── config.yaml                # 配置文件
├── mapping_rules.md           # Excel→平台步骤映射规则（人工维护）
├── requirements.txt           # Python 依赖
├── SKILL/                     # Agent Skill 描述（独立分发）
│   └── SKILL.md               # Skill 定义文件
├── prompts/                   # LLM Prompt 模板
│   ├── system.txt             # 系统提示词
│   ├── step_mapping.txt       # 步骤映射 Prompt
│   ├── summary_generation.txt # 摘要生成 Prompt
│   └── risk_analysis.txt      # 风险分析 Prompt
├── src/opsbutler/
│   ├── main.py                # CLI 入口与管道编排
│   ├── config.py              # 配置加载
│   ├── models.py              # Pydantic 数据模型
│   ├── excel_parser.py        # 动态 Excel 解析
│   ├── llm_client.py          # LLM 客户端（requests）
│   ├── plan_generator.py      # LLM 调用编排（3次调用）
│   ├── word_generator.py      # Word 文档构建
│   └── mcp_server.py          # MCP Server（Agent 工具接口）
├── sample/                    # 示例文件
│   ├── 上线checklist.xlsx     # 示例 Excel 上线清单
│   └── 测试文档.docx          # 目标 Word 格式参考
├── .mcp.json                  # Claude Code MCP Server 配置
├── .claude/skills/opsbutler-deployment-plan/
│   └── SKILL.md               # Agent Skill 描述
└── output/                    # 输出目录
```

## 处理流程

```
Excel 上线清单
    │
    ▼
动态解析（自动检测 Sheet/列名）
    │
    ▼
LLM 调用 1: 步骤映射（Excel JSON + mapping_rules.md → 步骤映射结果）
    │
    ▼
LLM 调用 2: 摘要生成（变更应用/原因/影响分析）
    │
    ▼
LLM 调用 3: 风险分析（验证计划/回退措施/风险规避）
    │
    ▼
程序化构建 Word 实施方案文档
```
