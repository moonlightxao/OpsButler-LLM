---
name: opsbutler-deployment-plan
description: 从 Excel 上线清单自动生成部署实施方案（Word 文档）。当用户提到上线清单、go-live checklist、实施方案、部署计划、Excel 变更清单等关键词时触发。使用 generate_deployment_plan 工具完成端到端流水线。
---

# OpsButler 部署方案生成

## 触发条件

当用户提到以下关键词或场景时使用此 Skill：
- 上线清单、go-live checklist、变更清单
- 部署实施方案、实施计划
- 从 Excel 生成部署文档
- 变更影响分析、风险评估

## 使用方法

调用 `generate_deployment_plan` 工具，提供两个参数：

- `excel_path`：Excel (.xlsx) 上线清单文件的绝对路径或相对路径
- `output_path`：输出 Word (.docx) 文件的路径

## 前置条件

1. `config.yaml` 已配置好 LLM provider（支持 OpenAI 兼容接口和 Ollama）
2. Excel 文件包含标准上线 checklist 结构（应用配置、容器配置、MQS 配置、ROMA 任务等 sheet）
3. `mapping_rules.md` 和 `prompts/` 目录存在

## 输出

工具返回结果摘要：
- `output_file`：生成的 Word 文件路径
- `task_count`：任务总数
- `module_count`：模块数量
- `step_count`：实施步骤数
- `verification_steps`：验证步骤数
- `rollback_steps`：回退步骤数
- `risk_count`：风险项数量

生成的 Word 文档包含 5 个章节：原因和目的、实施步骤和计划、实施后验证计划、应急回退措施、风险分析和规避措施。

## 注意事项

- 流程包含 3 次 LLM 调用，耗时 1-3 分钟
- LLM 调用失败时会自动重试（指数退避）
- Excel 列名自动检测，支持多种中文名称变体
- 生成的方案需人工审核确认后再用于生产环境
