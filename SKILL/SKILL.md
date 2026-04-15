---
name: opsbutler-deployment-plan
description: 从 Excel 上线清单自动生成部署实施方案（Word 文档）。当用户提到上线清单、go-live checklist、实施方案、部署计划、Excel 变更清单等关键词时触发。
---

# OpsButler 部署方案生成

## 触发条件

当用户提到以下关键词或场景时使用此 Skill：
- 上线清单、go-live checklist、变更清单
- 部署实施方案、实施计划
- 从 Excel 生成部署文档
- 变更影响分析、风险评估

## 使用方法

此 Skill 通过在本地执行 Python 命令运行 OpsButler 流水线。你需要向用户获取以下两个必要参数：

1. **`excel_path`**：用户本地的 Excel (.xlsx) 上线清单文件路径
2. **`output_path`**：用户指定的输出 Word (.docx) 文件路径

如果用户未提供完整路径，请主动询问。

获取参数后，使用 **Bash 工具** 执行以下命令：

```bash
python -m opsbutler -e <excel_path> -o <output_path>
```

命令需要在项目目录 `/root/projects/OpsButler-LLM` 下执行。如果用户的路径是相对路径，请注意转换。

可选参数：
- `-c / --config`：指定配置文件路径（默认 `config.yaml`）
- `--log-level`：日志级别 DEBUG/INFO/WARNING/ERROR（默认 INFO）

## 前置条件

1. `config.yaml` 已配置好 LLM provider（支持 OpenAI 兼容接口和 Ollama）
2. Excel 文件包含标准上线 checklist 结构（应用配置、容器配置、MQS 配置、ROMA 任务等 sheet）
3. `mapping_rules.md` 和 `prompts/` 目录存在

## 输出

命令执行成功后会在指定路径生成 Word 文档，终端会输出处理进度信息。

生成的 Word 文档包含 5 个章节：原因和目的、实施步骤和计划、实施后验证计划、应急回退措施、风险分析和规避措施。

## 注意事项

- 流程按 Sheet 拆分调用 LLM（每个 Sheet 2 次调用 + 1 次综合汇总 + 1 次风险分析），耗时视 Sheet 数量而定，通常 1-3 分钟
- LLM 调用失败时会自动重试（指数退避）
- Excel 列名自动检测，支持多种中文名称变体
- 生成的方案需人工审核确认后再用于生产环境
