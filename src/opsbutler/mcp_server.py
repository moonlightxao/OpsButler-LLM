"""MCP Server for OpsButler-LLM: exposes deployment plan generation as a tool."""

import asyncio
import logging
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import Context, FastMCP

from opsbutler.config import load_config
from opsbutler.excel_parser import load_excel, load_schedule_sheet
from opsbutler.llm_client import create_llm_client
from opsbutler.plan_generator import PlanGenerator
from opsbutler.word_generator import WordGenerator

logger = logging.getLogger(__name__)


@asynccontextmanager
async def server_lifespan(server: FastMCP):
    """Load config and create LLM client once at startup."""
    config = load_config("config.yaml")
    llm_client = create_llm_client(config.llm)
    yield {"config": config, "llm_client": llm_client}


mcp = FastMCP(
    name="OpsButler-LLM",
    instructions="从 Excel 上线清单生成部署实施方案（Word 文档）。部分 Sheet 页的详细数据可能被配置为 zip 附件（内含 Excel 表格），与 Word 文档一同输出到同一目录。",
    lifespan=server_lifespan,
)


@mcp.tool()
async def generate_deployment_plan(
    excel_path: str,
    output_path: str,
    ctx: Context,
) -> dict[str, Any]:
    """从 Excel 上线清单生成部署实施方案（Word 文档）。

    完整流水线：解析 Excel → 按 Sheet 拆分步骤映射（LLM）→ 按 Sheet 拆分汇总（LLM）→ 综合汇总（LLM）→ 风险分析（LLM）→ 生成 Word 文件。
    流程包含多轮 LLM 调用（每个 Sheet 2 次 + 1 次综合 + 1 次风险分析），可能需要 1-3 分钟。

    当 mapping_rules.md 中为某个 Sheet 配置了 `- 打包方式: zip` 时，该 Sheet 的详细表格数据不会写入 Word 文档，
    而是生成一个独立的 zip 文件（内含 Excel 表格），保存在输出目录下，文件名为 `{Sheet页名称}.zip`。
    主 Word 文档中该步骤仍保留标题和描述，并以提示文字指向附件。

    Args:
        excel_path: Excel (.xlsx) 上线清单文件的路径。
        output_path: 输出 Word (.docx) 文件的路径。

    Returns:
        包含生成结果摘要的字典：output_file（Word 路径）、zip_files（zip 附件路径列表，可选）、task_count、module_count、step_count 等。
    """
    config = ctx.request_context.lifespan_context["config"]
    llm_client = ctx.request_context.lifespan_context["llm_client"]

    # Validate input
    excel = Path(excel_path)
    if not excel.exists():
        raise ValueError(f"Excel 文件不存在: {excel_path}")

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    # Step 0: Parse Excel
    ctx.info("正在解析 Excel 文件...")
    excel_payload = await asyncio.to_thread(load_excel, str(excel), config)
    schedule_table = await asyncio.to_thread(load_schedule_sheet, str(excel))
    ctx.report_progress(1, 3, "Excel 解析完成")

    # Steps 1-4: Run full pipeline via PlanGenerator
    ctx.info("正在生成实施方案（按 Sheet 拆分 LLM 调用）...")
    generator = PlanGenerator(llm_client, config)
    plan = await asyncio.to_thread(generator.generate, excel_payload, schedule_table)
    ctx.report_progress(2, 3, "方案生成完成")

    # Step 5: Generate Word document
    ctx.info("正在生成 Word 文档...")
    word_result = await asyncio.to_thread(WordGenerator().generate, plan, str(output))
    ctx.report_progress(3, 3, "Word 文档生成完成")

    result = {
        "output_file": str(output),
        "task_count": plan.task_count,
        "module_count": plan.module_count,
        "step_count": len(plan.step_details),
        "verification_steps": len(plan.verification_plan.verification_steps),
        "rollback_steps": len(plan.rollback_plan.rollback_steps),
        "risk_count": len(plan.risk_analysis.risks),
    }
    if "zip_files" in word_result:
        result["zip_files"] = word_result["zip_files"]

    return result


if __name__ == "__main__":
    mcp.run()
