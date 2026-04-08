"""MCP Server for OpsButler-LLM: exposes deployment plan generation as a tool."""

import asyncio
import logging
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any

from mcp.server.fastmcp import Context, FastMCP

from opsbutler.config import load_config
from opsbutler.excel_parser import load_excel
from opsbutler.llm_client import create_llm_client
from opsbutler.models import ImplementationPlan
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
    instructions="从 Excel 上线清单生成部署实施方案（Word 文档）。",
    lifespan=server_lifespan,
)


@mcp.tool()
async def generate_deployment_plan(
    excel_path: str,
    output_path: str,
    ctx: Context,
) -> dict[str, Any]:
    """从 Excel 上线清单生成部署实施方案（Word 文档）。

    完整流水线：解析 Excel → 步骤映射（LLM）→ 汇总生成（LLM）→ 风险分析（LLM）→ 生成 Word 文件。
    流程包含 3 次 LLM 调用，可能需要 1-3 分钟。

    Args:
        excel_path: Excel (.xlsx) 上线清单文件的路径。
        output_path: 输出 Word (.docx) 文件的路径。

    Returns:
        包含生成结果摘要的字典：文件路径、任务数、模块数、步骤数等。
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
    ctx.report_progress(1, 5, "Excel 解析完成")

    # Step 1: Step mapping (LLM call 1)
    ctx.info("步骤 1/4: 生成步骤映射...")
    generator = PlanGenerator(llm_client, config)
    mapping_rules = generator._load_mapping_rules()
    excel_json = excel_payload.model_dump_json(indent=2, exclude={"parsed_at"})

    mapping_result = await asyncio.to_thread(
        generator._do_step_mapping, excel_json, mapping_rules
    )
    ctx.report_progress(2, 5, "步骤映射完成")

    # Step 2: Group data (no LLM)
    ctx.info("步骤 2/4: 分组数据...")
    step_details = await asyncio.to_thread(
        generator._group_data, excel_payload, mapping_result
    )
    ctx.report_progress(3, 5, "数据分组完成")

    # Step 3: Summary generation (LLM call 2)
    ctx.info("步骤 3/4: 生成汇总...")
    summary = await asyncio.to_thread(
        generator._do_summary_generation, excel_json, mapping_result
    )
    ctx.report_progress(4, 5, "汇总生成完成")

    # Step 4: Risk analysis (LLM call 3)
    ctx.info("步骤 4/4: 生成风险分析...")
    steps_summary = generator._build_steps_summary(step_details)
    verification, rollback, risk = await asyncio.to_thread(
        generator._do_risk_analysis, steps_summary
    )
    ctx.report_progress(5, 5, "风险分析完成")

    # Merge into ImplementationPlan
    plan = ImplementationPlan(
        summary=summary,
        task_count=len(mapping_result.task_table),
        module_count=len(set(t.task_name for t in mapping_result.task_table)),
        high_risk_count=0,
        task_table=mapping_result.task_table,
        step_details=step_details,
        verification_plan=verification,
        rollback_plan=rollback,
        risk_analysis=risk,
    )

    # Generate Word document
    ctx.info("正在生成 Word 文档...")
    await asyncio.to_thread(WordGenerator().generate, plan, str(output))

    return {
        "output_file": str(output),
        "task_count": plan.task_count,
        "module_count": plan.module_count,
        "step_count": len(plan.step_details),
        "verification_steps": len(plan.verification_plan.verification_steps),
        "rollback_steps": len(plan.rollback_plan.rollback_steps),
        "risk_count": len(plan.risk_analysis.risks),
    }


if __name__ == "__main__":
    mcp.run()
