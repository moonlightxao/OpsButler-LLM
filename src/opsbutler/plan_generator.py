import json
import logging
from pathlib import Path
from opsbutler.config import Config
from opsbutler.models import (
    ExcelPayload, StepMappingResult, StepMapping,
    SummarySection, VerificationPlan, RollbackPlan, RiskAnalysis,
    ImplementationPlan, TaskEntry, StepDetail, OperationGroup,
)
from opsbutler.llm_client import LLMClient

logger = logging.getLogger(__name__)


class PlanGenerator:
    """Orchestrates LLM calls to generate implementation plan."""

    def __init__(self, llm_client: LLMClient, config: Config):
        self.llm = llm_client
        self.config = config
        self.prompts_dir = Path(__file__).parent.parent.parent / "prompts"
        self.system_prompt = self._load_prompt("system.txt")

    def _load_prompt(self, filename: str) -> str:
        """Load a prompt template file."""
        path = self.prompts_dir / filename
        if path.exists():
            return path.read_text(encoding="utf-8")
        return ""

    def _load_mapping_rules(self) -> str:
        """Load the mapping rules markdown file."""
        rules_path = Path(self.config.mapping.rules_file)
        if rules_path.exists():
            return rules_path.read_text(encoding="utf-8")
        logger.warning(f"Mapping rules file not found: {rules_path}")
        return ""

    def generate(self, excel_payload: ExcelPayload) -> ImplementationPlan:
        """
        Full generation pipeline:
        1. LLM step mapping
        2. Group data by step -> operation_type
        3. LLM summary generation
        4. LLM risk analysis
        5. Merge into ImplementationPlan
        """
        mapping_rules = self._load_mapping_rules()
        excel_json = excel_payload.model_dump_json(indent=2, exclude={"parsed_at"})

        # Step 1: Step mapping
        logger.info("Step 1: Generating step mappings...")
        mapping_result = self._do_step_mapping(excel_json, mapping_rules)

        # Step 2: Group data
        logger.info("Step 2: Grouping data by step and operation type...")
        step_details = self._group_data(excel_payload, mapping_result)

        # Step 3: Summary generation
        logger.info("Step 3: Generating summary...")
        summary = self._do_summary_generation(excel_json, mapping_result)

        # Step 4: Risk analysis
        logger.info("Step 4: Generating risk analysis...")
        steps_summary = self._build_steps_summary(step_details)
        verification, rollback, risk = self._do_risk_analysis(steps_summary)

        # Step 5: Merge
        plan = ImplementationPlan(
            summary=summary,
            task_count=len(mapping_result.task_table),
            module_count=len(set(t.task_name for t in mapping_result.task_table)),
            high_risk_count=0,  # Could be calculated from risk analysis
            task_table=mapping_result.task_table,
            step_details=step_details,
            verification_plan=verification,
            rollback_plan=rollback,
            risk_analysis=risk,
        )
        logger.info("Plan generation complete.")
        return plan

    def _do_step_mapping(self, excel_json: str, mapping_rules: str) -> StepMappingResult:
        """LLM Call 1: Map Excel rows to platform steps."""
        template = self._load_prompt("step_mapping.txt")
        user_message = template.format(
            mapping_rules=mapping_rules,
            excel_data=excel_json,
        )

        messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": user_message},
        ]

        result = self.llm.chat_json(messages)
        # Parse with Pydantic
        return StepMappingResult(**result)

    def _do_summary_generation(self, excel_json: str, mapping_result: StepMappingResult) -> SummarySection:
        """LLM Call 2: Generate summary text."""
        template = self._load_prompt("summary_generation.txt")
        user_message = template.format(
            excel_data=excel_json,
            step_mappings=json.dumps(
                [m.model_dump() for m in mapping_result.step_mappings],
                ensure_ascii=False, indent=2
            ),
        )

        messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": user_message},
        ]

        result = self.llm.chat_json(messages)
        return SummarySection(**result["summary"])

    def _do_risk_analysis(self, steps_summary: str) -> tuple[VerificationPlan, RollbackPlan, RiskAnalysis]:
        """LLM Call 3: Generate risk analysis."""
        template = self._load_prompt("risk_analysis.txt")
        user_message = template.format(steps_summary=steps_summary)

        messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": user_message},
        ]

        result = self.llm.chat_json(messages)
        verification = VerificationPlan(**result["verification_plan"])
        rollback = RollbackPlan(**result["rollback_plan"])
        risk = RiskAnalysis(**result["risk_analysis"])
        return verification, rollback, risk

    def _group_data(self, excel_payload: ExcelPayload, mapping_result: StepMappingResult) -> list[StepDetail]:
        """
        Group Excel rows by step and operation type based on mapping result.
        """
        # Build a lookup: (sheet_name, row_index) -> row data
        sheet_rows = {}
        for sheet in excel_payload.sheets:
            sheet_rows[sheet.sheet_name] = sheet.rows

        step_details = []
        for mapping in mapping_result.step_mappings:
            # Get the rows for this step
            rows = sheet_rows.get(mapping.source_sheet, [])
            step_rows = [rows[i] for i in mapping.row_indices if i < len(rows)]

            # Group by operation type
            operation_groups = self._group_by_operation(step_rows, mapping.source_sheet, excel_payload)

            # Get step description from mapping or use default
            description = mapping.description or ""

            step_details.append(StepDetail(
                step_name=mapping.step_name,
                step_description=description,
                operation_groups=operation_groups,
            ))

        return step_details

    def _group_by_operation(self, rows: list[dict], sheet_name: str, excel_payload: ExcelPayload) -> list[OperationGroup]:
        """Group rows by their operation type (action column value)."""
        # Find the action column for this sheet
        action_col = None
        for sheet in excel_payload.sheets:
            if sheet.sheet_name == sheet_name:
                action_col = sheet.detected_action_column
                break

        if not action_col:
            # No action column, put all rows in one group
            return [OperationGroup(operation_type="操作", rows=rows)]

        groups = {}
        for row in rows:
            op_type = str(row.get(action_col, "其他"))
            if op_type not in groups:
                groups[op_type] = []
            groups[op_type].append(row)

        return [OperationGroup(operation_type=op, rows=rs) for op, rs in groups.items()]

    def _build_steps_summary(self, step_details: list[StepDetail]) -> str:
        """Build a text summary of steps for risk analysis prompt."""
        lines = []
        for step in step_details:
            lines.append(f"步骤: {step.step_name}")
            lines.append(f"描述: {step.step_description}")
            for group in step.operation_groups:
                lines.append(f"  操作类型: {group.operation_type}, 数据行数: {len(group.rows)}")
            lines.append("")
        return "\n".join(lines)
