import json
import logging
from pathlib import Path
from opsbutler.config import Config
from opsbutler.models import (
    ExcelPayload, SheetData, StepMappingResult, StepMapping,
    SummarySection, VerificationPlan, RollbackPlan, RiskAnalysis,
    ImplementationPlan, StepDetail, OperationGroup,
    SheetSummary, ScheduleTable,
)
from opsbutler.llm_client import LLMClient

logger = logging.getLogger(__name__)


class PlanGenerator:
    """Orchestrates LLM calls to generate implementation plan (per-sheet architecture)."""

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

    # ------------------------------------------------------------------
    # Mapping rules parsing
    # ------------------------------------------------------------------

    def _extract_sheet_rules(self, mapping_rules: str, sheet_name: str) -> str:
        """
        Extract the mapping rules section relevant to a specific sheet.

        mapping_rules.md is organized with `## Title` sections, each containing
        a `来源Sheet: SheetName` line. This method returns the general rules
        plus the matching section for the given sheet_name.
        Returns empty string if no matching section is found.
        """
        sections = mapping_rules.split("\n## ")
        general_rules = sections[0]  # content before first ## header

        for section in sections[1:]:
            for line in section.split("\n"):
                stripped = line.strip()
                if stripped.startswith("- 来源Sheet:") or stripped.startswith("- 来源 Sheet:"):
                    rule_sheet = stripped.split(":", 1)[1].strip()
                    if rule_sheet == sheet_name:
                        return general_rules + "\n## " + section
        return ""

    # ------------------------------------------------------------------
    # Per-sheet JSON helpers
    # ------------------------------------------------------------------

    def _serialize_sheet(self, sheet: SheetData) -> str:
        """Serialize a single sheet's rows to compact JSON."""
        return json.dumps(
            {"sheet_name": sheet.sheet_name, "rows": sheet.rows},
            ensure_ascii=False,
        )

    # ------------------------------------------------------------------
    # Phase 1: Per-sheet step mapping
    # ------------------------------------------------------------------

    def _do_step_mapping_for_sheet(
        self, sheet: SheetData, sheet_json: str, sheet_rules: str
    ) -> StepMappingResult:
        """LLM Call 1 (per-sheet): Map rows of one sheet to platform steps."""
        template = self._load_prompt("step_mapping_sheet.txt")
        user_message = template.format(
            mapping_rules=sheet_rules,
            sheet_name=sheet.sheet_name,
            sheet_data=sheet_json,
        )

        messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": user_message},
        ]

        result = self.llm.chat_json(messages)
        return StepMappingResult(**result)

    # ------------------------------------------------------------------
    # Phase 3a: Per-sheet summary
    # ------------------------------------------------------------------

    def _do_summary_for_sheet(
        self,
        sheet: SheetData,
        sheet_json: str,
        sheet_mappings: list[StepMapping],
    ) -> SheetSummary:
        """LLM Call 2 (per-sheet): Generate change summary for one sheet."""
        template = self._load_prompt("summary_sheet.txt")
        user_message = template.format(
            sheet_name=sheet.sheet_name,
            sheet_data=sheet_json,
            step_mappings=json.dumps(
                [m.model_dump() for m in sheet_mappings],
                ensure_ascii=False,
            ),
        )

        messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": user_message},
        ]

        result = self.llm.chat_json(messages)
        return SheetSummary(sheet_name=sheet.sheet_name, **result)

    # ------------------------------------------------------------------
    # Phase 3b: Synthesis
    # ------------------------------------------------------------------

    def _do_summary_synthesis(self, sheet_summaries: list[SheetSummary]) -> SummarySection:
        """LLM Call 2 (synthesis): Combine per-sheet summaries into final SummarySection."""
        summaries_text = "\n\n".join(
            f"### {s.sheet_name}\n- 涉及应用: {s.changed_apps}\n- 变更摘要: {s.changes_summary}"
            for s in sheet_summaries
        )

        template = self._load_prompt("summary_synthesis.txt")
        user_message = template.format(sheet_summaries=summaries_text)

        messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": user_message},
        ]

        result = self.llm.chat_json(messages)
        return SummarySection(**result["summary"])

    # ------------------------------------------------------------------
    # Phase 4: Risk analysis (unchanged from original)
    # ------------------------------------------------------------------

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

    # ------------------------------------------------------------------
    # Data grouping (no LLM)
    # ------------------------------------------------------------------

    def _group_data(self, excel_payload: ExcelPayload, mapping_result: StepMappingResult) -> list[StepDetail]:
        """Group Excel rows by step and operation type based on mapping result."""
        sheet_rows = {}
        for sheet in excel_payload.sheets:
            sheet_rows[sheet.sheet_name] = sheet.rows

        step_details = []
        for mapping in mapping_result.step_mappings:
            rows = sheet_rows.get(mapping.source_sheet, [])
            step_rows = [rows[i] for i in mapping.row_indices if i < len(rows)]
            operation_groups = self._group_by_operation(step_rows, mapping.source_sheet, excel_payload)
            description = mapping.description or ""

            step_details.append(StepDetail(
                step_name=mapping.step_name,
                step_description=description,
                operation_groups=operation_groups,
            ))

        return step_details

    def _group_by_operation(self, rows: list[dict], sheet_name: str, excel_payload: ExcelPayload) -> list[OperationGroup]:
        """Group rows by their operation type (action column value)."""
        action_col = None
        for sheet in excel_payload.sheets:
            if sheet.sheet_name == sheet_name:
                action_col = sheet.detected_action_column
                break

        if not action_col:
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

    # ------------------------------------------------------------------
    # Main pipeline
    # ------------------------------------------------------------------

    def generate(self, excel_payload: ExcelPayload, schedule_table: ScheduleTable | None = None) -> ImplementationPlan:
        """
        Full generation pipeline (per-sheet architecture):
        1. Per-sheet step mapping (Call 1 x N)
        2. Group data by step -> operation_type
        3. Per-sheet summary (Call 2 x N) + synthesis (Call 2 + 1)
        4. Risk analysis (Call 3)
        5. Merge into ImplementationPlan
        """
        mapping_rules = self._load_mapping_rules()

        # === Phase 1: Per-sheet step mapping ===
        logger.info("Phase 1: Per-sheet step mapping...")
        all_mappings: list[StepMapping] = []

        for sheet in excel_payload.sheets:
            sheet_rules = self._extract_sheet_rules(mapping_rules, sheet.sheet_name)
            if not sheet_rules:
                logger.info(f"  Skipping sheet '{sheet.sheet_name}': no matching rules")
                continue

            logger.info(f"  Mapping sheet '{sheet.sheet_name}' ({len(sheet.rows)} rows)...")
            sheet_json = self._serialize_sheet(sheet)
            result = self._do_step_mapping_for_sheet(sheet, sheet_json, sheet_rules)
            all_mappings.extend(result.step_mappings)

        mapping_result = StepMappingResult(
            step_mappings=all_mappings,
        )

        # === Phase 2: Group data ===
        logger.info("Phase 2: Grouping data by step and operation type...")
        step_details = self._group_data(excel_payload, mapping_result)

        # === Phase 3: Per-sheet summary + synthesis ===
        logger.info("Phase 3: Per-sheet summary generation...")
        sheet_summaries: list[SheetSummary] = []

        for sheet in excel_payload.sheets:
            sheet_mappings = [m for m in all_mappings if m.source_sheet == sheet.sheet_name]
            if not sheet_mappings:
                continue

            logger.info(f"  Summarizing sheet '{sheet.sheet_name}'...")
            sheet_json = self._serialize_sheet(sheet)
            summary = self._do_summary_for_sheet(sheet, sheet_json, sheet_mappings)
            sheet_summaries.append(summary)

        logger.info("  Synthesizing final summary...")
        summary = self._do_summary_synthesis(sheet_summaries)

        # === Phase 4: Risk analysis ===
        logger.info("Phase 4: Generating risk analysis...")
        steps_summary = self._build_steps_summary(step_details)
        verification, rollback, risk = self._do_risk_analysis(steps_summary)

        # === Phase 5: Merge ===
        task_count = len(schedule_table.rows) if schedule_table else len(step_details)
        module_count = len(set(
            row.get("APPID", "") for row in schedule_table.rows
        )) if schedule_table else len(excel_payload.summary.unique_apps)

        plan = ImplementationPlan(
            summary=summary,
            task_count=task_count,
            module_count=module_count,
            high_risk_count=0,
            schedule_table=schedule_table,
            step_details=step_details,
            verification_plan=verification,
            rollback_plan=rollback,
            risk_analysis=risk,
        )
        logger.info("Plan generation complete.")
        return plan
