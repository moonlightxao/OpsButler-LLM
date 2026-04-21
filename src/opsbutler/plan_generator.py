import json
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from opsbutler.config import Config
from opsbutler.models import (
    ExcelPayload, SheetData, StepMappingResult, StepMapping,
    SummarySection, VerificationPlan, RollbackPlan, RiskAnalysis,
    ImplementationPlan, StepDetail, OperationGroup,
    SheetSummary, ScheduleTable, LargeSheetMapping, PrepTable,
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

    def _load_zip_sheets(self) -> set[str]:
        """Parse mapping_rules.md and return set of sheet names marked with '- 打包方式: zip'."""
        mapping_rules = self._load_mapping_rules()
        if not mapping_rules:
            return set()

        sections = mapping_rules.split("\n## ")
        zip_sheets: set[str] = set()

        for section in sections[1:]:
            current_sheet = None
            is_zip = False
            for line in section.split("\n"):
                stripped = line.strip()
                if stripped.startswith("- 来源Sheet:") or stripped.startswith("- 来源 Sheet:"):
                    current_sheet = stripped.split(":", 1)[1].strip()
                elif stripped.startswith("- 打包方式:"):
                    mode = stripped.split(":", 1)[1].strip().lower()
                    is_zip = (mode == "zip")
            if current_sheet and is_zip:
                zip_sheets.add(current_sheet)

        return zip_sheets

    # ------------------------------------------------------------------
    # Per-sheet JSON helpers
    # ------------------------------------------------------------------

    def _serialize_sheet(self, sheet: SheetData) -> str:
        """Serialize a single sheet's rows to compact JSON."""
        return json.dumps(
            {"sheet_name": sheet.sheet_name, "rows": sheet.rows},
            ensure_ascii=False,
        )

    def _batch_sheet_rows(self, sheet: SheetData) -> list[tuple[int, SheetData]]:
        """Split a sheet's rows into batches by batch_size.

        Returns list of (offset, SheetData) tuples where offset is the
        starting row index in the original sheet.
        """
        batch_size = self.config.llm.batch_size
        rows = sheet.rows
        if len(rows) <= batch_size:
            return [(0, sheet)]

        batches = []
        for i in range(0, len(rows), batch_size):
            batch_rows = rows[i:i + batch_size]
            batch_sheet = SheetData(
                sheet_name=sheet.sheet_name,
                headers=sheet.headers,
                rows=batch_rows,
                detected_action_column=sheet.detected_action_column,
                detected_app_column=sheet.detected_app_column,
            )
            batches.append((i, batch_sheet))
        return batches

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
        # Normalize: LLM may return a bare list [{...}, ...] instead of {"step_mappings": [...]}
        if isinstance(result, list):
            result = {"step_mappings": result}
        return StepMappingResult(**result)

    def _do_step_mapping_for_large_sheet(
        self, sheet: SheetData, sheet_rules: str
    ) -> LargeSheetMapping:
        """LLM Call for large sheets: analyze deduplicated operations to determine platform steps."""
        template = self._load_prompt("step_mapping_large_sheet.txt")
        user_message = template.format(
            mapping_rules=sheet_rules,
            sheet_name=sheet.sheet_name,
            unique_operations=json.dumps(sheet.unique_operations, ensure_ascii=False),
        )

        messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": user_message},
        ]

        result = self.llm.chat_json(messages)
        # Normalize: LLM may return a bare list instead of {"step_mappings": [...]}
        if isinstance(result, list):
            result = {"step_mappings": result}
        return LargeSheetMapping(**result)

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
        # Defensive: LLM may return dict instead of string for these fields
        for key in ("changed_apps", "changes_summary"):
            if key in result and isinstance(result[key], dict):
                result[key] = str(result[key])
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
        # Defensive: LLM may return list instead of dict for these fields
        vp = result["verification_plan"]
        rp = result["rollback_plan"]
        ra = result["risk_analysis"]
        verification = VerificationPlan(**vp) if isinstance(vp, dict) else VerificationPlan(verification_steps=vp)
        rollback = RollbackPlan(**rp) if isinstance(rp, dict) else RollbackPlan(rollback_steps=rp)
        risk = RiskAnalysis(**ra) if isinstance(ra, dict) else RiskAnalysis(risks=ra)
        return verification, rollback, risk

    # ------------------------------------------------------------------
    # Data grouping (no LLM)
    # ------------------------------------------------------------------

    def _group_data(self, excel_payload: ExcelPayload, mapping_result: StepMappingResult, zip_sheets: set[str] | None = None, large_sheet_ops: dict | None = None) -> list[StepDetail]:
        """Group Excel rows by step and operation type based on mapping result."""
        if zip_sheets is None:
            zip_sheets = set()
        if large_sheet_ops is None:
            large_sheet_ops = {}
        sheet_rows = {}
        for sheet in excel_payload.sheets:
            sheet_rows[sheet.sheet_name] = sheet.rows

        step_details = []
        for mapping in mapping_result.step_mappings:
            rows = sheet_rows.get(mapping.source_sheet, [])

            if mapping.source_sheet in large_sheet_ops:
                # Large sheet: use all rows for ZIP content
                operation_groups = self._group_by_operation(rows, mapping.source_sheet, excel_payload)
                step_details.append(StepDetail(
                    step_name=mapping.step_name,
                    step_description=mapping.description or "",
                    operation_groups=operation_groups,
                    source_sheet=mapping.source_sheet,
                    is_zip=True,
                    is_large_sheet=True,
                    operation_descriptions=large_sheet_ops.get(mapping.source_sheet, {}),
                ))
            else:
                # Normal sheet: use mapped rows only
                step_rows = [rows[i] for i in mapping.row_indices if i < len(rows)]
                operation_groups = self._group_by_operation(step_rows, mapping.source_sheet, excel_payload)
                step_details.append(StepDetail(
                    step_name=mapping.step_name,
                    step_description=mapping.description or "",
                    operation_groups=operation_groups,
                    source_sheet=mapping.source_sheet,
                    is_zip=(mapping.source_sheet in zip_sheets),
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

    def generate(self, excel_payload: ExcelPayload, schedule_table: ScheduleTable | None = None, prep_table: PrepTable | None = None) -> ImplementationPlan:
        """
        Full generation pipeline (per-sheet architecture):
        1. Per-sheet step mapping (Call 1 x N)
        2. Group data by step -> operation_type
        3. Per-sheet summary (Call 2 x N) + synthesis (Call 2 + 1)
        4. Risk analysis (Call 3)
        5. Merge into ImplementationPlan
        """
        mapping_rules = self._load_mapping_rules()
        zip_sheets = self._load_zip_sheets()

        # === Phase 1: Per-sheet step mapping ===
        logger.info("Phase 1: Per-sheet step mapping...")
        all_mappings: list[StepMapping] = []
        large_sheet_ops: dict[str, dict[str, str]] = {}  # sheet_name -> {op -> desc}

        # Process "上线制品包" first so artifact steps appear first in the Word document
        sorted_sheets = sorted(
            excel_payload.sheets,
            key=lambda s: 0 if s.sheet_name == "上线制品包" else 1,
        )

        for sheet in sorted_sheets:
            sheet_rules = self._extract_sheet_rules(mapping_rules, sheet.sheet_name)
            if not sheet_rules:
                logger.info(f"  Skipping sheet '{sheet.sheet_name}': no matching rules")
                continue

            if sheet.is_large:
                # Large sheet mode: dedup operations -> LLM analysis
                logger.info(f"  Mapping large sheet '{sheet.sheet_name}' ({len(sheet.rows)} rows, {len(sheet.unique_operations)} unique ops)...")
                large_result = self._do_step_mapping_for_large_sheet(sheet, sheet_rules)
                all_mappings.extend(large_result.step_mappings)
                large_sheet_ops[sheet.sheet_name] = large_result.operation_descriptions
                continue

            batches = self._batch_sheet_rows(sheet)
            logger.info(f"  Mapping sheet '{sheet.sheet_name}' ({len(sheet.rows)} rows, {len(batches)} batch(es))...")

            if len(batches) == 1:
                # Single batch — no concurrency needed
                _, batch_sheet = batches[0]
                sheet_json = self._serialize_sheet(batch_sheet)
                result = self._do_step_mapping_for_sheet(batch_sheet, sheet_json, sheet_rules)
                all_mappings.extend(result.step_mappings)
            else:
                # Multiple batches — concurrent execution
                max_workers = self.config.llm.max_workers

                def _map_batch(offset: int, bs: SheetData) -> list[StepMapping]:
                    sj = self._serialize_sheet(bs)
                    r = self._do_step_mapping_for_sheet(bs, sj, sheet_rules)
                    # Adjust row_indices to global offset
                    for m in r.step_mappings:
                        m.row_indices = [idx + offset for idx in m.row_indices]
                    return r.step_mappings

                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    futures = {
                        executor.submit(_map_batch, offset, bs): offset
                        for offset, bs in batches
                    }
                    for future in as_completed(futures):
                        all_mappings.extend(future.result())

        mapping_result = StepMappingResult(
            step_mappings=all_mappings,
        )

        # === Phase 2: Group data ===
        logger.info("Phase 2: Grouping data by step and operation type...")
        step_details = self._group_data(excel_payload, mapping_result, zip_sheets, large_sheet_ops)

        # === Phase 3: Per-sheet summary + synthesis ===
        logger.info("Phase 3: Per-sheet summary generation...")
        sheet_summaries: list[SheetSummary] = []

        for sheet in excel_payload.sheets:
            sheet_mappings = [m for m in all_mappings if m.source_sheet == sheet.sheet_name]
            if not sheet_mappings:
                continue

            if sheet.is_large:
                # Large sheet: build summary directly from dedup operations (no LLM call)
                ops_text = ", ".join(sheet.unique_operations)
                sheet_summaries.append(SheetSummary(
                    sheet_name=sheet.sheet_name,
                    changed_apps=f"共{len(sheet.unique_operations)}种操作类型",
                    changes_summary=f"操作类型包括: {ops_text}",
                ))
                continue

            batches = self._batch_sheet_rows(sheet)
            logger.info(f"  Summarizing sheet '{sheet.sheet_name}' ({len(sheet.rows)} rows, {len(batches)} batch(es))...")

            if len(batches) == 1:
                _, batch_sheet = batches[0]
                sheet_json = self._serialize_sheet(batch_sheet)
                summary = self._do_summary_for_sheet(batch_sheet, sheet_json, sheet_mappings)
                sheet_summaries.append(summary)
            else:
                max_workers = self.config.llm.max_workers
                batch_summaries: list[SheetSummary] = []

                def _summarize_batch(offset: int, bs: SheetData) -> SheetSummary:
                    sj = self._serialize_sheet(bs)
                    # Filter mappings to those with row indices in this batch range
                    batch_end = offset + len(bs.rows)
                    batch_mappings = [
                        m for m in sheet_mappings
                        if any(offset <= idx < batch_end for idx in m.row_indices)
                    ]
                    return self._do_summary_for_sheet(bs, sj, batch_mappings)

                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                    futures = {
                        executor.submit(_summarize_batch, offset, bs): offset
                        for offset, bs in batches
                    }
                    for future in as_completed(futures):
                        batch_summaries.append(future.result())

                # Merge batch summaries into one
                merged_apps = "; ".join(s.changed_apps for s in batch_summaries)
                merged_desc = "\n".join(s.changes_summary for s in batch_summaries)
                sheet_summaries.append(SheetSummary(
                    sheet_name=sheet.sheet_name,
                    changed_apps=merged_apps,
                    changes_summary=merged_desc,
                ))

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
            prep_table=prep_table,
            step_details=step_details,
            verification_plan=verification,
            rollback_plan=rollback,
            risk_analysis=risk,
        )
        logger.info("Plan generation complete.")
        return plan
