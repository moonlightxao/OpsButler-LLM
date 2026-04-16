from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from opsbutler.models import ImplementationPlan


class WordGenerator:
    """Generate implementation plan Word document from ImplementationPlan data."""

    def generate(self, plan: ImplementationPlan, output_path: str) -> str:
        """Generate Word document and save to output_path. Returns output_path."""
        doc = Document()

        # Title
        doc.add_heading("上线checklist - 实施方案", level=0)

        # Section 1
        self._add_section1(doc, plan)

        # Summary info
        self._add_summary_info(doc, plan)

        # Section 2
        self._add_section2(doc, plan)

        # Section 3
        self._add_verification(doc, plan)

        # Section 4
        self._add_rollback(doc, plan)

        # Section 5
        self._add_risk(doc, plan)

        doc.save(output_path)
        return output_path

    def _add_section1(self, doc, plan: ImplementationPlan):
        doc.add_heading("1 原因和目的", level=1)

        doc.add_heading("1.1 变更应用", level=2)
        doc.add_paragraph(plan.summary.changed_apps)

        doc.add_heading("1.2 变更原因和目的", level=2)
        doc.add_paragraph(plan.summary.reason_and_purpose)

        doc.add_heading("1.3 变更影响", level=2)
        doc.add_paragraph(plan.summary.impact_analysis)

    def _add_summary_info(self, doc, plan: ImplementationPlan):
        # Add empty paragraph for spacing
        doc.add_paragraph("")

        # Bold summary header
        p = doc.add_paragraph()
        run = p.add_run("【摘要信息】")
        run.bold = True

        doc.add_paragraph(f"• 任务总数：{plan.task_count}")
        doc.add_paragraph(f"• 涉及模块：{plan.module_count} 个")
        doc.add_paragraph(f"• 高危操作：{plan.high_risk_count} 个")

    def _add_section2(self, doc, plan: ImplementationPlan):
        doc.add_heading("2 实施步骤和计划", level=1)

        # Add empty paragraph for spacing
        doc.add_paragraph("")

        # Task summary table (实施总表)
        self._add_task_table(doc, plan)

        # Detailed steps
        doc.add_heading("2.1 详细实施步骤", level=3)

        for idx, step in enumerate(plan.step_details, 1):
            self._add_step(doc, idx, step)

    def _add_task_table(self, doc, plan: ImplementationPlan):
        """Add the implementation summary table."""
        if plan.schedule_table:
            self._add_schedule_table(doc, plan.schedule_table)
        else:
            self._add_legacy_task_table(doc, plan.task_table)

    def _add_schedule_table(self, doc, schedule_table):
        """Render task table from '变更安排' sheet data."""
        headers = schedule_table.headers
        rows = schedule_table.rows

        table = doc.add_table(rows=1 + len(rows), cols=len(headers), style="Table Grid")

        # Header row
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = str(header)

        # Data rows
        for row_idx, row_data in enumerate(rows, 1):
            for col_idx, header in enumerate(headers):
                value = row_data.get(header, "")
                table.rows[row_idx].cells[col_idx].text = str(value) if value is not None else ""

    def _add_legacy_task_table(self, doc, task_table):
        """Legacy task table rendering (backward compat)."""
        headers = ["序号", "任务", "开始时间", "结束时间", "实施人", "复核人"]
        table = doc.add_table(rows=1, cols=len(headers), style="Table Grid")

        # Header row
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header

        # Data rows
        for task in task_table:
            row = table.add_row()
            row.cells[0].text = str(task.sequence)
            row.cells[1].text = task.task_name
            row.cells[2].text = task.start_time
            row.cells[3].text = task.end_time
            row.cells[4].text = task.operator
            row.cells[5].text = task.reviewer

    def _add_step(self, doc, step_index: int, step):
        """Add a single step section with heading, description, and operation tables."""
        doc.add_heading(f"2.1.{step_index} {step.step_name}", level=4)

        # Step description
        if step.step_description:
            doc.add_paragraph(step.step_description)

        # Operation groups
        for group in step.operation_groups:
            # Operation type label
            doc.add_paragraph(f"【{group.operation_type}】")
            doc.add_paragraph(f"执行以下{group.operation_type}操作：")

            # Data table
            if group.rows:
                self._add_data_table(doc, group.rows)

    def _add_data_table(self, doc, rows: list[dict]):
        """Add a data table from row dicts. Column headers come from dict keys."""
        if not rows:
            return

        headers = list(rows[0].keys())
        # Filter out None values from headers
        headers = [h for h in headers if h]

        table = doc.add_table(rows=1, cols=len(headers), style="Table Grid")

        # Header row
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = str(header)

        # Data rows
        for row_data in rows:
            row = table.add_row()
            for i, header in enumerate(headers):
                value = row_data.get(header, "")
                row.cells[i].text = str(value) if value is not None else ""

    def _add_verification(self, doc, plan: ImplementationPlan):
        doc.add_heading("3 实施后验证计划", level=1)
        text = "\n".join(f"{i+1}. {step}" for i, step in enumerate(plan.verification_plan.verification_steps))
        doc.add_paragraph(text)

    def _add_rollback(self, doc, plan: ImplementationPlan):
        doc.add_heading("4 应急回退措施", level=1)
        text = "\n".join(f"{i+1}. {step}" for i, step in enumerate(plan.rollback_plan.rollback_steps))
        doc.add_paragraph(text)

    def _add_risk(self, doc, plan: ImplementationPlan):
        doc.add_heading("5 风险分析和规避措施", level=1)

        if not plan.risk_analysis.risks:
            doc.add_paragraph("本次变更无高风险项。")
            return

        # Risk table
        headers = ["风险描述", "概率", "影响", "规避措施"]
        table = doc.add_table(rows=1, cols=len(headers), style="Table Grid")

        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header

        for risk in plan.risk_analysis.risks:
            row = table.add_row()
            row.cells[0].text = risk.risk_description
            row.cells[1].text = risk.probability
            row.cells[2].text = risk.impact
            row.cells[3].text = risk.mitigation
