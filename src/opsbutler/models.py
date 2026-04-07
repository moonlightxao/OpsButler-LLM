from pydantic import BaseModel, Field
from typing import Optional, Any
from datetime import datetime


class ExcelSummary(BaseModel):
    total_sheets: int
    total_rows: int
    unique_apps: list[str]
    unique_operation_types: list[str]
    sheet_names: list[str]


class SheetData(BaseModel):
    sheet_name: str
    headers: list[str]
    rows: list[dict[str, Any]]
    detected_action_column: Optional[str] = None
    detected_app_column: Optional[str] = None


class ExcelPayload(BaseModel):
    source_file: str
    sheets: list[SheetData]
    parsed_at: datetime = Field(default_factory=datetime.now)
    summary: ExcelSummary


# LLM Step Mapping output
class StepMapping(BaseModel):
    step_name: str
    source_sheet: str
    row_indices: list[int]
    description: Optional[str] = None
    notes: Optional[str] = None


class StepMappingResult(BaseModel):
    step_mappings: list[StepMapping]
    task_table: list["TaskEntry"]


# Operation grouping within a step
class OperationGroup(BaseModel):
    operation_type: str
    rows: list[dict[str, Any]]


class StepDetail(BaseModel):
    step_name: str
    step_description: str  # from mapping_rules.md
    operation_groups: list[OperationGroup]


# LLM Summary output (Section 1)
class SummarySection(BaseModel):
    changed_apps: str
    reason_and_purpose: str
    impact_analysis: str


# Task table entry (实施总表)
class TaskEntry(BaseModel):
    sequence: int
    task_name: str
    start_time: str = ""
    end_time: str = ""
    operator: str = ""
    reviewer: str = ""


# Risk analysis (Sections 3/4/5)
class VerificationPlan(BaseModel):
    verification_steps: list[str]


class RollbackPlan(BaseModel):
    rollback_steps: list[str]


class RiskEntry(BaseModel):
    risk_description: str
    probability: str
    impact: str
    mitigation: str


class RiskAnalysis(BaseModel):
    risks: list[RiskEntry]


# Complete implementation plan
class ImplementationPlan(BaseModel):
    summary: SummarySection
    task_count: int
    module_count: int
    high_risk_count: int
    task_table: list[TaskEntry]
    step_details: list[StepDetail]
    verification_plan: VerificationPlan
    rollback_plan: RollbackPlan
    risk_analysis: RiskAnalysis
