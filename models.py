"""
Pydantic models for structured outputs and data validation
"""

from typing import Dict, List, Any, Optional, Literal
from pydantic import BaseModel, Field, field_validator, ConfigDict
from enum import Enum


class QuestionType(str, Enum):
    """Types of questions/items in the document"""
    PARENT_HEADER = "parent_header"
    NUMBERED_REQUIREMENT = "numbered_requirement"
    LETTERED_REQUIREMENT = "lettered_requirement"
    SUB_LIST_REQUIREMENT = "sub_list_requirement"
    BULLET_ITEM = "bullet_item"
    GENERAL_QUESTION = "general_question"


class SheetType(str, Enum):
    """Types of sheets in the workbook"""
    CONTENT_SHEET = "content_sheet"
    QUESTION_SHEET = "question_sheet"
    REFERENCE_SHEET = "reference_sheet"
    SUMMARY_SHEET = "summary_sheet"


class HierarchicalPattern(BaseModel):
    """Patterns for detecting hierarchical structure"""
    parent_indicators: List[str] = Field(
        default_factory=lambda: ["ends with :", "contains 'following'"],
        description="Patterns indicating parent questions"
    )
    list_patterns: List[str] = Field(
        default_factory=lambda: ["starts with number)", "starts with letter)"],
        description="Patterns for list items"
    )
    requirement_patterns: List[str] = Field(
        default_factory=lambda: ["starts with letter."],
        description="Patterns for actual requirements"
    )


class SheetExtractionStrategy(BaseModel):
    """Strategy for extracting questions from a sheet"""
    question_columns: List[str] = Field(
        default_factory=list,
        description="Columns containing questions/requirements"
    )
    answer_columns: List[str] = Field(
        default_factory=list,
        description="Columns for responses"
    )
    hierarchy_column: Optional[str] = Field(
        None,
        description="Column containing hierarchy info (e.g., section numbers)"
    )
    start_row: int = Field(2, description="Row to start extraction")
    skip_rows: List[int] = Field(default_factory=list)
    special_patterns: List[str] = Field(default_factory=list)
    hierarchical_patterns: Optional[HierarchicalPattern] = None
    question_column: Optional[str] = None

    @field_validator('question_columns', mode='before')
    def validate_question_columns(cls, v, info):
        if not v and info.data.get('question_column'):
            return [info.data['question_column']]
        return v if v else []


class SheetAnalysis(BaseModel):
    """Analysis result for a single sheet"""
    sheet_type: SheetType
    purpose: str
    contains_questions: bool = False
    skip_extraction: bool = False
    extraction_strategy: Optional[SheetExtractionStrategy] = None
    relationship_to_content: Optional[str] = None
    confidence: Literal["high", "medium", "low"] = "medium"


class DocumentOverview(BaseModel):
    """Overview of the entire document"""
    document_type: str
    total_question_sheets: int
    common_structure: Optional[str] = None


class SheetsAnalysisResult(BaseModel):
    """Complete analysis of all sheets"""
    sheets_analysis: Dict[str, SheetAnalysis]
    document_overview: DocumentOverview


class ColumnDetectionResult(BaseModel):
    """Result of column detection for a sheet"""
    question_column: str
    answer_columns: List[str]
    hierarchy_column: Optional[str] = None
    column_purposes: Dict[str, str]
    hierarchical_patterns: Optional[HierarchicalPattern] = None
    start_row: int = 2
    skip_rows: List[int] = Field(default_factory=list)
    confidence: Literal["high", "medium", "low"] = "high"


class FillingInstructions(BaseModel):
    """Instructions for filling the document"""
    general: str
    by_section: Dict[str, str] = Field(default_factory=dict)


class AnswerGuidelines(BaseModel):
    """Guidelines for answers"""
    compliance_responses: List[str] = Field(
        default_factory=lambda: ["Yes", "No", "Partial", "N/A"]
    )
    detail_requirements: Optional[str] = None
    evidence_requirements: Optional[str] = None


class GlobalContext(BaseModel):
    """Global context extracted from content sheet"""
    document_type: str
    document_purpose: str
    filling_instructions: FillingInstructions
    sheet_relationships: Dict[str, str] = Field(default_factory=dict)
    answer_guidelines: AnswerGuidelines
    terminology: Dict[str, Dict[str, str]] = Field(default_factory=dict)
    evaluation_criteria: Optional[str] = None
    special_notes: List[str] = Field(default_factory=list)


class ColumnFillStrategy(BaseModel):
    """Strategy for filling a specific column"""
    purpose: str
    positive_values: List[str] = Field(default_factory=list)
    negative_values: List[str] = Field(default_factory=list)
    partial_values: List[str] = Field(default_factory=list)
    conditional_logic: Optional[str] = None
    empty_probability: float = 0.0


class FillDistribution(BaseModel):
    """Distribution of response types"""
    positive: int = 70
    negative: int = 15
    partial: int = 15

    @field_validator('positive', 'negative', 'partial')
    def validate_percentage(cls, v):
        if not 0 <= v <= 100:
            raise ValueError('Percentage must be between 0 and 100')
        return v


class FillStrategy(BaseModel):
    """Complete strategy for filling a sheet"""
    distribution: FillDistribution
    column_strategies: Dict[str, ColumnFillStrategy]
    cross_column_rules: List[str] = Field(default_factory=list)


class ExtractedQuestion(BaseModel):
    """A single extracted question with hierarchy info"""
    model_config = ConfigDict(use_enum_values=True)

    question_id: int
    row_id: int
    column_letter: str
    question: str
    answers: Dict[str, str] = Field(default_factory=dict)
    question_type: QuestionType = QuestionType.GENERAL_QUESTION
    is_parent: bool = False
    should_fill: bool = True
    parent_id: Optional[int] = None
    parent_text: Optional[str] = None
    hierarchy_level: Optional[int] = None


class HierarchyStats(BaseModel):
    """Statistics about hierarchical structure"""
    parent_headers: int = 0
    numbered_requirements: int = 0
    lettered_requirements: int = 0
    sub_list_requirements: int = 0
    bullet_items: int = 0
    total_fillable: int = 0


class ExtractionResult(BaseModel):
    """Result of question extraction from a sheet"""
    sheet_name: str
    document_structure: Dict[str, Any]
    total_questions_extracted: int
    questions: List[ExtractedQuestion]
    statistics: Dict[str, Any]
    hierarchy_stats: Optional[HierarchyStats] = None
    column_info: Optional[ColumnDetectionResult] = None