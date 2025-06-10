"""
Clean Gemini API client using latest Gen AI SDK with structured output
Uses Pydantic models for type-safe responses
"""

import os
from typing import Dict, List, Any, Optional
from google import genai

from models import (
    SheetsAnalysisResult, SheetAnalysis, SheetType, DocumentOverview,
    ColumnDetectionResult, GlobalContext,
    FillingInstructions, AnswerGuidelines, FillStrategy, FillDistribution,
    ColumnFillStrategy, SheetExtractionStrategy,
    SheetsAnalysisResponse, ColumnDetectionResponse, GlobalContextResponse,
    FillStrategyResponse
)


class GeminiStructuredClient:
    """Clean Gemini client using structured output"""

    def __init__(self):
        api_key = os.getenv('GEMINI_API_KEY')
        if not api_key:
            raise ValueError("GEMINI_API_KEY environment variable is required")

        self.client = genai.Client(api_key=api_key)
        self.model_name = 'gemini-2.5-flash-preview-05-20'
        
        # Base generation config
        self.base_config = {
            'temperature': 0,
        }

    def analyze_sheets_intelligently(self, sheets_info: List[Dict],
                                     global_context: Optional[GlobalContext] = None) -> SheetsAnalysisResult:
        """LLM-based sheet analysis using Gemini with structured output"""

        print("🧠 Using Gemini LLM-based intelligent sheet analysis...")
        print("📝 Analyzing actual content patterns to classify sheets...")

        analysis_data = self._prepare_sheet_analysis_data(sheets_info)

        try:
            llm_result = self._get_gemini_sheet_analysis(analysis_data, global_context)
            if llm_result:
                print("✅ Gemini sheet analysis completed successfully")
                return llm_result
        except Exception as e:
            print(f"⚠️  Gemini sheet analysis failed: {str(e)[:100]}")

        print("🔄 Using intelligent pattern-based analysis...")
        return self._intelligent_pattern_analysis(sheets_info)

    def _prepare_sheet_analysis_data(self, sheets_info: List[Dict]) -> Dict[str, Any]:
        """Prepare comprehensive data for LLM analysis"""

        analysis_data = {
            "document_info": {
                "total_sheets": len(sheets_info),
                "sheet_names": [sheet['name'] for sheet in sheets_info]
            },
            "sheets": []
        }

        for sheet in sheets_info:
            sheet_data = {
                "name": sheet['name'],
                "dimensions": f"{sheet.get('rows', 0)} rows x {sheet.get('columns', 0)} columns",
                "headers": [],
                "sample_content": []
            }

            headers = sheet.get('headers', [])
            for header in headers[:8]:
                if header and header.get('value'):
                    sheet_data["headers"].append({
                        "column": header.get('column', ''),
                        "header": str(header['value'])[:50]
                    })

            sample_data = sheet.get('sample_data', [])
            if sample_data and len(sample_data) > 1:
                for row_idx in range(1, min(6, len(sample_data))):
                    if row_idx < len(sample_data):
                        row = sample_data[row_idx]
                        row_content = []
                        for col_idx in range(min(4, len(row))):
                            if col_idx < len(row) and row[col_idx]:
                                content = str(row[col_idx])[:80]
                                if content.strip():
                                    row_content.append(content)
                        if row_content:
                            sheet_data["sample_content"].append(row_content)

            analysis_data["sheets"].append(sheet_data)

        return analysis_data

    def _get_gemini_sheet_analysis(self, analysis_data: Dict[str, Any],
                                   global_context: Optional[GlobalContext] = None) -> Optional[SheetsAnalysisResult]:
        """Get Gemini analysis of sheet types using structured output"""

        context_section = ""
        if global_context:
            context_section = f"""
DOCUMENT CONTEXT:
Type: {global_context.document_type}
Purpose: {global_context.document_purpose}
"""

        prompt = f"""Analyze these Excel sheets to intelligently classify them as CONTENT sheets vs QUESTION sheets.

{context_section}

SHEET ANALYSIS DATA:
{analysis_data}

CLASSIFICATION TASK:
Analyze each sheet's actual content, structure, and purpose to determine:

1. CONTENT/INSTRUCTION SHEETS:
   - Contain explanations, guidelines, or instructions
   - Have descriptive text about how to complete the document
   - Usually smaller, with explanatory content
   - Provide context or background information

2. QUESTION/REQUIREMENT SHEETS:
   - Contain actual questions, requirements, or items to be answered
   - Have structured format with questions and response areas
   - Usually larger, with systematic question-answer structure
   - Designed for data collection or responses

ANALYSIS PRINCIPLES:
- Base decisions on ACTUAL CONTENT and STRUCTURE, not just sheet names
- Consider the purpose and workflow of each sheet
- Look at headers, sample content, and overall organization
- Identify which sheet would serve as the instruction/context source

Provide a structured analysis with content sheet detection, classification reasoning, 
detailed sheet analysis for each sheet, and document overview."""

        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=prompt,
                config={
                    **self.base_config,
                    'response_mime_type': 'application/json',
                    'response_schema': SheetsAnalysisResponse,
                }
            )

            result = response.parsed

            sheets_analysis = {}
            content_sheet_name = result.content_sheet_detected

            print(f"🎯 Gemini detected content sheet: {content_sheet_name or 'None'}")
            print(f"📋 Analysis approach: {result.classification_reasoning[:80]}...")

            for sheet_name, analysis in result.sheets_analysis.items():
                sheet_type = SheetType.CONTENT_SHEET if analysis.sheet_type == "content_sheet" else SheetType.QUESTION_SHEET

                print(f"  📄 '{sheet_name}': {analysis.sheet_type} - {analysis.reasoning[:60]}...")

                extraction_strategy = None
                if sheet_type == SheetType.QUESTION_SHEET and analysis.extraction_strategy:
                    extraction_strategy = SheetExtractionStrategy(
                        question_columns=analysis.extraction_strategy.question_columns,
                        answer_columns=analysis.extraction_strategy.answer_columns,
                        start_row=analysis.extraction_strategy.start_row
                    )

                sheets_analysis[sheet_name] = SheetAnalysis(
                    sheet_type=sheet_type,
                    purpose=analysis.purpose,
                    contains_questions=analysis.contains_questions,
                    skip_extraction=analysis.skip_extraction,
                    extraction_strategy=extraction_strategy,
                    confidence=analysis.confidence
                )

            document_overview = DocumentOverview(
                document_type=result.document_overview.document_type,
                total_question_sheets=result.document_overview.total_question_sheets
            )

            return SheetsAnalysisResult(
                sheets_analysis=sheets_analysis,
                document_overview=document_overview
            )

        except Exception as e:
            print(f"🔴 Gemini sheet analysis failed: {str(e)[:100]}")
            return None

    def detect_columns_with_statistics(self, worksheet_data: Dict, sheet_name: str) -> ColumnDetectionResult:
        """Gemini-based intelligent column detection with structured output"""

        headers = worksheet_data.get('headers', [])
        samples = worksheet_data.get('samples', [])

        prompt = f"""Analyze this Excel sheet to intelligently detect question and answer columns.

SHEET: {sheet_name}
HEADERS: {headers}
SAMPLE DATA (first 6 rows): {samples[:6]}

INTELLIGENT ANALYSIS TASK:
Analyze the actual content, structure, and purpose of this sheet to identify:

1. QUESTION COLUMN: The column containing the actual questions, requirements, or items
   - Look for columns with substantial descriptive content
   - May contain questions, specifications, requirements, or criteria
   - Often has longer, varied text content
   - Could be ANY column position (A, B, G, etc.) - analyze the content!

2. ANSWER COLUMNS: Columns designed for responses or data entry
   - Look for columns intended for filling in responses
   - Often have shorter headers indicating response categories
   - Usually mostly empty in sample data (awaiting responses)
   - Should logically follow the question column in the workflow

CRITICAL ANALYSIS PRINCIPLES:
- Study ACTUAL CONTENT PATTERNS and SHEET PURPOSE
- The question column could be anywhere - don't assume position
- Answer columns should make sense for the data collection workflow
- Consider the logical flow: metadata → questions → responses

Provide structured column detection with question column, answer columns, hierarchy column if any,
column purposes, analysis reasoning, start row, and confidence level."""

        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=prompt,
                config={
                    **self.base_config,
                    'response_mime_type': 'application/json',
                    'response_schema': ColumnDetectionResponse,
                }
            )

            result = response.parsed

            print(f"    🧠 Gemini Column Analysis: {result.analysis_reasoning}")

            return ColumnDetectionResult(
                question_column=result.question_column,
                answer_columns=result.answer_columns,
                hierarchy_column=result.hierarchy_column,
                column_purposes=result.column_purposes,
                start_row=result.start_row,
                confidence=result.confidence
            )

        except Exception as e:
            print(f"    ⚠️  Gemini column detection failed: {str(e)[:50]}")
            return self._intelligent_statistical_fallback(worksheet_data)

    def extract_global_context(self, content_data: Dict) -> GlobalContext:
        """Gemini-based global context extraction with structured output"""

        prompt = f"""Analyze this content/instruction sheet to extract comprehensive global context.

CONTENT SHEET DATA:
{content_data}

COMPREHENSIVE ANALYSIS TASKS:
Analyze the actual content to understand:

1. Document Type & Purpose: What kind of document is this? (RFI, RFP, assessment, etc.)
2. Filling Instructions: How should responses be provided?
3. Sheet Relationships: How do different sections relate?
4. Answer Requirements: What format/style of answers are expected?
5. Key Terminology: Important terms and definitions
6. Evaluation Criteria: How will responses be assessed?

Base your analysis on ACTUAL CONTENT, not assumptions.

Provide structured context with document type, purpose, filling instructions,
sheet relationships, answer guidelines, terminology, evaluation criteria, and special notes."""

        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=prompt,
                config={
                    **self.base_config,
                    'response_mime_type': 'application/json',
                    'response_schema': GlobalContextResponse,
                }
            )

            result = response.parsed

            filling_instructions = FillingInstructions(
                general=result.filling_instructions.get('general', ''),
                by_section=result.filling_instructions.get('by_section', {})
            )

            answer_guidelines = AnswerGuidelines(
                compliance_responses=result.answer_guidelines.get('compliance_responses', []),
                detail_requirements=result.answer_guidelines.get('detail_requirements', ''),
                evidence_requirements=result.answer_guidelines.get('evidence_requirements', '')
            )

            return GlobalContext(
                document_type=result.document_type,
                document_purpose=result.document_purpose,
                filling_instructions=filling_instructions,
                sheet_relationships=result.sheet_relationships,
                answer_guidelines=answer_guidelines,
                terminology=result.terminology,
                evaluation_criteria=result.evaluation_criteria,
                special_notes=result.special_notes
            )

        except Exception as e:
            print(f"    ⚠️  Gemini context extraction failed: {str(e)[:50]}")
            return self._create_enhanced_fallback_context()

    def generate_intelligent_fill_strategy(self, sheet_info: Dict,
                                           global_context: Optional[GlobalContext] = None) -> FillStrategy:
        """Gemini-based intelligent filling strategy generation with structured output"""

        context_section = ""
        if global_context:
            context_section = f"""
DOCUMENT CONTEXT:
Type: {global_context.document_type}
Purpose: {global_context.document_purpose}
Instructions: {global_context.filling_instructions.general}
"""

        prompt = f"""Generate intelligent answer filling strategy with cross-column logic.
{context_section}

SHEET INFORMATION:
- Sheet Name: {sheet_info['sheet_name']}
- Fillable Questions: {sheet_info.get('fillable_questions', 0)}
- Answer Columns: {sheet_info['answer_columns']}
- Column Purposes: {sheet_info.get('column_purposes', {})}

STRATEGY REQUIREMENTS:
1. Analyze column purposes to understand relationships
2. Create cross-column rules for logical consistency
3. Generate realistic response distributions
4. Provide column-specific response values
5. Consider conditional logic and empty probabilities

Provide structured strategy with distribution percentages, column strategies, and cross-column rules."""

        try:
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=prompt,
                config={
                    **self.base_config,
                    'response_mime_type': 'application/json',
                    'response_schema': FillStrategyResponse,
                }
            )

            result = response.parsed

            distribution = FillDistribution(
                positive=result.distribution.get('positive', 70),
                negative=result.distribution.get('negative', 15),
                partial=result.distribution.get('partial', 15)
            )

            column_strategies = {}
            for col, strategy in result.column_strategies.items():
                column_strategies[col] = ColumnFillStrategy(
                    purpose=strategy.get('purpose', ''),
                    positive_values=strategy.get('positive_values', []),
                    negative_values=strategy.get('negative_values', []),
                    partial_values=strategy.get('partial_values', []),
                    conditional_logic=strategy.get('conditional_logic', ''),
                    empty_probability=strategy.get('empty_probability', 0.1)
                )

            return FillStrategy(
                distribution=distribution,
                column_strategies=column_strategies,
                cross_column_rules=result.cross_column_rules
            )

        except Exception as e:
            print(f"    ⚠️  Gemini strategy generation failed: {str(e)[:50]}")
            return self._create_intelligent_fallback_strategy(sheet_info)

    def _intelligent_pattern_analysis(self, sheets_info: List[Dict]) -> SheetsAnalysisResult:
        """Intelligent pattern-based analysis when Gemini fails"""

        print("📊 Analyzing sheets with intelligent pattern recognition...")

        sheets_analysis = {}

        for sheet in sheets_info:
            sheet_name = sheet['name']
            rows = sheet.get('rows', 0)
            columns = sheet.get('columns', 0)

            print(f"  🔍 Analyzing '{sheet_name}': {rows}r x {columns}c")

            if rows > 20:
                answer_columns = [chr(ord('B') + i) for i in range(min(columns - 1, 5))]

                sheets_analysis[sheet_name] = SheetAnalysis(
                    sheet_type=SheetType.QUESTION_SHEET,
                    purpose=f"Question sheet with {rows} items",
                    contains_questions=True,
                    skip_extraction=False,
                    extraction_strategy=SheetExtractionStrategy(
                        question_columns=['A'],
                        answer_columns=answer_columns,
                        start_row=2
                    ),
                    confidence="medium"
                )
                print(f"      ✅ QUESTION sheet (size-based heuristic)")
            else:
                sheets_analysis[sheet_name] = SheetAnalysis(
                    sheet_type=SheetType.CONTENT_SHEET,
                    purpose="Content/instruction sheet",
                    contains_questions=False,
                    skip_extraction=True,
                    confidence="medium"
                )
                print(f"      ❌ CONTENT sheet (size-based heuristic)")

        question_sheet_count = len([s for s in sheets_analysis.values()
                                    if s.sheet_type == SheetType.QUESTION_SHEET])

        return SheetsAnalysisResult(
            sheets_analysis=sheets_analysis,
            document_overview=DocumentOverview(
                document_type="Business Document",
                total_question_sheets=question_sheet_count
            )
        )

    def _intelligent_statistical_fallback(self, worksheet_data: Dict) -> ColumnDetectionResult:
        """Simple statistical fallback"""

        headers = worksheet_data.get('headers', [])

        if not headers:
            return ColumnDetectionResult(
                question_column='A',
                answer_columns=['B', 'C'],
                column_purposes={'A': 'Questions', 'B': 'Response', 'C': 'Comments'},
                start_row=2,
                confidence="low"
            )

        question_col = 'A'
        answer_cols = [chr(ord('A') + i + 1) for i in range(min(len(headers) - 1, 5))]

        column_purposes = {}
        for i, header in enumerate(headers):
            col_letter = chr(ord('A') + i)
            header_value = str(header.get('value', ''))
            if i == 0:
                column_purposes[col_letter] = f"Questions - {header_value}"
            else:
                column_purposes[col_letter] = f"Response - {header_value}"

        return ColumnDetectionResult(
            question_column=question_col,
            answer_columns=answer_cols,
            column_purposes=column_purposes,
            start_row=2,
            confidence="medium"
        )

    def _create_enhanced_fallback_context(self) -> GlobalContext:
        """Enhanced fallback context"""
        return GlobalContext(
            document_type="Request for Proposal (RFP)",
            document_purpose="Vendor capability and compliance assessment",
            filling_instructions=FillingInstructions(
                general="Provide accurate and detailed responses to all requirements based on actual capabilities"
            ),
            answer_guidelines=AnswerGuidelines(
                compliance_responses=["Yes", "No", "Partial", "N/A"],
                detail_requirements="Provide explanations for all responses, especially negative or partial ones"
            )
        )

    def _create_intelligent_fallback_strategy(self, sheet_info: Dict) -> FillStrategy:
        """Create intelligent fallback strategy based on analysis"""
        column_strategies = {}
        answer_columns = sheet_info.get('answer_columns', [])

        for col in answer_columns:
            column_strategies[col] = ColumnFillStrategy(
                purpose=f"Response column {col}",
                positive_values=["Yes", "Available", "Supported"],
                negative_values=["No", "Not Available", "Not Supported"],
                partial_values=["Partial", "Limited", "Conditional"],
                empty_probability=0.1
            )

        return FillStrategy(
            distribution=FillDistribution(positive=70, negative=15, partial=15),
            column_strategies=column_strategies,
            cross_column_rules=[
                "Ensure consistency across related columns",
                "Provide context for non-positive responses"
            ]
        )