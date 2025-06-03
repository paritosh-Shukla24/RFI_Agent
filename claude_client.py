"""
Claude API client for structured outputs and intelligent analysis - FIXED VERSION
Replace your claude_client.py with this file
"""

import os
import json
import anthropic
from typing import Dict, List, Any, Optional

from models import (
    SheetsAnalysisResult, SheetAnalysis, SheetType, DocumentOverview,
    ColumnDetectionResult, HierarchicalPattern, GlobalContext,
    FillingInstructions, AnswerGuidelines, FillStrategy, FillDistribution,
    ColumnFillStrategy, SheetExtractionStrategy
)


class ClaudeStructuredClient:
    """Enhanced client with dynamic prompts and smart detection"""

    def __init__(self):
        api_key = os.getenv('ANTHROPIC_API_KEY')
        if not api_key:
            raise ValueError("ANTHROPIC_API_KEY environment variable is required")
        self.client = anthropic.Anthropic(api_key=api_key)

    def analyze_sheets_intelligently(self, sheets_info: List[Dict],
                                     global_context: Optional[GlobalContext] = None) -> SheetsAnalysisResult:
        """Intelligent sheet analysis with dynamic detection - ALWAYS USE FALLBACK FOR RELIABILITY"""

        print("🔧 Using reliable fallback analysis (bypassing API for consistency)...")
        return self._create_intelligent_fallback_analysis(sheets_info)

    def _create_intelligent_fallback_analysis(self, sheets_info: List[Dict]) -> SheetsAnalysisResult:
        """FIXED: Intelligent fallback analysis based on content patterns"""
        sheets_analysis = {}

        print("📊 Analyzing sheets with improved logic...")

        for sheet in sheets_info:
            sheet_name = sheet['name']
            rows = sheet.get('rows', 0)
            columns = sheet.get('columns', 0)

            print(f"  🔍 Analyzing '{sheet_name}': {rows} rows x {columns} cols")

            # FIXED LOGIC: More aggressive question sheet detection
            is_question_sheet = False
            reasons = []

            # Rule 1: Content sheets by name (very specific)
            sheet_name_lower = sheet_name.lower()
            if sheet_name_lower in ['content', 'instruction', 'instructions', 'guidelines', 'notes', 'readme', 'overview', 'intro']:
                is_question_sheet = False
                reasons.append("content sheet by name")
            else:
                # Rule 2: RFI/RFP category sheets are ALWAYS question sheets
                rfi_categories = ['integration', 'supplier', 'general', 'analytics', 'sourcing', 'contract', 'management', 'technical', 'functional', 'commercial']
                if any(word in sheet_name_lower for word in rfi_categories):
                    is_question_sheet = True
                    reasons.append("RFI category sheet")

                # Rule 3: Large sheets are question sheets (unless proven otherwise)
                if rows > 30 and columns > 3:
                    is_question_sheet = True
                    reasons.append("large sheet")

                # Rule 4: Check headers for response indicators
                headers = sheet.get('headers', [])
                response_indicators = 0

                for header in headers:
                    if header and header.get('value'):
                        header_value = str(header['value']).lower()
                        # Look for response column indicators
                        response_words = ['compliance', 'response', 'answer', 'status', 'yes', 'no', 'comment', 'remark', 'notes', 'details', 'applicable', 'supported', 'available']
                        if any(word in header_value for word in response_words):
                            response_indicators += 1

                if response_indicators > 0:
                    is_question_sheet = True
                    reasons.append(f"{response_indicators} response columns")

            # IMPORTANT: Default to question sheet for unknown multi-column sheets
            if not is_question_sheet and columns > 5 and rows > 20:
                is_question_sheet = True
                reasons.append("multi-column sheet default")

            # Generate answer columns for question sheets
            if is_question_sheet:
                # Generate answer columns B through reasonable limit
                max_answer_cols = min(columns - 1, 10)  # B through K
                answer_columns = [chr(ord('B') + i) for i in range(max_answer_cols)]
            else:
                answer_columns = []

            print(f"    {'✅ QUESTION' if is_question_sheet else '❌ CONTENT'} '{sheet_name}': {', '.join(reasons)}")
            if is_question_sheet and answer_columns:
                print(f"      📋 Answer columns: {answer_columns[:5]}{'...' if len(answer_columns) > 5 else ''}")

            # Create the analysis result
            if is_question_sheet:
                sheets_analysis[sheet_name] = SheetAnalysis(
                    sheet_type=SheetType.QUESTION_SHEET,
                    purpose=f"Question/requirement sheet with {rows} items",
                    contains_questions=True,
                    skip_extraction=False,
                    extraction_strategy=SheetExtractionStrategy(
                        question_columns=['A'],
                        answer_columns=answer_columns,
                        start_row=2
                    ),
                    confidence="high"
                )
            else:
                sheets_analysis[sheet_name] = SheetAnalysis(
                    sheet_type=SheetType.CONTENT_SHEET,
                    purpose="Content or reference sheet",
                    contains_questions=False,
                    skip_extraction=True,
                    confidence="high"
                )

        question_sheet_count = len([s for s in sheets_analysis.values()
                                  if s.sheet_type == SheetType.QUESTION_SHEET])

        print(f"  📊 FINAL RESULT: {question_sheet_count} question sheets detected out of {len(sheets_info)} total sheets")

        document_overview = DocumentOverview(
            document_type="RFI/RFP Document",
            total_question_sheets=question_sheet_count
        )

        return SheetsAnalysisResult(
            sheets_analysis=sheets_analysis,
            document_overview=document_overview
        )

    def detect_columns_with_statistics(self, worksheet_data: Dict, sheet_name: str) -> ColumnDetectionResult:
        """Smart column detection using statistics and patterns"""

        # Add column statistics to the analysis
        column_stats = self._analyze_column_patterns(worksheet_data)

        prompt = f"""Detect question and answer columns in this sheet.

    SHEET: {sheet_name}
    HEADERS: {json.dumps(worksheet_data['headers'])}
    SAMPLE DATA: {json.dumps(worksheet_data['samples'][:10])}  
    COLUMN STATISTICS: {json.dumps(column_stats)}

    DETECTION RULES:
    1. Question column: The column with the actual questions/requirements/specifications
       - Usually has longer text
       - Contains descriptive content
       - Often column A or B

    2. Answer columns: Columns meant for responses
       - Headers like "Compliance", "Response", "Status", "Answer", "Yes/No"
       - Usually empty or sparsely filled
       - Can be just one column (like column B in a 2-column sheet)

    For a simple 2-column sheet with questions and a "Compliance" column:
    - Column A would be the question column
    - Column B would be the answer column

    RESPONSE FORMAT:
    {{
        "question_column": "letter of question column",
        "answer_columns": ["letter(s) of answer column(s)"],
        "hierarchy_column": null,
        "column_purposes": {{
            "A": "purpose based on header/content",
            "B": "purpose based on header/content"
        }},
        "start_row": 2,
        "confidence": "high" | "medium" | "low"
    }}"""

        try:
            response = self.client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=8192,
                temperature=0.1,
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )

            response_text = response.content[0].text.strip()
            if not response_text:
                print("    ⚠️  Empty response from Claude API")
                return self._statistical_fallback_detection(worksheet_data, column_stats)

            # Clean response text
            if response_text.startswith("```json"):
                response_text = response_text[7:]
            if response_text.startswith("```"):
                response_text = response_text[3:]
            if response_text.endswith("```"):
                response_text = response_text[:-3]

            result_json = json.loads(response_text)
            return ColumnDetectionResult(**result_json)
        except Exception as e:
            print(f"    ⚠️  Statistical column detection failed: {str(e)[:50]}")
            return self._statistical_fallback_detection(worksheet_data, column_stats)

    def extract_global_context(self, content_data: Dict) -> GlobalContext:
        """Extract global context with better analysis"""

        prompt = f"""Analyze this content/instruction sheet to extract comprehensive global context.

CONTENT SHEET DATA:
{json.dumps(content_data, indent=2)}

EXTRACTION TASKS - Analyze actual content (not assumptions):
1. What type of document is this? (RFI, RFP, compliance checklist, etc.)
2. What are the filling instructions?
3. How do different sections/sheets relate to each other?
4. What are the answer requirements and formats?
5. Key terminology and definitions
6. Evaluation or scoring criteria

BE THOROUGH - this context will guide the entire filling process.

RESPONSE FORMAT:
{{
    "document_type": "specific analyzed type",
    "document_purpose": "analyzed purpose from content",
    "filling_instructions": {{
        "general": "extracted general instructions",
        "by_section": {{
            "section_name": "specific instructions found"
        }}
    }},
    "sheet_relationships": {{
        "sheet_name": "analyzed relationship"
    }},
    "answer_guidelines": {{
        "compliance_responses": ["values found in content"],
        "detail_requirements": "requirements found",
        "evidence_requirements": "evidence requirements found"
    }},
    "terminology": {{
        "key_terms": {{"term": "definition found"}},
        "acronyms": {{"acronym": "expansion found"}}
    }},
    "evaluation_criteria": "criteria found or null",
    "special_notes": ["important notes found"]
}}"""

        try:
            response = self.client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=8192,
                temperature=0.1,
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )

            response_text = response.content[0].text.strip()
            if not response_text:
                print("    ⚠️  Empty response from Claude API")
                return self._create_enhanced_fallback_context()

            # Clean response text
            if response_text.startswith("```json"):
                response_text = response_text[7:]
            if response_text.startswith("```"):
                response_text = response_text[3:]
            if response_text.endswith("```"):
                response_text = response_text[:-3]

            result_json = json.loads(response_text)
            return GlobalContext(**result_json)
        except Exception as e:
            print(f"    ⚠️  Global context extraction failed: {str(e)[:50]}")
            return self._create_enhanced_fallback_context()

    def generate_intelligent_fill_strategy(self, sheet_info: Dict,
                                         global_context: Optional[GlobalContext] = None) -> FillStrategy:
        """Generate intelligent filling strategy with cross-column logic"""

        context_section = ""
        if global_context:
            context_section = f"""
GLOBAL CONTEXT:
Document Type: {global_context.document_type}
Filling Instructions: {global_context.filling_instructions.general}
Answer Guidelines: {json.dumps(global_context.answer_guidelines.model_dump())}
"""

        prompt = f"""Generate intelligent answer filling strategy with cross-column logic.
{context_section}

SHEET INFORMATION:
- Sheet Name: {sheet_info['sheet_name']}
- Fillable Questions: {sheet_info.get('fillable_questions', 0)}
- Answer Columns: {sheet_info['answer_columns']}
- Column Purposes: {json.dumps(sheet_info.get('column_purposes', {}))}

REQUIREMENTS FOR INTELLIGENT STRATEGY:
1. Analyze column purposes to determine relationships
2. Create cross-column rules (e.g., if compliance=No, then details=Not applicable)
3. Realistic distribution based on document type
4. Column-specific values that make sense for each purpose
5. Consider empty probability for optional columns

RESPONSE FORMAT:
{{
    "distribution": {{
        "positive": 70,
        "negative": 15,
        "partial": 15
    }},
    "column_strategies": {{
        "C": {{
            "purpose": "analyzed purpose",
            "positive_values": ["realistic positive values"],
            "negative_values": ["realistic negative values"],
            "partial_values": ["realistic partial values"],
            "conditional_logic": "when to use each value type",
            "empty_probability": 0.1
        }}
    }},
    "cross_column_rules": [
        "If column C is 'No', then column D should be 'Not applicable'",
        "Only one compliance column should be marked per row",
        "Remarks column should align with compliance status"
    ]
}}"""

        try:
            response = self.client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=8192,
                temperature=0.1,
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )

            response_text = response.content[0].text.strip()
            if not response_text:
                print("    ⚠️  Empty response from Claude API")
                return self._create_intelligent_fallback_strategy(sheet_info)

            # Clean response text
            if response_text.startswith("```json"):
                response_text = response_text[7:]
            if response_text.startswith("```"):
                response_text = response_text[3:]
            if response_text.endswith("```"):
                response_text = response_text[:-3]

            result_json = json.loads(response_text)
            return FillStrategy(**result_json)
        except Exception as e:
            print(f"    ⚠️  Strategy generation failed: {str(e)[:50]}")
            return self._create_intelligent_fallback_strategy(sheet_info)

    def _analyze_column_patterns(self, worksheet_data: Dict) -> Dict[str, Any]:
        """Analyze statistical patterns in each column"""
        patterns = {}
        samples = worksheet_data.get('samples', [])

        if not samples:
            return patterns

        # Analyze each column
        max_cols = len(samples[0]) if samples else 0

        for col_idx in range(max_cols):
            col_letter = chr(ord('A') + col_idx)

            filled_cells = 0
            long_text_cells = 0
            short_text_cells = 0
            numeric_cells = 0
            total_text_length = 0

            # Analyze column content
            for row in samples[1:]:  # Skip header row
                if col_idx < len(row):
                    value = row[col_idx]
                    if value and str(value).strip():
                        filled_cells += 1
                        text_len = len(str(value))
                        total_text_length += text_len

                        if text_len > 100:
                            long_text_cells += 1
                        elif text_len < 20:
                            short_text_cells += 1

                        # Check if numeric
                        try:
                            float(str(value).replace(',', ''))
                            numeric_cells += 1
                        except:
                            pass

            sample_count = len(samples) - 1  # Exclude header
            if sample_count > 0:
                patterns[col_letter] = {
                    'filled_ratio': filled_cells / sample_count,
                    'long_text_ratio': long_text_cells / sample_count,
                    'short_text_ratio': short_text_cells / sample_count,
                    'numeric_ratio': numeric_cells / sample_count,
                    'avg_text_length': total_text_length / max(filled_cells, 1),
                    'likely_question_column': long_text_cells > sample_count * 0.3,
                    'likely_answer_column': short_text_cells > sample_count * 0.3 and filled_cells > sample_count * 0.1
                }

        return patterns

    def _statistical_fallback_detection(self, worksheet_data: Dict, column_stats: Dict) -> ColumnDetectionResult:
        """Fallback using statistical analysis"""
        # Find question column - highest long text ratio
        question_col = 'A'  # default to A for most RFI sheets
        max_long_text = 0

        for col, stats in column_stats.items():
            if stats.get('likely_question_column', False) and stats.get('long_text_ratio', 0) > max_long_text:
                max_long_text = stats['long_text_ratio']
                question_col = col

        # Find answer columns - likely answer columns after question column
        answer_cols = []
        question_col_idx = ord(question_col) - ord('A')

        for col, stats in column_stats.items():
            col_idx = ord(col) - ord('A')
            if col_idx > question_col_idx:
                answer_cols.append(col)

        # If no answer columns found, use pattern after question column
        if not answer_cols:
            # Generate reasonable answer columns
            num_cols = len(worksheet_data.get('headers', []))
            max_answer_cols = min(num_cols - 1, 8)
            answer_cols = [chr(ord('B') + i) for i in range(max_answer_cols)]

        # Build column purposes from headers
        headers = worksheet_data.get('headers', [])
        column_purposes = {}
        for header in headers:
            col = header.get('column', '')
            value = header.get('value', '')
            if col and value:
                column_purposes[col] = str(value)

        return ColumnDetectionResult(
            question_column=question_col,
            answer_columns=answer_cols,
            column_purposes=column_purposes,
            hierarchical_patterns=HierarchicalPattern(
                parent_indicators=["ends with :", "contains 'following'", "includes:"],
                list_patterns=["starts with a)", "starts with b)", "starts with number)"],
                requirement_patterns=["starts with letter."]
            ),
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
        """Create intelligent fallback strategy based on column analysis"""
        column_strategies = {}
        answer_columns = sheet_info.get('answer_columns', [])
        column_purposes = sheet_info.get('column_purposes', {})

        for i, col in enumerate(answer_columns):
            purpose = column_purposes.get(col, f"Column {col}")
            purpose_lower = purpose.lower()

            # Intelligent strategy based on column purpose
            if any(word in purpose_lower for word in ['compli', 'yes', 'no', 'status']):
                column_strategies[col] = ColumnFillStrategy(
                    purpose="Compliance/Status",
                    positive_values=["Yes", "Compliant", "Supported", "✓"],
                    negative_values=["No", "Not Compliant", "Not Supported", "✗"],
                    partial_values=["Partial", "Limited", "With Conditions"],
                    empty_probability=0.05
                )
            elif any(word in purpose_lower for word in ['comment', 'remark', 'note', 'detail']):
                column_strategies[col] = ColumnFillStrategy(
                    purpose="Comments/Details",
                    positive_values=["Fully supported with standard configuration", "Available out-of-the-box", "Standard feature"],
                    negative_values=["Not available in current version", "Would require custom development", "Not supported"],
                    partial_values=["Available with customization", "Requires additional configuration", "Limited support available"],
                    empty_probability=0.15
                )
            elif any(word in purpose_lower for word in ['cost', 'price', 'fee']):
                column_strategies[col] = ColumnFillStrategy(
                    purpose="Cost/Pricing",
                    positive_values=["Included in base cost", "No additional charge", "Standard pricing"],
                    negative_values=["Additional licensing required", "Custom pricing", "Not available"],
                    partial_values=["Additional cost may apply", "Depends on configuration", "Quote required"],
                    empty_probability=0.3
                )
            else:
                # Generic strategy
                column_strategies[col] = ColumnFillStrategy(
                    purpose=purpose,
                    positive_values=["Available", "Supported", "Yes"],
                    negative_values=["Not Available", "Not Supported", "No"],
                    partial_values=["Limited", "Partial", "Conditional"],
                    empty_probability=0.2
                )

        return FillStrategy(
            distribution=FillDistribution(positive=70, negative=15, partial=15),
            column_strategies=column_strategies,
            cross_column_rules=[
                "Ensure consistency across compliance columns",
                "Comments should provide context for non-positive responses",
                "Only fill relevant columns based on response type"
            ]
        )