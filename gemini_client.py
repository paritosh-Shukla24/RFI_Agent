"""
Gemini API client for structured outputs and intelligent analysis
Compatible with current Google Generative AI SDK
"""

import os
import re
import json
import google.generativeai as genai
from typing import Dict, List, Any, Optional

from models import (
    SheetsAnalysisResult, SheetAnalysis, SheetType, DocumentOverview,
    ColumnDetectionResult, HierarchicalPattern, GlobalContext,
    FillingInstructions, AnswerGuidelines, FillStrategy, FillDistribution,
    ColumnFillStrategy, SheetExtractionStrategy,
    # Structured response models
    SheetsAnalysisResponse, ColumnDetectionResponse, GlobalContextResponse,
    FillStrategyResponse
)


class GeminiStructuredClient:
    """Enhanced Gemini client with full LLM-based intelligence - no hardcoding"""

    def __init__(self):
        api_key = os.getenv('GEMINI_API_KEY')
        if not api_key:
            raise ValueError("GEMINI_API_KEY environment variable is required")

        genai.configure(api_key=api_key)
        # Use the model that works best for structured output
        self.model = genai.GenerativeModel('gemini-2.5-flash-preview-05-20')
        self.detected_content_sheet = None

        # Configure generation settings for structured output
        self.generation_config = {
            "temperature": 0.1,
            "max_output_tokens": 8192,
        }

    def analyze_sheets_intelligently(self, sheets_info: List[Dict],
                                     global_context: Optional[GlobalContext] = None) -> SheetsAnalysisResult:
        """LLM-based sheet analysis using Gemini - analyzes actual content to classify sheets"""

        print("🧠 Using Gemini LLM-based intelligent sheet analysis...")
        print("📝 Analyzing actual content patterns to classify sheets...")

        # Prepare comprehensive data for LLM analysis
        analysis_data = self._prepare_sheet_analysis_data(sheets_info)

        # Try Gemini LLM analysis first
        try:
            llm_result, content_sheet_name = self._get_gemini_sheet_analysis(analysis_data, global_context)
            if llm_result:
                print("✅ Gemini LLM sheet analysis completed successfully")
                # Store detected content sheet for later use
                self.detected_content_sheet = content_sheet_name
                return llm_result
        except Exception as e:
            print(f"⚠️  Gemini LLM sheet analysis failed: {str(e)[:100]}")

        # Fallback to intelligent pattern analysis
        print("🔄 Using intelligent pattern-based analysis...")
        result, content_sheet_name = self._intelligent_pattern_analysis(sheets_info)
        # Store detected content sheet for later use
        self.detected_content_sheet = content_sheet_name
        return result

    def _prepare_sheet_analysis_data(self, sheets_info: List[Dict]) -> Dict[str, Any]:
        """Prepare comprehensive data for Gemini LLM sheet analysis"""

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
                "sample_content": [],
                "content_analysis": {}
            }

            # Extract headers (first 8 for analysis)
            headers = sheet.get('headers', [])
            for header in headers[:8]:
                if header and header.get('value'):
                    sheet_data["headers"].append({
                        "column": header.get('column', ''),
                        "header": str(header['value'])[:50]
                    })

            # Extract sample content (first 5 data rows, first 4 columns)
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

            # Basic content analysis
            all_text = ' '.join([
                ' '.join([h.get('header', '') for h in sheet_data["headers"]]),
                ' '.join([' '.join(row) for row in sheet_data["sample_content"]])
            ]).lower()

            sheet_data["content_analysis"] = {
                "total_text_length": len(all_text),
                "appears_structured": len(sheet_data["headers"]) > 3 and len(sheet_data["sample_content"]) > 2,
                "has_empty_columns": any('header' in h and len(h['header']) < 10 for h in sheet_data["headers"][-3:]) if sheet_data["headers"] else False
            }

            analysis_data["sheets"].append(sheet_data)

        return analysis_data

    def _get_gemini_sheet_analysis(self, analysis_data: Dict[str, Any],
                               global_context: Optional[GlobalContext] = None) -> Optional[SheetsAnalysisResult]:
        """Get Gemini analysis of sheet types based on actual content"""

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
{json.dumps(analysis_data, indent=2)}

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

RESPONSE FORMAT (JSON only):
{{
    "content_sheet_detected": "SheetName" or null,
    "classification_reasoning": "explanation of analysis approach",
    "sheets_analysis": {{
        "SheetName": {{
            "sheet_type": "content_sheet" | "question_sheet",
            "purpose": "specific purpose based on content analysis",
            "contains_questions": true/false,
            "skip_extraction": true/false,
            "reasoning": "detailed reasoning for this classification",
            "confidence": "high" | "medium" | "low",
            "extraction_strategy": {{
                "question_columns": ["A"],
                "answer_columns": ["B", "C"],
                "start_row": 2
            }}
        }}
    }},
    "document_overview": {{
        "document_type": "analyzed document type",
        "total_question_sheets": number
    }}
}}"""

        try:
            response = self.model.generate_content(
                prompt,
                generation_config=self.generation_config
            )

            response_text = response.text.strip()

            # Clean and parse response
            response_text = self._clean_json_response(response_text)

            # Parse as SheetsAnalysisResponse first
            result_json = json.loads(response_text)

            # Convert to our models
            sheets_analysis = {}
            content_sheet_name = result_json.get("content_sheet_detected")

            print(f"🎯 Gemini detected content sheet: {content_sheet_name or 'None'}")
            print(f"📋 Analysis approach: {result_json.get('classification_reasoning', 'Not provided')[:80]}...")

            for sheet_name, analysis in result_json["sheets_analysis"].items():
                sheet_type = SheetType.CONTENT_SHEET if analysis["sheet_type"] == "content_sheet" else SheetType.QUESTION_SHEET

                print(f"  📄 '{sheet_name}': {analysis['sheet_type']} - {analysis['reasoning'][:60]}...")

                # Create extraction strategy for question sheets
                extraction_strategy = None
                if sheet_type == SheetType.QUESTION_SHEET:
                    strategy_data = analysis.get("extraction_strategy", {})
                    extraction_strategy = SheetExtractionStrategy(
                        question_columns=strategy_data.get("question_columns", ["A"]),
                        answer_columns=strategy_data.get("answer_columns", ["B", "C"]),
                        start_row=strategy_data.get("start_row", 2)
                    )

                sheets_analysis[sheet_name] = SheetAnalysis(
                    sheet_type=sheet_type,
                    purpose=analysis["purpose"],
                    contains_questions=analysis["contains_questions"],
                    skip_extraction=analysis["skip_extraction"],
                    extraction_strategy=extraction_strategy,
                    confidence=analysis["confidence"]
                )

            document_overview = DocumentOverview(
                document_type=result_json["document_overview"]["document_type"],
                total_question_sheets=result_json["document_overview"]["total_question_sheets"]
            )

            result = SheetsAnalysisResult(
                sheets_analysis=sheets_analysis,
                document_overview=document_overview
            )

            return result, content_sheet_name

        except Exception as e:
            print(f"🔴 Gemini sheet analysis failed: {str(e)[:100]}")
            return None, None

    def detect_columns_with_statistics(self, worksheet_data: Dict, sheet_name: str) -> ColumnDetectionResult:
        """Gemini LLM-based intelligent column detection"""

        # Prepare comprehensive analysis data
        column_stats = self._analyze_column_patterns(worksheet_data)
        headers = worksheet_data.get('headers', [])
        samples = worksheet_data.get('samples', [])

        prompt = f"""Analyze this Excel sheet to intelligently detect question and answer columns.

SHEET: {sheet_name}
HEADERS: {json.dumps(headers, indent=1)}
SAMPLE DATA (first 6 rows): {json.dumps(samples[:6], indent=1)}
COLUMN STATISTICS: {json.dumps(column_stats, indent=1)}

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

3. METADATA COLUMNS: Reference columns (NOT answer columns)
   - ID numbers, categories, classifications
   - These provide context but aren't for responses

CRITICAL ANALYSIS PRINCIPLES:
- Study ACTUAL CONTENT PATTERNS and SHEET PURPOSE
- The question column could be anywhere - don't assume position
- Answer columns should make sense for the data collection workflow
- Consider the logical flow: metadata → questions → responses

RESPONSE FORMAT (JSON only):
{{
    "question_column": "X",
    "answer_columns": ["Y", "Z"],
    "hierarchy_column": null,
    "column_purposes": {{
        "A": "specific purpose based on analysis",
        "B": "specific purpose based on analysis"
    }},
    "analysis_reasoning": "explain your column detection logic",
    "start_row": 2,
    "confidence": "high"
}}"""

        try:
            response = self.model.generate_content(
                prompt,
                generation_config=self.generation_config
            )

            response_text = response.text.strip()
            response_text = self._clean_json_response(response_text)
            result_json = json.loads(response_text)

            # Print Gemini reasoning
            if 'analysis_reasoning' in result_json:
                print(f"    🧠 Gemini Column Analysis: {result_json['analysis_reasoning']}")

            return ColumnDetectionResult(**result_json)

        except Exception as e:
            print(f"    ⚠️  Gemini column detection failed: {str(e)[:50]}")
            return self._intelligent_statistical_fallback(worksheet_data, column_stats)

    def extract_global_context(self, content_data: Dict) -> GlobalContext:
        """Gemini LLM-based global context extraction"""

        prompt = f"""Analyze this content/instruction sheet to extract comprehensive global context.

CONTENT SHEET DATA:
{json.dumps(content_data, indent=2)}

COMPREHENSIVE ANALYSIS TASKS:
Analyze the actual content to understand:

1. Document Type & Purpose: What kind of document is this? (RFI, RFP, assessment, etc.)
2. Filling Instructions: How should responses be provided?
3. Sheet Relationships: How do different sections relate?
4. Answer Requirements: What format/style of answers are expected?
5. Key Terminology: Important terms and definitions
6. Evaluation Criteria: How will responses be assessed?

Base your analysis on ACTUAL CONTENT, not assumptions.

RESPONSE FORMAT (JSON only):
{{
    "document_type": "analyzed document type",
    "document_purpose": "analyzed purpose from content",
    "filling_instructions": {{
        "general": "extracted instructions",
        "by_section": {{
            "section_name": "specific instructions"
        }}
    }},
    "sheet_relationships": {{
        "sheet_name": "analyzed relationship"
    }},
    "answer_guidelines": {{
        "compliance_responses": ["extracted response types"],
        "detail_requirements": "detail requirements found",
        "evidence_requirements": "evidence requirements found"
    }},
    "terminology": {{
        "key_terms": {{"term": "definition"}},
        "acronyms": {{"acronym": "expansion"}}
    }},
    "evaluation_criteria": "criteria found or null",
    "special_notes": ["important findings"]
}}"""

        try:
            response = self.model.generate_content(
                prompt,
                generation_config=self.generation_config
            )

            response_text = response.text.strip()
            response_text = self._clean_json_response(response_text)
            result_json = json.loads(response_text)

            return GlobalContext(**result_json)

        except Exception as e:
            print(f"    ⚠️  Global context extraction failed: {str(e)[:50]}")
            return self._create_enhanced_fallback_context()

    def generate_intelligent_fill_strategy(self, sheet_info: Dict,
                                           global_context: Optional[GlobalContext] = None) -> FillStrategy:
        """Gemini LLM-based intelligent filling strategy generation with better error handling"""

        # Validate input - make sure we have answer columns
        if not sheet_info.get('answer_columns'):
            print("    ⚠️  No answer columns available for strategy generation")
            return self._create_intelligent_fallback_strategy(sheet_info)

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
- Column Purposes: {json.dumps(sheet_info.get('column_purposes', {}))}

STRATEGY REQUIREMENTS:
1. Analyze column purposes to understand relationships
2. Create cross-column rules for logical consistency
3. Generate realistic response distributions
4. Provide column-specific response values
5. IMPORTANT: Fill ALL cells - set empty_probability to 0

RESPONSE FORMAT (JSON only):
{{
    "distribution": {{
        "positive": 70,
        "negative": 15,
        "partial": 15
    }},
    "column_strategies": {{
        "C": {{
            "purpose": "analyzed purpose",
            "positive_values": ["realistic values"],
            "negative_values": ["realistic values"],
            "partial_values": ["realistic values"],
            "conditional_logic": "logic description",
            "empty_probability": 0.0
        }}
    }},
    "cross_column_rules": [
        "logical consistency rules",
        "conditional filling rules"
    ]
}}"""

        try:
            response = self.model.generate_content(
                prompt,
                generation_config=self.generation_config
            )

            response_text = response.text.strip()
            response_text = self._clean_json_response(response_text)

            try:
                result_json = json.loads(response_text)

                # Check if the result has the required fields
                if not all(key in result_json for key in ['distribution', 'column_strategies', 'cross_column_rules']):
                    print(f"    ⚠️  Incomplete strategy JSON - missing required fields")
                    return self._create_intelligent_fallback_strategy(sheet_info)

                # Validate column strategies - make sure all answer columns have strategies
                for col in sheet_info['answer_columns']:
                    if col not in result_json['column_strategies']:
                        print(f"    ⚠️  Missing strategy for column {col} - adding default")
                        result_json['column_strategies'][col] = {
                            "purpose": f"Response field {col}",
                            "positive_values": ["Yes", "Supported", "Available"],
                            "negative_values": ["No", "Not Supported", "Not Available"],
                            "partial_values": ["Partial", "Limited Support"],
                            "empty_probability": 0.0  # Force no empty cells
                        }
                    else:
                        # Ensure empty_probability is 0 for all strategies
                        result_json['column_strategies'][col]['empty_probability'] = 0.0

                # Create the distribution object
                distribution = FillDistribution(
                    positive=result_json['distribution'].get('positive', 70),
                    negative=result_json['distribution'].get('negative', 15),
                    partial=result_json['distribution'].get('partial', 15)
                )

                # Create column strategies
                column_strategies = {}
                for col, strat_data in result_json['column_strategies'].items():
                    column_strategies[col] = ColumnFillStrategy(
                        purpose=strat_data.get('purpose', f"Column {col}"),
                        positive_values=strat_data.get('positive_values', ["Yes"]),
                        negative_values=strat_data.get('negative_values', ["No"]),
                        partial_values=strat_data.get('partial_values', ["Partial"]),
                        conditional_logic=strat_data.get('conditional_logic'),
                        empty_probability=0.0  # Always set to 0 to fill all cells
                    )

                # Create the final strategy
                return FillStrategy(
                    distribution=distribution,
                    column_strategies=column_strategies,
                    cross_column_rules=result_json.get('cross_column_rules', [])
                )

            except Exception as parse_error:
                print(f"    ⚠️  Error parsing strategy JSON: {str(parse_error)[:100]}")
                return self._create_intelligent_fallback_strategy(sheet_info)

        except Exception as e:
            print(f"    ⚠️  Strategy generation failed: {str(e)[:100]}")
            return self._create_intelligent_fallback_strategy(sheet_info)

    def _clean_json_response(self, response_text: str) -> str:
        """Clean Gemini LLM response to extract valid JSON with improved error handling"""

        if not response_text.strip():
            raise ValueError("Empty response")

        # Remove markdown formatting
        if response_text.startswith("```json"):
            response_text = response_text[7:]
        elif response_text.startswith("```"):
            response_text = response_text[3:]

        if response_text.endswith("```"):
            response_text = response_text[:-3]

        # Find JSON boundaries
        lines = response_text.split('\n')
        json_start = -1

        for i, line in enumerate(lines):
            if line.strip().startswith('{'):
                json_start = i
                break

        if json_start > 0:
            response_text = '\n'.join(lines[json_start:])

        # Enhanced JSON cleaning
        try:
            # Basic JSON validation
            json.loads(response_text)
            return response_text.strip()
        except json.JSONDecodeError as e:
            # Apply fixes for common JSON errors
            print(f"    🔧 Attempting to fix JSON: {str(e)[:50]}")

            # Fix missing commas between properties
            response_text = re.sub(r'"\s*\n\s*"', '",\n"', response_text)

            # Fix trailing commas before closing brackets
            response_text = re.sub(r',(\s*[\]}])', r'\1', response_text)

            # Fix missing quotes around property names
            response_text = re.sub(r'([{,]\s*)([a-zA-Z_][a-zA-Z0-9_]*)(\s*:)', r'\1"\2"\3', response_text)

            # Try to validate again
            try:
                json.loads(response_text)
                print("    ✅ JSON fixed successfully")
                return response_text.strip()
            except json.JSONDecodeError:
                # Last resort: extract as much valid JSON as possible
                print("    ⚠️ Using partial JSON extraction")
                return self._extract_valid_json(response_text)

    def _extract_valid_json(self, text: str) -> str:
        """Extract valid JSON from potentially broken JSON string"""
        # Find the outermost brackets
        open_brace = text.find('{')
        if open_brace == -1:
            return "{}"  # Fallback to empty object

        # Track nested braces to find the matching closing brace
        brace_count = 0
        in_string = False
        escape_next = False

        for i in range(open_brace, len(text)):
            char = text[i]

            # Handle string literals and escaping
            if char == '\\' and not escape_next:
                escape_next = True
                continue

            if char == '"' and not escape_next:
                in_string = not in_string

            escape_next = False

            # Only count braces outside of strings
            if not in_string:
                if char == '{':
                    brace_count += 1
                elif char == '}':
                    brace_count -= 1
                    if brace_count == 0:
                        # We found a complete JSON object
                        result = text[open_brace:i+1]

                        # Final validation and correction
                        try:
                            return self._fix_and_validate_json(result)
                        except:
                            # If it still fails, return best-effort result
                            return result

        # If we can't find a complete JSON object, return a fallback
        return "{}"

    def _fix_and_validate_json(self, json_str: str) -> str:
        """Final JSON fixing and validation"""
        # Replace any \' with ' for proper string handling
        json_str = json_str.replace("\\'", "'")

        # Try to parse it
        try:
            json.loads(json_str)
            return json_str
        except:
            # Apply aggressive fixes
            # Remove any property that might be causing issues
            lines = json_str.split('\n')
            fixed_lines = []

            for i, line in enumerate(lines):
                if i == 0 or i == len(lines) - 1:
                    # Keep the opening and closing braces
                    fixed_lines.append(line)
                    continue

                # Skip any line that might be problematic
                if '"' in line and ':' in line and not line.strip().endswith(',') and i < len(lines) - 2:
                    fixed_lines.append(line + ',')
                else:
                    fixed_lines.append(line)

            fixed_json = '\n'.join(fixed_lines)

            # One final try
            try:
                json.loads(fixed_json)
                return fixed_json
            except:
                # Last resort - create minimal valid JSON
                return "{}"

    def _intelligent_pattern_analysis(self, sheets_info: List[Dict]) -> tuple:
        """Intelligent pattern-based analysis when Gemini LLM fails"""

        print("📊 Analyzing sheets with intelligent pattern recognition...")

        sheets_analysis = {}
        content_sheet_candidates = []

        for sheet in sheets_info:
            sheet_name = sheet['name']
            rows = sheet.get('rows', 0)
            columns = sheet.get('columns', 0)

            # Analyze actual content patterns
            sample_data = sheet.get('sample_data', [])
            headers = sheet.get('headers', [])

            # Calculate content characteristics
            characteristics = self._analyze_sheet_characteristics(sample_data, headers)

            print(f"  🔍 Analyzing '{sheet_name}': {rows}r x {columns}c")
            print(f"      📊 Scores - Content: {characteristics['content_score']}, Questions: {characteristics['question_score']}, Structure: {characteristics['structure_score']}")

            # IMPROVED LOGIC: If it has many rows and question-like content, it's likely a question sheet
            # Sheets with more than 25 rows are usually question sheets in RFIs/RFPs
            is_content_sheet = False
            if rows < 25 and characteristics['content_score'] > characteristics['question_score']:
                is_content_sheet = True
            elif characteristics['question_score'] == 0 and characteristics['content_score'] > 2:
                is_content_sheet = True

            # Override for sheets that look like question sheets
            if rows > 30 or characteristics['question_score'] > 2:
                is_content_sheet = False

            reasons = []
            if is_content_sheet:
                content_sheet_candidates.append((sheet_name, characteristics['content_score']))
                reasons.append(f"content-focused (score: {characteristics['content_score']})")
            else:
                if characteristics['structure_score'] > 3:
                    reasons.append("structured format")
                if rows > 30:
                    reasons.append("substantial size")
                if characteristics['question_score'] > 0:
                    reasons.append("question-like content")
                if not reasons:
                    reasons.append("default classification")

            # Create analysis result
            if is_content_sheet:
                sheets_analysis[sheet_name] = SheetAnalysis(
                    sheet_type=SheetType.CONTENT_SHEET,
                    purpose="Content/instruction sheet based on pattern analysis",
                    contains_questions=False,
                    skip_extraction=True,
                    confidence="medium"
                )
                print(f"      ❌ CONTENT sheet: {', '.join(reasons)}")
            else:
                # Generate intelligent answer columns
                answer_columns = self._generate_intelligent_answer_columns(headers, columns)

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
                print(f"      ✅ QUESTION sheet: {', '.join(reasons)}")

        # Select best content sheet candidate
        content_sheet_name = None
        if content_sheet_candidates:
            # Pick the one with highest content score
            content_sheet_candidates.sort(key=lambda x: x[1], reverse=True)
            content_sheet_name = content_sheet_candidates[0][0]
            print(f"🎯 Selected content sheet: {content_sheet_name}")

        question_sheet_count = len([s for s in sheets_analysis.values()
                                  if s.sheet_type == SheetType.QUESTION_SHEET])

        print(f"  📊 FINAL RESULT: {question_sheet_count} question sheets detected")

        return SheetsAnalysisResult(
            sheets_analysis=sheets_analysis,
            document_overview=DocumentOverview(
                document_type="Business Document",
                total_question_sheets=question_sheet_count
            )
        ), content_sheet_name

    def _intelligent_statistical_fallback(self, worksheet_data: Dict, column_stats: Dict) -> ColumnDetectionResult:
        """Intelligent statistical fallback based on patterns, not hardcoding"""

        print("    🔧 Using intelligent statistical analysis...")

        headers = worksheet_data.get('headers', [])
        samples = worksheet_data.get('samples', [])

        # Simple case: 2-column document (very common in RFIs/RFPs)
        if len(headers) == 2:
            print("    📋 2-column document detected - using standard Q&A layout")
            question_col = 'A'
            answer_cols = ['B']
            column_purposes = {
                'A': "Questions/Requirements",
                'B': "Response"
            }
            return ColumnDetectionResult(
                question_column=question_col,
                answer_columns=answer_cols,
                column_purposes=column_purposes,
                start_row=2,
                confidence="high"
            )

        if not headers:
            return self._basic_fallback_detection()

        # Analyze each column intelligently
        column_analysis = {}

        for i, header in enumerate(headers):
            col_letter = header.get('column', chr(ord('A') + i))
            header_value = str(header.get('value', '')).strip()
            stats = column_stats.get(col_letter, {})

            analysis = {
                'header': header_value,
                'avg_text_length': stats.get('avg_text_length', 0),
                'filled_ratio': stats.get('filled_ratio', 0),
                'long_text_ratio': stats.get('long_text_ratio', 0),
                'question_score': 0,
                'answer_score': 0,
                'metadata_score': 0
            }

            # Question column scoring
            if analysis['avg_text_length'] > 50:  # Substantial content
                analysis['question_score'] += 3
            if analysis['long_text_ratio'] > 0.3:  # High ratio of long text
                analysis['question_score'] += 2
            if len(header_value.split()) >= 2:  # Multi-word headers
                analysis['question_score'] += 1

            # Answer column scoring
            if len(header_value) < 20 and header_value:  # Short, meaningful headers
                analysis['answer_score'] += 2
            if analysis['filled_ratio'] < 0.3:  # Mostly empty (for responses)
                analysis['answer_score'] += 3
            if 5 < analysis['avg_text_length'] < 30:  # Response-sized content
                analysis['answer_score'] += 2

            # Answer column keyword scoring - NEW
            answer_keywords = ['compliance', 'response', 'answer', 'status', 'comment',
                              'remarks', 'details', 'vendor', 'note']
            if any(keyword in header_value.lower() for keyword in answer_keywords):
                analysis['answer_score'] += 5

            # Metadata column scoring
            if len(header_value) <= 5 and ('#' in header_value or header_value.lower() in ['id', 'no']):
                analysis['metadata_score'] += 3
            if analysis['filled_ratio'] > 0.8 and analysis['avg_text_length'] < 15:
                analysis['metadata_score'] += 2

            column_analysis[col_letter] = analysis

            max_score = max(analysis['question_score'], analysis['answer_score'], analysis['metadata_score'])
            print(f"      📊 {col_letter}: '{header_value}' - Q:{analysis['question_score']}, A:{analysis['answer_score']}, M:{analysis['metadata_score']}")

        # Find best question column
        question_candidates = [(col, data['question_score']) for col, data in column_analysis.items()
                             if data['question_score'] > 0]

        if question_candidates:
            question_candidates.sort(key=lambda x: x[1], reverse=True)
            question_col = question_candidates[0][0]
        else:
            # Fallback to longest content
            text_lengths = [(col, data['avg_text_length']) for col, data in column_analysis.items()]
            text_lengths.sort(key=lambda x: x[1], reverse=True)
            question_col = text_lengths[0][0] if text_lengths else 'A'

        # Find answer columns after question column
        question_col_idx = ord(question_col) - ord('A')
        answer_cols = []

        for col, data in column_analysis.items():
            col_idx = ord(col) - ord('A')
            if (col_idx > question_col_idx and
                data['answer_score'] > data['metadata_score'] and
                data['answer_score'] > 0):
                answer_cols.append(col)

        # Ensure at least one answer column
        if not answer_cols and len(headers) > 1:
            # Default to the second column if question column is first
            if question_col == 'A' and len(headers) > 1:
                answer_cols = ['B']
            # Otherwise find any available column
            else:
                for col, data in column_analysis.items():
                    col_idx = ord(col) - ord('A')
                    if col_idx != question_col_idx and col_idx < len(headers):
                        answer_cols.append(col)
                        break

        # Build purposes
        column_purposes = {}
        for col, data in column_analysis.items():
            max_score = max(data['question_score'], data['answer_score'], data['metadata_score'])
            if data['question_score'] == max_score and max_score > 0:
                column_purposes[col] = f"Questions/Requirements - {data['header']}"
            elif data['answer_score'] == max_score and max_score > 0:
                column_purposes[col] = f"Response Field - {data['header']}"
            elif data['metadata_score'] == max_score and max_score > 0:
                column_purposes[col] = f"Reference Data - {data['header']}"
            else:
                column_purposes[col] = data['header']

        print(f"    🎯 INTELLIGENT DETECTION RESULT:")
        print(f"      📊 Question column: {question_col}")
        print(f"      📋 Answer columns ({len(answer_cols)}): {answer_cols}")

        return ColumnDetectionResult(
            question_column=question_col,
            answer_columns=answer_cols,
            column_purposes=column_purposes,
            start_row=2,
            confidence="medium"
        )

    def _analyze_sheet_characteristics(self, sample_data: List[List[str]], headers: List[Dict]) -> Dict[str, int]:
        """Analyze sheet characteristics without hardcoded rules"""

        if not sample_data or len(sample_data) < 2:
            return {"content_score": 0, "question_score": 0, "structure_score": 0}

        all_text = []
        for row in sample_data[:8]:  # First 8 rows
            for cell in row[:6]:  # First 6 columns
                if cell and str(cell).strip():
                    all_text.append(str(cell).lower())

        combined_text = ' '.join(all_text)

        # Content indicators (instructional/explanatory language)
        content_score = 0
        content_patterns = [
            ('explain', 2), ('instruction', 3), ('guideline', 2), ('overview', 2),
            ('please', 1), ('complete', 1), ('fill', 1), ('ensure', 1),
            ('note:', 2), ('important', 1), ('background', 2), ('purpose', 2)
        ]

        for pattern, weight in content_patterns:
            content_score += combined_text.count(pattern) * weight

        # Question indicators (requirement/response language)
        question_score = 0
        question_patterns = [
            ('requirement', 3), ('must', 2), ('shall', 2), ('provide', 2),
            ('describe', 2), ('list', 1), ('specify', 2), ('detail', 1),
            ('compliance', 2), ('supported', 1), ('available', 1)
        ]

        for pattern, weight in question_patterns:
            question_score += combined_text.count(pattern) * weight

        # Structure indicators (organized Q&A format)
        structure_score = 0
        if len(headers) > 5:  # Multiple columns suggest structure
            structure_score += 2
        if len(sample_data) > 20:  # Many rows suggest systematic data
            structure_score += 2
        if any(len(str(cell)) > 50 for row in sample_data[:5] for cell in row[:3]):  # Long content
            structure_score += 1

        return {
            "content_score": content_score,
            "question_score": question_score,
            "structure_score": structure_score
        }

    def _generate_intelligent_answer_columns(self, headers: List[Dict], total_columns: int) -> List[str]:
        """Generate intelligent answer columns based on header analysis"""

        answer_columns = []

        # IMPROVED: If only 1-2 columns, the second one is almost always the answer column
        if total_columns <= 2 and total_columns > 1:
            # With just 2 columns, the second is almost always for answers
            return ['B']

        if not headers or len(headers) <= 1:
            # Basic fallback for limited headers
            return [chr(ord('A') + i + 1) for i in range(min(total_columns - 1, 3))]

        # Analyze headers for potential answer columns
        answer_keyword_indicators = ['compliance', 'response', 'answer', 'status',
                                   'comment', 'remarks', 'comments', 'meet', 'support',
                                   'details', 'vendor', 'yes', 'no', 'notes', 'evidence']

        # First pass - look for explicit answer column indicators in headers
        for i, header in enumerate(headers[1:], 1):  # Skip first column (usually questions)
            if header and header.get('value'):
                header_value = str(header['value']).lower()
                col_letter = header.get('column', chr(ord('A') + i))

                # Check for response-type headers using improved indicators
                if any(keyword in header_value for keyword in answer_keyword_indicators):
                    answer_columns.append(col_letter)
                # Otherwise if it's a short header and not obviously a question column
                elif (len(header_value) < 25 and
                     not any(q_term in header_value for q_term in ['question', 'requirement', 'function'])):
                    answer_columns.append(col_letter)

        # If no answer columns found yet, use at least one column after the question column
        if not answer_columns and total_columns > 1:
            answer_columns = [chr(ord('A') + 1)]  # Use the second column

        # Limit to reasonable number
        if len(answer_columns) > 5:
            answer_columns = answer_columns[:5]

        # Ensure we have at least one answer column
        if not answer_columns and total_columns > 1:
            answer_columns = ['B']  # Default to second column

        return answer_columns

    def _analyze_column_patterns(self, worksheet_data: Dict) -> Dict[str, Any]:
        """Analyze statistical patterns in each column"""
        patterns = {}
        samples = worksheet_data.get('samples', [])

        if not samples:
            return patterns

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

                        try:
                            float(str(value).replace(',', ''))
                            numeric_cells += 1
                        except:
                            pass

            sample_count = len(samples) - 1
            if sample_count > 0:
                patterns[col_letter] = {
                    'filled_ratio': filled_cells / sample_count,
                    'long_text_ratio': long_text_cells / sample_count,
                    'short_text_ratio': short_text_cells / sample_count,
                    'numeric_ratio': numeric_cells / sample_count,
                    'avg_text_length': total_text_length / max(filled_cells, 1)
                }

        return patterns

    def _basic_fallback_detection(self) -> ColumnDetectionResult:
        """Basic fallback when no analysis possible"""
        return ColumnDetectionResult(
            question_column='A',
            answer_columns=['B', 'C'],
            column_purposes={'A': 'Questions', 'B': 'Response', 'C': 'Comments'},
            start_row=2,
            confidence="low"
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
        """Create intelligent fallback strategy based on analysis with no empty cells"""
        column_strategies = {}
        answer_columns = sheet_info.get('answer_columns', [])
        column_purposes = sheet_info.get('column_purposes', {})

        # Handle zero answer columns (error case)
        if not answer_columns and sheet_info.get('fillable_questions', 0) > 0:
            # Force default answer column
            if len(column_purposes) > 1:
                # Use second column as answer column
                second_col = chr(ord('A') + 1)
                answer_columns = [second_col]
                print(f"    ⚠️ No answer columns detected but questions found. Using column {second_col} as answer column.")
            else:
                # Create a default answer column
                answer_columns = ['B']
                print(f"    ⚠️ No answer columns detected but questions found. Using default column B.")

        for col in answer_columns:
            purpose = column_purposes.get(col, f"Column {col}").lower()

            # Set appropriate strategy based on column purpose/name
            if any(word in purpose for word in ['compliance', 'status', 'support', 'meet']):
                column_strategies[col] = ColumnFillStrategy(
                    purpose="Compliance/Status",
                    positive_values=["Yes", "Compliant", "Supported", "Available"],
                    negative_values=["No", "Not Compliant", "Not Supported", "Unavailable"],
                    partial_values=["Partial", "Limited", "With Conditions"],
                    empty_probability=0.0  # ENSURE NO EMPTY CELLS
                )
            elif any(word in purpose for word in ['comment', 'detail', 'note', 'remark', 'description']):
                column_strategies[col] = ColumnFillStrategy(
                    purpose="Comments/Details",
                    positive_values=[
                        "Fully supported as required",
                        "Available as standard functionality",
                        "Provided out-of-the-box with our solution",
                        "Implemented in current version",
                        "Standard feature in our product"
                    ],
                    negative_values=[
                        "Not available in current version",
                        "Not supported in our solution",
                        "Requires custom development",
                        "Not part of standard offering",
                        "Not implemented at this time"
                    ],
                    partial_values=[
                        "Available with configuration changes",
                        "Partially supported with limitations",
                        "Supported with third-party integration",
                        "Limited functionality available",
                        "Planned for future release"
                    ],
                    empty_probability=0.0  # ENSURE NO EMPTY CELLS
                )
            else:
                column_strategies[col] = ColumnFillStrategy(
                    purpose=f"Response field {col}",
                    positive_values=["Yes", "Available", "Supported", "Compliant"],
                    negative_values=["No", "Not Available", "Not Supported", "Non-Compliant"],
                    partial_values=["Partial", "Limited", "With Conditions", "In Development"],
                    empty_probability=0.0  # ENSURE NO EMPTY CELLS
                )

        # Create default cross-column rules
        cross_rules = [
            "Ensure consistency between compliance status and explanation fields",
            "Provide detailed justification for non-compliant or partially compliant items",
            "For positive responses, include specific capabilities that exceed requirements"
        ]

        # Default distribution favoring positive responses
        distribution = FillDistribution(positive=70, negative=15, partial=15)

        print("    💡 Created fallback fill strategy with NO empty cells")

        return FillStrategy(
            distribution=distribution,
            column_strategies=column_strategies,
            cross_column_rules=cross_rules
        )