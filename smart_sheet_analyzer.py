"""
Smart Sheet Analyzer - Fixed for Gemini client compatibility
"""

import json
from typing import Dict, List, Any, Optional, Tuple
from models import SheetsAnalysisResult, SheetAnalysis, SheetType, DocumentOverview, SheetExtractionStrategy


class SmartSheetAnalyzer:
    """Smart sheet analyzer compatible with Gemini client"""

    def __init__(self, gemini_client):
        self.gemini_client = gemini_client

    def analyze_all_sheets_with_llm(self, sheets_info: List[Dict]) -> Tuple[SheetsAnalysisResult, Optional[str]]:
        """Analyze all sheets using Gemini to determine content vs question sheets"""

        print("🧠 Using Gemini-based intelligent sheet analysis...")
        print("📝 Analyzing actual content patterns, not just sheet names...")

        try:
            # Use Gemini client for analysis
            llm_result = self.gemini_client.analyze_sheets_intelligently(sheets_info)

            if llm_result:
                print("✅ Gemini analysis completed successfully")

                # Find content sheet from analysis
                content_sheet_name = self._find_content_sheet_from_analysis(llm_result)

                return llm_result, content_sheet_name
            else:
                print("⚠️  Gemini analysis failed, using smart fallback")
                return self._smart_fallback_analysis(sheets_info)

        except Exception as e:
            print(f"⚠️  Gemini analysis error: {str(e)[:100]}")
            print("🔄 Using smart fallback analysis")
            return self._smart_fallback_analysis(sheets_info)

    def _find_content_sheet_from_analysis(self, analysis_result: SheetsAnalysisResult) -> Optional[str]:
        """Find the content sheet from the analysis results"""

        content_sheet_candidates = []

        for sheet_name, analysis in analysis_result.sheets_analysis.items():
            if analysis.sheet_type == SheetType.CONTENT_SHEET:
                # Score based on purpose and confidence
                score = 1
                if 'instruction' in analysis.purpose.lower():
                    score += 2
                if 'content' in analysis.purpose.lower():
                    score += 2
                if analysis.confidence == 'high':
                    score += 1

                content_sheet_candidates.append((sheet_name, score))

        if content_sheet_candidates:
            # Return the highest scoring content sheet
            content_sheet_candidates.sort(key=lambda x: x[1], reverse=True)
            content_sheet_name = content_sheet_candidates[0][0]
            print(f"🎯 Selected content sheet: {content_sheet_name}")
            return content_sheet_name

        return None

    def _smart_fallback_analysis(self, sheets_info: List[Dict]) -> Tuple[SheetsAnalysisResult, Optional[str]]:
        """Smart fallback when Gemini fails - still better than keyword matching"""

        print("🔧 Using smart content-based fallback analysis...")

        sheets_analysis = {}
        content_sheet_candidates = []

        for sheet in sheets_info:
            sheet_name = sheet['name']
            rows = sheet.get('rows', 0)
            columns = sheet.get('columns', 0)

            # Analyze actual content, not just names
            sample_data = sheet.get('sample_data', [])
            content_indicators = self._analyze_content_patterns(sample_data)

            print(f"  🔍 Analyzing '{sheet_name}': {rows}r x {columns}c")
            print(f"      📊 Content score: {content_indicators['content_score']}, Question score: {content_indicators['question_score']}")

            # Determine if it's a content sheet based on actual content
            is_content_sheet = (
                content_indicators['content_score'] > content_indicators['question_score'] and
                content_indicators['content_score'] > 3 and
                rows < 50  # Content sheets are usually smaller
            )

            if is_content_sheet:
                content_sheet_candidates.append((sheet_name, content_indicators['content_score']))

            # Classification
            if is_content_sheet:
                sheets_analysis[sheet_name] = SheetAnalysis(
                    sheet_type=SheetType.CONTENT_SHEET,
                    purpose="Content/instruction sheet based on content analysis",
                    contains_questions=False,
                    skip_extraction=True,
                    confidence="medium"
                )
                print(f"      ❌ CONTENT sheet (content analysis)")
            else:
                # Generate reasonable answer columns
                max_answer_cols = min(columns - 1, 8)
                answer_columns = [chr(ord('B') + i) for i in range(max_answer_cols)]

                sheets_analysis[sheet_name] = SheetAnalysis(
                    sheet_type=SheetType.QUESTION_SHEET,
                    purpose=f"Question sheet with {rows} requirements",
                    contains_questions=True,
                    skip_extraction=False,
                    extraction_strategy=SheetExtractionStrategy(
                        question_columns=['A'],
                        answer_columns=answer_columns,
                        start_row=2
                    ),
                    confidence="medium"
                )
                print(f"      ✅ QUESTION sheet (size: {rows}x{columns})")

        # Select best content sheet candidate
        content_sheet_name = None
        if content_sheet_candidates:
            # Pick the one with highest content score
            content_sheet_candidates.sort(key=lambda x: x[1], reverse=True)
            content_sheet_name = content_sheet_candidates[0][0]
            print(f"🎯 Selected content sheet: {content_sheet_name}")

        question_sheet_count = len([s for s in sheets_analysis.values()
                                  if s.sheet_type == SheetType.QUESTION_SHEET])

        print(f"📊 FINAL RESULT: {question_sheet_count} question sheets, content sheet: {content_sheet_name}")

        analysis_result = SheetsAnalysisResult(
            sheets_analysis=sheets_analysis,
            document_overview=DocumentOverview(
                document_type="RFI/RFP Document",
                total_question_sheets=question_sheet_count
            )
        )

        return analysis_result, content_sheet_name
    
    def _analyze_content_patterns(self, sample_data: List[List[str]]) -> Dict[str, int]:
        """Analyze content patterns to determine sheet type"""
        
        if not sample_data or len(sample_data) < 2:
            return {"content_score": 0, "question_score": 0}
        
        all_text = []
        for row in sample_data[:10]:  # First 10 rows
            for cell in row[:5]:  # First 5 columns
                if cell and str(cell).strip():
                    all_text.append(str(cell).lower())
        
        # Content indicators (instructional language)
        content_indicators = [
            'instruction', 'guideline', 'overview', 'introduction', 'please fill',
            'complete the', 'provide information', 'note:', 'important:',
            'how to', 'please ensure', 'this document', 'the purpose',
            'background', 'context', 'explanation'
        ]
        
        # Question indicators (requirement language)
        question_indicators = [
            'requirement', 'compliance', 'must', 'shall', 'provide',
            'describe your', 'list all', 'specify', 'detail',
            'yes/no', 'supported', 'available', 'capability'
        ]
        
        content_score = 0
        question_score = 0
        
        for text in all_text:
            for indicator in content_indicators:
                if indicator in text:
                    content_score += 1
            
            for indicator in question_indicators:
                if indicator in text:
                    question_score += 1
        
        return {
            "content_score": content_score,
            "question_score": question_score
        }