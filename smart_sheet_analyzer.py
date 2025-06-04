"""
Smart Sheet Analyzer - LLM-based content vs question sheet detection
"""

import json
from typing import Dict, List, Any, Optional
from models import SheetsAnalysisResult, SheetAnalysis, SheetType, DocumentOverview, SheetExtractionStrategy


class SmartSheetAnalyzer:
    """LLM-based intelligent sheet analyzer"""
    
    def __init__(self, claude_client):
        self.claude_client = claude_client
    
    def analyze_all_sheets_with_llm(self, sheets_info: List[Dict]) -> SheetsAnalysisResult:
        """Analyze all sheets using LLM to determine content vs question sheets"""
        
        print("🧠 Using LLM-based intelligent sheet analysis...")
        print("📝 Analyzing actual content patterns, not just sheet names...")
        
        # Prepare comprehensive data for LLM analysis
        analysis_data = self._prepare_analysis_data(sheets_info)
        
        # Get LLM analysis
        try:
            llm_result = self._get_llm_sheet_analysis(analysis_data)
            
            if llm_result:
                print("✅ LLM analysis completed successfully")
                return llm_result
            else:
                print("⚠️  LLM analysis failed, using smart fallback")
                return self._smart_fallback_analysis(sheets_info)
                
        except Exception as e:
            print(f"⚠️  LLM analysis error: {str(e)[:100]}")
            print("🔄 Using smart fallback analysis")
            return self._smart_fallback_analysis(sheets_info)
    
    def _prepare_analysis_data(self, sheets_info: List[Dict]) -> Dict[str, Any]:
        """Prepare comprehensive data for LLM analysis"""
        
        analysis_data = {
            "total_sheets": len(sheets_info),
            "sheets": []
        }
        
        for sheet in sheets_info:
            sheet_data = {
                "name": sheet['name'],
                "rows": sheet.get('rows', 0),
                "columns": sheet.get('columns', 0),
                "headers": [],
                "sample_content": [],
                "content_indicators": []
            }
            
            # Extract headers
            headers = sheet.get('headers', [])
            for header in headers[:10]:  # First 10 headers
                if header and header.get('value'):
                    sheet_data["headers"].append(str(header['value'])[:50])
            
            # Extract sample content from first few rows
            sample_data = sheet.get('sample_data', [])
            if sample_data and len(sample_data) > 1:
                for row_idx in range(1, min(6, len(sample_data))):  # Skip header, get 5 rows
                    if row_idx < len(sample_data):
                        row = sample_data[row_idx]
                        row_content = []
                        for col_idx in range(min(3, len(row))):  # First 3 columns
                            if col_idx < len(row) and row[col_idx]:
                                content = str(row[col_idx])[:100]  # First 100 chars
                                if content.strip():
                                    row_content.append(content)
                        if row_content:
                            sheet_data["sample_content"].append(row_content)
            
            # Analyze content indicators
            all_text = []
            if sample_data:
                for row in sample_data[:10]:
                    for cell in row[:5]:
                        if cell and str(cell).strip():
                            all_text.append(str(cell).lower())
            
            # Content sheet indicators
            content_words = ['instruction', 'guideline', 'overview', 'introduction', 
                           'please', 'fill', 'complete', 'respond', 'note', 'important']
            question_words = ['requirement', 'compliance', 'must', 'shall', 'provide', 'describe']
            
            content_score = sum(1 for text in all_text for word in content_words if word in text)
            question_score = sum(1 for text in all_text for word in question_words if word in text)
            
            sheet_data["content_indicators"] = {
                "content_score": content_score,
                "question_score": question_score,
                "has_questions": question_score > 0,
                "appears_instructional": content_score > question_score
            }
            
            analysis_data["sheets"].append(sheet_data)
        
        return analysis_data
    
    def _get_llm_sheet_analysis(self, analysis_data: Dict[str, Any]) -> Optional[SheetsAnalysisResult]:
        """Get LLM analysis of sheet types"""
        
        prompt = f"""Analyze these Excel sheets to determine which are CONTENT/INSTRUCTION sheets vs QUESTION/REQUIREMENT sheets.

DOCUMENT ANALYSIS:
{json.dumps(analysis_data, indent=2)}

ANALYSIS TASK:
Look at the actual content, headers, and patterns in each sheet to determine:

1. CONTENT SHEETS contain:
   - Instructions on how to fill the document
   - Guidelines and explanations
   - Overview information
   - Reference material
   - Generally smaller with explanatory text

2. QUESTION SHEETS contain:
   - Actual questions or requirements to be answered
   - Forms with response columns
   - Technical specifications
   - Compliance checklists
   - Generally larger with structured Q&A format

CRITICAL: Base your decision on ACTUAL CONTENT ANALYSIS, not just sheet names!

RESPONSE FORMAT (JSON only, no other text):
{{
    "content_sheet_detected": "SheetName" or null,
    "sheets_analysis": {{
        "SheetName": {{
            "sheet_type": "content_sheet" | "question_sheet",
            "purpose": "specific purpose based on actual content analysis",
            "contains_questions": true/false,
            "skip_extraction": true/false,
            "reasoning": "why this classification was chosen",
            "confidence": "high" | "medium" | "low",
            "extraction_strategy": {{
                "question_columns": ["A"],
                "answer_columns": ["B", "C", "D"],
                "start_row": 2
            }}
        }}
    }},
    "document_overview": {{
        "document_type": "detected type based on content",
        "total_question_sheets": number
    }}
}}"""

        try:
            response = self.claude_client.client.messages.create(
                model="claude-3-5-sonnet-20241022",
                max_tokens=8192,
                temperature=0.1,
                messages=[{"role": "user", "content": prompt}]
            )
            
            response_text = response.content[0].text.strip()
            
            # Clean response
            if response_text.startswith("```json"):
                response_text = response_text[7:]
            if response_text.startswith("```"):
                response_text = response_text[3:]
            if response_text.endswith("```"):
                response_text = response_text[:-3]
            
            # Parse JSON
            result = json.loads(response_text.strip())
            
            # Convert to our models
            sheets_analysis = {}
            content_sheet_name = result.get("content_sheet_detected")
            
            print(f"🎯 LLM detected content sheet: {content_sheet_name or 'None'}")
            
            for sheet_name, analysis in result["sheets_analysis"].items():
                sheet_type = SheetType.CONTENT_SHEET if analysis["sheet_type"] == "content_sheet" else SheetType.QUESTION_SHEET
                
                print(f"  📋 '{sheet_name}': {analysis['sheet_type']} - {analysis['reasoning'][:60]}...")
                
                # Create extraction strategy for question sheets
                extraction_strategy = None
                if sheet_type == SheetType.QUESTION_SHEET:
                    strategy_data = analysis.get("extraction_strategy", {})
                    extraction_strategy = SheetExtractionStrategy(
                        question_columns=strategy_data.get("question_columns", ["A"]),
                        answer_columns=strategy_data.get("answer_columns", ["B", "C", "D"]),
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
                document_type=result["document_overview"]["document_type"],
                total_question_sheets=result["document_overview"]["total_question_sheets"]
            )
            
            return SheetsAnalysisResult(
                sheets_analysis=sheets_analysis,
                document_overview=document_overview
            ), content_sheet_name
            
        except Exception as e:
            print(f"🔴 LLM analysis parsing failed: {str(e)[:100]}")
            return None, None
    
    def _smart_fallback_analysis(self, sheets_info: List[Dict]) -> SheetsAnalysisResult:
        """Smart fallback when LLM fails - still better than keyword matching"""
        
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
        
        return SheetsAnalysisResult(
            sheets_analysis=sheets_analysis,
            document_overview=DocumentOverview(
                document_type="RFI/RFP Document",
                total_question_sheets=question_sheet_count
            )
        ), content_sheet_name
    
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