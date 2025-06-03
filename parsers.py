"""
Question parsing utilities for multi-line questions and hierarchical structure
"""

import os
import re
import json
import time
import anthropic
from typing import Dict, List, Any, Optional

from models import HierarchicalPattern, QuestionType


class MultiLineQuestionParser:
    """Parse questions that contain multiple sub-questions within a single cell"""

    @staticmethod
    def parse_cell_content(text: str) -> List[Dict[str, Any]]:
        """Parse a cell that might contain multiple questions"""
        if not text:
            return []

        # Check if this cell contains multiple questions
        lines = text.strip().split('\n')

        # Patterns for sub-questions within cells
        sub_patterns = [
            r'^[a-z]\)\s+',  # a) pattern
            r'^[a-z]\.\s+',  # a. pattern
            r'^\d+\)\s+',    # 1) pattern
            r'^[•\-\*]\s+',  # bullet points
        ]

        # If cell has multiple lines with sub-patterns, parse them
        sub_questions = []
        current_question = []

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # Check if this line starts a new sub-question
            is_sub_question = any(re.match(pattern, line) for pattern in sub_patterns)

            if is_sub_question and current_question:
                # Save previous question
                sub_questions.append(' '.join(current_question))
                current_question = [line]
            else:
                current_question.append(line)

        # Add the last question
        if current_question:
            sub_questions.append(' '.join(current_question))

        # If we found multiple sub-questions, return them
        if len(sub_questions) > 1:
            return [{'text': q, 'is_sub_question': True} for q in sub_questions]

        # Otherwise, return the original text as a single question
        return [{'text': text, 'is_sub_question': False}]


class HierarchicalQuestionParser:
    """Parse questions understanding their hierarchical structure using LLM for context"""

    def __init__(self, patterns: Optional[HierarchicalPattern] = None):
        self.patterns = patterns or HierarchicalPattern()
        api_key = os.getenv('ANTHROPIC_API_KEY')
        if not api_key:
            raise ValueError("ANTHROPIC_API_KEY environment variable is required")
        self.claude_client = anthropic.Anthropic(api_key=api_key)

    def _build_hierarchy_prompt(self, chunk: List[Dict]) -> str:
        """Build improved hierarchy parsing prompt with better instructions"""
        truncated_chunk = []
        for q in chunk:
            truncated_q = {
                'row': q['row'],
                'text': q['text'][:500] + '...' if len(q['text']) > 500 else q['text']
            }
            truncated_chunk.append(truncated_q)

        return f"""Analyze these questions for hierarchical structure. Return a JSON array with one result for each input question provided.

    TASK: Analyze the {len(truncated_chunk)} questions below and return exactly {len(truncated_chunk)} results in a JSON array.

    REQUIRED FIELDS for each item:
    - row: the exact row number from input
    - question_type: one of "parent_header", "numbered_requirement", "lettered_requirement", "sub_list_requirement", "general_question"
    - is_parent: true if it's a header that introduces a list, false otherwise
    - should_fill: false ONLY for parent headers, true for everything else
    - hierarchy_level: 0 for top level, 1 for sub-items, 2 for sub-sub-items
    - parent_row: the row number of the parent, or null

    HIERARCHY DETECTION RULES:
    1. Parent headers: End with ":" AND contain words like "following", "includes", "comprises"
    2. Numbered requirements: Start with "1)", "2)", etc. - These ARE requirements that need answers
    3. Lettered requirements: Start with "a.", "b.", etc. - These ARE requirements unless they also introduce lists
    4. Sub-list items: Items under parents - These ARE requirements that need answers
    5. When in doubt, mark as should_fill=true (it's better to fill than skip)

    INPUT QUESTIONS:
    {json.dumps(truncated_chunk, indent=1)}

    Return only a valid JSON array with {len(truncated_chunk)} items. Do not include any other text or explanations."""

    def _get_llm_response(self, prompt: str) -> Any:
        """Get response from LLM with improved retry logic and error handling"""
        max_retries = 5
        base_delay = 1

        for attempt in range(max_retries):
            try:
                # Add progressive delay
                if attempt > 0:
                    delay = base_delay * (2 ** (attempt - 1))
                    print(f"    ⏳ Waiting {delay}s before retry {attempt + 1}...")
                    time.sleep(delay)

                response = self.claude_client.messages.create(
                    model="claude-3-5-sonnet-20241022",
                    max_tokens=8192,
                    temperature=0.1,
                    messages=[
                        {"role": "user", "content": prompt}
                    ]
                )

                # Check if response was successful
                if not response:
                    raise ValueError(f"No response object returned on attempt {attempt + 1}")

                if not response.content or not response.content[0].text.strip():
                    raise ValueError(f"Empty response text on attempt {attempt + 1}")

                return response

            except Exception as e:
                error_msg = str(e)[:100]
                print(f"    ⚠️  LLM attempt {attempt + 1}/{max_retries} failed: {error_msg}")

                # If it's the last attempt, raise the exception
                if attempt == max_retries - 1:
                    print(f"    ❌ All {max_retries} attempts failed, will use fallback")
                    return None

        return None

    def _parse_llm_response(self, response_text: str) -> List[Dict]:
        """Parse and clean the LLM response with better error handling"""
        if not response_text or not response_text.strip():
            raise ValueError("Empty response text")

        response_text = response_text.strip()

        # Clean up response
        if response_text.startswith("```json"):
            response_text = response_text[7:]
        if response_text.startswith("```"):
            response_text = response_text[3:]
        if response_text.endswith("```"):
            response_text = response_text[:-3]

        # Remove any trailing commas before closing brackets/braces
        response_text = re.sub(r',(\s*[}\]])', r'\1', response_text)

        try:
            parsed = json.loads(response_text.strip())
            if not isinstance(parsed, list):
                raise ValueError(f"Expected list, got {type(parsed)}")
            return parsed
        except json.JSONDecodeError as e:
            print(f"    ⚠️  JSON parse error: {str(e)}")
            print(f"    📝 Response text: {response_text[:500]}...")
            raise

    def parse_questions_contextually_simplified(self, questions: List[Dict[str, Any]],
                                                chunk_size: int = 50,
                                                overlap: int = 10) -> List[Dict[str, Any]]:
        """Improved version with better error handling and chunk management"""
        all_results = {}
        total_questions = len(questions)

        print(f"    📊 Processing {total_questions} items in chunks of up to {chunk_size} with {overlap} overlap")

        chunk_num = 1
        start = 0
        max_chunks = 50
        consecutive_failures = 0
        max_consecutive_failures = 3

        while start < total_questions and chunk_num <= max_chunks:
            # Dynamically adjust chunk size if we're having issues
            current_chunk_size = chunk_size
            if consecutive_failures >= 2:
                current_chunk_size = max(20, chunk_size // 2)

            end = min(start + current_chunk_size, total_questions)
            chunk = questions[start:end]
            actual_chunk_size = len(chunk)

            print(f"    🔍 Processing chunk {chunk_num}: items {start}-{end - 1} ({actual_chunk_size} items)")

            # Safety check for empty chunks
            if actual_chunk_size == 0:
                print(f"    ⚠️  Empty chunk detected, moving to next")
                start += 1
                continue

            try:
                prompt = self._build_hierarchy_prompt(chunk)
                response = self._get_llm_response(prompt)

                if not response or not response.content[0].text:
                    print(f"    ⚠️  Empty response for chunk {chunk_num}, using fallback")
                    self._apply_fallback_to_chunk(chunk, start, all_results)
                    consecutive_failures += 1
                else:
                    try:
                        result_json = self._parse_llm_response(response.content[0].text)

                        # Improved result mapping with better error handling
                        processed_count = self._process_chunk_results(
                            result_json, chunk, start, end, all_results, total_questions
                        )

                        print(f"    ✅ Chunk {chunk_num} processed: {processed_count} items mapped successfully")
                        consecutive_failures = 0

                    except Exception as parse_error:
                        print(f"    ⚠️  Parse error in chunk {chunk_num}: {str(parse_error)[:100]}")
                        self._apply_fallback_to_chunk(chunk, start, all_results)
                        consecutive_failures += 1

            except Exception as e:
                print(f"    ❌ LLM processing failed for chunk {chunk_num}: {str(e)[:100]}")
                print(f"    🔄 Applying fallback classification...")
                self._apply_fallback_to_chunk(chunk, start, all_results)
                consecutive_failures += 1

            # Check if we should bail out
            if consecutive_failures >= max_consecutive_failures:
                print(f"    ⚠️  Too many consecutive failures, applying fallback to remaining items")
                # Apply fallback to all remaining items
                for i in range(start, total_questions):
                    if i not in all_results:
                        all_results[i] = self._create_fallback_item()
                break

            # Calculate next start position more carefully
            start = self._calculate_next_start(start, current_chunk_size, overlap, total_questions)
            chunk_num += 1

        # Fill any remaining gaps with fallback
        self._fill_remaining_gaps(all_results, total_questions, questions)

        # Create final ordered list
        final_parsed = []
        for i in range(total_questions):
            if i in all_results:
                final_parsed.append(all_results[i])
            else:
                print(f"    ⚠️  Question {i} missing, adding fallback")
                final_parsed.append(self._create_fallback_item())

        print(f"    📈 Total processed: {len(final_parsed)} items ({len(all_results)} successfully parsed)")
        return final_parsed

    def _process_chunk_results(self, result_json: List[Dict], chunk: List[Dict],
                              start: int, end: int, all_results: Dict, total_questions: int) -> int:
        """Process chunk results with improved error handling"""
        processed_count = 0

        # Create mapping from row numbers to results
        row_to_result = {}
        for item in result_json:
            if isinstance(item, dict) and 'row' in item:
                row_to_result[item['row']] = item

        # Process each question in the chunk
        for i in range(len(chunk)):
            global_index = start + i

            # Safety check
            if global_index >= total_questions:
                break

            question = chunk[i]
            question_row = question.get('row')

            # Skip if already processed (overlap handling)
            if global_index in all_results:
                continue

            # Try to find matching result
            if question_row in row_to_result:
                item = row_to_result[question_row]

                try:
                    parsed_item = {
                        'question_type': QuestionType(item.get('question_type', 'general_question')),
                        'is_parent': item.get('is_parent', False),
                        'should_fill': item.get('should_fill', True),
                        'hierarchy_level': item.get('hierarchy_level', 0),
                        'parent_id': item.get('parent_row'),
                        'parent_text': None
                    }

                    # Add parent text if available
                    if parsed_item['parent_id'] is not None:
                        parent_q = next((q for q in chunk if q['row'] == parsed_item['parent_id']), None)
                        if parent_q:
                            parsed_item['parent_text'] = parent_q['text']

                    all_results[global_index] = parsed_item
                    processed_count += 1

                except Exception as e:
                    print(f"    ⚠️  Error processing item {global_index}: {str(e)[:50]}")
                    all_results[global_index] = self._create_fallback_item()

            else:
                # No matching result found - use fallback
                all_results[global_index] = self._create_fallback_item()

        return processed_count

    def _apply_fallback_to_chunk(self, chunk: List[Dict], start: int, all_results: Dict):
        """Apply fallback classification to entire chunk"""
        for i, question in enumerate(chunk):
            global_index = start + i
            if global_index not in all_results:
                all_results[global_index] = self._create_fallback_item()

    def _create_fallback_item(self) -> Dict[str, Any]:
        """Create a fallback item for failed processing"""
        return {
            'question_type': QuestionType.GENERAL_QUESTION,
            'is_parent': False,
            'should_fill': True,
            'hierarchy_level': 0,
            'parent_id': None,
            'parent_text': None
        }

    def _calculate_next_start(self, current_start: int, chunk_size: int, overlap: int, total_questions: int) -> int:
        """Calculate next start position with safety checks"""
        next_start = current_start + chunk_size - overlap

        # Ensure progress is made
        if next_start <= current_start:
            next_start = current_start + 1

        # Don't exceed total
        if next_start >= total_questions:
            return total_questions

        return next_start

    def _fill_remaining_gaps(self, all_results: Dict, total_questions: int, questions: List[Dict]):
        """Fill any remaining gaps in results"""
        gaps_filled = 0
        for i in range(total_questions):
            if i not in all_results:
                all_results[i] = self._create_fallback_item()
                gaps_filled += 1

        if gaps_filled > 0:
            print(f"    🔧 Filled {gaps_filled} remaining gaps with fallback classification")