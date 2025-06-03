"""
Utility functions for file operations and analysis
"""

import os
import json
from datetime import datetime
from typing import Dict, Any


def save_extraction_results(results: Dict[str, Any], file_path: str) -> str:
    """Save extraction results with enhanced analysis"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = os.path.splitext(os.path.basename(file_path))[0]
    output_dir = f"enhanced_extraction_{base_name}_{timestamp}"
    os.makedirs(output_dir, exist_ok=True)

    # Save main extraction results
    output_file = os.path.join(output_dir, "extraction_results.json")
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, default=str, ensure_ascii=False)

    # Save enhanced analysis report
    analysis_file = os.path.join(output_dir, "analysis_report.json")
    analysis_data = create_analysis_report(results, file_path, timestamp)

    with open(analysis_file, 'w', encoding='utf-8') as f:
        json.dump(analysis_data, f, indent=2, ensure_ascii=False)

    return output_dir, output_file, analysis_file


def create_analysis_report(results: Dict[str, Any], file_path: str, timestamp: str) -> Dict[str, Any]:
    """Create enhanced analysis report"""
    analysis_data = {
        'file': file_path,
        'timestamp': timestamp,
        'extraction_method': 'Enhanced with Statistical Analysis',
        'global_context': results.get('global_context'),
        'sheet_summary': {}
    }

    total_questions = 0
    total_fillable = 0

    for sheet_name, sheet_data in results.get('sheet_results', {}).items():
        sheet_total = sheet_data.get('total_questions_extracted', 0)
        sheet_fillable = sheet_data.get('statistics', {}).get('fillable_questions', 0)

        total_questions += sheet_total
        total_fillable += sheet_fillable

        analysis_data['sheet_summary'][sheet_name] = {
            'total_items': sheet_total,
            'fillable_requirements': sheet_fillable,
            'structural_items': sheet_total - sheet_fillable,
            'hierarchy_stats': sheet_data.get('hierarchy_stats', {}),
            'column_detection': {
                'question_column': sheet_data.get('column_info', {}).get('question_column'),
                'answer_columns': sheet_data.get('column_info', {}).get('answer_columns', []),
                'confidence': sheet_data.get('column_info', {}).get('confidence')
            }
        }

    analysis_data['overall_summary'] = {
        'total_items_extracted': total_questions,
        'total_fillable_requirements': total_fillable,
        'total_structural_items': total_questions - total_fillable,
        'fillable_percentage': round((total_fillable / max(total_questions, 1)) * 100, 2)
    }

    return analysis_data


def print_extraction_summary(results: Dict[str, Any], use_llm_hierarchy: bool):
    """Print enhanced extraction summary"""
    total_questions = 0
    total_fillable = 0

    for sheet_data in results.get('sheet_results', {}).values():
        total_questions += sheet_data.get('total_questions_extracted', 0)
        total_fillable += sheet_data.get('statistics', {}).get('fillable_questions', 0)

    print(f"\n📊 Enhanced Extraction Summary:")
    print(f"  🎯 Intelligent detection used: Content sheets, statistical column analysis")
    print(f"  📋 Total items extracted: {total_questions}")
    print(f"  ✅ Fillable requirements: {total_fillable}")
    print(f"  🏗️  Structural items: {total_questions - total_fillable}")
    print(f"  📈 Fillable percentage: {round((total_fillable / max(total_questions, 1)) * 100, 2)}%")
    print(f"  🧠 Hierarchy parsing: {'LLM-based' if use_llm_hierarchy else 'Rule-based'}")


def load_extraction_results(extraction_path: str) -> Dict[str, Any]:
    """Load extraction results from file or directory"""
    if os.path.isdir(extraction_path):
        extraction_file = os.path.join(extraction_path, "extraction_results.json")
    else:
        extraction_file = extraction_path

    if not os.path.exists(extraction_file):
        raise FileNotFoundError(f"Extraction file not found: {extraction_file}")

    with open(extraction_file, 'r', encoding='utf-8') as f:
        return json.load(f)


def validate_file_path(file_path: str) -> bool:
    """Validate if the Excel file exists and is readable"""
    if not os.path.exists(file_path):
        print(f"❌ File not found: {file_path}")
        return False

    if not file_path.lower().endswith(('.xlsx', '.xls')):
        print(f"❌ File must be an Excel file (.xlsx or .xls): {file_path}")
        return False

    return True