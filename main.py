"""
Enhanced Excel Question Extraction System with Dynamic Detection
Main entry point for the application

Usage:
    python main.py extract <excel_file> [--no-llm-hierarchy] [--content-sheet=SheetName]
    python main.py fill <excel_file> <extraction_json>
    python main.py both <excel_file> [--no-llm-hierarchy] [--content-sheet=SheetName]
"""

import sys
import os
from typing import Optional

# Add current directory to path for imports
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from config import validate_environment
from extractor import EnhancedExcelExtractor
from filler import EnhancedExcelFiller
from utils import (
    save_extraction_results, print_extraction_summary, 
    load_extraction_results, validate_file_path
)


def print_usage():
    """Print usage information"""
    print("🚀 Enhanced Excel Question Extraction & Filling System")
    print("=" * 60)
    print("Features:")
    print("  🔍 Intelligent content sheet detection (not hardcoded)")
    print("  📊 Statistical column analysis")
    print("  🧠 Smart hierarchy parsing")
    print("  ✏️  Cross-column intelligent filling")
    print("  📚 Global context awareness")
    print()
    print("Usage:")
    print("  python main.py extract <excel_file> [--no-llm-hierarchy] [--content-sheet=SheetName]")
    print("  python main.py fill <excel_file> <extraction_json>")
    print("  python main.py both <excel_file> [--no-llm-hierarchy] [--content-sheet=SheetName]")
    print()
    print("Options:")
    print("  --no-llm-hierarchy: Use rule-based hierarchy parsing instead of LLM")
    print("  --content-sheet=SheetName: Manually specify content sheet name")
    print()
    print("Examples:")
    print("  python main.py extract document.xlsx")
    print("  python main.py extract document.xlsx --content-sheet=Instructions")
    print("  python main.py both document.xlsx --no-llm-hierarchy")


def parse_arguments():
    """Parse command line arguments"""
    if len(sys.argv) < 3:
        return None, None, None, None, None

    command = sys.argv[1]
    file_path = sys.argv[2]

    # Parse options
    use_llm_hierarchy = "--no-llm-hierarchy" not in sys.argv
    content_sheet_name = None
    extraction_json = None

    for arg in sys.argv:
        if arg.startswith("--content-sheet="):
            content_sheet_name = arg.split("=", 1)[1]

    # For fill command, get extraction JSON
    if command == "fill" and len(sys.argv) >= 4:
        extraction_json = sys.argv[3]

    return command, file_path, use_llm_hierarchy, content_sheet_name, extraction_json


def run_extraction(file_path: str, use_llm_hierarchy: bool = True, 
                  content_sheet_name: Optional[str] = None) -> str:
    """Run extraction process"""
    print("📋 Step 1: Enhanced Question Extraction")
    print(f"   🔍 Intelligent detection: {'Enabled' if use_llm_hierarchy else 'Rule-based only'}")
    print(f"   📊 Statistical analysis: Enabled")
    print(f"   📚 Global context: Enabled")
    
    if content_sheet_name:
        print(f"   📄 Manual content sheet: {content_sheet_name}")

    extractor = EnhancedExcelExtractor(
        file_path, 
        use_llm_hierarchy=use_llm_hierarchy,
        content_sheet_name=content_sheet_name
    )
    
    results = extractor.extract_all()
    
    # Save results
    output_dir, extraction_file, analysis_file = save_extraction_results(results, file_path)
    
    print(f"\n📁 Enhanced results saved to: {output_dir}")
    print(f"  📄 Full extraction: {extraction_file}")
    print(f"  📊 Analysis report: {analysis_file}")
    
    # Print summary
    print_extraction_summary(results, use_llm_hierarchy)
    
    return output_dir


def run_filling(file_path: str, extraction_json: str) -> str:
    """Run filling process"""
    print("\n✏️  Step 2: Enhanced Intelligent Filling")
    print("   🤖 Cross-column logic: Enabled")
    print("   📊 Statistical strategies: Enabled")
    print("   🧠 Context-aware responses: Enabled")

    # Load extraction results
    extraction_results = load_extraction_results(extraction_json)
    
    filler = EnhancedExcelFiller(file_path)
    output_file = filler.fill_all(extraction_results)
    
    return output_file


def main():
    """Main execution function"""
    try:
        # Validate environment
        validate_environment()
        
        # Parse arguments
        command, file_path, use_llm_hierarchy, content_sheet_name, extraction_json = parse_arguments()
        
        if not command or not file_path:
            print_usage()
            sys.exit(1)
        
        # Validate file path
        if not validate_file_path(file_path):
            sys.exit(1)
        
        if command == "extract":
            run_extraction(file_path, use_llm_hierarchy, content_sheet_name)
            
        elif command == "fill":
            if not extraction_json:
                print("❌ Please provide extraction JSON file")
                print("Usage: python main.py fill <excel_file> <extraction_json>")
                sys.exit(1)
            
            run_filling(file_path, extraction_json)
            
        elif command == "both":
            # Run extraction first
            output_dir = run_extraction(file_path, use_llm_hierarchy, content_sheet_name)
            
            # Then run filling
            extraction_json = os.path.join(output_dir, "extraction_results.json")
            output_file = run_filling(file_path, extraction_json)
            
            print(f"\n🎉 Enhanced Processing Complete!")
            print(f"   ✅ Intelligent detection and analysis applied")
            print(f"   ✅ Cross-column logic and context awareness used")
            print(f"   ✅ Final enhanced file: {output_file}")
            print(f"   📁 Complete results in: {output_dir}")
            
        else:
            print(f"❌ Unknown command: {command}")
            print("Use: extract, fill, or both")
            sys.exit(1)
            
    except KeyboardInterrupt:
        print("\n⚠️  Process interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()