"""
Test Script for Excel Documentation AI Agent
Run this to test the system without the Streamlit UI
"""

import sys
from pathlib import Path
import argparse
import json

# Add src to path
sys.path.insert(0, str(Path(__file__).parent))

from src.excel_analyzer import ExcelAnalyzer
from src.llm_integration import get_llm_integration
from src.documentation_generator import DocumentationGenerator
from src.utils import validate_excel_file, get_output_path


def test_excel_analysis(file_path: str):
    """Test the Excel analysis component"""
    print("\n" + "=" * 60)
    print("Testing Excel Analysis")
    print("=" * 60)

    # Validate file
    is_valid, message = validate_excel_file(file_path)
    print(f"File Validation: {message}")

    if not is_valid:
        print("❌ File validation failed")
        return False

    # Analyze
    print("\nAnalyzing Excel file...")
    analyzer = ExcelAnalyzer(file_path)

    try:
        analysis_report = analyzer.analyze()
        print("✓ Analysis complete")
        print(f"  Rows: {analysis_report['dimensions']['rows']}")
        print(f"  Columns: {analysis_report['dimensions']['columns']}")
        print(f"  Sheet: {analysis_report['sheet_name']}")

        # Show columns
        print("\nColumns detected:")
        for col in analysis_report["columns"]:
            formula_marker = " [FORMULA]" if col["has_formula"] else ""
            link_marker = f" [{col['link_type']}]" if col["contains_link"] else ""
            print(f"  - {col['name']} ({col['data_type']}){formula_marker}{link_marker}")

        return True, analyzer, analysis_report

    except Exception as e:
        print(f"❌ Analysis failed: {str(e)}")
        import traceback
        traceback.print_exc()
        return False, None, None


def test_llm_connection():
    """Test LLM connection"""
    print("\n" + "=" * 60)
    print("Testing LLM Connection")
    print("=" * 60)

    try:
        print("Connecting to Ollama LLM...")
        llm = get_llm_integration()
        print("✓ Connected to Ollama")

        # Test a simple query
        print("\nTesting LLM with sample query...")
        response = llm.llm.predict(text="What is Excel?")
        print(f"✓ LLM Response: {response[:100]}...")

        return True, llm

    except ConnectionError as e:
        print(f"❌ Connection failed: {str(e)}")
        return False, None
    except Exception as e:
        print(f"❌ LLM test failed: {str(e)}")
        import traceback
        traceback.print_exc()
        return False, None


def test_documentation_generation(analyzer, llm, file_path: str):
    """Test documentation generation"""
    print("\n" + "=" * 60)
    print("Testing Documentation Generation")
    print("=" * 60)

    try:
        print("Generating documentation...")
        doc_gen = DocumentationGenerator(analyzer, llm)
        documentation = doc_gen.generate_documentation()

        print(f"✓ Documentation generated ({len(documentation)} characters)")

        # Save to file
        output_path = get_output_path(file_path, output_dir="./output")
        print(f"\nSaving to: {output_path}")
        doc_gen.save_to_file(output_path)
        print("✓ Documentation saved")

        # Show preview
        lines = documentation.split("\n")
        print("\nDocumentation Preview (first 20 lines):")
        print("-" * 60)
        for line in lines[:20]:
            print(line)
        print("...")

        return True

    except Exception as e:
        print(f"❌ Documentation generation failed: {str(e)}")
        import traceback
        traceback.print_exc()
        return False


def main():
    """Main test function"""
    parser = argparse.ArgumentParser(
        description="Test the Excel Documentation AI Agent"
    )
    parser.add_argument(
        "file",
        help="Path to Excel file to test"
    )
    parser.add_argument(
        "--test-analysis-only",
        action="store_true",
        help="Only test Excel analysis, skip LLM"
    )
    parser.add_argument(
        "--test-llm-only",
        action="store_true",
        help="Only test LLM connection"
    )

    args = parser.parse_args()
    file_path = args.file

    print("\n" + "=" * 60)
    print("Excel Documentation AI Agent - Test Suite")
    print("=" * 60)

    # Test Excel analysis
    analysis_result = test_excel_analysis(file_path)

    if not analysis_result[0]:
        print("\n❌ Tests failed")
        return 1

    if args.test_analysis_only:
        print("\n✓ Excel analysis test passed")
        return 0

    analyzer, analysis_report = analysis_result[1], analysis_result[2]

    # Test LLM
    llm_result = test_llm_connection()

    if not llm_result[0]:
        print("\n❌ LLM connection failed")
        print("Make sure Ollama is running: ollama serve")
        return 1

    if args.test_llm_only:
        print("\n✓ LLM connection test passed")
        return 0

    llm = llm_result[1]

    # Test documentation generation
    if not test_documentation_generation(analyzer, llm, file_path):
        print("\n❌ Documentation generation failed")
        return 1

    print("\n" + "=" * 60)
    print("✓ All tests passed!")
    print("=" * 60)
    print("\nYou can now run: streamlit run app.py")

    return 0


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
