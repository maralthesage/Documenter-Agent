"""
Documentation Generator Module
Generates well-formatted markdown documentation from analyzed Excel data
"""

from typing import Dict, List, Any, Optional
from datetime import datetime
from src.excel_analyzer import ExcelAnalyzer, ColumnInfo
from src.llm_integration import LLMIntegration
import re
import pandas as pd


class DocumentationGenerator:
    """Generates markdown documentation from Excel analysis"""

    def __init__(self, analyzer: ExcelAnalyzer, llm: LLMIntegration):
        self.analyzer = analyzer
        self.llm = llm
        self.analysis_report = None

    def generate_documentation(self) -> str:
        """Generate complete markdown documentation"""
        # Get analysis report
        self.analysis_report = self.analyzer.analyze()

        # Build markdown document
        doc = self._generate_header()
        doc += "\n\n"
        doc += self._generate_body(include_toc=True)
        doc += "\n\n"
        doc += self._generate_footer()

        return doc
    
    def generate_sheet_documentation(self, include_toc: bool = False) -> str:
        """Generate documentation for a sheet, intended for workbook-level output"""
        self.analysis_report = self.analyzer.analyze()

        doc = self._generate_sheet_header()
        doc += "\n\n"
        doc += self._shift_headings(self._generate_body(include_toc=include_toc), 1)

        return doc

    def _generate_body(self, include_toc: bool = True) -> str:
        """Generate the core sections for a sheet"""
        parts = [
            self._generate_overview(),
        ]

        if include_toc:
            parts.append(self._generate_table_of_contents())

        parts.extend([
            self._generate_data_structure(),
            self._generate_columns_documentation(),
            self._generate_data_relationships(),
            self._generate_formatting_analysis(),
        ])

        formulas = self._generate_formulas_section()
        if formulas:
            parts.append(formulas)

        return "\n\n".join(parts)

    def _generate_sheet_header(self) -> str:
        """Generate a per-sheet header for workbook documentation"""
        sheet_name = self.analysis_report.get("sheet_name", "Unknown")
        rows = self.analysis_report["dimensions"]["rows"]
        cols = self.analysis_report["dimensions"]["columns"]

        header = f"""## Sheet: {sheet_name}

**Rows:** {rows}  
**Columns:** {cols}"""

        return header

    def _generate_header(self) -> str:
        """Generate document header"""
        file_name = self.analysis_report["file_path"].split("/")[-1]
        sheet_name = self.analysis_report.get("sheet_name", "Unknown")

        header = f"""# Excel Documentation Report

**File:** `{file_name}`  
**Sheet:** `{sheet_name}`  
**Generated:** {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}  
**Total Rows:** {self.analysis_report['dimensions']['rows']}  
**Total Columns:** {self.analysis_report['dimensions']['columns']}"""

        return header

    def _generate_overview(self) -> str:
        """Generate overview section with LLM"""
        columns_info = self.analyzer.get_column_info()

        overview = "## Overview\n\n"

        # Get LLM-generated overview
        analysis_data = []
        for col in columns_info:
            analysis_data.append({
                "name": str(col.name),
                "data_type": str(col.data_type),
                "description": f"{int(col.unique_values)} unique values, {int(col.non_null_count)} non-null",
            })

        try:
            llm_overview = self.llm.generate_comprehensive_documentation(
                analysis_data, str(self.analysis_report.get("sheet_name", "Sheet"))
            )
            overview += llm_overview
        except Exception as e:
            overview += f"*Unable to generate LLM overview: {str(e)}*"

        return overview

    def _generate_table_of_contents(self) -> str:
        """Generate table of contents"""
        columns = self.analyzer.get_column_info()

        toc = "## Table of Contents\n\n"
        toc += "- [Data Structure](#data-structure)\n"
        toc += "- [Column Descriptions](#column-descriptions)\n"

        for i, col in enumerate(columns, 1):
            col_anchor = col.name.lower().replace(" ", "-").replace("_", "-")
            toc += f"  - [{col.name}](#{col_anchor})\n"

        toc += "- [Data Relationships](#data-relationships)\n"
        toc += "- [Formatting Analysis](#formatting-analysis)\n"

        if any(col.has_formula for col in columns):
            toc += "- [Formulas](#formulas)\n"

        return toc

    def _generate_data_structure(self) -> str:
        """Generate data structure section"""
        report = self.analysis_report
        dims = report["dimensions"]

        structure = "## Data Structure\n\n"
        structure += f"| Metric | Value |\n"
        structure += f"|--------|-------|\n"
        structure += f"| Number of Rows | {dims['rows']} |\n"
        structure += f"| Number of Columns | {dims['columns']} |\n"
        structure += f"| Total Cells | {dims['rows'] * dims['columns']} |\n"

        # Data type distribution
        data_types = {}
        for col in report["columns"]:
            dt = col["data_type"]
            data_types[dt] = data_types.get(dt, 0) + 1

        structure += f"| Data Type Distribution | "
        dist_str = ", ".join(f"{dt}: {count}" for dt, count in sorted(data_types.items()))
        structure += f"{dist_str} |\n"

        return structure

    def _generate_columns_documentation(self) -> str:
        """Generate detailed column documentation"""
        columns = self.analyzer.get_column_info()

        doc = "## Column Descriptions\n\n"

        for col in columns:
            doc += self._generate_column_section(col)
            doc += "\n\n"

        return doc

    def _generate_column_section(self, col: ColumnInfo) -> str:
        """Generate documentation for a single column"""
        col_anchor = col.name.lower().replace(" ", "-").replace("_", "-")

        section = f"### {col.name} {{#{col_anchor}}}\n\n"

        # Basic metadata
        section += f"| Property | Value |\n"
        section += f"|----------|-------|\n"
        section += f"| Data Type | `{col.data_type}` |\n"
        section += f"| Non-Null Count | {col.non_null_count} |\n"
        section += f"| Unique Values | {col.unique_values} |\n"
        section += f"| Column Index | {col.index + 1} |\n"

        if col.numeric_range:
            section += f"| Range | {col.numeric_range[0]} to {col.numeric_range[1]} |\n"

        section += "\n"

        # Column description from LLM
        try:
            # Ensure sample values are strings to avoid JSON serialization errors
            safe_samples = [str(v) if v is not None else "" for v in col.sample_values[:5]]
            description = self.llm.infer_data_meaning(
                str(col.name),
                str(col.data_type),
                safe_samples
            )
            section += f"**Description:**\n\n{description}\n\n"
        except Exception as e:
            section += f"*Unable to generate description: {str(e)}*\n\n"

        # Sample values
        if col.sample_values:
            section += f"**Sample Values:**\n\n"
            for i, val in enumerate(col.sample_values[:5], 1):
                section += f"{i}. `{val}`\n"
            section += "\n"

        # Formatting information
        if col.fill_color or col.font_color:
            section += f"**Formatting:**\n\n"
            if col.fill_color:
                section += f"- Fill Color: `{col.fill_color}`\n"
            if col.font_color:
                section += f"- Font Color: `{col.font_color}`\n"

            try:
                # Ensure formatting values are strings to avoid JSON serialization errors
                formatting = {}
                if col.fill_color:
                    formatting["fill_color"] = str(col.fill_color)
                if col.font_color:
                    formatting["font_color"] = str(col.font_color)

                significance = self.llm.explain_formatting_significance(
                    str(col.name), formatting
                )
                section += f"\n**Formatting Significance:**\n\n{significance}\n"
            except Exception as e:
                section += f"*Unable to analyze formatting: {str(e)}*\n"

            section += "\n"

        # Formula information
        if col.has_formula:
            section += f"**Formulas Detected:** Yes\n\n"
            for formula in col.formulas[:3]:  # Limit to first 3 formulas
                section += f"- Formula: `{formula}`\n"

            section += "\n"

            # LLM analysis of formulas
            try:
                for formula in col.formulas[:2]:
                    explanation = self.llm.analyze_formula(
                        formula, col.name,
                        f"Column contains {col.unique_values} unique values"
                    )
                    section += f"**Formula Explanation:**\n\n{explanation}\n\n"
            except Exception as e:
                section += f"*Unable to analyze formulas: {str(e)}*\n\n"

        # Link information
        if col.contains_link:
            section += f"**External References:** `{col.link_type}`\n\n"
            section += f"This column contains {col.link_type} references to external data.\n"

        return section

    def _generate_data_relationships(self) -> str:
        """Generate section about data relationships"""
        columns = self.analyzer.get_column_info()

        section = "## Data Relationships\n\n"

        # Identify formula columns and their dependencies
        formula_columns = [col for col in columns if col.has_formula]

        if formula_columns:
            section += "### Computed Columns\n\n"
            section += "The following columns contain formulas and derive their values from other columns:\n\n"

            for col in formula_columns:
                section += f"- **{col.name}**: Contains {len(col.formulas)} formula(s)\n"

            section += "\n"

        # Identify linked columns
        linked_columns = [col for col in columns if col.contains_link]

        if linked_columns:
            section += "### Linked Data\n\n"
            section += "The following columns contain references to external data:\n\n"

            for col in linked_columns:
                section += f"- **{col.name}**: `{col.link_type}`\n"

            section += "\n"

        # Formatting groups
        formatting_groups = self.analyzer.extract_shared_formatting_groups()

        if formatting_groups:
            section += "### Data Groups (by Formatting)\n\n"
            section += "Columns with similar formatting may represent related data:\n\n"

            for i, group in enumerate(formatting_groups, 1):
                section += f"**Group {i}:** {', '.join(f'`{col}`' for col in group)}\n"

            section += "\n"

        return section

    def _generate_formatting_analysis(self) -> str:
        """Generate formatting analysis section"""
        columns = self.analyzer.get_column_info()
        formatted_cols = [col for col in columns if col.fill_color or col.font_color]

        section = "## Formatting Analysis\n\n"

        if not formatted_cols:
            section += "No special formatting detected in this sheet.\n"
            return section

        section += "The following columns have special formatting that may indicate data grouping or significance:\n\n"

        for col in formatted_cols:
            section += f"### {col.name}\n\n"

            if col.fill_color:
                section += f"- **Fill Color:** `{col.fill_color}`\n"
            if col.font_color:
                section += f"- **Font Color:** `{col.font_color}`\n"

            section += "\n"

        return section

    def _generate_formulas_section(self) -> str:
        """Generate detailed formulas section"""
        columns = self.analyzer.get_column_info()
        formula_cols = [col for col in columns if col.has_formula]

        if not formula_cols:
            return ""

        section = "## Formulas\n\n"
        section += "This sheet contains formulas in the following columns:\n\n"

        for col in formula_cols:
            section += f"### {col.name}\n\n"

            for i, formula in enumerate(col.formulas, 1):
                # Truncate long formulas
                display_formula = formula if len(formula) < 100 else formula[:97] + "..."
                section += f"**Formula {i}:**\n\n```excel\n{display_formula}\n```\n\n"

            section += "\n"

        return section

    def _generate_footer(self) -> str:
        """Generate document footer"""
        return self.default_footer()

    @staticmethod
    def default_footer() -> str:
        """Default footer text for documentation"""
        footer = """---

## Notes

- This documentation was automatically generated from the Excel file.
- The analysis includes column types, formatting, formulas, and external references.
- AI-generated descriptions are based on column names, data types, and sample values.
- For sensitive data, please ensure proper access controls are in place.
- Generated using Excel Documentation AI Agent
"""
        return footer

    def save_to_file(self, output_path: str) -> str:
        """Generate documentation and save to markdown file"""
        documentation = self.generate_documentation()

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(documentation)

        return output_path

    def _shift_headings(self, text: str, shift: int) -> str:
        """Shift markdown heading levels while preserving code blocks."""
        if shift == 0:
            return text

        lines = text.splitlines()
        in_code_block = False
        shifted_lines = []

        for line in lines:
            if line.startswith("```"):
                in_code_block = not in_code_block
                shifted_lines.append(line)
                continue

            if in_code_block or not line.startswith("#"):
                shifted_lines.append(line)
                continue

            hashes = len(line) - len(line.lstrip("#"))
            new_level = max(1, min(6, hashes + shift))
            shifted_lines.append("#" * new_level + line[hashes:])

        return "\n".join(shifted_lines)


class WorkbookDocumentationGenerator:
    """Generates markdown documentation for all sheets in a workbook"""

    def __init__(self, file_path: str, llm: LLMIntegration):
        self.file_path = file_path
        self.llm = llm

    def generate_documentation(self) -> str:
        """Generate documentation for all sheets in a workbook"""
        xls = pd.ExcelFile(self.file_path)
        sheet_names = xls.sheet_names

        doc = self._generate_workbook_header(sheet_names)

        sections = []
        for sheet_name in sheet_names:
            analyzer = ExcelAnalyzer(self.file_path, sheet_name=sheet_name)
            generator = DocumentationGenerator(analyzer, self.llm)
            sections.append(generator.generate_sheet_documentation(include_toc=False))

        doc += "\n\n" + "\n\n".join(sections)
        doc += "\n\n" + DocumentationGenerator.default_footer()

        return doc

    def _generate_workbook_header(self, sheet_names: List[str]) -> str:
        """Generate the workbook-level header"""
        file_name = self.file_path.split("/")[-1]
        sheet_list = "\n".join(f"- {name}" for name in sheet_names)

        header = f"""# Excel Documentation Report

**File:** `{file_name}`  
**Generated:** {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}  
**Total Sheets:** {len(sheet_names)}

## Sheets

{sheet_list}"""

        return header
