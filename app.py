"""
Excel Documentation AI Agent - Streamlit Application
Main web interface for the documentation generator
"""

import streamlit as st
import os
import sys
from pathlib import Path
import tempfile
import time

# Add src directory to path
sys.path.insert(0, str(Path(__file__).parent))

from src.excel_analyzer import ExcelAnalyzer
from src.llm_integration import get_llm_integration
from src.documentation_generator import DocumentationGenerator, WorkbookDocumentationGenerator
from src.utils import validate_excel_file, get_output_path, ProgressTracker
from config.settings import PAGE_TITLE, PAGE_ICON, OLLAMA_BASE_URL, OLLAMA_MODEL


# Page configuration
st.set_page_config(
    page_title=PAGE_TITLE,
    page_icon=PAGE_ICON,
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom styling
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stTabs [data-baseweb="tab-list"] button {
        font-size: 16px;
    }
    .success-box {
        background-color: #d4edda;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #28a745;
    }
    .info-box {
        background-color: #d1ecf1;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #17a2b8;
    }
    </style>
""", unsafe_allow_html=True)


def initialize_session_state():
    """Initialize session state variables"""
    if "file_uploaded" not in st.session_state:
        st.session_state.file_uploaded = False
    if "analysis_complete" not in st.session_state:
        st.session_state.analysis_complete = False
    if "documentation_generated" not in st.session_state:
        st.session_state.documentation_generated = False
    if "progress_messages" not in st.session_state:
        st.session_state.progress_messages = []


def main():
    """Main application function"""
    initialize_session_state()

    # Sidebar
    with st.sidebar:
        st.header("⚙️ Configuration")

        st.subheader("LLM Settings")
        st.info(f"🔗 Ollama URL: `{OLLAMA_BASE_URL}`")
        st.info(f"🤖 Model: `{OLLAMA_MODEL}`")

        with st.expander("Ollama Setup Instructions"):
            st.markdown("""
            ### Make sure Ollama is running:
            
            ```bash
            ollama serve
            ```
            
            ### Pull a German-friendly model:
            
            ```bash
            # Mistral (lightweight, good for most tasks)
            ollama pull mistral
            
            # Or Mixtral (more capable)
            ollama pull mixtral
            
            # Or other German-friendly models
            ollama pull neural-chat
            ```
            """)

        st.divider()

        st.subheader("About")
        st.markdown("""
        **Excel Documentation AI Agent**
        
        Automatically generates comprehensive documentation for Excel files using:
        - Advanced Excel analysis
        - AI-powered insights from Ollama
        - Beautiful markdown formatting
        
        Version: 1.0.0
        """)

    # Main content
    st.title(f"{PAGE_ICON} {PAGE_TITLE}")

    st.markdown("""
    This application analyzes Excel files and generates comprehensive documentation including:
    - **Column Analysis**: Data types, statistics, and sample values
    - **Formula Detection**: Identifies and explains formulas
    - **Formatting Analysis**: Interprets colors and styles for data meaning
    - **External References**: Detects VLOOKUP and other linked data
    - **AI Insights**: Uses Ollama LLM to generate intelligent descriptions
    """)

    st.divider()

    # File input section
    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader("📁 Input File")

        input_method = st.radio(
            "Choose input method:",
            ["Upload File", "Use File Path"],
            horizontal=True,
            key="input_method"
        )

        if input_method == "Upload File":
            uploaded_file = st.file_uploader(
                "Upload an Excel file",
                type=["xlsx", "xls", "xlsm"],
                key="file_uploader"
            )

            if uploaded_file is not None:
                # Save uploaded file to temp location
                with tempfile.NamedTemporaryFile(
                    delete=False,
                    suffix=Path(uploaded_file.name).suffix,
                ) as tmp_file:
                    tmp_file.write(uploaded_file.getbuffer())
                    file_path = tmp_file.name
                    st.session_state.file_uploaded = True
                    st.session_state.current_file = file_path
                    st.success(f"✓ File uploaded: {uploaded_file.name}")
        else:
            file_path = st.text_input(
                "Enter full path to Excel file:",
                placeholder="/path/to/your/file.xlsx",
                key="file_path_input"
            )

            if file_path:
                is_valid, message = validate_excel_file(file_path)
                if is_valid:
                    st.session_state.file_uploaded = True
                    st.session_state.current_file = file_path
                    st.success(f"✓ {message}")
                else:
                    st.error(f"✗ {message}")
                    st.session_state.file_uploaded = False

    with col2:
        st.subheader("📊 File Info")
        if st.session_state.file_uploaded and "current_file" in st.session_state:
            try:
                file_size = os.path.getsize(st.session_state.current_file)
                st.metric("File Size", f"{file_size / 1024:.2f} KB")
            except:
                pass

    st.divider()

    # Analysis section
    if st.session_state.file_uploaded and "current_file" in st.session_state:
        st.subheader("🔍 Analysis")

        col1, col2, col3 = st.columns(3)

        with col1:
            analyze_button = st.button(
                "▶ Start Analysis",
                use_container_width=True,
                type="primary",
                key="analyze_button"
            )

        with col2:
            st.button(
                "🔄 Reset",
                use_container_width=True,
                key="reset_button",
                on_click=lambda: st.session_state.clear()
            )

        if analyze_button:
            try:
                progress_container = st.container()
                status_container = st.container()

                with progress_container:
                    progress_bar = st.progress(0)
                    status_text = st.empty()

                with status_container:
                    st.info("🔄 Initializing LLM connection...")

                # Step 1: Initialize LLM
                progress_bar.progress(10)
                status_text.info("🔄 Connecting to Ollama...")
                try:
                    llm = get_llm_integration()
                    status_text.success("✓ Connected to Ollama")
                except ConnectionError as e:
                    st.error(f"❌ {str(e)}")
                    st.stop()

                # Step 2: Analyze Excel and generate documentation
                progress_bar.progress(30)
                status_text.info("🔄 Analyzing Excel file...")

                try:
                    import pandas as pd
                    sheet_names = pd.ExcelFile(st.session_state.current_file).sheet_names
                except Exception:
                    sheet_names = []

                progress_bar.progress(60)
                status_text.info("🔄 Generating documentation...")

                if len(sheet_names) > 1:
                    doc_generator = WorkbookDocumentationGenerator(
                        st.session_state.current_file, llm
                    )
                else:
                    analyzer = ExcelAnalyzer(st.session_state.current_file, sheet_name=None)
                    doc_generator = DocumentationGenerator(analyzer, llm)

                progress_bar.progress(80)
                status_text.info("🔄 Creating markdown output...")
                documentation = doc_generator.generate_documentation()
                progress_bar.progress(95)

                # Step 4: Save file
                progress_bar.progress(98)
                status_text.info("🔄 Saving documentation...")
                output_path = get_output_path(
                    st.session_state.current_file,
                    output_dir="./output"
                )
                with open(output_path, "w", encoding="utf-8") as f:
                    f.write(documentation)

                progress_bar.progress(100)
                status_text.success("✓ Documentation generated successfully!")

                st.session_state.documentation_generated = True
                st.session_state.documentation_content = documentation
                st.session_state.output_file_path = output_path

                # Show analysis summary
                st.divider()
                st.subheader("📈 Analysis Summary")

                cols = st.columns(4)
                with cols[0]:
                    st.metric("Rows", analysis_report["dimensions"]["rows"])
                with cols[1]:
                    st.metric("Columns", analysis_report["dimensions"]["columns"])
                with cols[2]:
                    formula_cols = sum(1 for col in analysis_report["columns"] if col["has_formula"])
                    st.metric("Formula Columns", formula_cols)
                with cols[3]:
                    linked_cols = sum(1 for col in analysis_report["columns"] if col["contains_link"])
                    st.metric("Linked Columns", linked_cols)

                # Show column list
                with st.expander("📋 Column Details", expanded=False):
                    col_data = []
                    for col in analysis_report["columns"]:
                        col_data.append({
                            "Name": col["name"],
                            "Type": col["data_type"],
                            "Non-Null": col["non_null_count"],
                            "Unique": col["unique_values"],
                            "Formula": "✓" if col["has_formula"] else "-",
                            "Link": col.get("link_type", "-") if col["contains_link"] else "-",
                        })

                    st.dataframe(col_data, use_container_width=True, hide_index=True)

            except Exception as e:
                st.error(f"❌ Error during analysis: {str(e)}")
                import traceback
                with st.expander("Error Details"):
                    st.code(traceback.format_exc())

    # Documentation display and download
    if st.session_state.documentation_generated and "documentation_content" in st.session_state:
        st.divider()
        st.subheader("📄 Generated Documentation")

        # Tabs for different views
        tab1, tab2, tab3 = st.tabs(["Preview", "Raw Markdown", "Download"])

        with tab1:
            st.markdown(st.session_state.documentation_content)

        with tab2:
            st.text_area(
                "Markdown Source:",
                value=st.session_state.documentation_content,
                height=600,
                disabled=True,
                key="markdown_source"
            )

        with tab3:
            # Download button
            st.download_button(
                label="📥 Download Markdown File",
                data=st.session_state.documentation_content,
                file_name=Path(st.session_state.output_file_path).name,
                mime="text/markdown",
                use_container_width=True,
                type="primary"
            )

            st.info(f"📂 File also saved to: `{st.session_state.output_file_path}`")

    # Help section
    with st.expander("❓ Help & FAQ"):
        st.markdown("""
        ### How does it work?
        
        1. **Upload** your Excel file
        2. **Click** "Start Analysis"
        3. The system will:
           - Parse the Excel structure
           - Analyze columns, formulas, and formatting
           - Use AI to generate descriptions
           - Create comprehensive documentation
        
        ### What does it analyze?
        
        - **Column Names**: Infers meaning from column names
        - **Data Types**: Automatically detects data types
        - **Formulas**: Identifies and explains Excel formulas
        - **Formatting**: Analyzes colors to infer data grouping
        - **External Links**: Detects VLOOKUP and references
        
        ### Requirements
        
        - Ollama must be running locally
        - A language model must be installed (Mistral, Mixtral, etc.)
        - Excel file must be valid (.xlsx, .xls, or .xlsm)
        
        ### Troubleshooting
        
        **Q: "Cannot connect to Ollama"**
        A: Start Ollama with `ollama serve` in another terminal
        
        **Q: "Model not found"**
        A: Pull the model with `ollama pull mistral`
        
        **Q: "Documentation is incomplete"**
        A: Check that Ollama is running and has enough memory
        """)


if __name__ == "__main__":
    main()
