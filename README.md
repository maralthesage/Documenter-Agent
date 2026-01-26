# Excel Documentation AI Agent

An intelligent AI agent that automatically generates comprehensive documentation for Excel files using Ollama, LangChain, and Streamlit.

## Features

- **Column Analysis**: Automatically analyzes and documents every column
- **Data Type Detection**: Identifies data types and patterns
- **Formula Analysis**: Detects formulas and traces back to original data
- **Formatting Inference**: Analyzes colors and formatting to infer data groupings
- **Cross-Reference Resolution**: Attempts to resolve VLOOKUP and other external data references
- **Markdown Documentation**: Generates well-structured, formatted markdown documentation

## Prerequisites

- Python 3.9+
- Ollama installed and running locally
- A German-speaking LLM (Mistral or Mixtral) pulled in Ollama

## Installation

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Set Up Ollama

Make sure Ollama is running:

```bash
ollama serve
```

In another terminal, pull a German-speaking model:

```bash
# Mistral (smaller, efficient)
ollama pull mistral

# Or Mixtral (more capable)
ollama pull mixtral
```

### 3. Configure Environment

Copy `.env.example` to `.env` and update if needed:

```bash
cp .env.example .env
```

## Usage

Run the Streamlit application:

```bash
streamlit run app.py
```

Then:
1. Open your browser to `http://localhost:8501`
2. Upload or specify the path to your Excel file
3. Click "Generate Documentation"
4. Download the generated markdown documentation

## Project Structure

```
├── src/
│   ├── __init__.py
│   ├── excel_analyzer.py      # Excel parsing and analysis
│   ├── llm_integration.py     # Ollama and LangChain integration
│   ├── documentation_generator.py  # Documentation generation
│   └── utils.py               # Utility functions
├── config/
│   ├── __init__.py
│   └── settings.py            # Configuration settings
├── app.py                      # Main Streamlit application
├── requirements.txt            # Python dependencies
└── .env.example               # Environment variables template
```

## How It Works

1. **Excel Analysis**: Reads the Excel file and extracts:
   - Column names and data types
   - Cell formatting (colors, styles)
   - Formulas and their components
   - Data samples for context

2. **LLM Processing**: Uses Ollama with LangChain to:
   - Analyze column names and infer meaning
   - Describe data patterns
   - Explain formula logic
   - Interpret formatting significance

3. **Documentation Generation**: Creates a structured markdown file with:
   - Overview of the sheet
   - Detailed column documentation
   - Data relationships and dependencies
   - Computed fields explanations

## Configuration

Edit `config/settings.py` to customize:
- Ollama URL and model
- Processing parameters
- Documentation formatting

## Troubleshooting

- **Connection Error**: Ensure Ollama is running (`ollama serve`)
- **Model Not Found**: Pull the model with `ollama pull <model-name>`
- **Memory Issues**: Use a smaller model like `mistral`
- **Slow Processing**: Consider reducing the sample size in settings

## License

MIT
