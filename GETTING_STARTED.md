# Getting Started Guide

## Quick Start (5 minutes)

### 1. Install Ollama
- Visit https://ollama.ai
- Download and install for your OS
- Start Ollama: `ollama serve`

### 2. Pull a Model
In another terminal:
```bash
ollama pull mistral
```

### 3. Install Python Dependencies
```bash
pip install -r requirements.txt
```

### 4. Generate Sample Data
```bash
python create_samples.py
```

### 5. Run the Application
```bash
streamlit run app.py
```

Then open http://localhost:8501 in your browser.

---

## What You Can Do

1. **Upload your own Excel file** and get instant documentation
2. **See column analysis** with data types, statistics, and samples
3. **View formula explanations** for computed columns
4. **Understand formatting** - the AI explains what colors might mean
5. **Detect links** - identifies VLOOKUP and external references
6. **Download documentation** as a professional markdown file

---

## Project Structure

```
Documenter-Agent/
├── src/
│   ├── excel_analyzer.py       # Excel parsing & analysis
│   ├── llm_integration.py      # Ollama/LangChain integration
│   ├── documentation_generator.py  # Creates markdown docs
│   └── utils.py                # Helper functions
├── config/
│   └── settings.py             # Configuration
├── sample_data/                # Sample Excel files
├── output/                     # Generated documentation
├── app.py                      # Main Streamlit app
├── test.py                     # Test suite
├── create_samples.py           # Generate sample Excel files
├── run.sh / run.bat            # Quick start scripts
└── requirements.txt            # Python dependencies
```

---

## Features in Detail

### Excel Analysis
- **Automatic type detection**: int, float, string, datetime, boolean, URL
- **Formatting extraction**: Colors, styles, and their significance
- **Formula detection**: Identifies all formulas in the spreadsheet
- **Reference detection**: VLOOKUP, INDEX/MATCH, external links
- **Statistics**: Min, max, unique values, null counts

### AI-Powered Insights
- **Column naming inference**: What does the name tell us?
- **Formula explanation**: What does this formula do?
- **Formatting significance**: Why these colors/styles?
- **Data relationship mapping**: How columns relate to each other
- **Overview generation**: What's the purpose of this sheet?

### Documentation Output
- Professional markdown format
- Table of contents with anchors
- Detailed column documentation
- Formula breakdowns
- Data relationship diagrams
- Formatting analysis
- Formatting groups identification

---

## Configuration

Edit `config/settings.py` to customize:

```python
OLLAMA_MODEL = "mistral"      # Change LLM model
MAX_SAMPLE_ROWS = 50           # Sample size for analysis
INCLUDE_DATA_SAMPLES = True    # Include example values
SAMPLE_SIZE = 10               # Number of samples per column
```

Or edit `.env`:

```
OLLAMA_BASE_URL=http://localhost:11434
OLLAMA_MODEL=mistral
```

---

## Common Use Cases

### 1. Document Legacy Excel Files
Upload old Excel files with complex formulas and get instant documentation of what everything does.

### 2. Onboard New Team Members
Generate documentation to help new employees understand your data structures.

### 3. Data Audit
Analyze Excel files to understand data quality, relationships, and calculations.

### 4. Knowledge Transfer
Convert implicit knowledge about Excel files into explicit documentation.

---

## Model Recommendations

| Model | Size | Speed | Quality | Best For |
|-------|------|-------|---------|----------|
| Mistral | 7B | Fast | Good | General use |
| Mixtral | 8x7B | Moderate | Excellent | Complex analysis |
| Neural-Chat | 7B | Fast | Good | Conversational |
| Llama 2 | 7B/13B | Variable | Good | General use |

For German language support, all modern models work well. Mistral is recommended for best speed/quality balance.

---

## Troubleshooting

### Connection Issues
```bash
# Check if Ollama is running
curl http://localhost:11434/api/tags

# List available models
ollama list

# Pull a model if missing
ollama pull mistral
```

### Performance Issues
- Use `mistral` instead of `mixtral` for faster processing
- Reduce `MAX_SAMPLE_ROWS` in settings
- Close other applications
- Ensure sufficient RAM (4GB minimum, 8GB recommended)

### File Issues
- Ensure file is not open in Excel or other applications
- Check file permissions
- Try with a sample file first

---

## Next Steps

1. Try the application with sample data:
   ```bash
   python create_samples.py
   streamlit run app.py
   ```

2. Upload your own Excel files

3. Customize the prompts in `src/llm_integration.py` for your needs

4. Modify output format in `src/documentation_generator.py`

5. Integrate into your workflow or CI/CD pipeline

---

## Support

- Check SETUP.md for detailed installation instructions
- Review README.md for project overview
- Run `python test.py <file>` to test individual components
- Check error messages in the Streamlit terminal

---

## License

MIT License - Feel free to use and modify for your needs!
