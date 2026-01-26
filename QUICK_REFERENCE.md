# Quick Reference Guide

## 🚀 Installation (5 minutes)

```bash
# 1. Install Ollama
# Download from https://ollama.ai

# 2. Start Ollama (Terminal 1)
ollama serve

# 3. Pull a model (Terminal 2)
ollama pull mistral

# 4. Install dependencies
pip install -r requirements.txt

# 5. Run the app
streamlit run app.py

# Opens: http://localhost:8501
```

---

## 📊 Usage

### Via Web Interface
1. Open http://localhost:8501
2. Upload Excel file
3. Click "Start Analysis"
4. Download markdown documentation

### Via Command Line
```bash
python test.py 'path/to/file.xlsx'
```

---

## 🔧 Configuration

Edit `.env`:
```
OLLAMA_BASE_URL=http://localhost:11434
OLLAMA_MODEL=mistral
```

Edit `config/settings.py` for advanced options.

---

## 📁 Project Files

| File | Purpose |
|------|---------|
| `app.py` | Main web application |
| `src/excel_analyzer.py` | Excel parsing |
| `src/llm_integration.py` | AI integration |
| `src/documentation_generator.py` | Documentation |
| `src/utils.py` | Helper functions |
| `test.py` | Testing |
| `create_samples.py` | Sample data |

---

## 🆘 Troubleshooting

| Problem | Solution |
|---------|----------|
| Cannot connect to Ollama | `ollama serve` in terminal |
| Model not found | `ollama pull mistral` |
| Python not found | `pip install -r requirements.txt` |
| Port already in use | Change port in Streamlit |

---

## 📚 Documentation

- `README.md` - Overview
- `GETTING_STARTED.md` - Detailed start guide
- `SETUP.md` - Installation instructions
- `API.md` - API reference
- `DEPLOYMENT.md` - Production guide
- `PROJECT_SUMMARY.md` - Complete project info

---

## 🎯 Features

- ✅ Automatic Excel analysis
- ✅ Formula detection & explanation
- ✅ Formatting analysis
- ✅ AI-powered insights
- ✅ Professional markdown output
- ✅ Web UI interface
- ✅ Supports German & English

---

## 💡 Tips

1. **Performance**: Use Mistral for speed, Mixtral for quality
2. **Memory**: Large files? Reduce MAX_SAMPLE_ROWS
3. **Timeout**: Still processing? Increase in settings
4. **Models**: List available: `ollama list`

---

## 🔗 Links

- [Ollama](https://ollama.ai)
- [LangChain](https://python.langchain.com/)
- [Streamlit](https://streamlit.io/)
- [pandas](https://pandas.pydata.org/)

---

## ⌨️ Keyboard Shortcuts (Streamlit)

- `Ctrl+C` - Stop application
- `R` - Rerun app
- `C` - Clear cache

---

## 📞 Common Commands

```bash
# List Ollama models
ollama list

# Pull model
ollama pull <model-name>

# Remove model
ollama rm <model-name>

# Generate sample files
python create_samples.py

# Run tests
python test.py './file.xlsx'

# Start web app
streamlit run app.py

# Check Python version
python --version

# Create virtual environment
python -m venv venv

# Activate venv (macOS/Linux)
source venv/bin/activate

# Activate venv (Windows)
venv\Scripts\activate
```

---

## 🎓 First Steps

1. **Setup (5 min)**
   - Install Ollama
   - Pull Mistral model
   - Install Python dependencies

2. **Test (2 min)**
   - Generate sample files
   - Run test.py

3. **Use (2 min)**
   - Start Streamlit
   - Upload Excel file
   - Get documentation!

---

## 📊 Performance

| File Size | Time |
|-----------|------|
| < 1MB | 1-2 minutes |
| 1-10MB | 2-5 minutes |
| 10-50MB | 5-10 minutes |
| > 50MB | 10+ minutes |

*Depends on model and system resources*

---

## 🔐 Security

- Local processing (no cloud)
- File validation
- Size limits
- Supports authentication

---

## 🌍 Supported Models

- ✅ Mistral (fast, recommended)
- ✅ Mixtral (powerful)
- ✅ Neural-Chat (conversational)
- ✅ Llama 2 (general)
- ✅ Any Ollama model

---

## 💾 Output Files

Documentation saved to: `./output/`

Format: `{filename}_documentation_{timestamp}.md`

Example: `sales_data_documentation_20240126_143022.md`

---

## 🎨 Customization

**Change LLM prompts**: Edit `src/llm_integration.py`

**Modify output format**: Edit `src/documentation_generator.py`

**Add new analysis**: Edit `src/excel_analyzer.py`

---

## 📈 Scaling

- Single file: Use web UI
- Multiple files: Use loop in Python
- Batch processing: Use Celery
- Production: See DEPLOYMENT.md

---

## ✅ Checklist

- [ ] Ollama installed
- [ ] Model pulled
- [ ] Python 3.9+
- [ ] Dependencies installed
- [ ] .env configured
- [ ] App started
- [ ] Excel file ready

---

**Ready to go!** 🚀
