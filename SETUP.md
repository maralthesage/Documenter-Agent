# Installation and Setup Guide

## Prerequisites

- Python 3.9 or higher
- Ollama installed (https://ollama.ai)
- A German-friendly LLM (Mistral, Mixtral, or Neural-Chat)
- At least 2GB of free memory
- macOS, Linux, or Windows with WSL

## Step 1: Clone/Set Up the Project

```bash
cd /path/to/Documenter-Agent
```

## Step 2: Install Python Dependencies

### On macOS/Linux:

```bash
# Create virtual environment (optional but recommended)
python3 -m venv venv
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### On Windows:

```bash
# Create virtual environment (optional but recommended)
python -m venv venv
venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

## Step 3: Set Up Ollama

### Install Ollama

Visit https://ollama.ai and download for your OS.

### Start Ollama Service

**macOS/Linux:**
```bash
ollama serve
```

This will start Ollama on `http://localhost:11434`

**Windows:**
- Open Ollama application
- It runs in the background and is automatically available

### Pull a Language Model

In a new terminal (while Ollama is running):

**Option 1: Mistral (Recommended - Lightweight)**
```bash
ollama pull mistral
```

**Option 2: Mixtral (More Powerful)**
```bash
ollama pull mixtral
```

**Option 3: Neural-Chat (Good for German)**
```bash
ollama pull neural-chat
```

**Option 4: Llama 2**
```bash
ollama pull llama2
```

Check available models:
```bash
ollama list
```

## Step 4: Configure Environment

```bash
# Copy example configuration
cp .env.example .env

# Edit .env if needed (default values work fine)
# nano .env  # or open in your text editor
```

The default configuration:
```
OLLAMA_BASE_URL=http://localhost:11434
OLLAMA_MODEL=mistral
```

## Step 5: Run the Application

### On macOS/Linux:

```bash
# Make script executable
chmod +x run.sh

# Run
./run.sh
```

### On Windows:

```bash
run.bat
```

### Manual run:

```bash
streamlit run app.py
```

This will:
1. Check that Ollama is running
2. Check that the model is available
3. Start the Streamlit web interface
4. Open browser to `http://localhost:8501`

## Usage

1. **Open the web interface** in your browser (automatically opens)

2. **Upload Excel file**:
   - Click "Upload File" or
   - Enter file path directly

3. **Click "Start Analysis"**

4. **Wait for completion** (may take 1-5 minutes depending on file size)

5. **Download documentation** as markdown

6. **View preview** in the app or in any markdown viewer

## Troubleshooting

### "Cannot connect to Ollama"

```bash
# Make sure Ollama is running
ollama serve

# Or check if it's already running
curl http://localhost:11434/api/tags
```

### "Model not found"

```bash
# Pull the model
ollama pull mistral

# Or check installed models
ollama list
```

### "Connection refused"

- Check firewall settings
- Ensure Ollama port (11434) is not blocked
- Try changing `OLLAMA_BASE_URL` in `.env`

### "Out of memory"

- Close other applications
- Use a smaller model: `ollama pull neural-chat`
- Reduce `MAX_SAMPLE_ROWS` in `config/settings.py`

### "Streamlit not found"

```bash
# Make sure virtual environment is activated
source venv/bin/activate  # macOS/Linux
venv\Scripts\activate     # Windows

# Install dependencies again
pip install -r requirements.txt
```

### "Excel file not readable"

- Ensure file is not open in Excel or other application
- Check file permissions
- Try a different Excel file

## Performance Tips

1. **Use Mistral for speed** - lighter model, faster processing
2. **Use Mixtral for quality** - more detailed analysis
3. **Reduce sample size** in `config/settings.py` for large files
4. **Allocate more RAM** to Ollama if available

## Next Steps

- Try with a sample Excel file
- Customize settings in `config/settings.py`
- Modify prompts in `src/llm_integration.py`
- Customize output format in `src/documentation_generator.py`

## Getting Help

Check the README.md for project overview and features.
