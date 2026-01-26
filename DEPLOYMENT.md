# Deployment & Advanced Configuration

## Docker Deployment (Optional)

Create a `Dockerfile`:

```dockerfile
FROM python:3.11-slim

WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    libssl-dev \
    && rm -rf /var/lib/apt/lists/*

# Copy requirements
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application
COPY . .

# Expose Streamlit port
EXPOSE 8501

# Health check
HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health

# Run application
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

Build and run:

```bash
docker build -t excel-doc-agent .
docker run -p 8501:8501 -v /path/to/files:/app/data excel-doc-agent
```

---

## Production Deployment

### 1. Server Setup

```bash
# Install Python and dependencies
sudo apt-get install python3.11 python3.11-venv

# Clone/download project
git clone <repo> /opt/excel-doc-agent
cd /opt/excel-doc-agent

# Create virtual environment
python3.11 -m venv venv
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### 2. Ollama Setup

```bash
# Install Ollama on separate or same server
curl https://ollama.ai/install.sh | sh

# Create systemd service
sudo systemctl enable ollama
sudo systemctl start ollama

# Pull models
ollama pull mistral
ollama pull mixtral
```

### 3. Streamlit Configuration

Create `.streamlit/config.toml`:

```toml
[server]
port = 8501
headless = true
runOnSave = true

[logger]
level = "info"

[client]
showErrorDetails = false

[theme]
primaryColor = "#4472C4"
backgroundColor = "#FFFFFF"
secondaryBackgroundColor = "#F0F2F6"
textColor = "#262730"
```

### 4. Nginx Reverse Proxy

```nginx
server {
    listen 80;
    server_name excel-doc.example.com;

    location / {
        proxy_pass http://localhost:8501;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
```

### 5. Systemd Service

Create `/etc/systemd/system/excel-doc-agent.service`:

```ini
[Unit]
Description=Excel Documentation AI Agent
After=network.target ollama.service

[Service]
Type=simple
User=www-data
WorkingDirectory=/opt/excel-doc-agent
Environment="PATH=/opt/excel-doc-agent/venv/bin"
ExecStart=/opt/excel-doc-agent/venv/bin/streamlit run app.py
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
```

Enable and start:

```bash
sudo systemctl enable excel-doc-agent
sudo systemctl start excel-doc-agent
sudo systemctl status excel-doc-agent
```

---

## Performance Optimization

### 1. Model Selection

| Use Case | Model | Memory | Speed |
|----------|-------|--------|-------|
| Quick analysis | mistral | 4GB | Fast |
| Detailed analysis | mixtral | 8GB | Medium |
| Best quality | neural-chat | 6GB | Medium |

### 2. Caching

Add to `llm_integration.py`:

```python
from functools import lru_cache

@lru_cache(maxsize=128)
def analyze_column_name(self, column_name: str) -> str:
    # Cached analysis
    pass
```

### 3. Batch Processing

```python
# Process multiple files
files = ['file1.xlsx', 'file2.xlsx', 'file3.xlsx']
for file in files:
    analyzer = ExcelAnalyzer(file)
    doc_gen = DocumentationGenerator(analyzer, llm)
    doc_gen.save_to_file(f'./output/{Path(file).stem}.md')
```

### 4. Resource Limits

In `.env`:

```
MAX_ROWS_TO_ANALYZE=1000
MAX_COLUMNS_TO_ANALYZE=50
ANALYSIS_TIMEOUT=600  # 10 minutes
```

---

## Security Considerations

### 1. File Upload Validation

```python
MAX_FILE_SIZE = 100 * 1024 * 1024  # 100MB
ALLOWED_EXTENSIONS = {'.xlsx', '.xls', '.xlsm'}

def validate_upload(file_path):
    # Check size
    if os.path.getsize(file_path) > MAX_FILE_SIZE:
        raise ValueError("File too large")
    
    # Check extension
    if Path(file_path).suffix not in ALLOWED_EXTENSIONS:
        raise ValueError("Invalid file type")
```

### 2. Sensitive Data Handling

```python
# Mask sensitive data in output
SENSITIVE_PATTERNS = [
    r'[0-9]{3}-[0-9]{2}-[0-9]{4}',  # SSN
    r'[0-9]{16}',  # Credit card
    r'[^@]+@[^@]+\.[^@]+',  # Email
]

def sanitize_output(text):
    for pattern in SENSITIVE_PATTERNS:
        text = re.sub(pattern, '[REDACTED]', text)
    return text
```

### 3. Access Control

Use Streamlit authentication:

```python
import streamlit as st
from streamlit_authenticator import Authenticate

authenticator = Authenticate(
    names=['John Smith', 'Jane Doe'],
    usernames=['jsmith', 'jdoe'],
    passwords=['password123', 'password456']
)

name, authentication_status, username = authenticator.login()

if authentication_status:
    # Show app
    main()
elif authentication_status == False:
    st.error('Invalid credentials')
elif authentication_status == None:
    st.warning('Enter credentials')
```

---

## Monitoring

### 1. Logging

Add to `config/settings.py`:

```python
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/excel_doc_agent.log'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)
```

### 2. Health Checks

```python
def health_check():
    """Check if all components are healthy"""
    checks = {
        'ollama': check_ollama(),
        'disk_space': check_disk_space(),
        'memory': check_memory(),
    }
    return all(checks.values())
```

### 3. Metrics Collection

```python
import time
from prometheus_client import Counter, Histogram

analysis_counter = Counter('analysis_total', 'Total analyses')
analysis_duration = Histogram('analysis_duration_seconds', 'Analysis duration')

@analysis_duration.time()
def analyze_with_metrics(file_path):
    analyzer = ExcelAnalyzer(file_path)
    result = analyzer.analyze()
    analysis_counter.inc()
    return result
```

---

## Troubleshooting Production Issues

### Issue: Slow Performance

```bash
# Check Ollama memory usage
ollama info

# Monitor system resources
top
free -h
df -h

# Reduce analysis scope
# Edit MAX_SAMPLE_ROWS in config/settings.py
```

### Issue: Memory Leaks

```python
# Add garbage collection
import gc

def analyze_with_cleanup(file_path):
    try:
        result = ExcelAnalyzer(file_path).analyze()
        return result
    finally:
        gc.collect()
```

### Issue: Connection Timeouts

```python
# Increase timeout in llm_integration.py
response = requests.post(
    url,
    timeout=600,  # Increase from 300
    ...
)
```

---

## Scaling

### 1. Load Balancing

Use multiple instances behind load balancer:

```nginx
upstream excel_doc_backend {
    server localhost:8501;
    server localhost:8502;
    server localhost:8503;
}

server {
    listen 80;
    server_name excel-doc.example.com;
    
    location / {
        proxy_pass http://excel_doc_backend;
    }
}
```

### 2. Distributed Processing

```python
from celery import Celery

app = Celery('excel_doc_agent')

@app.task
def analyze_file(file_path):
    analyzer = ExcelAnalyzer(file_path)
    doc_gen = DocumentationGenerator(analyzer, llm)
    doc_gen.save_to_file(...)
```

### 3. Database for Results

```python
import sqlite3

def save_analysis_result(file_name, analysis_json):
    conn = sqlite3.connect('analyses.db')
    cursor = conn.cursor()
    cursor.execute(
        'INSERT INTO analyses (filename, result) VALUES (?, ?)',
        (file_name, analysis_json)
    )
    conn.commit()
```

---

## Backup & Recovery

### 1. Backup Configuration

```bash
#!/bin/bash
# Backup.sh

BACKUP_DIR="/backups/excel-doc-agent"
SOURCE="/opt/excel-doc-agent"

# Create backup
tar -czf "$BACKUP_DIR/backup-$(date +%Y%m%d).tar.gz" "$SOURCE"

# Keep last 30 days
find "$BACKUP_DIR" -name "backup-*.tar.gz" -mtime +30 -delete
```

### 2. Database Backup

```bash
# Backup analyses database
sqlite3 /opt/excel-doc-agent/analyses.db ".backup '/backups/analyses-$(date +%Y%m%d).db'"
```

---

## Maintenance

### Regular Tasks

```bash
# Update dependencies
pip install --upgrade -r requirements.txt

# Clean up old outputs
find ./output -name "*.md" -mtime +30 -delete

# Update models
ollama pull mistral
ollama pull mixtral

# Check logs
tail -f logs/excel_doc_agent.log
```

### Model Updates

```bash
# Pull latest version of model
ollama pull mistral:latest

# Roll back if needed
# Models are stored in ~/.ollama/models/
```

---

## License

MIT License - See LICENSE file for details
