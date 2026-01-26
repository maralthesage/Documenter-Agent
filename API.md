# API Documentation

## Module: `excel_analyzer.py`

### Class: `ExcelAnalyzer`

Handles all Excel file parsing and analysis.

#### Methods

##### `__init__(file_path: str)`
Initialize the analyzer with an Excel file path.

```python
analyzer = ExcelAnalyzer('/path/to/file.xlsx')
```

##### `analyze() -> Dict[str, Any]`
Run complete analysis and return results.

```python
report = analyzer.analyze()
# Returns: {
#   'file_path': str,
#   'sheet_name': str,
#   'dimensions': {'rows': int, 'columns': int},
#   'columns': [ColumnInfo dict, ...]
# }
```

##### `get_column_info() -> List[ColumnInfo]`
Get the list of analyzed column information objects.

```python
columns = analyzer.get_column_info()
for col in columns:
    print(f"{col.name}: {col.data_type}")
```

##### `extract_shared_formatting_groups() -> List[List[str]]`
Identify columns with shared formatting (potential data grouping).

```python
groups = analyzer.extract_shared_formatting_groups()
# Returns: [['col1', 'col2'], ['col3', 'col4']]
```

### Dataclass: `ColumnInfo`

Stores information about a column.

```python
@dataclass
class ColumnInfo:
    name: str
    index: int
    data_type: str  # 'int', 'float', 'string', 'datetime', etc.
    non_null_count: int
    unique_values: int
    has_formula: bool
    formulas: List[str]
    sample_values: List[Any]
    min_value: Optional[Any]
    max_value: Optional[Any]
    fill_color: Optional[str]
    font_color: Optional[str]
    contains_link: bool
    link_type: Optional[str]
    numeric_range: Optional[Tuple[float, float]]
```

---

## Module: `llm_integration.py`

### Class: `LLMIntegration`

Handles communication with Ollama LLM through LangChain.

#### Methods

##### `__init__(model: Optional[str] = None, base_url: Optional[str] = None)`
Initialize LLM integration with Ollama.

```python
llm = LLMIntegration(model='mistral')
```

##### `analyze_column_name(column_name: str, context: str = "") -> str`
Use LLM to infer meaning from a column name.

```python
description = llm.analyze_column_name('customer_id', 'sales data')
# Returns: "A unique identifier for each customer..."
```

##### `analyze_formula(formula: str, column_name: str, context: str = "") -> str`
Explain what a formula does.

```python
explanation = llm.analyze_formula('=VLOOKUP(B2,Sheet2!$A:$C,3,0)', 'Price')
# Returns: "This formula looks up the value in B2..."
```

##### `infer_data_meaning(column_name: str, data_type: str, samples: List[str]) -> str`
Infer column meaning from name, type, and samples.

```python
meaning = llm.infer_data_meaning('revenue', 'float', ['1500.50', '2000.00', '1200.25'])
# Returns: "This column represents revenue amounts..."
```

##### `explain_formatting_significance(column_name: str, formatting: dict) -> str`
Explain what special formatting might mean.

```python
explanation = llm.explain_formatting_significance('status', {'fill_color': '00FF00'})
# Returns: "The green color likely indicates positive or completed status..."
```

##### `generate_comprehensive_documentation(columns_analysis: List[dict], sheet_name: str = "Sheet") -> str`
Generate overall sheet documentation.

```python
overview = llm.generate_comprehensive_documentation(analysis_data, 'Sales')
# Returns: "This spreadsheet tracks daily sales transactions..."
```

### Class: `OllamaLLM`

Custom LangChain LLM wrapper for Ollama.

```python
llm_model = OllamaLLM(model='mistral')
response = llm_model.predict(text='What is this?')
```

---

## Module: `documentation_generator.py`

### Class: `DocumentationGenerator`

Generates markdown documentation from Excel analysis.

#### Methods

##### `__init__(analyzer: ExcelAnalyzer, llm: LLMIntegration)`
Initialize the documentation generator.

```python
doc_gen = DocumentationGenerator(analyzer, llm)
```

##### `generate_documentation() -> str`
Generate complete markdown documentation.

```python
markdown = doc_gen.generate_documentation()
# Returns: Complete markdown document as string
```

##### `save_to_file(output_path: str) -> str`
Generate documentation and save to file.

```python
file_path = doc_gen.save_to_file('./output/documentation.md')
```

#### Generated Markdown Structure

The generated documentation includes:

1. **Header** - File info, dimensions, generation timestamp
2. **Overview** - AI-generated overview of the data
3. **Table of Contents** - Navigable TOC with anchors
4. **Data Structure** - Dimensions and data type distribution
5. **Column Descriptions** - Detailed info for each column:
   - Data type
   - Non-null count
   - Unique values
   - Sample values
   - Formatting info
   - Formulas (if present)
   - AI-generated description
6. **Data Relationships** - Computed columns, linked data, formatting groups
7. **Formatting Analysis** - What special formatting might mean
8. **Formulas** - Detailed formula explanations
9. **Footer** - Metadata and generation info

---

## Module: `utils.py`

### Functions

#### `validate_excel_file(file_path: str) -> Tuple[bool, str]`
Validate an Excel file.

```python
is_valid, message = validate_excel_file('/path/to/file.xlsx')
if is_valid:
    print("File is valid")
```

#### `get_output_path(input_file: str, output_dir: Optional[str] = None) -> str`
Generate output path for documentation file.

```python
output = get_output_path('./data.xlsx', './output')
# Returns: './output/data_documentation_20240126_143022.md'
```

#### `truncate_string(s: str, max_length: int = 50, suffix: str = "...") -> str`
Truncate string to max length.

```python
short = truncate_string("This is a long string", 10)
# Returns: "This is..."
```

#### `clean_column_name(name: str) -> str`
Clean and normalize column names.

```python
clean = clean_column_name("Col-Name_123!")
# Returns: "Col-Name 123"
```

#### `create_markdown_table(headers: list, rows: list) -> str`
Create markdown table from data.

```python
table = create_markdown_table(
    ['Name', 'Age'],
    [['Alice', '30'], ['Bob', '25']]
)
# Returns: Markdown table string
```

### Class: `ProgressTracker`

Track progress of operations.

```python
tracker = ProgressTracker(total_steps=10)
for i in range(10):
    tracker.update(f"Step {i}")
    
progress = tracker.get_progress()  # Returns: 0-100
status = tracker.get_status()      # Returns: dict with progress info
```

---

## Module: `config/settings.py`

Configuration constants for the application.

```python
OLLAMA_BASE_URL = "http://localhost:11434"
OLLAMA_MODEL = "mistral"
MAX_SAMPLE_ROWS = 50
MAX_COLUMN_NAME_LENGTH = 100
CHUNK_SIZE = 4
MAX_FORMULA_LENGTH = 200
INCLUDE_DATA_SAMPLES = True
SAMPLE_SIZE = 10
```

---

## Streamlit App: `app.py`

Web interface for the application.

### Components

1. **Sidebar**
   - Configuration display
   - Ollama setup instructions
   - About section

2. **Main Content**
   - File upload/input
   - Analysis button
   - Progress tracking
   - Results display

3. **Output**
   - Documentation preview
   - Raw markdown view
   - Download button

### Functions

#### `initialize_session_state()`
Initialize Streamlit session state variables.

#### `main()`
Main application entry point.

---

## Usage Examples

### Basic Usage

```python
from src.excel_analyzer import ExcelAnalyzer
from src.llm_integration import get_llm_integration
from src.documentation_generator import DocumentationGenerator

# Analyze Excel file
analyzer = ExcelAnalyzer('sales_data.xlsx')
analysis = analyzer.analyze()

# Initialize LLM
llm = get_llm_integration()

# Generate documentation
doc_gen = DocumentationGenerator(analyzer, llm)
markdown = doc_gen.generate_documentation()

# Save to file
doc_gen.save_to_file('./output/sales_documentation.md')
```

### Custom Analysis

```python
# Get column information
columns = analyzer.get_column_info()

for col in columns:
    print(f"Column: {col.name}")
    print(f"  Type: {col.data_type}")
    print(f"  Has Formula: {col.has_formula}")
    if col.has_formula:
        print(f"  Formulas: {col.formulas}")
    
    # Get AI analysis
    description = llm.infer_data_meaning(
        col.name,
        col.data_type,
        col.sample_values
    )
    print(f"  Description: {description}")
```

### Advanced: Custom Documentation

```python
from src.documentation_generator import DocumentationGenerator

class CustomDocumentationGenerator(DocumentationGenerator):
    def _generate_overview(self) -> str:
        # Custom overview generation
        return "Custom overview..."

gen = CustomDocumentationGenerator(analyzer, llm)
custom_md = gen.generate_documentation()
```

---

## Error Handling

### Common Exceptions

```python
# Connection error
try:
    llm = get_llm_integration()
except ConnectionError as e:
    print(f"Cannot connect to Ollama: {e}")

# File validation error
is_valid, message = validate_excel_file(file_path)
if not is_valid:
    print(f"Invalid file: {message}")

# Analysis error
try:
    analysis = analyzer.analyze()
except Exception as e:
    print(f"Analysis failed: {e}")
```

---

## Performance Tips

1. **Use Mistral for speed** - Lightweight, fast processing
2. **Reduce sample size** for large files in settings
3. **Increase chunk size** for more parallel processing
4. **Cache results** to avoid re-analysis
5. **Use SSD storage** for better performance

---

## Testing

Run the test suite:

```bash
# Test with specific file
python test.py 'path/to/file.xlsx'

# Test analysis only
python test.py 'path/to/file.xlsx' --test-analysis-only

# Test LLM only
python test.py 'path/to/file.xlsx' --test-llm-only
```

---

## Version History

### v1.0.0
- Initial release
- Excel analysis
- LLM integration
- Documentation generation
- Streamlit UI

---

## License

MIT License
