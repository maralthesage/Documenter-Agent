"""
LLM Integration Module
Handles communication with Ollama through LangChain
"""

import requests
from langchain.llms.base import LLM
from langchain.callbacks.manager import CallbackManagerForLLMRun
from typing import Any, List, Optional
from pydantic import Field
import os
from config.settings import OLLAMA_BASE_URL, OLLAMA_MODEL


class OllamaLLM(LLM):
    """Custom LangChain LLM wrapper for Ollama"""

    model: str = Field(default=OLLAMA_MODEL)
    base_url: str = Field(default=OLLAMA_BASE_URL)
    temperature: float = Field(default=0.7)
    top_p: float = Field(default=0.9)

    @property
    def _llm_type(self) -> str:
        return "ollama"

    def _call(
        self,
        prompt: str,
        stop: Optional[List[str]] = None,
        run_manager: Optional[CallbackManagerForLLMRun] = None,
        **kwargs: Any,
    ) -> str:
        """Call the Ollama API"""
        try:
            payload = {
                "model": str(self.model),
                "prompt": str(prompt),
                "temperature": float(self.temperature),
                "top_p": float(self.top_p),
                "stream": False,
            }
            response = requests.post(
                f"{self.base_url}/api/generate",
                json=payload,
                timeout=300,
            )
            response.raise_for_status()
            result = response.json()
            return result.get("response", "").strip()
        except requests.exceptions.ConnectionError:
            raise ConnectionError(
                f"Cannot connect to Ollama at {self.base_url}. "
                "Make sure Ollama is running with: ollama serve"
            )
        except Exception as e:
            raise Exception(f"Error calling Ollama: {str(e)}")


class LLMIntegration:
    """Main class for LLM integration"""

    def __init__(self, model: Optional[str] = None, base_url: Optional[str] = None):
        self.model = str(model or OLLAMA_MODEL)
        self.base_url = str(base_url or OLLAMA_BASE_URL)
        self.llm = OllamaLLM(model=self.model, base_url=self.base_url)
        self._verify_connection()

    def _verify_connection(self):
        """Verify that Ollama is running and model is available"""
        try:
            response = requests.get(f"{self.base_url}/api/tags", timeout=10)
            response.raise_for_status()
            models = response.json().get("models", [])
            model_names = [m["name"] for m in models]

            if not any(self.model in name for name in model_names):
                raise ValueError(
                    f"Model '{self.model}' not found in Ollama. "
                    f"Available models: {model_names}. "
                    f"Pull the model with: ollama pull {self.model}"
                )
        except requests.exceptions.ConnectionError:
            raise ConnectionError(
                f"Cannot connect to Ollama at {self.base_url}. "
                "Make sure Ollama is running with: ollama serve"
            )

    def analyze_column_name(self, column_name: str, context: str = "") -> str:
        """
        Use LLM to infer meaning from column name and context
        """
        prompt = f"""Analyze the following Excel column name and provide a brief explanation of what data it likely contains.

Column Name: {column_name}
Context: {context}

Provide a concise explanation (1-2 sentences) of what this column likely represents. Be practical and consider common business contexts.
Write in English and avoid using the word 'Anzahl'."""

        return self.llm.predict(text=prompt)

    def analyze_formula(self, formula: str, column_name: str, context: str = "") -> str:
        """
        Use LLM to explain what a formula does
        """
        prompt = f"""Analyze the following Excel formula and explain what it does in plain language.

Column Name: {column_name}
Formula: {formula}
Context: {context}

Provide a clear, concise explanation (2-3 sentences) of:
1. What the formula calculates
2. What data it uses as input
3. What the result represents
Write in English and avoid using the word 'Anzahl'."""

        return self.llm.predict(text=prompt)

    def infer_data_meaning(self, column_name: str, data_type: str, samples: List[str]) -> str:
        """
        Use LLM to infer the meaning of a column based on samples
        """
        # Ensure all inputs are strings and handle None values
        column_name = str(column_name) if column_name is not None else "Unknown"
        data_type = str(data_type) if data_type is not None else "unknown"
        
        # Convert all samples to strings, filtering out None values
        safe_samples = []
        for s in samples[:5]:
            if s is not None:
                safe_samples.append(str(s)[:50])  # Limit to 50 chars per sample
        
        samples_str = ", ".join(safe_samples) if safe_samples else "(No data)"

        prompt = f"""Based on the column name, data type, and sample values, infer what this Excel column represents.

Column Name: {column_name}
Data Type: {data_type}
Sample Values: {samples_str}

Provide a detailed but concise explanation of:
1. What this column represents
2. What the data type tells us
3. Any patterns or characteristics observed in the samples
4. How this might be used in the context of the spreadsheet
Write in English, and do not use the word 'Anzahl'. Use the column name exactly as given."""

        return self.llm.predict(text=prompt)

    def explain_formatting_significance(self, column_name: str, formatting: dict) -> str:
        """
        Use LLM to explain why certain cells might have specific formatting
        """
        # Ensure column_name is a string
        column_name = str(column_name) if column_name is not None else "Unknown"
        
        # Convert all formatting values to strings to ensure JSON serialization
        safe_formatting = {}
        for k, v in formatting.items():
            if v is not None:
                safe_formatting[str(k)] = str(v)
        
        formatting_str = ", ".join(f"{k}: {v}" for k, v in safe_formatting.items())

        prompt = f"""In Excel, formatting (colors, styles) often indicates something meaningful. 
Analyze what the following formatting on a column might signify.

Column Name: {column_name}
Formatting: {formatting_str}

Consider:
1. Is this formatting grouping related data?
2. Does the color indicate a category or status?
3. Could it indicate importance or a data quality issue?
4. What business logic might explain this formatting?

Provide your analysis in 2-3 sentences.
Write in English and avoid using the word 'Anzahl'."""

        return self.llm.predict(text=prompt)

    def generate_comprehensive_documentation(
        self, columns_analysis: List[dict], sheet_name: str = "Sheet"
    ) -> str:
        """
        Use LLM to generate comprehensive documentation for the entire sheet
        """
        # Prepare column summaries with safe type conversion
        col_summaries = []
        for col in columns_analysis:
            col_name = str(col.get('name', 'Unknown'))
            col_type = str(col.get('data_type', 'unknown'))
            col_desc = str(col.get('description', 'No description'))
            
            summary = f"- **{col_name}** ({col_type}): {col_desc}"
            col_summaries.append(summary)

        columns_text = "\n".join(col_summaries)
        sheet_name = str(sheet_name) if sheet_name is not None else "Sheet"

        prompt = f"""Create a comprehensive overview documentation for an Excel spreadsheet.

Sheet Name: {sheet_name}

Columns:
{columns_text}

Write a professional markdown-formatted overview (5-7 sentences) that:
1. Describes the purpose of this sheet
2. Explains the overall data structure
3. Highlights key relationships between columns
4. Suggests how this data might be used
5. Notes any data quality observations

Format it as a professional business document.
Write in English and avoid using the word 'Anzahl'."""

        return self.llm.predict(text=prompt)

    def generate_markdown_summary(self, analysis_report: dict) -> str:
        """
        Generate a markdown summary of the analysis
        """
        columns = analysis_report.get("columns", [])
        col_list = "\n".join(
            [f"- {col['name']} ({col['data_type']})" for col in columns]
        )

        prompt = f"""Create a brief markdown summary (3-4 bullets) of this Excel data structure:

Sheet: {analysis_report.get('sheet_name', 'Unknown')}
Dimensions: {analysis_report.get('dimensions', {})}

Columns:
{col_list}

Provide a concise markdown summary that captures the essence of what this spreadsheet contains.
Write in English and avoid using the word 'Anzahl'."""

        return self.llm.predict(text=prompt)


def get_llm_integration(
    model: Optional[str] = None, base_url: Optional[str] = None
) -> LLMIntegration:
    """Factory function to get or create LLM integration"""
    return LLMIntegration(model=model, base_url=base_url)
