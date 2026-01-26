@echo off
REM Excel Documentation AI Agent - Run Script for Windows
REM Starts the Streamlit application

echo.
echo ================================================================
echo Excel Documentation AI Agent
echo ================================================================
echo.

REM Check if .env file exists
if not exist .env (
    echo Creating .env file from .env.example...
    copy .env.example .env
    echo Done.
)

echo.
echo Checking prerequisites...
echo.

REM Check if Ollama is running
echo Checking Ollama connection...
powershell -Command "(Invoke-WebRequest -Uri 'http://localhost:11434/api/tags' -UseBasicParsing -ErrorAction SilentlyContinue) | Out-Null" 2>nul
if %errorlevel% equ 0 (
    echo OK - Connected to Ollama
) else (
    echo ERROR - Cannot connect to Ollama
    echo Please start Ollama with: ollama serve
    pause
    exit /b 1
)

echo.
echo Starting Streamlit application...
echo.
echo Open your browser to: http://localhost:8501
echo.

streamlit run app.py
