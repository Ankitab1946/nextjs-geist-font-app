@echo off
echo === Fuzzy Column Matcher Setup and Run Script ===

:: Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Python is not installed. Please install Python and try again.
    pause
    exit /b 1
)

:: Check if pip is installed
pip --version >nul 2>&1
if errorlevel 1 (
    echo pip is not installed. Please install pip and try again.
    pause
    exit /b 1
)

:: Create virtual environment
echo Creating virtual environment...
python -m venv venv

:: Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate.bat

:: Install requirements
echo Installing requirements...
pip install -r requirements.txt

:: Run tests
echo Running tests...
python test_matcher.py

:: Start Streamlit app
echo Starting Streamlit app...
set PYTHONPATH=%PYTHONPATH%;%CD%
streamlit run --server.port 8000 python_backend/streamlit_app.py

:: Note: The script will keep running until the Streamlit app is closed
:: The virtual environment will remain active until you close the window
pause
