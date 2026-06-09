@echo off
setlocal EnableExtensions

cd /d "%~dp0"
title GRE Vocabulary Trainer - First-Time Setup

set APP_FILE=gre_vocab_trainer_pro_full.py
set DB_FILE=GRE_VOCAB_DATABASE_FINAL_CLEAN.xlsx
set USER_FILE=GRE_USER_DATA.xlsx

echo ==========================================
echo   GRE Vocabulary Trainer - First-Time Setup
echo ==========================================
echo.

echo Checking required files...
echo.

if not exist "%APP_FILE%" (
    echo Missing: %APP_FILE%
    echo Put this setup file in the same folder as your Python trainer.
    pause
    exit /b 1
)

if not exist "%DB_FILE%" (
    echo Missing: %DB_FILE%
    echo Put the vocabulary database in this folder.
    pause
    exit /b 1
)

if not exist "%USER_FILE%" (
    echo Missing: %USER_FILE%
    echo Put GRE_USER_DATA.xlsx in this folder.
    pause
    exit /b 1
)

echo Required files found.
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set PY_CMD=py
    goto python_ok
)

where python >nul 2>nul
if %errorlevel%==0 (
    set PY_CMD=python
    goto python_ok
)

echo Python is not installed or not added to PATH.
echo.
echo Install Python 3.11 or newer from python.org
echo During installation, check: Add Python to PATH
echo.
pause
exit /b 1

:python_ok
echo Python found: %PY_CMD%
echo.

if not exist ".venv" (
    echo Creating virtual environment...
    %PY_CMD% -m venv .venv
    if %errorlevel% neq 0 (
        echo Failed to create virtual environment.
        pause
        exit /b 1
    )
) else (
    echo Virtual environment already exists.
)

echo.
echo Activating virtual environment...
call ".venv\Scripts\activate.bat"
if %errorlevel% neq 0 (
    echo Failed to activate virtual environment.
    pause
    exit /b 1
)

echo.
echo Upgrading pip...
python -m pip install --upgrade pip

echo.
echo Installing required packages...
python -m pip install pandas openpyxl

if %errorlevel% neq 0 (
    echo.
    echo Package installation failed.
    echo Check your internet connection and try again.
    pause
    exit /b 1
)

echo.
echo Creating launcher file: Open_GRE_Vocab_Trainer.bat

(
echo @echo off
echo cd /d "%%~dp0"
echo title GRE Vocabulary Trainer
echo.
echo if exist ".venv\Scripts\activate.bat" ^(
echo     call ".venv\Scripts\activate.bat"
echo ^)
echo.
echo if not exist "%APP_FILE%" ^(
echo     echo Python file not found: %APP_FILE%
echo     pause
echo     exit /b 1
echo ^)
echo.
echo python "%APP_FILE%"
echo.
echo if %%errorlevel%% neq 0 ^(
echo     echo.
echo     echo The app closed with an error.
echo     pause
echo ^)
) > "Open_GRE_Vocab_Trainer.bat"

echo.
echo ==========================================
echo Setup complete.
echo.
echo From now on, open the app using:
echo Open_GRE_Vocab_Trainer.bat
echo ==========================================
echo.

pause
endlocal