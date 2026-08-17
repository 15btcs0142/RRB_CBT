@echo off
title RRB CBT v1.13 - Exam Manager
color 0A
setlocal enabledelayedexpansion

:MENU
cls
echo.
echo  ============================================================
echo         RRB CBT v1.13 - School Exam System
echo  ============================================================
echo.
echo   SERVER CONTROLS
echo   ---------------
echo   [1]  Start Server (Network - All Students Can Access)
echo   [2]  Start Server (Localhost Only)
echo.
echo   OPEN IN BROWSER
echo   ---------------
echo   [3]  Open Admin Panel
echo   [4]  Open Teacher Portal
echo   [5]  Open Student Login
echo.
echo   SETUP ^& TOOLS
echo   ---------------
echo   [6]  Install / Update Dependencies
echo   [7]  Set AI API Keys (Gemini, DeepSeek, ChatGPT, Claude)
echo   [8]  Show My IP Address
echo   [9]  Backup Database
echo  [10]  View Logs
echo  [11]  Start RQ Redis Worker
echo.
echo   [0]  Exit
echo  ============================================================
echo.
set /p choice="  Enter your choice: "

if "%choice%"=="1"  goto SERVER_NETWORK
if "%choice%"=="2"  goto SERVER_LOCAL
if "%choice%"=="3"  goto OPEN_ADMIN
if "%choice%"=="4"  goto OPEN_TEACHER
if "%choice%"=="5"  goto OPEN_STUDENT
if "%choice%"=="6"  goto INSTALL_DEPS
if "%choice%"=="7"  goto SET_API_KEY
if "%choice%"=="8"  goto SHOW_IP
if "%choice%"=="9"  goto BACKUP_DB
if "%choice%"=="10" goto VIEW_LOGS
if "%choice%"=="11" goto START_RQ_WORKER
if "%choice%"=="0"  goto EXIT

echo   [!] Invalid choice. Please try again.
timeout /t 2 >nul
goto MENU

:SERVER_NETWORK
cls
echo.
echo  ============================================================
echo       Starting RRB CBT Server (Network Mode)
echo  ============================================================
echo.
echo   Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo   [ERROR] Python is not installed or not in PATH!
    echo   Please install Python 3.9+ from https://python.org
    echo   Make sure to check "Add Python to PATH" during install.
    pause
    goto MENU
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo   Found: %%v
echo   Checking packages...
pip show PyJWT >nul 2>&1
if errorlevel 1 (
    echo   Installing packages...
    pip install -r requirements.txt
) else (
    echo   Packages OK.
)
echo.
echo  ============================================================
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4 Address"') do (
    set RAW_IP=%%a
    set RAW_IP=!RAW_IP: =!
    echo   Network URL  ^: http://!RAW_IP!:5000
    goto IP_FOUND
)
:IP_FOUND
echo   Local URL    ^: http://127.0.0.1:5000
echo   Admin        ^: http://127.0.0.1:5000/admin
echo   Teacher      ^: http://127.0.0.1:5000/teacher
echo   Student      ^: http://127.0.0.1:5000/
echo  ============================================================
echo   Developed by Gaurav Shukla
echo.
echo   Press Ctrl+C to stop the server.
echo.
if exist ".api_key" (
    set /p SAVED_KEY=<.api_key
    set GEMINI_API_KEY=!SAVED_KEY!
    echo   Gemini API Key loaded.
    echo.
)
python app.py
goto MENU

:SERVER_LOCAL
cls
echo.
echo  ============================================================
echo       Starting RRB CBT Server (Localhost Mode)
echo  ============================================================
echo.
echo   Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo   [ERROR] Python is not installed or not in PATH!
    echo   Please install Python 3.9+ from https://python.org
    echo   Make sure to check "Add Python to PATH" during install.
    pause
    goto MENU
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo   Found: %%v
echo   Checking packages...
pip show PyJWT >nul 2>&1
if errorlevel 1 (
    echo   Installing packages...
    pip install -r requirements.txt
) else (
    echo   Packages OK.
)
echo.
echo  ============================================================
echo   Local URL    ^: http://127.0.0.1:5000
echo   Admin        ^: http://127.0.0.1:5000/admin
echo   Teacher      ^: http://127.0.0.1:5000/teacher
echo   Student      ^: http://127.0.0.1:5000/
echo  ============================================================
echo   Developed by Gaurav Shukla
echo.
echo   Press Ctrl+C to stop the server.
echo.
if exist ".api_key" (
    set /p SAVED_KEY=<.api_key
    set GEMINI_API_KEY=!SAVED_KEY!
    echo   Gemini API Key loaded.
    echo.
)
python app.py
goto MENU

:OPEN_ADMIN
cls
echo   Opening Admin Panel...
start http://127.0.0.1:5000/admin
timeout /t 2 >nul
goto MENU

:OPEN_TEACHER
cls
echo   Opening Teacher Portal...
start http://127.0.0.1:5000/teacher
timeout /t 2 >nul
goto MENU

:OPEN_STUDENT
cls
echo   Opening Student Login...
start http://127.0.0.1:5000/
timeout /t 2 >nul
goto MENU

:INSTALL_DEPS
cls
echo.
echo  ============================================================
echo       Installing / Updating Dependencies
echo  ============================================================
echo.
python --version >nul 2>&1
if errorlevel 1 (
    echo   [ERROR] Python not found!
    echo   Download from: https://python.org
    pause
    goto MENU
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do echo   Found: %%v
echo.
echo   Installing packages from requirements.txt...
echo   (This may take a few minutes the first time)
echo.
pip install -r requirements.txt
echo.
if errorlevel 1 (
    echo   [WARNING] Some packages failed.
    echo   Try manually: pip install Flask openpyxl weasyprint PyJWT python-dotenv
) else (
    echo   [SUCCESS] All packages installed successfully!
)
echo.
pause
goto MENU

:SET_API_KEY
cls
echo.
echo  ============================================================
echo       Set AI API Keys (Gemini, DeepSeek, ChatGPT, Claude)
echo  ============================================================
echo.
echo   This enables AI-powered test generation for teachers.
echo   Gemini:   https://aistudio.google.com/
echo   DeepSeek: https://platform.deepseek.com/
echo   OpenAI:   https://platform.openai.com/
echo   Claude:   https://console.anthropic.com/
echo.
if exist ".api_key" (
    set /p CURRENT_KEY=<.api_key
    echo   Current Gemini key: !CURRENT_KEY:~0,20!...
    echo.
)
set /p NEW_KEY="  Paste your Gemini API key (or press Enter to skip): "
if "!NEW_KEY!"=="" (
    echo   No changes made.
) else (
    echo !NEW_KEY!>.api_key
    set GEMINI_API_KEY=!NEW_KEY!
    echo GEMINI_API_KEY=!NEW_KEY!>apikey.env
    echo   [SUCCESS] Gemini API key saved to .api_key and apikey.env!
)
echo.
echo   NOTE: You can also edit apikey.env directly to add DEEPSEEK_API_KEY,
echo   OPENAI_API_KEY, and CLAUDE_API_KEY for multi-provider fallback.
echo.
pause
goto MENU

:SHOW_IP
cls
echo.
echo  ============================================================
echo       Your Network Information
echo  ============================================================
echo.
ipconfig | findstr /c:"IPv4 Address"
echo.
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /c:"IPv4 Address"') do (
    set RAW_IP=%%a
    set RAW_IP=!RAW_IP: =!
    echo   Share with students  ^: http://!RAW_IP!:5000
    echo   Admin URL            ^: http://!RAW_IP!:5000/admin
    echo   Teacher URL          ^: http://!RAW_IP!:5000/teacher
    goto SHOW_IP_DONE
)
:SHOW_IP_DONE
echo.
echo   NOTE: All devices must be on the same WiFi/LAN network.
echo.
pause
goto MENU

:BACKUP_DB
cls
echo.
echo  ============================================================
echo       Backup Database
echo  ============================================================
echo.
if not exist "database.db" (
    echo   [ERROR] database.db not found!
    pause
    goto MENU
)
if not exist "backups" mkdir backups
set TIMESTAMP=%date:~10,4%%date:~4,2%%date:~7,2%_%time:~0,2%%time:~3,2%
set TIMESTAMP=%TIMESTAMP: =0%
set BACKUP_FILE=backups\database_%TIMESTAMP%.db
copy "database.db" "%BACKUP_FILE%" >nul
if errorlevel 1 (
    echo   [ERROR] Backup failed!
) else (
    echo   [SUCCESS] Backup saved to: %BACKUP_FILE%
)
echo.
echo   Existing backups:
dir /b backups\*.db 2>nul || echo   (none yet)
echo.
pause
goto MENU

:VIEW_LOGS
cls
echo.
echo  ============================================================
echo       Server Logs
echo  ============================================================
echo.
if exist "logs\app.log" (
    echo   Last 30 lines:
    echo.
    powershell "Get-Content 'logs\app.log' -Tail 30"
) else (
    echo   No log file found.
    echo   Logs appear in the console window when server is running.
)
echo.
pause
goto MENU

:START_RQ_WORKER
cls
echo.
echo  ============================================================
echo       Starting RRB CBT RQ Redis Worker
echo  ============================================================
echo.
python rq_worker.py
pause
goto MENU

:EXIT
cls
echo.
echo   Thank you for using RRB CBT v1.13. Goodbye!
echo.
timeout /t 2 >nul
exit