@echo off
echo Starting local HTTP server...
echo.
echo Choose an option:
echo 1. Node.js HTTP Server (recommended)
echo 2. Python HTTP Server
echo 3. Install and run Node.js HTTP Server
echo.
set /p choice=Enter your choice (1-3): 

if "%choice%"=="1" (
    echo Starting Node.js HTTP Server...
    npx http-server . -p 8000 -c-1 --cors
) else if "%choice%"=="2" (
    echo Starting Python HTTP Server...
    python -m http.server 8000
) else if "%choice%"=="3" (
    echo Installing Node.js HTTP Server...
    npm install -g http-server
    echo Starting Node.js HTTP Server...
    npx http-server . -p 8000 -c-1 --cors
) else (
    echo Invalid choice. Starting Node.js HTTP Server...
    npx http-server . -p 8000 -c-1 --cors
)

echo.
echo Server started! Open your browser and go to:
echo http://localhost:8000
echo.
echo Press Ctrl+C to stop the server
pause 