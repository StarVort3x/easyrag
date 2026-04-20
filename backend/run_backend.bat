@echo off
REM EasyRAG 后端启动脚本

echo ========================================
echo EasyRAG Backend Startup
echo ========================================
echo.

REM 获取当前目录
cd /d "%~dp0"

REM 检查Python环境
echo Checking Python environment...
python --version

REM 安装依赖（如果需要）
echo.
echo Installing dependencies...
python -m pip install -q -r requirements.txt

REM 启动后端服务
echo.
echo Starting backend service...
echo Server will run on http://0.0.0.0:8000
echo.
echo Press Ctrl+C to stop the server
echo.

python start.py

pause
