@echo off
:: 设置控制台标题
title EasyRAG 启动器

:: 设置颜色
color 0A

echo =======================================
echo    EasyRAG 知识库系统 - 一键启动程序
echo =======================================
echo.

:: 设置项目路径
set "PROJECT_ROOT=%~dp0"
set "BACKEND_DIR=%PROJECT_ROOT%backend"
set "FRONTEND_DIR=%PROJECT_ROOT%ruoyi-ui"
set "CONDA_ENV=easyrag"

:: 检查conda是否可用
where conda >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo [错误] 未检测到Anaconda/Miniconda，请先安装
    pause
    exit /b 1
)

:: 检查conda环境是否存在
call conda env list | findstr /c:"%CONDA_ENV%" >nul
if %ERRORLEVEL% neq 0 (
    echo [错误] 未找到conda环境: %CONDA_ENV%
    echo 请先创建环境: conda create -n %CONDA_ENV% python=3.9
    pause
    exit /b 1
)

:: 检查目录是否存在
if not exist "%BACKEND_DIR%" (
    echo [错误] 后端目录不存在: %BACKEND_DIR%
    pause
    exit /b 1
)

if not exist "%FRONTEND_DIR%" (
    echo [错误] 前端目录不存在: %FRONTEND_DIR%
    pause
    exit /b 1
)

:: 创建符号链接（如果不存在）
if not exist "%BACKEND_DIR%\easy-local-rag-main" (
    echo [信息] 正在创建符号链接...
    mklink /D "%BACKEND_DIR%\easy-local-rag-main" "%PROJECT_ROOT%easy-local-rag-main"
)

echo.
echo [信息] 正在启动服务...
echo.

:: 启动后端服务
start "EasyRAG Backend" /D"%BACKEND_DIR%" cmd /k "call conda activate %CONDA_ENV% && python start.py"

:: 等待后端服务启动
timeout /t 5 >nul

:: 启动前端服务
start "EasyRAG Frontend" /D"%FRONTEND_DIR%" cmd /k "npm run serve"

:: 等待前端服务启动
timeout /t 10 >nul

:: 打开浏览器
start http://localhost:8080

echo.
echo =======================================
echo    服务启动完成！
echo    前端地址: http://localhost:8080
echo    后端API:  http://localhost:8000
echo    API文档:  http://localhost:8000/docs
echo =======================================
echo.
echo [提示] 按任意键关闭所有服务...
pause >nul

:: 关闭所有服务
taskkill /FI "WINDOWTITLE eq EasyRAG Backend*" /F >nul 2>&1
taskkill /FI "WINDOWTITLE eq EasyRAG Frontend*" /F >nul 2>&1

exit /b 0