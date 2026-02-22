@echo off
chcp 65001 >nul
echo ==========================================
echo    DeepExtract 启动脚本
echo ==========================================
echo.

REM 获取脚本所在目录
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo [1/3] 检查 Python 环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8+
    pause
    exit /b 1
)
echo [OK] Python 已安装

echo.
echo [2/3] 安装依赖...
if exist "backend\requirements.txt" (
    pip install -r backend\requirements.txt -q
    echo [OK] 依赖安装完成
) else (
    echo [警告] 未找到 requirements.txt
)

echo.
echo [3/3] 启动服务...
echo.
echo ==========================================
echo   访问地址: http://localhost:5000
echo ==========================================
echo.

REM 启动 Flask 服务
cd backend
python app.py

pause
