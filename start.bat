@echo off
chcp 65001 >nul
title 发货单处理工具 v2.0 (Web 版)

echo ==============================================
echo    发货单处理工具 v2.0 (Web 版) 启动器
echo ==============================================
echo.

REM 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未检测到 Python，请先安装 Python 3.9+
    echo        下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

REM 安装依赖（首次会自动安装）
echo [步骤 1/2] 检查并安装依赖...
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --quiet
if errorlevel 1 (
    echo [警告] 清华源安装失败，尝试默认源...
    pip install -r requirements.txt --quiet
)

echo.
echo [步骤 2/2] 启动 Web 服务...
echo.
echo ==============================================
echo    浏览器会自动打开 http://localhost:8501
echo    如未自动打开，请手动访问该地址
echo    按 Ctrl+C 可关闭服务
echo ==============================================
echo.

streamlit run app.py

pause
