@echo off
chcp 65001 >nul
echo =========================================
echo   KiKi的摸魚系統 - Flask 服務器啟動腳本
echo =========================================
echo.
echo 系統說明:
echo   首頁 (N8N 控制台):  http://localhost:5000/
echo   郵件工具頁:         http://localhost:5000/email-tool
echo.
echo 服務端口配置:
echo   Flask 服務:         端口 5000
echo   N8N 數據拉取服務:   端口 8080 (需要另外啟動)

echo.
echo =========================================
echo.

D:
cd D:\PythonCode\KiKiAutoPullDataScript

REM 檢查是否已安裝依賴
echo 正在檢查 Python 環境...
python --version
if %errorlevel% neq 0 (
    echo ? 錯誤: 找不到 Python，請先安裝 Python
    pause
    exit /b 1
)

echo.
echo 正在檢查依賴套件...

REM 檢查 Flask 是否安裝
python -c "import flask" 2>nul
if %errorlevel% neq 0 (
    echo ?? Flask 未安裝，正在安裝依賴...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo ? 依賴安裝失敗！
        pause
        exit /b 1
    )
)

echo ? 依賴檢查完成
echo.
echo =========================================
echo ?? 正在啟動 Flask 服務器...
echo =========================================
echo.
echo 請在瀏覽器中訪問: http://localhost:5000
echo 按 Ctrl+C 可停止服務器
echo.

REM 啟動 Flask 服務器
python app.py

pause
