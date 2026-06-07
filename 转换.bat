@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
echo =============================================
echo   AI2DOCX — AI 内容一键转 Word
echo   正在从剪贴板读取...
echo =============================================
echo.

:: 用 PowerShell 获取可靠日期（YYYYMMDD）
for /f %%i in ('powershell -c "Get-Date -Format yyyyMMdd"') do set TODAY=%%i

C:\Anaconda3\python "%~dp0ai2docx.py" -o "AI内容导出_%TODAY%.docx"

echo.
if %errorlevel% equ 0 (
    echo 按任意键退出...
    pause >nul
) else (
    echo 按任意键关闭...
    pause >nul
)
