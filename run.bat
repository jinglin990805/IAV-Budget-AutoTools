@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

rem Prefer an interpreter that has the workbench dependencies installed.
set "PYTHON_CMD=python"
if exist ".paddle-venv\Scripts\python.exe" (
    ".paddle-venv\Scripts\python.exe" -c "import flask, openpyxl, pdfplumber, pypdfium2, PIL" >nul 2>&1
    if not errorlevel 1 set "PYTHON_CMD=.paddle-venv\Scripts\python.exe"
)
if "%PYTHON_CMD%"=="python" (
    python -c "import flask, openpyxl, pdfplumber, pypdfium2, PIL" >nul 2>&1
    if errorlevel 1 (
        py -3.12 -c "import flask, openpyxl, pdfplumber, pypdfium2, PIL" >nul 2>&1
        if not errorlevel 1 set "PYTHON_CMD=py -3.12"
    )
)

:menu
cls
echo ============================================
echo   IAV TV 项目管理工作台
echo ============================================
echo   [1] PO 工时规划 (V1)
echo   [2] PO 工时规划 (V2 - 分析报告)
echo   [3] PO 工时规划 (V2 - 预览差异)
echo   [4] PO 工时规划 (V2 - 回写Budget)
echo   [5] Budget 月度更新
echo   [6] 差旅汇总
echo   [7] PO/GR 状态追踪
echo   [8] 运行全部工具
echo   [W] 启动工作台 (浏览器)
echo   [0] 退出
echo ============================================
set /p choice=请选择 (0-8/W):

if "%choice%"=="1" goto run_po_v1
if "%choice%"=="2" goto run_po_v2
if "%choice%"=="3" goto run_po_preview
if "%choice%"=="4" goto run_po_write
if "%choice%"=="5" goto run_budget
if "%choice%"=="6" goto run_travel
if "%choice%"=="7" goto run_pogr
if "%choice%"=="8" goto run_all
if /i "%choice%"=="w" goto workbench
if "%choice%"=="0" goto end
echo 无效选择。
pause
goto menu

:run_po_v1
echo.
echo === PO Hour Allocator (V1) ===
%PYTHON_CMD% po_allocator.py
echo.
echo === 完成 ===
pause
goto menu

:run_po_v2
echo.
echo === PO Hour Allocator V2 (分析报告) ===
%PYTHON_CMD% po_allocator_v2.py --mode report
echo.
echo === 完成 ===
pause
goto menu

:run_po_preview
echo.
echo === PO Hour Allocator V2 (预览差异) ===
echo 比较当前数据和分析结果，不修改 Budget Report
echo.
%PYTHON_CMD% po_allocator_v2.py --mode preview
echo.
echo === 完成 ===
pause
goto menu

:run_po_write
echo.
echo === PO Hour Allocator V2 (回写Budget Report) ===
echo 警告: 此操作将修改 Budget Report 文件！
echo 系统会自动备份，但仍请确认后再继续。
echo.
%PYTHON_CMD% po_allocator_v2.py --mode write --generate-usage-status
echo.
echo === 完成 ===
pause
goto menu

:run_budget
echo.
echo === Budget Report Updater ===
%PYTHON_CMD% budget_updater.py
echo.
echo === 完成 ===
pause
goto menu

:run_travel
echo.
echo === Travel Cost Summary ===
%PYTHON_CMD% travel_summary.py
echo.
echo === 完成 ===
pause
goto menu

:run_pogr
echo.
echo === PO/GR Status Tracker ===
%PYTHON_CMD% po_gr_tracker.py list
echo.
echo --- Alerts ---
%PYTHON_CMD% po_gr_tracker.py alerts
echo.
pause
goto menu

:run_all
echo.
echo === [1/4] PO Hour Allocator V2 ===
%PYTHON_CMD% po_allocator_v2.py --mode report
echo.
echo === [2/4] Budget Report Updater ===
%PYTHON_CMD% budget_updater.py
echo.
echo === [3/4] Travel Cost Summary ===
%PYTHON_CMD% travel_summary.py
echo.
echo === [4/4] PO/GR Status ===
%PYTHON_CMD% po_gr_tracker.py list
echo.
echo === 全部完成 ===
pause
goto menu

:workbench
echo.
echo === 启动 IAV TV 工作台 ===
echo 正在启动 API 服务器（浏览器将自动打开）...
echo 按 Ctrl+C 可停止服务器返回菜单。
echo.
%PYTHON_CMD% api_server.py --port 8765
echo 工作台已关闭。
pause
goto menu

:end
echo 再见!
exit /b 0
