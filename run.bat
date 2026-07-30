@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

:menu
cls
echo ============================================
echo   IAV TV 项目管理工作台
echo ============================================
echo   [1] PO 工时规划 (V1)
echo   [2] PO 工时规划 (V2 - Level 配额)
echo   [3] Budget 月度更新
echo   [4] 差旅汇总
echo   [5] PO/GR 状态追踪
echo   [6] 运行全部工具
echo   [W] 启动工作台 (浏览器)
echo   [0] 退出
echo ============================================
set /p choice=请选择 (0-6/W):

if "%choice%"=="1" goto run_po_v1
if "%choice%"=="2" goto run_po_v2
if "%choice%"=="3" goto run_budget
if "%choice%"=="4" goto run_travel
if "%choice%"=="5" goto run_pogr
if "%choice%"=="6" goto run_all
if /i "%choice%"=="w" goto workbench
if "%choice%"=="0" goto end
echo 无效选择。
pause
goto menu

:run_po_v1
echo.
echo === PO Hour Allocator (V1) ===
python po_allocator.py
echo.
echo === 完成 ===
pause
goto menu

:run_po_v2
echo.
echo === PO Hour Allocator V2 (Level 配额) ===
python po_allocator_v2.py
echo.
echo === 完成 ===
pause
goto menu

:run_budget
echo.
echo === Budget Report Updater ===
python budget_updater.py
echo.
echo === 完成 ===
pause
goto menu

:run_travel
echo.
echo === Travel Cost Summary ===
python travel_summary.py
echo.
echo === 完成 ===
pause
goto menu

:run_pogr
echo.
echo === PO/GR Status Tracker ===
python po_gr_tracker.py list
echo.
echo --- Alerts ---
python po_gr_tracker.py alerts
echo.
pause
goto menu

:run_all
echo.
echo === [1/5] PO Hour Allocator V2 ===
python po_allocator_v2.py
echo.
echo === [2/5] Budget Report Updater ===
python budget_updater.py
echo.
echo === [3/5] Travel Cost Summary ===
python travel_summary.py
echo.
echo === [4/5] PO/GR Status ===
python po_gr_tracker.py list
echo.
echo === [5/5] PO/GR Alerts ===
python po_gr_tracker.py alerts
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
python api_server.py --port 8765
echo 工作台已关闭。
pause
goto menu

:end
echo 再见!
exit /b 0
