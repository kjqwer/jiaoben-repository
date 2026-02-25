@echo off
chcp 65001 >nul
echo 正在获取当前活动窗口的进程信息...
echo.

echo 当前系统中有窗口的所有进程:
echo ===============================
echo.

REM 显示所有有窗口的进程
powershell "Get-Process | Where-Object {$_.MainWindowTitle -ne ''} | Select-Object Id,ProcessName,MainWindowTitle,@{Name='Memory(MB)';Expression={[math]::Round($_.WorkingSet64/1MB,2)}} | Format-Table -AutoSize"

echo.
echo ===============================
echo 提示: 上面列出的是所有有活动窗口的进程
echo 当前最前端的窗口通常是最近使用的进程
echo.

REM 使用简单的PowerShell命令获取进程详细信息
echo 详细进程信息:
echo ===============================
powershell "Get-Process | Where-Object {$_.MainWindowTitle -ne ''} | ForEach-Object { Write-Host ('进程: ' + $_.ProcessName + ' (ID: ' + $_.Id + ')') -ForegroundColor Green; Write-Host ('窗口: ' + $_.MainWindowTitle) -ForegroundColor Yellow; Write-Host ('路径: ' + $_.Path) -ForegroundColor Cyan; Write-Host '---' }"

echo.
echo 按任意键继续...
pause >nul 