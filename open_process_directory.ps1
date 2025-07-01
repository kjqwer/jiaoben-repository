# 获取当前活动窗口的程序目录并在文件资源管理器中打开
# 需要管理员权限才能获取某些系统进程的路径

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Collections.Generic;

public class Win32 {
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
    
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
    
    [DllImport("user32.dll")]
    public static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);
    
    [DllImport("user32.dll")]
    public static extern bool IsWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
    
    [DllImport("user32.dll")]
    public static extern bool GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
    
    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern IntPtr GetShellWindow();
    
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    
    public const uint GW_HWNDPREV = 3;
    public const uint GW_HWNDNEXT = 2;
}
"@

# 全局变量存储窗口列表
$script:windowList = @()

# 枚举窗口的回调函数
$enumWindowsCallback = {
    param([IntPtr]$hWnd, [IntPtr]$lParam)
    
    if ([Win32]::IsWindow($hWnd) -and [Win32]::IsWindowVisible($hWnd) -and -not [Win32]::IsIconic($hWnd)) {
        $title = New-Object System.Text.StringBuilder 256
        [Win32]::GetWindowText($hWnd, $title, 256)
        
        if ($title.ToString().Length -gt 0) {
            $processId = 0
            [Win32]::GetWindowThreadProcessId($hWnd, [ref]$processId)
            
            if ($processId -gt 0) {
                $windowInfo = @{
                    Handle = $hWnd
                    ProcessId = $processId
                    Title = $title.ToString()
                }
                $script:windowList += $windowInfo
            }
        }
    }
    return $true
}

try {
    # 短暂延迟，确保获取到快捷键触发前的窗口
    Start-Sleep -Milliseconds 200
    
    # 枚举所有可见窗口
    [Win32]::EnumWindows($enumWindowsCallback, [IntPtr]::Zero)
    
    # 获取当前活动窗口
    $foregroundWindow = [Win32]::GetForegroundWindow()
    $foregroundProcessId = 0
    [Win32]::GetWindowThreadProcessId($foregroundWindow, [ref]$foregroundProcessId)
    
    # 找到合适的窗口
    $targetWindow = $null
    $targetProcessId = 0
    
    # 首先尝试找到非控制台/终端的前一个窗口
    foreach ($window in $script:windowList) {
        if ($window.ProcessId -eq $foregroundProcessId) {
            continue  # 跳过当前活动窗口
        }
        
        try {
            $process = Get-Process -Id $window.ProcessId -ErrorAction SilentlyContinue
            if ($process) {
                # 跳过控制台、终端、脚本相关进程
                if ($process.ProcessName -notin @("powershell", "cmd", "conhost", "WindowsTerminal", "wt") -and
                    $process.ProcessName -notlike "*OpenProcessDirectory*" -and
                    $process.ProcessName -notlike "*ps2exe*") {
                    
                    $targetWindow = $window.Handle
                    $targetProcessId = $window.ProcessId
                    break
                }
            }
        } catch {
            continue
        }
    }
    
    # 如果没找到合适的窗口，使用当前活动窗口
    if (-not $targetWindow) {
        $targetWindow = $foregroundWindow
        $targetProcessId = $foregroundProcessId
    }
    
    if ($targetProcessId -eq 0) {
        Write-Host "无法获取当前窗口的进程ID" -ForegroundColor Red
        exit 1
    }
    
    # 获取进程信息
    $process = Get-Process -Id $targetProcessId -ErrorAction SilentlyContinue
    
    if (-not $process) {
        Write-Host "无法获取进程信息 (PID: $targetProcessId)" -ForegroundColor Red
        exit 1
    }
    
    # 获取进程的可执行文件路径
    $processPath = $process.MainModule.FileName
    
    if (-not $processPath) {
        Write-Host "无法获取进程路径 (进程: $($process.ProcessName))" -ForegroundColor Red
        exit 1
    }
    
    # 获取文件所在目录
    $directory = Split-Path -Parent $processPath
    
    Write-Host "当前活动窗口进程: $($process.ProcessName)" -ForegroundColor Green
    Write-Host "进程路径: $processPath" -ForegroundColor Green
    Write-Host "程序目录: $directory" -ForegroundColor Green
    
    # 检查目录是否存在
    if (Test-Path $directory) {
        Write-Host "正在打开目录..." -ForegroundColor Yellow
        # 在文件资源管理器中打开目录
        Start-Process "explorer.exe" -ArgumentList $directory
        Write-Host "目录已打开!" -ForegroundColor Green
    } else {
        Write-Host "目录不存在: $directory" -ForegroundColor Red
        exit 1
    }
    
} catch {
    Write-Host "发生错误: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "可能需要管理员权限来获取某些系统进程的路径" -ForegroundColor Yellow
    exit 1
}

# 等待用户按键后退出
# Write-Host "`n按任意键退出..." -ForegroundColor Cyan
# $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") 