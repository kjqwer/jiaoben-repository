# 工具集

简单的个人脚本集合，以后脚本都存入此处。

## 脚本说明

### MediaControl.ps1
全局媒体控制脚本，可以控制任何支持 Windows Media Session 的播放器。

**功能：**
- 播放/暂停
- 上一首/下一首
- 停止播放
- 自动启动媒体播放器

**用法：**
```powershell
.\MediaControl.ps1              # 播放/暂停
.\MediaControl.ps1 next         # 下一首
.\MediaControl.ps1 prev         # 上一首
.\MediaControl.ps1 stop         # 停止
```

### open_process_directory.ps1
快速打开当前活动窗口进程的工作目录（主要通过exe快捷键映射后使用）。

**功能：**
- 自动获取当前活动窗口的进程
- 打开该进程的可执行文件所在目录
- 智能跳过控制台/终端窗口

**用法：**
```powershell
.\open_process_directory.ps1 
```

## 编译版本

项目包含 PowerShell 脚本的编译版本（.exe），可以直接双击运行，但主要是通过快捷键映射之后使用。

## 编译工具

### 安装 ps2exe
```powershell
Install-Module -Name ps2exe -Force
```

### 编译脚本
```powershell
# 编译 open_process_directory.ps1
Invoke-ps2exe -InputFile "open_process_directory.ps1" -OutputFile "OpenProcessDirectory.exe" -RequireAdmin

# 编译 MediaControl.ps1  
Invoke-ps2exe -InputFile "MediaControl.ps1" -OutputFile "MediaControl.exe" -RequireAdmin
```

编译后的 .exe 文件可以直接运行，无需 PowerShell 环境。 