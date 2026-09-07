$ErrorActionPreference = 'Stop'

Write-Host '========================================' -ForegroundColor Cyan
Write-Host ' Word 加密文档自动打开器 - 安装' -ForegroundColor Cyan
Write-Host '========================================' -ForegroundColor Cyan

$toolsDir     = Join-Path $env:USERPROFILE 'tools'
$pwFile       = Join-Path $env:USERPROFILE '.docpw'
$scriptName   = 'open-doc.ps1'
$targetScript = Join-Path $toolsDir $scriptName

# ---------- 第 1 步：确保密码已加密保存 ----------
function Set-DocPassword {
    Write-Host ''
    Write-Host '[设置密码] 请输入 Word 文档的统一密码（屏幕显示为星号，仅加密保存在本机）' -ForegroundColor Yellow
    try {
        $secure = Read-Host -AsSecureString '请输入密码'
        if ($null -eq $secure -or $secure.Length -eq 0) {
            Write-Host '[!] 密码为空，已取消。' -ForegroundColor Red
            exit 1
        }
        $secure | ConvertFrom-SecureString | Set-Content $pwFile -Encoding Ascii
    } catch {
        Write-Host ('[!] 密码保存失败：{0}' -f $_.Exception.Message) -ForegroundColor Red
        exit 1
    }
    Write-Host '[+] 密码已加密保存到本机。' -ForegroundColor Green
}

if (Test-Path $pwFile) {
    try {
        $null = (Get-Content $pwFile -Raw).Trim() | ConvertTo-SecureString
        Write-Host '[i] 已存在有效的密码文件，跳过设置。' -ForegroundColor DarkGray
        Write-Host ('    如需更换密码：删除 {0} 后重新运行本安装。' -f $pwFile) -ForegroundColor DarkGray
    } catch {
        Write-Host '[!] 密码文件无法解密（可能损坏或由其他 Windows 账户生成），将重新设置。' -ForegroundColor Yellow
        Remove-Item $pwFile -Force
        Set-DocPassword
    }
} else {
    Set-DocPassword
}

# ---------- 第 2 步：检测 MS Office Word COM ----------
Write-Host ''
Write-Host '[检测] 检查 Microsoft Word 组件...' -ForegroundColor Yellow
$wordType = $null
try {
    $wordType = [Type]::GetTypeFromProgID('Word.Application')
} catch { }
if ($null -eq $wordType) {
    Write-Host '[!] 未检测到 Microsoft Word COM 组件。' -ForegroundColor Red
    Write-Host '    本工具依赖 MS Office Word，WPS 暂不支持。' -ForegroundColor Red
    exit 1
}
Write-Host '[+] Word COM 可用。' -ForegroundColor Green

# ---------- 第 3 步：复制核心脚本到固定位置 ----------
if (-not (Test-Path $toolsDir)) {
    New-Item -ItemType Directory -Path $toolsDir | Out-Null
}
$sourceScript = Join-Path $PSScriptRoot $scriptName
if (-not (Test-Path $sourceScript)) {
    Write-Host ('[!] 找不到核心脚本：{0}' -f $sourceScript) -ForegroundColor Red
    Write-Host '    请确认 open-doc.ps1 与本安装脚本在同一目录。' -ForegroundColor Red
    exit 1
}
Copy-Item $sourceScript $targetScript -Force
Write-Host ('[+] 核心脚本已安装：{0}' -f $targetScript) -ForegroundColor Green

# ---------- 第 4 步：创建“发送到”快捷方式 ----------
$sendTo = Join-Path $env:APPDATA 'Microsoft\Windows\SendTo'
if (-not (Test-Path $sendTo)) {
    Write-Host ('[!] 未找到 SendTo 目录：{0}' -f $sendTo) -ForegroundColor Red
    exit 1
}
$lnkPath = Join-Path $sendTo '打开需求文档.lnk'
$ws = New-Object -ComObject WScript.Shell
$lnk = $ws.CreateShortcut($lnkPath)
$lnk.TargetPath = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
$lnk.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$targetScript`""
$lnk.Description = '自动输入密码打开加密的 Word 需求文档'
$lnk.IconLocation = "$env:WINDIR\System32\shell32.dll,13"
$lnk.Save()
Write-Host '[+] 右键菜单已创建：发送到 -> 打开需求文档' -ForegroundColor Green

# ---------- 完成 ----------
Write-Host ''
Write-Host '安装完成！使用方法：' -ForegroundColor Cyan
Write-Host '  1. 在文件夹/共享目录中找到加密的 Word 文档' -ForegroundColor White
Write-Host '  2. 右键 -> 发送到 -> 打开需求文档' -ForegroundColor White
Write-Host ''
Write-Host '安全提示：若公司制度明确禁止自动化处理受控文档，请先与主管/IT 确认后再使用。' -ForegroundColor DarkGray
Write-Host ''
