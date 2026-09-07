param([string]$File)
$ErrorActionPreference = 'Stop'

$pwFile = Join-Path $env:USERPROFILE '.docpw'

function Get-DocPassword {
    if (-not (Test-Path $pwFile)) {
        Write-Host '[!] 未找到密码文件。请先执行保存密码的步骤。' -ForegroundColor Yellow
        exit 1
    }
    try {
        $secure = (Get-Content $pwFile -Raw).Trim() | ConvertTo-SecureString
    } catch {
        Write-Host '[!] 密码文件解析失败（可能由其他 Windows 账户生成或已损坏）。请重新保存密码。' -ForegroundColor Yellow
        exit 1
    }
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

if (-not $File) {
    Write-Host '用法：将加密的 Word 文档“发送到 -> 打开需求文档”' -ForegroundColor Cyan
    exit 1
}

# 解析完整路径；SendTo 传入含空格路径时可能被拆分，用剩余参数拼接还原
$full = $null
try {
    $full = (Resolve-Path $File -ErrorAction Stop).Path
} catch {
    try {
        $full = (Resolve-Path ($args -join ' ') -ErrorAction Stop).Path
    } catch {
        Write-Host "[!] 找不到文件: $File" -ForegroundColor Yellow
        exit 1
    }
}

$pw = Get-DocPassword

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, ...)
    $doc = $word.Documents.Open($full, $false, $false, $false, $pw)
    Write-Host "[+] 已打开（自动填入密码）: $full" -ForegroundColor Green
} catch {
    Write-Host "[!] 自动打开失败: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host '    可能原因：密码已变更。若如此，请重新执行保存密码步骤；或手动打开该文档。' -ForegroundColor Yellow
}
