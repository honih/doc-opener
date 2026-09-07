$ErrorActionPreference = 'SilentlyContinue'

Write-Host '========================================' -ForegroundColor Cyan
Write-Host ' Word 加密文档自动打开器 - 卸载' -ForegroundColor Cyan
Write-Host '========================================' -ForegroundColor Cyan

$lnkPath = Join-Path $env:APPDATA 'Microsoft\Windows\SendTo\打开需求文档.lnk'
Remove-Item $lnkPath -Force
Write-Host '[+] 已移除右键菜单项（发送到 -> 打开需求文档）。' -ForegroundColor Green

$targetScript = Join-Path $env:USERPROFILE 'tools\open-doc.ps1'
Remove-Item $targetScript -Force
Write-Host '[+] 已移除核心脚本。' -ForegroundColor Green

$pwFile = Join-Path $env:USERPROFILE '.docpw'
if (Test-Path $pwFile) {
    Write-Host ('[i] 密码文件已保留：{0}' -f $pwFile) -ForegroundColor DarkGray
    Write-Host '    如需彻底清除本地密码，请手动删除该文件。' -ForegroundColor DarkGray
}

Write-Host ''
Write-Host '卸载完成。' -ForegroundColor Cyan
Write-Host ''
