Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot

Write-Host "==> Project root: $ProjectRoot"

if (-not (Get-Command uv -ErrorAction SilentlyContinue)) {
    throw "未找到 uv，请先安装 uv。"
}

if (-not (Test-Path ".venv")) {
    Write-Host "==> 创建 .venv"
    uv venv
}

Write-Host "==> 安装运行和打包依赖"
uv pip install --python ".venv\\Scripts\\python.exe" -e ".[build]"

Write-Host "==> 清理旧产物"
if (Test-Path "build") {
    Remove-Item "build" -Recurse -Force
}
if (Test-Path "dist") {
    Remove-Item "dist" -Recurse -Force
}

Write-Host "==> 开始 PyInstaller 打包"
.venv\Scripts\python.exe -m PyInstaller ".\\bom_compare.spec" --noconfirm --clean

Write-Host ""
Write-Host "打包完成，产物路径："
Write-Host "  $ProjectRoot\\dist\\BOMCompare.exe"
