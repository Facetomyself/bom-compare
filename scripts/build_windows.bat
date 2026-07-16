@echo off
setlocal

set "PROJECT_ROOT=%~dp0.."
cd /d "%PROJECT_ROOT%"

where uv >nul 2>nul
if errorlevel 1 (
  echo 未找到 uv，请先安装 uv。
  exit /b 1
)

if not exist ".venv" (
  echo ==> 创建 .venv
  uv venv
)

echo ==> 安装运行和打包依赖
uv pip install --python ".venv\Scripts\python.exe" -e ".[build]"

echo ==> 清理旧产物
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo ==> 开始 PyInstaller 打包
".venv\Scripts\python.exe" -m PyInstaller ".\bom_compare.spec" --noconfirm --clean
if errorlevel 1 exit /b 1

echo.
echo 打包完成，产物路径：
echo   %PROJECT_ROOT%\dist\BOMCompare.exe
