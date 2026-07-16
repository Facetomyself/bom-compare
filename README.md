# BOM Compare UI

一个支持跨平台使用的 BOM `.xlsx` 比对工具，支持：

- 选择任意两个 `.xlsx` BOM 文件
- 自动识别可比对的工作表、表头行和主键列
- 忽略物料行顺序
- 只对物料信息做比对
- 差异单元格标红
- 对仅存在于某一侧的物料整行标红
- 支持图形界面和命令行两种使用方式
- 支持打包为 Windows 可执行文件

## 运行环境

- Python 3.10+
- 推荐使用 `uv`

## 安装

### macOS / Linux

```bash
cd "/Users/mengma/pythonProject/Bom_compare"
uv venv
source ".venv/bin/activate"
uv pip install -e .
```

### Windows PowerShell

```powershell
cd "你的项目目录\Bom_compare"
uv venv
.\.venv\Scripts\Activate.ps1
uv pip install -e .
```

## 启动图形界面

```bash
cd "/Users/mengma/pythonProject/Bom_compare"
source ".venv/bin/activate"
python "compare_bom.py" --gui
```

如果不带参数直接运行，也会默认打开 GUI：

```bash
python "compare_bom.py"
```

## 命令行使用

### 自动比对所有可识别的同名工作表

```bash
python "compare_bom.py" "左侧.xlsx" "右侧.xlsx"
```

## GUI 使用说明

1. 选择左侧和右侧的 BOM `.xlsx` 文件
2. 选择输出目录
3. 点击“开始比对”
4. 程序会自动识别可比对的同名工作表，并在输出目录生成两份带标红结果的工作簿

## 比对规则

- 默认把 `组件编码 / 物料编码 / 料号 / Part Number` 等识别为主键列
- 默认忽略 `序号 / 层次` 这类顺序列
- 对左右都存在且主键唯一的物料，逐字段比较并标红差异单元格
- 对只存在于某一侧的物料，整行物料信息标红
- 如果同一个主键在某个工作表里重复出现，会退化为“按整行签名匹配”，无法精确定位到单个字段时，会整行标红兜底

## 输出结果

输出目录里会生成两份文件，例如：

- `左侧文件名-差异标红-时间戳.xlsx`
- `右侧文件名-差异标红-时间戳.xlsx`

## 说明

- 当前版本只处理 `.xlsx` 文件
- GUI 使用的是 Python 自带 `tkinter`，不依赖额外桌面框架
- 如果 BOM 表头命名特别野，自动识别可能失败，这时建议直接改代码里的识别规则后重新打包

## Windows 打包

### 先说结论

通常**不能指望在 macOS 上直接稳定打出 Windows `.exe`**。

原因很直接：

- `PyInstaller` 打的是**目标操作系统原生可执行文件**
- Windows `.exe` 需要 Windows 环境下的 bootloader 和依赖收集
- 在 macOS 上即便你硬折腾 Wine 或别的旁路，稳定性也一般，容易整出“能打包、不能运行”的假成功

最稳的路子就两种：

1. **在 Windows 机器本地打包**
2. **用 GitHub Actions 的 Windows runner 自动打包**

### 本项目已经准备好的打包文件

- PyInstaller 配置：`bom_compare.spec`
- PowerShell 打包脚本：`scripts/build_windows.ps1`
- 批处理打包脚本：`scripts/build_windows.bat`
- GitHub Actions 工作流：`.github/workflows/build-windows.yml`

### 在 Windows 本地打包

#### PowerShell

```powershell
cd "你的项目目录\Bom_compare"
powershell -ExecutionPolicy Bypass -File ".\scripts\build_windows.ps1"
```

#### CMD

```bat
cd /d "你的项目目录\Bom_compare"
scripts\build_windows.bat
```

打包完成后，产物默认在：

```text
dist\BOMCompare.exe
```

### 用 GitHub Actions 自动打包

把项目推到 GitHub 后：

1. 打开仓库的 `Actions`
2. 运行 `Build Windows Package`
3. 等待任务完成
4. 在 artifacts 里下载 `BOMCompare-windows`

### 打包依赖

项目已经在 `pyproject.toml` 里加了：

```toml
[project.optional-dependencies]
build = ["pyinstaller>=6.15.0"]
```

所以 Windows 环境下直接执行：

```powershell
uv pip install --python ".venv\Scripts\python.exe" -e ".[build]"
```

就能把打包依赖一起装上。
