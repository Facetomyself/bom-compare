# BOM Compare

Windows-oriented BOM `.xlsx` 比对工具，提供图形界面与命令行两种使用方式。

## 功能特性

- 选择任意两个 `.xlsx` BOM 文件进行比对
- 自动识别可比对的同名工作表
- 自动识别常见主键列，如 `组件编码`、`物料编码`、`Part Number`
- 忽略物料顺序，仅比对物料信息
- 差异单元格标红
- 单侧独有物料整行标红
- 支持 GitHub Actions 自动构建 Windows 可执行文件

## 环境要求

- Python 3.10 及以上
- `uv`

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

## 启动方式

### 图形界面

```bash
python "compare_bom.py"
```

或：

```bash
python "compare_bom.py" --gui
```

### 命令行

```bash
python "compare_bom.py" "左侧.xlsx" "右侧.xlsx"
```

## 图形界面使用说明

1. 选择左侧 BOM 文件
2. 选择右侧 BOM 文件
3. 选择输出目录
4. 点击“开始比对”
5. 在输出目录获取两份带差异标红的结果文件

## 输出结果

输出目录默认生成两份结果文件：

- `左侧文件名-差异标红-时间戳-运行标记-left.xlsx`
- `右侧文件名-差异标红-时间戳-运行标记-right.xlsx`

两份结果共用一次运行的标记。不同目录的同名表或同一秒重复运行会产生独立结果；已存在文件保持原样。CLI 比对只需核心 Python 依赖，Tk 仅在启动 GUI 时载入。

## 项目结构

```text
.
├── bom_compare_core.py          # 核心比对逻辑
├── bom_compare_gui.py           # 图形界面
├── compare_bom.py               # 程序入口
├── bom_compare.spec             # PyInstaller 配置
├── scripts/
│   ├── build_windows.bat        # Windows CMD 打包脚本
│   └── build_windows.ps1        # Windows PowerShell 打包脚本
└── .github/workflows/
    ├── build-windows.yml        # Windows 构建流程
    └── release-windows.yml      # Windows Release 发布流程
```

## Windows 本地打包

### PowerShell

```powershell
cd "你的项目目录\Bom_compare"
powershell -ExecutionPolicy Bypass -File ".\scripts\build_windows.ps1"
```

### CMD

```bat
cd /d "你的项目目录\Bom_compare"
scripts\build_windows.bat
```

默认产物路径：

```text
dist\BOMCompare.exe
```

## GitHub Actions

### 构建

- 工作流：`.github/workflows/build-windows.yml`
- 用途：在 `main` 或 `master` 分支更新后自动构建 Windows 可执行文件并上传 artifact

### 发布 Release

- 工作流：`.github/workflows/release-windows.yml`
- 触发方式：
  - 推送标签：`v*`
  - 手动触发：`workflow_dispatch`
- 发布产物：
  - `BOMCompare-<version>-windows-x64.exe`

## 核心与命令行回归

提前下载并核对公开 wheel 后，在隔离环境内按锁安装，再运行测试：

```bash
python -m pip install --no-index --only-binary=:all: --require-hashes --find-links WHEEL_DIR -r tests/requirements-acceptance.lock
python -m unittest discover -s tests -v
```

用例覆盖无 Tk 的 CLI、两个同名工作簿、同一秒重复执行、目标已存在及第二份写入失败。Windows GUI 和 EXE 另行验收。

## 开发说明

- 当前版本仅处理 `.xlsx` 文件
- GUI 使用 Python 自带 `tkinter`
- 默认忽略 `序号`、`层次` 等顺序字段
- 若需调整自动识别规则，请修改 `bom_compare_core.py`
