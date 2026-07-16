from __future__ import annotations

import argparse
from pathlib import Path

from bom_compare_core import CompareError, ComparisonRequest, compare_workbooks
from bom_compare_gui import run_app


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="BOM xlsx 比对工具，可通过 GUI 或 CLI 使用。",
    )
    parser.add_argument("left", nargs="?", help="左侧 BOM xlsx 文件路径")
    parser.add_argument("right", nargs="?", help="右侧 BOM xlsx 文件路径")
    parser.add_argument(
        "--gui",
        action="store_true",
        help="启动桌面图形界面",
    )
    parser.add_argument(
        "--output-dir",
        default=str(Path(__file__).resolve().parent / "outputs"),
        help="结果输出目录",
    )
    parser.add_argument(
        "--manual",
        action="store_true",
        help="手动指定左右工作表和表头参数，而不是自动比对所有可识别的同名工作表",
    )
    parser.add_argument("--left-sheet", help="左侧工作表名称")
    parser.add_argument("--right-sheet", help="右侧工作表名称")
    parser.add_argument("--left-header-row", type=int, help="左侧表头所在行号")
    parser.add_argument("--right-header-row", type=int, help="右侧表头所在行号")
    parser.add_argument("--left-key", help="左侧主键列名称")
    parser.add_argument("--right-key", help="右侧主键列名称")
    parser.add_argument(
        "--ignore-headers",
        default="序号,层次",
        help="不参与比对的列名，使用英文逗号分隔",
    )
    return parser


def build_request(args: argparse.Namespace) -> ComparisonRequest:
    if not args.left or not args.right:
        raise CompareError("CLI 模式下必须提供左右两个 xlsx 文件路径，或者直接使用 --gui。")

    ignore_headers = tuple(
        item.strip()
        for item in args.ignore_headers.replace("，", ",").split(",")
        if item.strip()
    )

    return ComparisonRequest(
        left_path=Path(args.left).expanduser().resolve(),
        right_path=Path(args.right).expanduser().resolve(),
        output_dir=Path(args.output_dir).expanduser().resolve(),
        compare_all_common_sheets=not args.manual,
        left_sheet_name=args.left_sheet,
        right_sheet_name=args.right_sheet,
        left_header_row=args.left_header_row,
        right_header_row=args.right_header_row,
        left_key_name=args.left_key,
        right_key_name=args.right_key,
        ignore_headers=ignore_headers,
    )


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.gui or (not args.left and not args.right):
        run_app()
        return

    request = build_request(args)
    result = compare_workbooks(request)
    print(result.summary_text())


if __name__ == "__main__":
    main()
