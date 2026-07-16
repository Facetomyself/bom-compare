from __future__ import annotations

import re
from collections import Counter, defaultdict
from copy import copy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet


RED_FILL = PatternFill(fill_type="solid", fgColor="FFC7CE")
RED_FONT_COLOR = "9C0006"
DEFAULT_OUTPUT_DIR_NAME = "outputs"


HEADER_ALIAS_GROUPS: dict[str, set[str]] = {
    "组件编码": {
        "组件编码",
        "物料编码",
        "物料号",
        "料号",
        "part number",
        "part no",
        "item code",
        "component code",
        "material code",
        "pn",
        "p/n",
        "编码",
    },
    "组件描述": {
        "组件描述",
        "物料描述",
        "描述",
        "part description",
        "description",
        "item description",
    },
    "规格型号": {
        "规格型号",
        "规格",
        "型号",
        "spec",
        "specification",
        "model",
    },
    "单位": {
        "单位",
        "uom",
        "unit",
    },
    "用量": {
        "用量",
        "数量",
        "qty",
        "quantity",
        "usage",
    },
    "扩展用量": {
        "扩展用量",
        "总用量",
        "扩展数量",
        "extended qty",
        "total qty",
    },
    "组件位置": {
        "组件位置",
        "位号",
        "位号位置",
        "位置",
        "refdes",
        "reference designator",
        "location",
        "position",
    },
    "装配属性": {
        "装配属性",
        "assembly attribute",
        "mounting attribute",
    },
    "专用属性": {
        "专用属性",
        "special attribute",
    },
    "物料类型": {
        "物料类型",
        "material type",
    },
    "供应类型": {
        "供应类型",
        "supply type",
    },
    "供应仓库": {
        "供应仓库",
        "warehouse",
    },
    "版本": {
        "版本",
        "version",
    },
    "版本标签": {
        "版本标签",
        "version label",
    },
    "替代物料编码": {
        "替代物料编码",
        "替代编码",
        "alternative code",
        "substitute code",
    },
    "替代物料描述": {
        "替代物料描述",
        "替代描述",
        "alternative description",
        "substitute description",
    },
    "替代料位号": {
        "替代料位号",
        "替代位号",
        "alternative position",
    },
    "替代分摊比编码": {
        "替代分摊比编码",
        "替代分摊比",
    },
    "序号": {
        "序号",
        "no",
        "no.",
        "item",
        "item no",
        "line no",
        "index",
        "项次",
        "行号",
    },
    "层次": {
        "层次",
        "level",
        "hierarchy",
    },
}


def _simplify_text(value: str) -> str:
    return re.sub(r"[\s\r\n\t:：\-_./\\()[\]{}]+", "", value).lower()


SIMPLIFIED_ALIAS_TO_CANONICAL: dict[str, str] = {}
for canonical_name, aliases in HEADER_ALIAS_GROUPS.items():
    SIMPLIFIED_ALIAS_TO_CANONICAL[_simplify_text(canonical_name)] = canonical_name
    for alias in aliases:
        SIMPLIFIED_ALIAS_TO_CANONICAL[_simplify_text(alias)] = canonical_name


DEFAULT_IGNORE_CANONICAL_HEADERS = {"序号", "层次"}
DEFAULT_KEY_CANONICAL_HEADERS = {"组件编码"}


class CompareError(Exception):
    """Raised when workbook comparison cannot proceed."""


@dataclass(frozen=True)
class HeaderCell:
    col_index: int
    raw_name: str
    canonical_name: str
    occurrence: int
    unique_key: str
    display_name: str


@dataclass(frozen=True)
class WorksheetSchema:
    sheet_name: str
    header_row: int
    headers: tuple[HeaderCell, ...]
    key_header: HeaderCell
    material_headers: tuple[HeaderCell, ...]


@dataclass(frozen=True)
class WorksheetInspection:
    sheet_name: str
    detected_header_row: int | None
    detected_key_name: str | None
    header_display_names: tuple[str, ...]
    is_candidate: bool
    detection_error: str | None = None


@dataclass(frozen=True)
class WorkbookInspection:
    workbook_path: Path
    sheets: tuple[WorksheetInspection, ...]

    @property
    def sheet_names(self) -> list[str]:
        return [sheet.sheet_name for sheet in self.sheets]

    def get_sheet(self, sheet_name: str) -> WorksheetInspection | None:
        for sheet in self.sheets:
            if sheet.sheet_name == sheet_name:
                return sheet
        return None


@dataclass(frozen=True)
class RowRecord:
    row_index: int
    key: str
    values: dict[str, str]


@dataclass(frozen=True)
class SheetComparisonSummary:
    left_sheet_name: str
    right_sheet_name: str
    only_left_count: int
    only_right_count: int
    changed_count: int
    duplicate_group_count: int
    duplicate_mismatch_row_count: int
    skipped_left_only_columns: tuple[str, ...]
    skipped_right_only_columns: tuple[str, ...]

    def to_lines(self) -> list[str]:
        lines = [
            (
                f"{self.left_sheet_name} ↔ {self.right_sheet_name}："
                f"左侧独有 {self.only_left_count} 行，"
                f"右侧独有 {self.only_right_count} 行，"
                f"逐字段差异 {self.changed_count} 行"
            )
        ]
        if self.duplicate_group_count:
            lines.append(
                f"  - 检测到重复主键分组 {self.duplicate_group_count} 组，"
                f"其中 {self.duplicate_mismatch_row_count} 行按整行高亮处理"
            )
        if self.skipped_left_only_columns:
            lines.append(
                "  - 仅左侧存在、未参与逐字段比对的列："
                + "、".join(self.skipped_left_only_columns)
            )
        if self.skipped_right_only_columns:
            lines.append(
                "  - 仅右侧存在、未参与逐字段比对的列："
                + "、".join(self.skipped_right_only_columns)
            )
        return lines


@dataclass(frozen=True)
class ComparisonRequest:
    left_path: Path
    right_path: Path
    output_dir: Path
    compare_all_common_sheets: bool = True
    left_sheet_name: str | None = None
    right_sheet_name: str | None = None
    left_header_row: int | None = None
    right_header_row: int | None = None
    left_key_name: str | None = None
    right_key_name: str | None = None
    ignore_headers: tuple[str, ...] = ()


@dataclass(frozen=True)
class ComparisonResult:
    left_output_path: Path
    right_output_path: Path
    summaries: tuple[SheetComparisonSummary, ...]
    compared_sheet_pairs: tuple[tuple[str, str], ...]

    def summary_text(self) -> str:
        lines = ["比对完成："]
        for summary in self.summaries:
            lines.extend(summary.to_lines())
        lines.append(f"左侧结果：{self.left_output_path}")
        lines.append(f"右侧结果：{self.right_output_path}")
        return "\n".join(lines)


def normalize_cell_value(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else format(value, "g")
    if isinstance(value, datetime):
        return value.isoformat(sep=" ", timespec="seconds")
    return str(value).strip()


def canonicalize_header_name(raw_name: str) -> str:
    simplified = _simplify_text(raw_name)
    return SIMPLIFIED_ALIAS_TO_CANONICAL.get(simplified, raw_name.strip())


def parse_ignore_headers(text: str) -> tuple[str, ...]:
    if not text.strip():
        return ()
    parts = re.split(r"[,，;；\n]+", text)
    return tuple(part.strip() for part in parts if part.strip())


def build_headers(ws: Worksheet, header_row: int) -> tuple[HeaderCell, ...]:
    occurrences: Counter[str] = Counter()
    headers: list[HeaderCell] = []

    for col_index in range(1, ws.max_column + 1):
        raw_name = normalize_cell_value(ws.cell(header_row, col_index).value)
        if not raw_name:
            continue
        canonical_name = canonicalize_header_name(raw_name)
        occurrences[canonical_name] += 1
        occurrence = occurrences[canonical_name]
        unique_key = f"{canonical_name}__{occurrence}"
        display_name = raw_name if occurrence == 1 else f"{raw_name} ({occurrence})"
        headers.append(
            HeaderCell(
                col_index=col_index,
                raw_name=raw_name,
                canonical_name=canonical_name,
                occurrence=occurrence,
                unique_key=unique_key,
                display_name=display_name,
            )
        )

    return tuple(headers)


def detect_header_row(ws: Worksheet, max_scan_rows: int = 40) -> int:
    best_row_index: int | None = None
    best_score = -1
    scan_limit = min(ws.max_row, max_scan_rows)

    for row_index in range(1, scan_limit + 1):
        raw_values = [
            normalize_cell_value(ws.cell(row_index, col_index).value)
            for col_index in range(1, ws.max_column + 1)
        ]
        values = [value for value in raw_values if value]
        if len(values) < 3:
            continue

        canonical_values = {canonicalize_header_name(value) for value in values}
        key_hits = len(canonical_values & DEFAULT_KEY_CANONICAL_HEADERS)
        known_hits = len(canonical_values & set(HEADER_ALIAS_GROUPS))

        score = key_hits * 100 + known_hits * 10 + min(len(values), 15)
        if score > best_score and key_hits:
            best_score = score
            best_row_index = row_index

    if best_row_index is None:
        raise CompareError(f"工作表 {ws.title} 未识别到 BOM 表头，请手动指定表头行。")

    return best_row_index


def find_key_header(
    headers: Iterable[HeaderCell],
    manual_key_name: str | None = None,
) -> HeaderCell:
    header_list = list(headers)
    if not header_list:
        raise CompareError("表头为空，无法识别主键列。")

    if manual_key_name:
        normalized_manual = _simplify_text(manual_key_name)
        for header in header_list:
            candidates = {
                _simplify_text(header.raw_name),
                _simplify_text(header.display_name),
                _simplify_text(header.canonical_name),
            }
            if normalized_manual in candidates:
                return header
        raise CompareError(f"未找到手动指定的主键列：{manual_key_name}")

    for header in header_list:
        if header.canonical_name in DEFAULT_KEY_CANONICAL_HEADERS:
            return header

    raise CompareError("未识别到主键列，请手动指定。")


def build_schema(
    ws: Worksheet,
    *,
    manual_header_row: int | None = None,
    manual_key_name: str | None = None,
    ignore_headers: tuple[str, ...] = (),
) -> WorksheetSchema:
    header_row = manual_header_row or detect_header_row(ws)
    headers = build_headers(ws, header_row)
    if not headers:
        raise CompareError(f"工作表 {ws.title} 的表头行 {header_row} 没有可用列。")

    key_header = find_key_header(headers, manual_key_name)
    ignore_canonical_headers = {
        canonicalize_header_name(header_name)
        for header_name in ignore_headers
        if header_name.strip()
    }
    ignore_canonical_headers |= DEFAULT_IGNORE_CANONICAL_HEADERS

    material_headers = tuple(
        header
        for header in headers
        if header.canonical_name not in ignore_canonical_headers
    )
    if not any(header.unique_key == key_header.unique_key for header in material_headers):
        material_headers = (key_header,) + material_headers

    return WorksheetSchema(
        sheet_name=ws.title,
        header_row=header_row,
        headers=headers,
        key_header=key_header,
        material_headers=material_headers,
    )


def inspect_workbook(path: Path) -> WorkbookInspection:
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        sheets: list[WorksheetInspection] = []
        for ws in workbook.worksheets:
            try:
                schema = build_schema(ws)
                sheets.append(
                    WorksheetInspection(
                        sheet_name=ws.title,
                        detected_header_row=schema.header_row,
                        detected_key_name=schema.key_header.display_name,
                        header_display_names=tuple(
                            header.display_name for header in schema.headers
                        ),
                        is_candidate=True,
                    )
                )
            except Exception as exc:  # noqa: BLE001
                sheets.append(
                    WorksheetInspection(
                        sheet_name=ws.title,
                        detected_header_row=None,
                        detected_key_name=None,
                        header_display_names=(),
                        is_candidate=False,
                        detection_error=str(exc),
                    )
                )
        return WorkbookInspection(workbook_path=path, sheets=tuple(sheets))
    finally:
        workbook.close()


def extract_rows(ws: Worksheet, schema: WorksheetSchema) -> list[RowRecord]:
    rows: list[RowRecord] = []
    material_headers = schema.material_headers
    key_unique_key = schema.key_header.unique_key

    for row_index in range(schema.header_row + 1, ws.max_row + 1):
        key = normalize_cell_value(ws.cell(row_index, schema.key_header.col_index).value)
        if not key:
            continue

        values: dict[str, str] = {}
        non_key_has_value = False
        for header in material_headers:
            value = normalize_cell_value(ws.cell(row_index, header.col_index).value)
            values[header.unique_key] = value
            if header.unique_key != key_unique_key and value:
                non_key_has_value = True

        if not non_key_has_value and not key:
            continue

        rows.append(RowRecord(row_index=row_index, key=key, values=values))

    return rows


def apply_red_style(cell) -> None:
    cell.fill = copy(RED_FILL)
    base_font = copy(cell.font)
    base_font.color = RED_FONT_COLOR
    cell.font = base_font


def highlight_row(ws: Worksheet, row: RowRecord, headers: Iterable[HeaderCell]) -> None:
    for header in headers:
        apply_red_style(ws.cell(row.row_index, header.col_index))


def compare_row_signatures(
    rows: list[RowRecord],
    compare_keys: tuple[str, ...],
) -> Counter[tuple[str, ...]]:
    return Counter(
        tuple(row.values.get(compare_key, "") for compare_key in compare_keys)
        for row in rows
    )


def compare_sheet(
    left_ws: Worksheet,
    right_ws: Worksheet,
    *,
    left_schema: WorksheetSchema,
    right_schema: WorksheetSchema,
) -> SheetComparisonSummary:
    left_rows = extract_rows(left_ws, left_schema)
    right_rows = extract_rows(right_ws, right_schema)

    left_groups: dict[str, list[RowRecord]] = defaultdict(list)
    right_groups: dict[str, list[RowRecord]] = defaultdict(list)
    for row in left_rows:
        left_groups[row.key].append(row)
    for row in right_rows:
        right_groups[row.key].append(row)

    left_headers_by_key = {header.unique_key: header for header in left_schema.material_headers}
    right_headers_by_key = {
        header.unique_key: header for header in right_schema.material_headers
    }

    shared_compare_keys = tuple(
        unique_key
        for unique_key in left_headers_by_key
        if unique_key in right_headers_by_key
        and unique_key != left_schema.key_header.unique_key
    )

    skipped_left_only_columns = tuple(
        header.display_name
        for header in left_schema.material_headers
        if header.unique_key not in right_headers_by_key
        and header.unique_key != left_schema.key_header.unique_key
    )
    skipped_right_only_columns = tuple(
        header.display_name
        for header in right_schema.material_headers
        if header.unique_key not in left_headers_by_key
        and header.unique_key != right_schema.key_header.unique_key
    )

    only_left_count = 0
    only_right_count = 0
    changed_count = 0
    duplicate_group_count = 0
    duplicate_mismatch_row_count = 0

    all_keys = sorted(set(left_groups) | set(right_groups))
    for key in all_keys:
        left_group = left_groups.get(key, [])
        right_group = right_groups.get(key, [])

        if not left_group:
            only_right_count += len(right_group)
            for row in right_group:
                highlight_row(right_ws, row, right_schema.material_headers)
            continue

        if not right_group:
            only_left_count += len(left_group)
            for row in left_group:
                highlight_row(left_ws, row, left_schema.material_headers)
            continue

        if len(left_group) == 1 and len(right_group) == 1:
            left_row = left_group[0]
            right_row = right_group[0]
            diff_headers = [
                (
                    left_headers_by_key[compare_key],
                    right_headers_by_key[compare_key],
                )
                for compare_key in shared_compare_keys
                if left_row.values.get(compare_key, "") != right_row.values.get(compare_key, "")
            ]
            if diff_headers:
                changed_count += 1
                for left_header, right_header in diff_headers:
                    apply_red_style(left_ws.cell(left_row.row_index, left_header.col_index))
                    apply_red_style(right_ws.cell(right_row.row_index, right_header.col_index))
            continue

        duplicate_group_count += 1
        left_signature_counter = compare_row_signatures(left_group, shared_compare_keys)
        right_signature_counter = compare_row_signatures(right_group, shared_compare_keys)

        matched_signatures = left_signature_counter & right_signature_counter
        left_remaining = left_signature_counter - matched_signatures
        right_remaining = right_signature_counter - matched_signatures

        if not left_remaining and not right_remaining:
            continue

        for row in left_group:
            signature = tuple(row.values.get(compare_key, "") for compare_key in shared_compare_keys)
            if left_remaining[signature] > 0:
                left_remaining[signature] -= 1
                duplicate_mismatch_row_count += 1
                highlight_row(left_ws, row, left_schema.material_headers)

        for row in right_group:
            signature = tuple(row.values.get(compare_key, "") for compare_key in shared_compare_keys)
            if right_remaining[signature] > 0:
                right_remaining[signature] -= 1
                duplicate_mismatch_row_count += 1
                highlight_row(right_ws, row, right_schema.material_headers)

    return SheetComparisonSummary(
        left_sheet_name=left_schema.sheet_name,
        right_sheet_name=right_schema.sheet_name,
        only_left_count=only_left_count,
        only_right_count=only_right_count,
        changed_count=changed_count,
        duplicate_group_count=duplicate_group_count,
        duplicate_mismatch_row_count=duplicate_mismatch_row_count,
        skipped_left_only_columns=skipped_left_only_columns,
        skipped_right_only_columns=skipped_right_only_columns,
    )


def build_sheet_pairs(
    left_workbook,
    right_workbook,
    request: ComparisonRequest,
) -> list[tuple[Worksheet, Worksheet, WorksheetSchema, WorksheetSchema]]:
    pairs: list[tuple[Worksheet, Worksheet, WorksheetSchema, WorksheetSchema]] = []

    if request.compare_all_common_sheets:
        common_sheet_names = [
            sheet_name
            for sheet_name in left_workbook.sheetnames
            if sheet_name in right_workbook.sheetnames
        ]
        for sheet_name in common_sheet_names:
            left_ws = left_workbook[sheet_name]
            right_ws = right_workbook[sheet_name]
            try:
                left_schema = build_schema(
                    left_ws,
                    ignore_headers=request.ignore_headers,
                )
                right_schema = build_schema(
                    right_ws,
                    ignore_headers=request.ignore_headers,
                )
            except CompareError:
                continue
            pairs.append((left_ws, right_ws, left_schema, right_schema))

        if not pairs:
            raise CompareError("没有找到可自动识别并可比对的同名工作表。")
        return pairs

    if not request.left_sheet_name or not request.right_sheet_name:
        raise CompareError("手动模式下必须选择左右两个工作表。")

    left_ws = left_workbook[request.left_sheet_name]
    right_ws = right_workbook[request.right_sheet_name]
    left_schema = build_schema(
        left_ws,
        manual_header_row=request.left_header_row,
        manual_key_name=request.left_key_name,
        ignore_headers=request.ignore_headers,
    )
    right_schema = build_schema(
        right_ws,
        manual_header_row=request.right_header_row,
        manual_key_name=request.right_key_name,
        ignore_headers=request.ignore_headers,
    )
    pairs.append((left_ws, right_ws, left_schema, right_schema))
    return pairs


def build_output_paths(
    *,
    left_path: Path,
    right_path: Path,
    output_dir: Path,
) -> tuple[Path, Path]:
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    left_output_path = output_dir / f"{left_path.stem}-差异标红-{timestamp}.xlsx"
    right_output_path = output_dir / f"{right_path.stem}-差异标红-{timestamp}.xlsx"
    return left_output_path, right_output_path


def compare_workbooks(request: ComparisonRequest) -> ComparisonResult:
    if not request.left_path.exists():
        raise CompareError(f"左侧文件不存在：{request.left_path}")
    if not request.right_path.exists():
        raise CompareError(f"右侧文件不存在：{request.right_path}")

    left_workbook = load_workbook(request.left_path)
    right_workbook = load_workbook(request.right_path)

    try:
        sheet_pairs = build_sheet_pairs(left_workbook, right_workbook, request)
        summaries: list[SheetComparisonSummary] = []
        compared_sheet_pairs: list[tuple[str, str]] = []

        for left_ws, right_ws, left_schema, right_schema in sheet_pairs:
            summaries.append(
                compare_sheet(
                    left_ws,
                    right_ws,
                    left_schema=left_schema,
                    right_schema=right_schema,
                )
            )
            compared_sheet_pairs.append((left_ws.title, right_ws.title))

        request.output_dir.mkdir(parents=True, exist_ok=True)
        left_output_path, right_output_path = build_output_paths(
            left_path=request.left_path,
            right_path=request.right_path,
            output_dir=request.output_dir,
        )
        left_workbook.save(left_output_path)
        right_workbook.save(right_output_path)

        return ComparisonResult(
            left_output_path=left_output_path,
            right_output_path=right_output_path,
            summaries=tuple(summaries),
            compared_sheet_pairs=tuple(compared_sheet_pairs),
        )
    finally:
        left_workbook.close()
        right_workbook.close()
