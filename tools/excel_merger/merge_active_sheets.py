#!/usr/bin/env python3
"""Merge the active worksheet from every supported workbook in a folder."""

from __future__ import annotations

import argparse
from collections import OrderedDict
from collections.abc import Callable, Iterable, Sequence
from dataclasses import dataclass
from datetime import datetime, timezone
import io
import os
from pathlib import Path
import posixpath
import re
import sys
import tempfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape
import zipfile


REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SUPPORTED_EXTENSIONS = frozenset({".xlsx", ".xlsm"})
SKIPPED_EXTENSIONS = frozenset({".xls", ".xlsb"})
CELL_REF_RE = re.compile(r"^([A-Z]+)", re.IGNORECASE)
EXCEL_MAX_ROWS = 1_048_576
EXCEL_MAX_COLUMNS = 16_384

ProgressCallback = Callable[[int, int], None]


@dataclass(frozen=True)
class MergeSummary:
    """Result details for one active-sheet merge operation."""

    output_path: Path
    supported_file_count: int
    skipped_file_count: int
    failed_file_count: int
    merged_row_count: int
    error_log_path: Path | None


def column_letters_to_index(letters: str) -> int:
    value = 0
    for char in letters.upper():
        value = (value * 26) + (ord(char) - ord("A") + 1)
    return value


def column_index_to_letters(index: int) -> str:
    letters: list[str] = []
    remaining = index
    while remaining > 0:
        remaining -= 1
        letters.append(chr(ord("A") + (remaining % 26)))
        remaining //= 26
    return "".join(reversed(letters))


def local_name(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def extract_excel_text(string_item: ET.Element) -> str:
    """Read plain or rich-text content without XML formatting whitespace."""

    parts: list[str] = []
    for child in string_item:
        child_name = local_name(child.tag)
        if child_name == "t":
            parts.append(child.text or "")
        elif child_name == "r":
            for run_child in child:
                if local_name(run_child.tag) == "t":
                    parts.append(run_child.text or "")
    return "".join(parts)


def iter_excel_files(folder_path: Path) -> tuple[list[Path], list[Path]]:
    """Recursively list supported workbooks and known unsupported formats."""

    recognized_extensions = SUPPORTED_EXTENSIONS | SKIPPED_EXTENSIONS
    all_files = sorted(
        path
        for path in folder_path.rglob("*")
        if path.is_file()
        and not path.name.startswith("~$")
        and path.suffix.lower() in recognized_extensions
    )
    supported = [
        path
        for path in all_files
        if path.suffix.lower() in SUPPORTED_EXTENSIONS
    ]
    skipped = [
        path
        for path in all_files
        if path.suffix.lower() in SKIPPED_EXTENSIONS
    ]
    return supported, skipped


def read_xml(zip_file: zipfile.ZipFile, entry_path: str) -> ET.Element:
    try:
        data = zip_file.read(entry_path)
    except KeyError as exc:
        raise RuntimeError(f"Excel 文件缺少条目: {entry_path}") from exc
    return ET.fromstring(data)


def read_optional_xml(
    zip_file: zipfile.ZipFile,
    entry_path: str,
) -> ET.Element | None:
    try:
        data = zip_file.read(entry_path)
    except KeyError:
        return None
    return ET.fromstring(data)


def read_shared_strings(zip_file: zipfile.ZipFile) -> list[str]:
    root = read_optional_xml(zip_file, "xl/sharedStrings.xml")
    if root is None:
        return []

    return [
        extract_excel_text(child)
        for child in root
        if local_name(child.tag) == "si"
    ]


def resolve_active_sheet(zip_file: zipfile.ZipFile) -> tuple[str, str]:
    """Return the active worksheet title and its package-relative XML path."""

    workbook = read_xml(zip_file, "xl/workbook.xml")
    relationships = read_xml(zip_file, "xl/_rels/workbook.xml.rels")

    active_index = 0
    for node in workbook.iter():
        if local_name(node.tag) != "workbookView":
            continue
        raw_active_tab = node.attrib.get("activeTab")
        if raw_active_tab:
            try:
                active_index = int(raw_active_tab)
            except ValueError:
                active_index = 0
        break

    sheets = [
        node
        for node in workbook.iter()
        if local_name(node.tag) == "sheet"
    ]
    if not sheets:
        raise RuntimeError("工作簿中没有工作表。")
    if active_index < 0 or active_index >= len(sheets):
        active_index = 0

    sheet_node = sheets[active_index]
    relationship_id = sheet_node.attrib.get(f"{{{REL_NS}}}id")
    if not relationship_id:
        raise RuntimeError("找不到活动工作表的关系 ID。")

    target = None
    for relationship in relationships.iter():
        if local_name(relationship.tag) != "Relationship":
            continue
        if relationship.attrib.get("Id") == relationship_id:
            target = relationship.attrib.get("Target")
            break
    if not target:
        raise RuntimeError(f"找不到工作表关系: {relationship_id}")

    sheet_name = sheet_node.attrib.get("name", "Sheet1")
    normalized_target = target.replace("\\", "/").lstrip("/")
    if normalized_target.startswith("xl/"):
        sheet_path = posixpath.normpath(normalized_target)
    else:
        sheet_path = posixpath.normpath(
            posixpath.join("xl", normalized_target)
        )
    return sheet_name, sheet_path


def get_cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t", "")
    if cell_type == "inlineStr":
        for child in cell:
            if local_name(child.tag) == "is":
                return extract_excel_text(child)
        return ""

    value_text = ""
    for child in cell:
        if local_name(child.tag) == "v":
            value_text = child.text or ""
            break
    if not value_text:
        return ""

    if cell_type == "s":
        try:
            index = int(value_text)
        except ValueError:
            return ""
        if 0 <= index < len(shared_strings):
            return shared_strings[index]
        return ""
    if cell_type == "b":
        return "TRUE" if value_text == "1" else "FALSE"
    return value_text


def read_sheet_rows(
    zip_file: zipfile.ZipFile,
    sheet_path: str,
    shared_strings: list[str],
) -> list[list[str]]:
    try:
        raw_xml = zip_file.read(sheet_path)
    except KeyError as exc:
        raise RuntimeError(f"Excel 文件缺少工作表条目: {sheet_path}") from exc

    rows: list[list[str]] = []
    current_row_values: dict[int, str] | None = None
    for event, element in ET.iterparse(
        io.BytesIO(raw_xml),
        events=("start", "end"),
    ):
        tag = local_name(element.tag)
        if event == "start" and tag == "row":
            current_row_values = {}
            continue
        if event == "end" and tag == "c" and current_row_values is not None:
            reference = element.attrib.get("r", "")
            match = CELL_REF_RE.match(reference)
            if match:
                column_index = column_letters_to_index(match.group(1))
                current_row_values[column_index] = get_cell_value(
                    element,
                    shared_strings,
                )
            element.clear()
            continue
        if event == "end" and tag == "row":
            if current_row_values:
                row_values = [""] * max(current_row_values)
                for column_index, value in current_row_values.items():
                    row_values[column_index - 1] = value
                rows.append(row_values)
            current_row_values = None
            element.clear()
    return rows


def build_default_output_path(
    folder_path: str | Path,
    *,
    now: datetime | None = None,
) -> Path:
    folder = Path(folder_path).resolve()
    timestamp = (now or datetime.now()).strftime("%Y%m%d_%H%M%S")
    return folder.parent / (
        f"{folder.name}_merged_active_sheet_{timestamp}.xlsx"
    )


def register_shared_string(
    value: str,
    shared_string_map: OrderedDict[str, int],
) -> int:
    if value not in shared_string_map:
        shared_string_map[value] = len(shared_string_map)
    return shared_string_map[value]


def build_merged_rows(
    excel_files: Iterable[Path],
    keep_all_headers: bool,
    *,
    progress_callback: ProgressCallback | None = None,
) -> tuple[list[list[str]], OrderedDict[str, int], list[str]]:
    """Read workbooks and build rows for the merged worksheet."""

    files = list(excel_files)
    merged_rows: list[list[str]] = []
    shared_string_map: OrderedDict[str, int] = OrderedDict()
    failed_files: list[str] = []
    copied_sheet_count = 0

    for completed, file_path in enumerate(files, start=1):
        sheet_name = "<未解析>"
        sheet_path = "<未解析>"
        try:
            with zipfile.ZipFile(file_path, "r") as zip_file:
                shared_strings = read_shared_strings(zip_file)
                sheet_name, sheet_path = resolve_active_sheet(zip_file)
                sheet_rows = read_sheet_rows(
                    zip_file,
                    sheet_path,
                    shared_strings,
                )
        except Exception as exc:  # noqa: BLE001
            failed_files.append(
                f"读取失败: {file_path} "
                f"(活动工作表: {sheet_name}, 路径: {sheet_path}): {exc}"
            )
            if progress_callback is not None:
                progress_callback(completed, len(files))
            continue

        if sheet_rows:
            start_index = (
                0 if copied_sheet_count == 0 or keep_all_headers else 1
            )
            for row_index in range(start_index, len(sheet_rows)):
                source_row = list(sheet_rows[row_index])
                if row_index == 0:
                    output_row = ["SourceFile", *source_row]
                else:
                    output_row = [file_path.name, *source_row]
                merged_rows.append(output_row)
                for value in output_row:
                    if value != "":
                        register_shared_string(value, shared_string_map)
            copied_sheet_count += 1

        if progress_callback is not None:
            progress_callback(completed, len(files))

    if not merged_rows:
        raise RuntimeError("没有可合并的工作表数据。")
    return merged_rows, shared_string_map, failed_files


def write_text_entry(
    zip_file: zipfile.ZipFile,
    entry_name: str,
    text: str,
) -> None:
    zip_file.writestr(entry_name, text.encode("utf-8"))


def write_shared_strings(
    zip_file: zipfile.ZipFile,
    shared_string_map: OrderedDict[str, int],
) -> None:
    with zip_file.open("xl/sharedStrings.xml", "w") as stream:
        stream.write(
            (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<sst xmlns="http://schemas.openxmlformats.org/'
                'spreadsheetml/2006/main" '
                f'count="{len(shared_string_map)}" '
                f'uniqueCount="{len(shared_string_map)}">'
            ).encode("utf-8")
        )
        for value in shared_string_map:
            stream.write(
                (
                    f'<si><t xml:space="preserve">'
                    f"{escape(value)}</t></si>"
                ).encode("utf-8")
            )
        stream.write(b"</sst>")


def write_sheet_xml(
    zip_file: zipfile.ZipFile,
    merged_rows: list[list[str]],
    shared_string_map: OrderedDict[str, int],
) -> None:
    max_columns = max(len(row) for row in merged_rows)
    last_cell = (
        f"{column_index_to_letters(max_columns)}{len(merged_rows)}"
    )
    with zip_file.open("xl/worksheets/sheet1.xml", "w") as stream:
        stream.write(
            (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<worksheet xmlns="http://schemas.openxmlformats.org/'
                'spreadsheetml/2006/main">'
                f'<dimension ref="A1:{last_cell}"/><sheetData>'
            ).encode("utf-8")
        )
        for row_number, row in enumerate(merged_rows, start=1):
            stream.write(f'<row r="{row_number}">'.encode("utf-8"))
            for column_number, value in enumerate(row, start=1):
                if value == "":
                    continue
                shared_index = shared_string_map[value]
                cell_ref = (
                    f"{column_index_to_letters(column_number)}{row_number}"
                )
                stream.write(
                    (
                        f'<c r="{cell_ref}" t="s">'
                        f"<v>{shared_index}</v></c>"
                    ).encode("utf-8")
                )
            stream.write(b"</row>")
        stream.write(b"</sheetData></worksheet>")


def validate_merged_dimensions(
    merged_rows: Sequence[Sequence[str]],
) -> None:
    row_count = len(merged_rows)
    if row_count > EXCEL_MAX_ROWS:
        raise ValueError(
            f"合并结果共有 {row_count} 行，超过 Excel 上限 "
            f"{EXCEL_MAX_ROWS} 行。"
        )

    max_columns = max((len(row) for row in merged_rows), default=0)
    if max_columns > EXCEL_MAX_COLUMNS:
        raise ValueError(
            f"添加 SourceFile 后合并结果共有 {max_columns} 列，超过 Excel 上限 "
            f"{EXCEL_MAX_COLUMNS} 列。"
        )


def _write_output_xlsx_archive(
    output_path: Path,
    merged_rows: list[list[str]],
    shared_string_map: OrderedDict[str, int],
) -> None:
    core_timestamp = datetime.now(timezone.utc).strftime(
        "%Y-%m-%dT%H:%M:%SZ"
    )
    with zipfile.ZipFile(
        output_path,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=6,
    ) as zip_file:
        write_text_entry(
            zip_file,
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>""",
        )
        write_text_entry(
            zip_file,
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>""",
        )
        write_text_entry(
            zip_file,
            "docProps/app.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>QAtools Active Sheet Merger</Application>
</Properties>""",
        )
        write_text_entry(
            zip_file,
            "docProps/core.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>QAtools</dc:creator>
  <cp:lastModifiedBy>QAtools</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{core_timestamp}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{core_timestamp}</dcterms:modified>
</cp:coreProperties>""",
        )
        write_text_entry(
            zip_file,
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <bookViews><workbookView activeTab="0"/></bookViews>
  <sheets><sheet name="MergedData" sheetId="1" r:id="rId1"/></sheets>
</workbook>""",
        )
        write_text_entry(
            zip_file,
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>""",
        )
        write_text_entry(
            zip_file,
            "xl/styles.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>""",
        )
        write_shared_strings(zip_file, shared_string_map)
        write_sheet_xml(zip_file, merged_rows, shared_string_map)


def write_output_xlsx(
    output_path: Path,
    merged_rows: list[list[str]],
    shared_string_map: OrderedDict[str, int],
) -> None:
    """Atomically replace the output only after a valid workbook is complete."""

    validate_merged_dimensions(merged_rows)
    with tempfile.NamedTemporaryFile(
        dir=output_path.parent,
        prefix=f".{output_path.name}.",
        suffix=".tmp",
        delete=False,
    ) as temporary_file:
        temporary_path = Path(temporary_file.name)

    try:
        _write_output_xlsx_archive(
            temporary_path,
            merged_rows,
            shared_string_map,
        )
        os.replace(temporary_path, output_path)
    except Exception:
        temporary_path.unlink(missing_ok=True)
        raise


def write_failure_log(
    output_path: Path,
    failed_messages: list[str],
) -> Path | None:
    if not failed_messages:
        return None
    log_path = output_path.with_name(f"{output_path.stem}_errors.txt")
    lines = [
        "QAtools 合并表格错误日志",
        f"生成时间: {datetime.now().isoformat(timespec='seconds')}",
        "",
        *failed_messages,
    ]
    log_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return log_path


def merge_active_sheets(
    folder_path: str | Path,
    *,
    output_path: str | Path | None = None,
    keep_all_headers: bool = False,
    progress_callback: ProgressCallback | None = None,
) -> MergeSummary:
    """Merge active sheets and write a new workbook next to the input folder."""

    folder = Path(folder_path).resolve()
    if not folder.is_dir():
        raise ValueError(f"目录不存在: {folder}")

    supported_files, skipped_files = iter_excel_files(folder)
    resolved_output = (
        Path(output_path).resolve()
        if output_path is not None
        else build_default_output_path(folder)
    )
    if resolved_output.suffix.lower() != ".xlsx":
        raise ValueError("输出文件必须使用 .xlsx 扩展名。")
    source_paths = {path.resolve() for path in supported_files}
    if resolved_output in source_paths:
        raise ValueError(
            f"输出文件不能覆盖待合并的源工作簿: {resolved_output}"
        )
    if not supported_files and not skipped_files:
        raise ValueError(f"目录中没有 Excel 文件: {folder}")
    if not supported_files:
        raise ValueError(
            f"目录中没有支持的 .xlsx/.xlsm 文件: {folder}"
        )
    if not resolved_output.parent.is_dir():
        raise ValueError(f"输出目录不存在: {resolved_output.parent}")

    merged_rows, shared_string_map, failed_files = build_merged_rows(
        supported_files,
        keep_all_headers,
        progress_callback=progress_callback,
    )
    write_output_xlsx(resolved_output, merged_rows, shared_string_map)
    error_log_path = write_failure_log(resolved_output, failed_files)
    return MergeSummary(
        output_path=resolved_output,
        supported_file_count=len(supported_files),
        skipped_file_count=len(skipped_files),
        failed_file_count=len(failed_files),
        merged_row_count=len(merged_rows),
        error_log_path=error_log_path,
    )


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "递归合并目录内所有 .xlsx/.xlsm 文件的活动工作表，"
            "并添加 SourceFile 列。"
        )
    )
    parser.add_argument(
        "folder",
        nargs="?",
        help="包含待合并 Excel 文件的目录",
    )
    parser.add_argument(
        "--folder-path",
        dest="folder_path",
        help="目录路径（兼容原 mergesSheets 参数）",
    )
    parser.add_argument(
        "-o",
        "--output",
        "--output-path",
        dest="output_path",
        help="输出 .xlsx；默认生成在输入目录的上一级",
    )
    parser.add_argument(
        "--keep-all-headers",
        action="store_true",
        help="保留每个活动工作表的第一行；默认只保留第一份表头",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_argument_parser()
    args = parser.parse_args(argv)
    folder_path = args.folder_path or args.folder
    if not folder_path:
        parser.error("请提供待合并 Excel 文件所在目录。")

    try:
        summary = merge_active_sheets(
            folder_path,
            output_path=args.output_path,
            keep_all_headers=args.keep_all_headers,
        )
    except (OSError, RuntimeError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1
    print(f"合并完成: {summary.output_path}")
    print(f"支持的输入文件: {summary.supported_file_count}")
    print(f"输出行数: {summary.merged_row_count}")
    if summary.skipped_file_count:
        print(
            f"跳过 .xls/.xlsb 文件: {summary.skipped_file_count}",
            file=sys.stderr,
        )
    if summary.failed_file_count:
        print(
            f"读取失败的文件: {summary.failed_file_count}",
            file=sys.stderr,
        )
        print(f"错误日志: {summary.error_log_path}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
