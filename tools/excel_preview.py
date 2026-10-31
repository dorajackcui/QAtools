"""Small OOXML previews without loading workbook styles or all shared strings."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import posixpath
import random
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from tools.excel_file_ops import excel_files

MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PREVIEW_ROWS = 6
SHEET_BYTE_LIMIT = 2 * 1024 * 1024
STRINGS_BYTE_LIMIT = 8 * 1024 * 1024
TEXT_LIMIT = 1000


@dataclass(frozen=True)
class ExcelPreview:
    path: Path
    sheet: str
    sheet_names: tuple[str, ...]
    cells: dict[tuple[int, int], str]
    columns: int
    note: str = ""
    headers: dict[int, str] = field(default_factory=dict)
    sheet_fallback: bool = False


class _ReadLimit(Exception):
    pass


def _events(archive, name, limit):
    """Read a bounded prefix, even when a single XML row/string is enormous."""
    parser = ET.XMLPullParser(events=("start", "end"))
    with archive.open(name) as stream:
        consumed = 0
        while consumed < limit:
            chunk = stream.read(min(16384, limit - consumed))
            if not chunk:
                parser.close()
                return
            consumed += len(chunk)
            parser.feed(chunk)
            yield from parser.read_events()
        raise _ReadLimit


def _text(node):
    if node is None:
        return ""
    # Rich text uses r/t; phonetic annotations are not displayed cell text.
    return (node.findtext(f"{MAIN}t", "") +
            "".join(run.findtext(f"{MAIN}t", "") for run in node.findall(f"{MAIN}r")))


def _header_text(value, kind, formula):
    if not formula:
        if kind == "b":
            return "True" if value == "TRUE" else "False"
        if kind in {None, "n"}:
            try:
                return str(float(value) if any(char in value for char in ".eE") else int(value))
            except ValueError:
                pass
    return value


def _column(reference):
    value = 0
    for char in reference.upper():
        if not "A" <= char <= "Z":
            break
        value = value * 26 + ord(char) - ord("A") + 1
    return value


def read_excel_preview(path: str, sheet: str = "", *, first_sheet: bool = False,
                       folder: bool = False, exclude: tuple[str, ...] = (),
                       exclude_names: tuple[str, ...] = (), skip_links: bool = False,
                       allow_sheet_fallback: bool = False) -> ExcelPreview:
    source = Path(path).expanduser().resolve()
    if folder:
        files, _ = excel_files(source, extensions={".xlsx", ".xlsm"},
                               exclude=tuple(Path(item) for item in exclude if item), skip_links=skip_links)
        files = [item for item in files if item.name.casefold() not in exclude_names]
        if not files:
            raise ValueError("目录中没有可预览的 .xlsx/.xlsm 文件。")
        source = random.choice(files)
    if source.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError("轻量预览支持 .xlsx/.xlsm 文件。")

    cells, headers, shared, notes = {}, {}, {}, []
    sheet_fallback = False
    with ZipFile(source) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        all_sheets = workbook.findall(f"{MAIN}sheets/{MAIN}sheet")
        relations = {node.attrib["Id"]: node for node in
                     ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))}
        # Chartsheets are tabs too, but column selection requires a data sheet.
        sheets = [node for node in all_sheets
                  if node.attrib[f"{REL}id"] in relations and
                  relations[node.attrib[f"{REL}id"]].attrib.get("Type", "").endswith("/worksheet")]
        names = tuple(node.attrib["name"] for node in sheets)
        if not sheets:
            raise ValueError("工作簿没有数据工作表。")
        if sheet and sheet not in names:
            if not allow_sheet_fallback:
                raise ValueError(f"工作表不存在：{sheet}")
            notes.append(f"原配置工作表不存在：{sheet}。已显示活动数据表，请重新确认工作表和列。")
            sheet, sheet_fallback = "", True
        if sheet:
            selected = sheets[names.index(sheet)]
        elif first_sheet:
            selected = sheets[0]
        else:
            view = workbook.find(f"{MAIN}bookViews/{MAIN}workbookView")
            try:
                active = int(view.attrib.get("activeTab", "0")) if view is not None else 0
            except ValueError:
                active = 0
            visible = [node for node in sheets if node.attrib.get("state", "visible") == "visible"]
            selected = all_sheets[active] if 0 <= active < len(all_sheets) else sheets[0]
            if selected not in visible:
                selected = next(iter(visible), sheets[0])
        relationship = relations[selected.attrib[f"{REL}id"]]
        target = relationship.attrib["Target"]
        part = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
        row_number = 0
        events = _events(archive, part, SHEET_BYTE_LIMIT)
        try:
            for event, node in events:
                if node.tag != f"{MAIN}row":
                    continue
                if event == "start":
                    row_number = int(node.attrib.get("r", row_number + 1))
                    if row_number > PREVIEW_ROWS:
                        break
                    continue
                column = 0
                for cell in node.findall(f"{MAIN}c"):
                    column = _column(cell.attrib.get("r", "")) or column + 1
                    if not 1 <= column <= 16384:
                        continue
                    value = cell.findtext(f"{MAIN}v", "")
                    formula = cell.find(f"{MAIN}f")
                    kind = cell.attrib.get("t")
                    key = (row_number - 1, column - 1)
                    if formula is not None:
                        value = "=" + (formula.text or "")
                    elif kind == "s":
                        shared[key] = int(value)
                        continue
                    elif kind == "inlineStr":
                        value = _text(cell.find(f"{MAIN}is"))
                    elif kind == "b":
                        value = "TRUE" if value == "1" else "FALSE"
                    if value:
                        cells[key] = value[:TEXT_LIMIT]
                        if row_number == 1:
                            headers[column - 1] = _header_text(value, kind, formula is not None)
                node.clear()
                if row_number == PREVIEW_ROWS:
                    break
        except _ReadLimit:
            notes.append("数据较大，仅展示已读取的部分。")
        finally:
            events.close()

        if shared:
            wanted = set(shared.values())
            strings = {}
            events = _events(archive, "xl/sharedStrings.xml", STRINGS_BYTE_LIMIT)
            try:
                index = 0
                for event, node in events:
                    if event == "end" and node.tag == f"{MAIN}si":
                        if index in wanted:
                            strings[index] = _text(node)
                            wanted.remove(index)
                        node.clear()
                        index += 1
                        if not wanted:
                            break
            except _ReadLimit:
                pass
            finally:
                events.close()
            for key, index in shared.items():
                value = strings.get(index, "[文本未展开]")
                cells[key] = value[:TEXT_LIMIT]
                if key[0] == 0 and index in strings:
                    headers[key[1]] = value
            if wanted:
                notes.append("共享文本较大，部分文本未展开。")

    return ExcelPreview(source, selected.attrib["name"], names, cells,
                        max((column + 1 for _, column in cells), default=0), " ".join(notes),
                        headers, sheet_fallback)
