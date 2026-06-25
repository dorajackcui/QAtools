#!/usr/bin/env python3
"""Unified GUI launcher for all Excel tools in this repository."""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass
import tkinter as tk
from tkinter import ttk

from tools.chinese_target_checker.check_chinese_target_gui import ChineseTargetCheckerApp
from tools.excel_line_splitter.split_excel_lines_gui import SplitExcelLinesApp
from tools.french_nbsp_restorer.restore_french_nbsp_gui import FrenchNbspRestorerApp
from tools.llm_term_extractor.extract_llm_terms_gui import LlmTermExtractorApp
from tools.tag_placeholder_checker.check_tags_and_placeholders_gui import TagPlaceholderCheckerApp
from tools.term_glossary_checker.check_terms_against_glossary_gui import TermGlossaryCheckerApp
from tools.term_pair_checker.extract_terms_gui import ExtractTermsApp
from tools.workflow.workflow_gui import WorkflowRunnerApp
from tools.xbench_report_transformer.transform_xbench_report_gui import XbenchReportTransformerApp


ToolFactory = Callable[[tk.Misc], ttk.Frame]


@dataclass(frozen=True)
class ToolItem:
    key: str
    title: str
    description: str
    factory: ToolFactory


@dataclass(frozen=True)
class ToolGroup:
    title: str
    tools: tuple[ToolItem, ...]


TOOL_GROUPS = (
    ToolGroup(
        title="常用流程",
        tools=(
            ToolItem(
                key="workflow",
                title="Workflow 编排",
                description="按顺序执行术语对检查和 Tag 检查，统一写入输出 Excel。",
                factory=WorkflowRunnerApp,
            ),
        ),
    ),
    ToolGroup(
        title="术语处理",
        tools=(
            ToolItem(
                key="term_pair",
                title="术语对检查",
                description="检查 source 和 target 中的术语对应关系，并结合历史 TB 优先匹配。",
                factory=ExtractTermsApp,
            ),
            ToolItem(
                key="llm_terms",
                title="LLM 术语提取",
                description="从 Excel 文本中批量抽取术语，并可结合历史 TB 做冲突复核。",
                factory=LlmTermExtractorApp,
            ),
            ToolItem(
                key="glossary",
                title="术语表命中检查",
                description="用术语表检查文本中的 source / target 术语命中情况。",
                factory=TermGlossaryCheckerApp,
            ),
        ),
    ),
    ToolGroup(
        title="质量检查",
        tools=(
            ToolItem(
                key="tag_checker",
                title="Tag 检查",
                description="检查 tag、placeholder、换行标记和数字 tag 在 source / target 中是否一致。",
                factory=TagPlaceholderCheckerApp,
            ),
            ToolItem(
                key="chinese_target",
                title="Target 中文检查",
                description="扫描 target 文本中的中文字符，定位未翻译或混入中文的问题。",
                factory=ChineseTargetCheckerApp,
            ),
            ToolItem(
                key="xbench_report",
                title="Xbench QA 转换",
                description="把 Xbench QA Report 整理为按 key/source 聚合的行级 Excel。",
                factory=XbenchReportTransformerApp,
            ),
        ),
    ),
    ToolGroup(
        title="文本修复",
        tools=(
            ToolItem(
                key="line_splitter",
                title="分行拆列",
                description="把单元格中的多行文本拆分到指定结果列，便于后续检查。",
                factory=SplitExcelLinesApp,
            ),
            ToolItem(
                key="french_nbsp",
                title="法语 NBSP 恢复",
                description="恢复法语标点前的 NBSP，修正常见空格丢失问题。",
                factory=FrenchNbspRestorerApp,
            ),
        ),
    ),
)


class ToolshubApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.withdraw()
        self.root.title("Toolshub")
        self.root.resizable(True, True)
        self.tool_groups = TOOL_GROUPS
        self.tools_by_key = {
            tool.key: tool
            for group in self.tool_groups
            for tool in group.tools
        }
        self.selected_tool_key = tk.StringVar()
        self.current_tool_title = tk.StringVar()
        self.current_tool_description = tk.StringVar()
        self.tool_frames: dict[str, ttk.Frame] = {}
        self.nav_buttons: dict[str, ttk.Radiobutton] = {}
        self.current_tool_frame: ttk.Frame | None = None
        self._build_ui()
        self._fit_window_to_content()

    def _build_ui(self) -> None:
        self._configure_style()
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        shell = ttk.Frame(self.root, padding=16)
        shell.grid(row=0, column=0, sticky="nsew")
        shell.columnconfigure(1, weight=1)
        shell.rowconfigure(0, weight=1)

        sidebar = ttk.Frame(shell, padding=(0, 0, 16, 0))
        sidebar.grid(row=0, column=0, sticky="nsw")
        self._build_sidebar(sidebar)

        workspace = ttk.Frame(shell)
        workspace.grid(row=0, column=1, sticky="nsew")
        workspace.columnconfigure(0, weight=1)
        workspace.rowconfigure(3, weight=1)

        ttk.Label(
            workspace,
            textvariable=self.current_tool_title,
            style="Toolshub.Title.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            workspace,
            textvariable=self.current_tool_description,
            style="Toolshub.Description.TLabel",
            wraplength=760,
        ).grid(row=1, column=0, sticky="ew", pady=(4, 12))
        ttk.Separator(workspace).grid(row=2, column=0, sticky="ew", pady=(0, 12))

        self.content_frame = ttk.Frame(workspace)
        self.content_frame.grid(row=3, column=0, sticky="nsew")
        self.content_frame.columnconfigure(0, weight=1)
        self.content_frame.rowconfigure(0, weight=1)
        self._build_tool_pages()

        first_tool = self.tool_groups[0].tools[0]
        self.select_tool(first_tool.key)

    def _configure_style(self) -> None:
        style = ttk.Style(self.root)
        style.configure("Toolshub.AppTitle.TLabel", font=("TkDefaultFont", 15, "bold"))
        style.configure("Toolshub.Category.TLabel", font=("TkDefaultFont", 10, "bold"))
        style.configure("Toolshub.Title.TLabel", font=("TkDefaultFont", 16, "bold"))
        style.configure("Toolshub.Description.TLabel", foreground="#555555")
        style.configure("Toolshub.Nav.TRadiobutton", padding=(10, 5))

    def _build_sidebar(self, parent: ttk.Frame) -> None:
        ttk.Label(
            parent,
            text="Excel 工具箱",
            style="Toolshub.AppTitle.TLabel",
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))
        ttk.Label(
            parent,
            text="Toolshub",
            style="Toolshub.Description.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(0, 16))

        row = 2
        for group in self.tool_groups:
            ttk.Label(
                parent,
                text=group.title,
                style="Toolshub.Category.TLabel",
            ).grid(row=row, column=0, sticky="w", pady=(12 if row > 2 else 0, 4))
            row += 1
            for tool in group.tools:
                button = ttk.Radiobutton(
                    parent,
                    text=tool.title,
                    value=tool.key,
                    variable=self.selected_tool_key,
                    command=lambda key=tool.key: self.select_tool(key),
                    style="Toolshub.Nav.TRadiobutton",
                )
                button.grid(row=row, column=0, sticky="ew", pady=1)
                self.nav_buttons[tool.key] = button
                row += 1

        parent.columnconfigure(0, minsize=170)

    def _build_tool_pages(self) -> None:
        for group in self.tool_groups:
            for tool in group.tools:
                frame = tool.factory(self.content_frame)
                frame.grid(row=0, column=0, sticky="nsew")
                self.tool_frames[tool.key] = frame

    def select_tool(self, key: str) -> None:
        tool = self.tools_by_key[key]
        frame = self.tool_frames[key]
        self.selected_tool_key.set(key)
        self.current_tool_title.set(tool.title)
        self.current_tool_description.set(tool.description)
        self.current_tool_frame = frame
        frame.tkraise()

    def _fit_window_to_content(self) -> None:
        self.root.update_idletasks()
        width = self.root.winfo_reqwidth()
        height = self.root.winfo_reqheight()
        x = max((self.root.winfo_screenwidth() - width) // 2, 0)
        y = max((self.root.winfo_screenheight() - height) // 2, 30)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.minsize(min(width, 980), min(height, 720))
        self.root.deiconify()


def main() -> None:
    root = tk.Tk()
    ToolshubApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
