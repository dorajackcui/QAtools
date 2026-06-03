#!/usr/bin/env python3
"""Unified GUI launcher for all Excel tools in this repository."""

from __future__ import annotations

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


class ToolshubApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.withdraw()
        self.root.title("Toolshub")
        self.root.resizable(False, False)
        self._build_ui()
        self._fit_window_to_content()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")

        ttk.Label(
            frame,
            text="Excel 工具箱",
            font=("TkDefaultFont", 12, "bold"),
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            frame,
            text="统一入口管理 workflow 编排、术语对检查、LLM术语提取、术语表命中检查、tag检查、target中文检查、分行拆列和法语 NBSP 恢复。",
        ).grid(row=1, column=0, sticky="w", pady=(4, 10))

        notebook = ttk.Notebook(frame)
        notebook.grid(row=2, column=0, sticky="nsew")

        workflow_tab = WorkflowRunnerApp(notebook)
        term_pair_tab = ExtractTermsApp(notebook)
        llm_term_tab = LlmTermExtractorApp(notebook)
        glossary_tab = TermGlossaryCheckerApp(notebook)
        tag_checker_tab = TagPlaceholderCheckerApp(notebook)
        chinese_target_tab = ChineseTargetCheckerApp(notebook)
        splitter_tab = SplitExcelLinesApp(notebook)
        french_nbsp_tab = FrenchNbspRestorerApp(notebook)

        notebook.add(workflow_tab, text="Workflow编排")
        notebook.add(term_pair_tab, text="术语对检查")
        notebook.add(llm_term_tab, text="LLM术语提取")
        notebook.add(glossary_tab, text="术语表命中检查")
        notebook.add(tag_checker_tab, text="Tag检查")
        notebook.add(chinese_target_tab, text="Target中文检查")
        notebook.add(splitter_tab, text="分行拆列")
        notebook.add(french_nbsp_tab, text="法语NBSP恢复")

        frame.columnconfigure(0, weight=1)

    def _fit_window_to_content(self) -> None:
        self.root.update_idletasks()
        width = self.root.winfo_reqwidth()
        height = self.root.winfo_reqheight()
        x = max((self.root.winfo_screenwidth() - width) // 2, 0)
        y = max((self.root.winfo_screenheight() - height) // 2, 30)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.deiconify()


def main() -> None:
    root = tk.Tk()
    ToolshubApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
