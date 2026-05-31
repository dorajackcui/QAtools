from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from openpyxl import Workbook

from tools.chinese_target_checker.check_chinese_target_gui import ChineseTargetCheckerApp
from tools.excel_line_splitter.split_excel_lines_gui import SplitExcelLinesApp
from tools.french_nbsp_restorer.restore_french_nbsp_gui import FrenchNbspRestorerApp
from tools.llm_term_extractor.extract_llm_terms_gui import LlmTermExtractorApp
from tools.tag_placeholder_checker.check_tags_and_placeholders_gui import TagPlaceholderCheckerApp
from tools.term_glossary_checker.check_terms_against_glossary_gui import TermGlossaryCheckerApp
from tools.term_pair_checker.extract_terms_gui import ExtractTermsApp
from tools.workflow.workflow_gui import WorkflowRunnerApp


class FakeVar:
    def __init__(self, value: str = "") -> None:
        self.value = value

    def get(self) -> str:
        return self.value

    def set(self, value: str) -> None:
        self.value = value


class FakeBoolVar:
    def __init__(self, value: bool = False) -> None:
        self.value = value

    def get(self) -> bool:
        return self.value

    def set(self, value: bool) -> None:
        self.value = value


class FakeCombobox(dict):
    def __init__(self) -> None:
        super().__init__()
        self["values"] = ()


class GuiSheetSelectionTests(unittest.TestCase):
    def create_term_pair_workbook(self, path: Path) -> None:
        workbook = Workbook()
        data_sheet = workbook.active
        data_sheet.title = "Data"
        data_sheet["A1"] = "source"
        data_sheet["B1"] = "target"

        alternate_sheet = workbook.create_sheet("Alternate")
        alternate_sheet["C1"] = " SOURCE "
        alternate_sheet["E1"] = " target "

        missing_header_sheet = workbook.create_sheet("NoHeader")
        missing_header_sheet["A1"] = "foo"
        missing_header_sheet["B1"] = "bar"

        workbook.active = 1
        workbook.save(path)

    def create_history_tb_workbook(self, path: Path) -> None:
        workbook = Workbook()
        data_sheet = workbook.active
        data_sheet.title = "Raw"
        data_sheet["A1"] = "source"
        data_sheet["B1"] = "target"

        term_sheet = workbook.create_sheet("术语表")
        term_sheet["A1"] = "source术语"
        term_sheet["B1"] = "target术语"
        term_sheet["C1"] = "source术语（无mark）"
        term_sheet["D1"] = "target术语（无mark）"

        workbook.active = 0
        workbook.save(path)

    def create_offset_history_tb_workbook(self, path: Path) -> None:
        workbook = Workbook()
        term_sheet = workbook.active
        term_sheet.title = "术语表"
        term_sheet["A1"] = "metadata"
        term_sheet["B1"] = "notes"
        term_sheet["C2"] = "source术语（无mark）"
        term_sheet["D2"] = "target术语（无mark）"
        term_sheet["C3"] = "花艺"
        term_sheet["D3"] = "Art Floral"
        workbook.save(path)

    def create_glossary_workbook(self, path: Path) -> None:
        workbook = Workbook()
        glossary_sheet = workbook.active
        glossary_sheet.title = "Glossary"
        glossary_sheet["B1"] = "source"
        glossary_sheet["D1"] = "target"

        backup_sheet = workbook.create_sheet("Backup")
        backup_sheet["A1"] = "source"
        backup_sheet["C1"] = "target"

        workbook.active = 0
        workbook.save(path)

    def create_data_workbook(self, path: Path) -> None:
        workbook = Workbook()
        data_sheet = workbook.active
        data_sheet.title = "Data"
        data_sheet["C1"] = "source"
        data_sheet["F1"] = "target"

        archive_sheet = workbook.create_sheet("Archive")
        archive_sheet["A1"] = "source"
        archive_sheet["B1"] = "target"

        workbook.active = 1
        workbook.save(path)

    def create_source_only_workbook(self, path: Path) -> None:
        workbook = Workbook()
        data_sheet = workbook.active
        data_sheet.title = "SourceOnly"
        data_sheet["A1"] = "source"
        data_sheet["B1"] = "note"
        workbook.save(path)

    def create_splitter_workbook(self, path: Path) -> None:
        workbook = Workbook()
        data_sheet = workbook.active
        data_sheet.title = "Split"
        workbook.create_sheet("Other")
        workbook.active = 0
        workbook.save(path)

    def create_tag_checker_workbook(self, path: Path) -> None:
        workbook = Workbook()
        data_sheet = workbook.active
        data_sheet.title = "Tags"
        data_sheet["D1"] = "source"
        data_sheet["F1"] = "target"

        archive_sheet = workbook.create_sheet("Archive")
        archive_sheet["A1"] = "source"
        archive_sheet["B1"] = "target"

        workbook.active = 0
        workbook.save(path)

    def build_extract_terms_app(self, input_path: Path) -> ExtractTermsApp:
        app = ExtractTermsApp.__new__(ExtractTermsApp)
        app.input_file_var = FakeVar(str(input_path))
        app.source_column_var = FakeVar("A")
        app.target_column_var = FakeVar("B")
        app.sheet_var = FakeVar("")
        app.sheet_combobox = FakeCombobox()
        return app

    def build_extract_terms_app_with_history(self, history_path: Path) -> ExtractTermsApp:
        app = ExtractTermsApp.__new__(ExtractTermsApp)
        app.history_tb_file_var = FakeVar(str(history_path))
        app.history_sheet_var = FakeVar("")
        app.history_source_column_var = FakeVar("")
        app.history_target_column_var = FakeVar("")
        app.history_sheet_combobox = FakeCombobox()
        return app

    def build_glossary_checker_app(self, glossary_path: Path, data_path: Path) -> TermGlossaryCheckerApp:
        app = TermGlossaryCheckerApp.__new__(TermGlossaryCheckerApp)
        app.glossary_file_var = FakeVar(str(glossary_path))
        app.data_file_var = FakeVar(str(data_path))
        app.glossary_sheet_var = FakeVar("")
        app.data_sheet_var = FakeVar("")
        app.glossary_source_column_var = FakeVar("A")
        app.glossary_target_column_var = FakeVar("B")
        app.data_source_column_var = FakeVar("A")
        app.data_target_column_var = FakeVar("B")
        app.glossary_sheet_combobox = FakeCombobox()
        app.data_sheet_combobox = FakeCombobox()
        return app

    def build_splitter_app(self, input_path: Path) -> SplitExcelLinesApp:
        app = SplitExcelLinesApp.__new__(SplitExcelLinesApp)
        app.input_file_var = FakeVar(str(input_path))
        app.output_file_var = FakeVar("")
        app.sheet_var = FakeVar("")
        app.sheet_combobox = FakeCombobox()
        return app

    def build_tag_checker_app(self, input_path: Path) -> TagPlaceholderCheckerApp:
        app = TagPlaceholderCheckerApp.__new__(TagPlaceholderCheckerApp)
        app.input_file_var = FakeVar(str(input_path))
        app.output_file_var = FakeVar("")
        app.sheet_var = FakeVar("")
        app.source_column_var = FakeVar("A")
        app.target_column_var = FakeVar("B")
        app.sheet_combobox = FakeCombobox()
        return app

    def build_french_nbsp_restorer_app(self, input_path: Path) -> FrenchNbspRestorerApp:
        app = FrenchNbspRestorerApp.__new__(FrenchNbspRestorerApp)
        app.input_file_var = FakeVar(str(input_path))
        app.output_file_var = FakeVar("")
        app.sheet_var = FakeVar("")
        app.target_column_var = FakeVar("B")
        app.result_column_var = FakeVar("")
        app.sheet_combobox = FakeCombobox()
        return app

    def build_chinese_target_checker_app(self, input_path: Path) -> ChineseTargetCheckerApp:
        app = ChineseTargetCheckerApp.__new__(ChineseTargetCheckerApp)
        app.input_file_var = FakeVar(str(input_path))
        app.output_file_var = FakeVar("")
        app.sheet_var = FakeVar("")
        app.target_column_var = FakeVar("B")
        app.result_column_var = FakeVar("")
        app.sheet_combobox = FakeCombobox()
        return app

    def build_workflow_app(self, input_path: Path) -> WorkflowRunnerApp:
        app = WorkflowRunnerApp.__new__(WorkflowRunnerApp)
        app.input_file_var = FakeVar(str(input_path))
        app.output_file_var = FakeVar("")
        app.sheet_var = FakeVar("")
        app.source_column_var = FakeVar("A")
        app.target_column_var = FakeVar("B")
        app.sheet_combobox = FakeCombobox()
        return app

    def build_workflow_app_with_history(self, history_path: Path) -> WorkflowRunnerApp:
        app = WorkflowRunnerApp.__new__(WorkflowRunnerApp)
        app.term_history_tb_file_var = FakeVar(str(history_path))
        app.term_history_sheet_var = FakeVar("")
        app.term_history_source_column_var = FakeVar("")
        app.term_history_target_column_var = FakeVar("")
        app.term_history_sheet_combobox = FakeCombobox()
        return app

    def build_llm_term_extractor_app(self, input_path: Path) -> LlmTermExtractorApp:
        app = LlmTermExtractorApp.__new__(LlmTermExtractorApp)
        app.input_file_var = FakeVar(str(input_path))
        app.output_file_var = FakeVar("")
        app.sheet_var = FakeVar("")
        app.source_column_var = FakeVar("A")
        app.target_column_var = FakeVar("")
        app.sheet_combobox = FakeCombobox()
        return app

    def build_llm_term_extractor_app_with_history(self, history_path: Path) -> LlmTermExtractorApp:
        app = LlmTermExtractorApp.__new__(LlmTermExtractorApp)
        app.history_tb_file_var = FakeVar(str(history_path))
        app.history_sheet_var = FakeVar("")
        app.history_source_column_var = FakeVar("")
        app.history_target_column_var = FakeVar("")
        app.history_start_row_var = FakeVar("2")
        app.history_sheet_combobox = FakeCombobox()
        return app

    def test_term_pair_refresh_populates_sheet_choices_and_detects_columns(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "terms.xlsx"
            self.create_term_pair_workbook(workbook_path)
            app = self.build_extract_terms_app(workbook_path)

            app.refresh_sheet_choices(show_error=False)

            self.assertEqual(app.sheet_combobox["values"], ("Data", "Alternate", "NoHeader"))
            self.assertEqual(app.sheet_var.get(), "Alternate")
            self.assertEqual(app.source_column_var.get(), "C")
            self.assertEqual(app.target_column_var.get(), "E")

    def test_term_pair_switching_sheet_redetects_columns(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "terms.xlsx"
            self.create_term_pair_workbook(workbook_path)
            app = self.build_extract_terms_app(workbook_path)

            app.refresh_sheet_choices(show_error=False)
            app.sheet_var.set("Data")
            app.handle_sheet_selected(show_error=False)

            self.assertEqual(app.source_column_var.get(), "A")
            self.assertEqual(app.target_column_var.get(), "B")

    def test_term_pair_missing_detection_keeps_manual_columns(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "terms.xlsx"
            self.create_term_pair_workbook(workbook_path)
            app = self.build_extract_terms_app(workbook_path)
            app.source_column_var.set("X")
            app.target_column_var.set("Y")

            app.sheet_var.set("NoHeader")
            app.handle_sheet_selected(show_error=False)

            self.assertEqual(app.source_column_var.get(), "X")
            self.assertEqual(app.target_column_var.get(), "Y")

    def test_term_pair_history_tb_defaults_to_term_sheet_and_detects_source_target_columns(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            history_path = Path(tmp_dir) / "history.xlsx"
            self.create_history_tb_workbook(history_path)
            app = self.build_extract_terms_app_with_history(history_path)

            app.refresh_history_sheet_choices(show_error=False)

            self.assertEqual(app.history_sheet_combobox["values"], ("Raw", "术语表"))
            self.assertEqual(app.history_sheet_var.get(), "术语表")
            self.assertEqual(app.history_source_column_var.get(), "A")
            self.assertEqual(app.history_target_column_var.get(), "B")

    def test_glossary_checker_refreshes_each_file_independently(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            glossary_path = Path(tmp_dir) / "glossary.xlsx"
            data_path = Path(tmp_dir) / "data.xlsx"
            self.create_glossary_workbook(glossary_path)
            self.create_data_workbook(data_path)
            app = self.build_glossary_checker_app(glossary_path, data_path)

            app.refresh_glossary_sheet_choices(show_error=False)
            app.refresh_data_sheet_choices(show_error=False)

            self.assertEqual(app.glossary_sheet_combobox["values"], ("Glossary", "Backup"))
            self.assertEqual(app.glossary_sheet_var.get(), "Glossary")
            self.assertEqual(app.glossary_source_column_var.get(), "B")
            self.assertEqual(app.glossary_target_column_var.get(), "D")

            self.assertEqual(app.data_sheet_combobox["values"], ("Data", "Archive"))
            self.assertEqual(app.data_sheet_var.get(), "Archive")
            self.assertEqual(app.data_source_column_var.get(), "A")
            self.assertEqual(app.data_target_column_var.get(), "B")

    def test_glossary_checker_switching_sheet_redetects_columns(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            glossary_path = Path(tmp_dir) / "glossary.xlsx"
            data_path = Path(tmp_dir) / "data.xlsx"
            self.create_glossary_workbook(glossary_path)
            self.create_data_workbook(data_path)
            app = self.build_glossary_checker_app(glossary_path, data_path)

            app.refresh_data_sheet_choices(show_error=False)
            app.data_sheet_var.set("Data")
            app.handle_data_sheet_selected(show_error=False)

            self.assertEqual(app.data_source_column_var.get(), "C")
            self.assertEqual(app.data_target_column_var.get(), "F")

    def test_splitter_refresh_populates_sheet_choices(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "split.xlsx"
            self.create_splitter_workbook(workbook_path)
            app = self.build_splitter_app(workbook_path)

            app.refresh_sheet_choices(show_error=False)

            self.assertEqual(app.sheet_combobox["values"], ("Split", "Other"))
            self.assertEqual(app.sheet_var.get(), "Split")

    def test_tag_checker_refresh_populates_sheet_choices_and_detects_columns(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "tags.xlsx"
            self.create_tag_checker_workbook(workbook_path)
            app = self.build_tag_checker_app(workbook_path)

            app.refresh_sheet_choices(show_error=False)

            self.assertEqual(app.sheet_combobox["values"], ("Tags", "Archive"))
            self.assertEqual(app.sheet_var.get(), "Tags")
            self.assertEqual(app.source_column_var.get(), "D")
            self.assertEqual(app.target_column_var.get(), "F")

    def test_tag_checker_switching_sheet_redetects_columns(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "tags.xlsx"
            self.create_tag_checker_workbook(workbook_path)
            app = self.build_tag_checker_app(workbook_path)

            app.refresh_sheet_choices(show_error=False)
            app.sheet_var.set("Archive")
            app.handle_sheet_selected(show_error=False)

            self.assertEqual(app.source_column_var.get(), "A")
            self.assertEqual(app.target_column_var.get(), "B")

    def test_french_nbsp_restorer_refresh_populates_sheet_choices_and_detects_target_column(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "nbsp.xlsx"
            self.create_tag_checker_workbook(workbook_path)
            app = self.build_french_nbsp_restorer_app(workbook_path)

            app.refresh_sheet_choices(show_error=False)

            self.assertEqual(app.sheet_combobox["values"], ("Tags", "Archive"))
            self.assertEqual(app.sheet_var.get(), "Tags")
            self.assertEqual(app.target_column_var.get(), "F")

    def test_chinese_target_checker_refresh_populates_sheet_choices_and_detects_target_column(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "chinese.xlsx"
            self.create_tag_checker_workbook(workbook_path)
            app = self.build_chinese_target_checker_app(workbook_path)

            app.refresh_sheet_choices(show_error=False)

            self.assertEqual(app.sheet_combobox["values"], ("Tags", "Archive"))
            self.assertEqual(app.sheet_var.get(), "Tags")
            self.assertEqual(app.target_column_var.get(), "F")

    def test_chinese_target_checker_choose_input_keeps_output_blank_for_in_place_default(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "chinese.xlsx"
            self.create_tag_checker_workbook(workbook_path)
            app = self.build_chinese_target_checker_app(Path(""))

            with patch(
                "tools.chinese_target_checker.check_chinese_target_gui.filedialog.askopenfilename",
                return_value=str(workbook_path),
            ):
                app.choose_input_file()

            self.assertEqual(app.input_file_var.get(), str(workbook_path))
            self.assertEqual(app.output_file_var.get(), "")
            self.assertEqual(app.sheet_var.get(), "Tags")

    def test_chinese_target_checker_has_no_problem_sheet_toggle(self) -> None:
        app = ChineseTargetCheckerApp.__new__(ChineseTargetCheckerApp)
        with (
            patch.object(ChineseTargetCheckerApp, "_build_ui", lambda self: None),
            patch("tools.chinese_target_checker.check_chinese_target_gui.ttk.Frame.__init__", lambda self, master=None, padding=None: None),
            patch("tools.chinese_target_checker.check_chinese_target_gui.tk.StringVar", FakeVar),
            patch("tools.chinese_target_checker.check_chinese_target_gui.tk.BooleanVar", FakeBoolVar),
        ):
            ChineseTargetCheckerApp.__init__(app, object())

        self.assertFalse(hasattr(app, "problem_sheet_var"))

    def test_french_nbsp_restorer_switching_sheet_redetects_target_column(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "nbsp.xlsx"
            self.create_tag_checker_workbook(workbook_path)
            app = self.build_french_nbsp_restorer_app(workbook_path)

            app.refresh_sheet_choices(show_error=False)
            app.sheet_var.set("Archive")
            app.handle_sheet_selected(show_error=False)

            self.assertEqual(app.target_column_var.get(), "B")

    def test_workflow_refresh_populates_sheet_choices_and_detects_columns(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "workflow.xlsx"
            self.create_tag_checker_workbook(workbook_path)
            app = self.build_workflow_app(workbook_path)

            app.refresh_sheet_choices(show_error=False)

            self.assertEqual(app.sheet_combobox["values"], ("Tags", "Archive"))
            self.assertEqual(app.sheet_var.get(), "Tags")
            self.assertEqual(app.source_column_var.get(), "D")
            self.assertEqual(app.target_column_var.get(), "F")

    def test_workflow_history_tb_defaults_to_term_sheet_and_detects_source_target_columns(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            history_path = Path(tmp_dir) / "history.xlsx"
            self.create_history_tb_workbook(history_path)
            app = self.build_workflow_app_with_history(history_path)

            app.refresh_term_history_sheet_choices(show_error=False)

            self.assertEqual(app.term_history_sheet_combobox["values"], ("Raw", "术语表"))
            self.assertEqual(app.term_history_sheet_var.get(), "术语表")
            self.assertEqual(app.term_history_source_column_var.get(), "A")
            self.assertEqual(app.term_history_target_column_var.get(), "B")

    def test_llm_term_extractor_refresh_populates_sheet_choices_and_detects_columns(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "llm_terms.xlsx"
            self.create_data_workbook(workbook_path)
            app = self.build_llm_term_extractor_app(workbook_path)

            app.refresh_sheet_choices(show_error=False)

            self.assertEqual(app.sheet_combobox["values"], ("Data", "Archive"))
            self.assertEqual(app.sheet_var.get(), "Archive")
            self.assertEqual(app.source_column_var.get(), "A")
            self.assertEqual(app.target_column_var.get(), "B")

    def test_llm_term_extractor_switching_sheet_redetects_columns(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "llm_terms.xlsx"
            self.create_data_workbook(workbook_path)
            app = self.build_llm_term_extractor_app(workbook_path)

            app.refresh_sheet_choices(show_error=False)
            app.sheet_var.set("Data")
            app.handle_sheet_selected(show_error=False)

            self.assertEqual(app.source_column_var.get(), "C")
            self.assertEqual(app.target_column_var.get(), "F")

    def test_llm_term_extractor_source_only_refresh_clears_target_column(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "source_only.xlsx"
            self.create_source_only_workbook(workbook_path)
            app = self.build_llm_term_extractor_app(workbook_path)
            app.target_column_var.set("B")

            app.refresh_sheet_choices(show_error=False)

            self.assertEqual(app.sheet_combobox["values"], ("SourceOnly",))
            self.assertEqual(app.sheet_var.get(), "SourceOnly")
            self.assertEqual(app.source_column_var.get(), "A")
            self.assertEqual(app.target_column_var.get(), "")

    def test_llm_term_extractor_history_tb_defaults_to_term_sheet_and_detects_nomark_columns(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            history_path = Path(tmp_dir) / "history.xlsx"
            self.create_history_tb_workbook(history_path)
            app = self.build_llm_term_extractor_app_with_history(history_path)

            app.refresh_history_sheet_choices(show_error=False)

            self.assertEqual(app.history_sheet_combobox["values"], ("Raw", "术语表"))
            self.assertEqual(app.history_sheet_var.get(), "术语表")
            self.assertEqual(app.history_source_column_var.get(), "C")
            self.assertEqual(app.history_target_column_var.get(), "D")

    def test_llm_term_extractor_history_tb_detection_uses_history_start_row(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            history_path = Path(tmp_dir) / "history_offset.xlsx"
            self.create_offset_history_tb_workbook(history_path)
            app = self.build_llm_term_extractor_app_with_history(history_path)
            app.history_start_row_var.set("3")

            app.refresh_history_sheet_choices(show_error=False)

            self.assertEqual(app.history_sheet_combobox["values"], ("术语表",))
            self.assertEqual(app.history_sheet_var.get(), "术语表")
            self.assertEqual(app.history_source_column_var.get(), "C")
            self.assertEqual(app.history_target_column_var.get(), "D")

    def test_llm_term_extractor_blank_target_and_history_fields_pass_to_processor(self) -> None:
        app = LlmTermExtractorApp.__new__(LlmTermExtractorApp)
        app.input_file_var = FakeVar("/tmp/input.xlsx")
        app.output_file_var = FakeVar("/tmp/output.xlsx")
        app.sheet_var = FakeVar("Data")
        app.source_column_var = FakeVar("A")
        app.target_column_var = FakeVar("")
        app.start_row_var = FakeVar("3")
        app.batch_size_var = FakeVar("25")
        app.codex_model_var = FakeVar("codex-model")
        app.codex_reasoning_effort_var = FakeVar("high")
        app.extract_prompt_file_var = FakeVar("")
        app.conflict_prompt_file_var = FakeVar("/tmp/conflict.md")
        app.history_tb_file_var = FakeVar("/tmp/history.xlsx")
        app.history_sheet_var = FakeVar("术语表")
        app.history_source_column_var = FakeVar("C")
        app.history_target_column_var = FakeVar("D")
        app.history_start_row_var = FakeVar("4")
        app.keep_raw_codex_output_var = FakeBoolVar(True)
        summary = SimpleNamespace(
            output_path=Path("/tmp/output.xlsx"),
            worksheet_title="Data",
            source_column="A",
            target_column="",
            start_row=3,
            scanned_row_count=10,
            batch_count=1,
            term_count=4,
            conflict_count=2,
        )

        with (
            patch(
                "tools.llm_term_extractor.extract_llm_terms_gui.process_excel",
                return_value=summary,
            ) as process_excel_mock,
            patch("tools.llm_term_extractor.extract_llm_terms_gui.messagebox.showinfo"),
        ):
            app.run_extraction()

        process_excel_mock.assert_called_once()
        self.assertIsNone(process_excel_mock.call_args.kwargs["target_column"])
        self.assertEqual(process_excel_mock.call_args.kwargs["batch_size"], 25)
        self.assertEqual(process_excel_mock.call_args.kwargs["extract_prompt_file"], None)
        self.assertEqual(process_excel_mock.call_args.kwargs["conflict_prompt_file"], "/tmp/conflict.md")
        self.assertEqual(process_excel_mock.call_args.kwargs["history_sheet"], "术语表")
        self.assertEqual(process_excel_mock.call_args.kwargs["history_source_column"], "C")
        self.assertEqual(process_excel_mock.call_args.kwargs["history_target_column"], "D")
        self.assertEqual(process_excel_mock.call_args.kwargs["history_start_row"], 4)
        self.assertEqual(process_excel_mock.call_args.kwargs["keep_raw_codex_output"], True)

    def test_llm_term_extractor_ignores_invalid_history_start_without_history_file(self) -> None:
        app = LlmTermExtractorApp.__new__(LlmTermExtractorApp)
        app.input_file_var = FakeVar("/tmp/input.xlsx")
        app.output_file_var = FakeVar("/tmp/output.xlsx")
        app.sheet_var = FakeVar("Data")
        app.source_column_var = FakeVar("A")
        app.target_column_var = FakeVar("")
        app.start_row_var = FakeVar("3")
        app.batch_size_var = FakeVar("25")
        app.codex_model_var = FakeVar("codex-model")
        app.codex_reasoning_effort_var = FakeVar("high")
        app.extract_prompt_file_var = FakeVar("")
        app.conflict_prompt_file_var = FakeVar("")
        app.history_tb_file_var = FakeVar("")
        app.history_sheet_var = FakeVar("术语表")
        app.history_source_column_var = FakeVar("C")
        app.history_target_column_var = FakeVar("D")
        app.history_start_row_var = FakeVar("not-an-int")
        app.keep_raw_codex_output_var = FakeBoolVar(False)
        summary = SimpleNamespace(
            output_path=Path("/tmp/output.xlsx"),
            worksheet_title="Data",
            source_column="A",
            target_column="",
            start_row=3,
            scanned_row_count=10,
            batch_count=1,
            term_count=4,
            conflict_count=2,
        )

        with (
            patch(
                "tools.llm_term_extractor.extract_llm_terms_gui.process_excel",
                return_value=summary,
            ) as process_excel_mock,
            patch("tools.llm_term_extractor.extract_llm_terms_gui.messagebox.showerror") as showerror_mock,
            patch("tools.llm_term_extractor.extract_llm_terms_gui.messagebox.showinfo"),
        ):
            app.run_extraction()

        showerror_mock.assert_not_called()
        process_excel_mock.assert_called_once()
        self.assertIsNone(process_excel_mock.call_args.kwargs["history_tb_file"])
        self.assertIsNone(process_excel_mock.call_args.kwargs["history_sheet"])
        self.assertIsNone(process_excel_mock.call_args.kwargs["history_source_column"])
        self.assertIsNone(process_excel_mock.call_args.kwargs["history_target_column"])
        self.assertEqual(process_excel_mock.call_args.kwargs["history_start_row"], 2)

    def test_term_pair_gui_defaults_to_square_and_angle_marks(self) -> None:
        app = ExtractTermsApp.__new__(ExtractTermsApp)
        with (
            patch.object(ExtractTermsApp, "_build_ui", lambda self: None),
            patch("tools.term_pair_checker.extract_terms_gui.ttk.Frame.__init__", lambda self, master=None, padding=None: None),
            patch("tools.term_pair_checker.extract_terms_gui.tk.StringVar", FakeVar),
            patch("tools.term_pair_checker.extract_terms_gui.tk.BooleanVar", FakeBoolVar),
        ):
            ExtractTermsApp.__init__(app, object())

        self.assertFalse(app.mark_style_vars["【】"].get())
        self.assertTrue(app.mark_style_vars["[]"].get())
        self.assertTrue(app.mark_style_vars["<>"].get())

    def test_workflow_gui_defaults_to_square_and_angle_term_marks(self) -> None:
        app = WorkflowRunnerApp.__new__(WorkflowRunnerApp)
        with (
            patch.object(WorkflowRunnerApp, "_build_ui", lambda self: None),
            patch("tools.workflow.workflow_gui.ttk.Frame.__init__", lambda self, master=None, padding=None: None),
            patch("tools.workflow.workflow_gui.tk.StringVar", FakeVar),
            patch("tools.workflow.workflow_gui.tk.BooleanVar", FakeBoolVar),
        ):
            WorkflowRunnerApp.__init__(app, object())

        self.assertFalse(app.term_mark_style_vars["【】"].get())
        self.assertTrue(app.term_mark_style_vars["[]"].get())
        self.assertTrue(app.term_mark_style_vars["<>"].get())


if __name__ == "__main__":
    unittest.main()
