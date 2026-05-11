import shutil
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from phraseloom.workflow import (
    fill_target_column_workbook,
    generate_tm_pairs,
    generate_workbook,
)


TESTFILES = Path("testfiles")


def ensure_tag_testfiles():
    TESTFILES.mkdir(exist_ok=True)

    tm = Workbook()
    tm_ws = tm.active
    tm_ws.append(["source", "target"])
    tm_ws.append(['<a href="shop">VIP10 Pack</a>', '<a href="shop">Pack VIP10 FR</a>'])
    tm_ws.append(['<a href="shop">VIP20 Pack</a>', '<a href="shop">Pack VIP20 FR</a>'])
    tm_ws.append(["Login failed", "Login failed FR"])
    tm.save(TESTFILES / "tag_tm.xlsx")

    source = Workbook()
    source_ws = source.active
    source_ws.append(["source", "target"])
    source_ws.append(['<a href="shop">VIP30 Pack</a>', None])
    source_ws.append(['<a href="shop">VIP40 Pack</a>', None])
    source_ws.append(['<img src="coin.png"/>', None])
    source_ws.append(["Brand new line", None])
    source.save(TESTFILES / "tag_source.xlsx")


def clean_generated_outputs():
    for folder in (TESTFILES / "tag_tm_l10n", TESTFILES / "tag_source_l10n"):
        if folder.exists():
            shutil.rmtree(folder)


def sheet_rows_by_header(workbook_path, sheet_name):
    workbook = load_workbook(workbook_path, data_only=True)
    worksheet = workbook[sheet_name]
    headers = [cell.value for cell in worksheet[1]]
    return [
        dict(zip(headers, row))
        for row in worksheet.iter_rows(min_row=2, values_only=True)
    ]


class TagWorkflowTestfilesTests(unittest.TestCase):
    def setUp(self):
        ensure_tag_testfiles()
        clean_generated_outputs()

    def tearDown(self):
        clean_generated_outputs()

    def test_tm_extract_and_fill_use_testfiles_with_tags(self):
        tm_workbook = TESTFILES / "tag_tm.xlsx"
        tm_pairs = TESTFILES / "tag_tm_l10n" / "tag_tm_reusable_units.xlsx"
        source_workbook = TESTFILES / "tag_source.xlsx"
        source_pack = TESTFILES / "tag_source_l10n" / "tag_source_pack.xlsx"
        standalone_todo = TESTFILES / "tag_source_l10n" / "tag_source_translator_todo.xlsx"
        filled_workbook = TESTFILES / "tag_source_l10n" / "tag_source_filled.xlsx"

        tm_stats = generate_tm_pairs(
            tm_workbook,
            tm_pairs,
            source_col="source",
            target_col="target",
            min_group_size=2,
        )

        self.assertEqual(tm_stats["template_pair_count"], 1)

        stats = generate_workbook(
            source_workbook,
            source_pack,
            source_col="source",
            target_col="target",
            tm_workbook=tm_pairs,
            min_group_size=2,
            use_existing_targets=False,
        )

        self.assertEqual(stats["prefilled_translation_unit_count"], 2)
        self.assertEqual(stats["untranslated_translation_unit_count"], 1)

        todo_workbook = load_workbook(standalone_todo)
        todo = todo_workbook["to_translate"]
        todo_headers = [cell.value for cell in todo[1]]
        source_idx = todo_headers.index("source_unit") + 1
        target_idx = todo_headers.index("target_unit") + 1
        todo_sources = [
            row[source_idx - 1].value for row in todo.iter_rows(min_row=2)
        ]
        self.assertEqual(todo_sources, ["Brand new line"])
        todo.cell(row=2, column=target_idx).value = "Brand new line FR"
        todo_workbook.save(standalone_todo)

        fill_stats = fill_target_column_workbook(
            source_workbook,
            filled_workbook,
            source_col="source",
            target_col="target",
            template_workbook=standalone_todo,
            min_group_size=2,
        )

        self.assertEqual(fill_stats["autofilled_count"], 4)

        filled = load_workbook(filled_workbook, data_only=True)
        rows = list(filled.active.iter_rows(values_only=True))
        self.assertEqual(rows[1][1], '<a href="shop">Pack VIP30 FR</a>')
        self.assertEqual(rows[2][1], '<a href="shop">Pack VIP40 FR</a>')
        self.assertEqual(rows[3][1], '<img src="coin.png"/>')
        self.assertEqual(rows[4][1], "Brand new line FR")

    def test_fill_writes_target_when_tag_mismatch_warning_exists(self):
        source_workbook = TESTFILES / "tag_source.xlsx"
        pack = TESTFILES / "tag_source_l10n" / "tag_source_pack.xlsx"
        report = TESTFILES / "tag_source_l10n" / "tag_source_report.xlsx"

        generate_workbook(
            source_workbook,
            pack,
            source_col="source",
            target_col="target",
            min_group_size=2,
            use_existing_targets=False,
        )

        pack_workbook = load_workbook(pack)
        units = pack_workbook["translation_units"]
        headers = [cell.value for cell in units[1]]
        source_idx = headers.index("source_unit") + 1
        target_idx = headers.index("target_unit") + 1
        for row in units.iter_rows(min_row=2):
            source_unit = row[source_idx - 1].value
            if source_unit == "{t1_op}VIP{num1} Pack{t1_cl}":
                row[target_idx - 1].value = "Pack VIP{num1} FR"
            elif source_unit == "Brand new line":
                row[target_idx - 1].value = "Brand new line FR"
        pack_workbook.save(pack)

        generate_workbook(
            source_workbook,
            report,
            source_col="source",
            target_col="target",
            template_workbook=pack,
            min_group_size=2,
            use_existing_targets=False,
        )

        rows = sheet_rows_by_header(report, "source_map")
        self.assertEqual(rows[0]["auto_target"], "Pack VIP30 FR")
        self.assertEqual(rows[1]["auto_target"], "Pack VIP40 FR")
        self.assertTrue(
            any("tag_mismatch" in str(row["warning"] or "") for row in rows)
        )


if __name__ == "__main__":
    unittest.main()
