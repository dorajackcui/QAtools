from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from tools.tb_projects import TbProject, TbProjectStore


class TbProjectStoreTests(unittest.TestCase):
    def make_project(
        self,
        name: str,
        file_path: Path,
        *,
        sheet: str = "术语表",
        source_column: str = "A",
        target_column: str = "B",
        start_row: int = 2,
    ) -> TbProject:
        return TbProject(
            name=name,
            file_path=str(file_path),
            sheet=sheet,
            source_column=source_column,
            target_column=target_column,
            start_row=start_row,
        )

    def test_projects_are_shared_across_store_instances(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            temp_path = Path(tmp_dir)
            config_path = temp_path / "tb_projects.json"
            tb_path = temp_path / "history.xlsx"
            tb_path.touch()
            first_store = TbProjectStore(config_path)
            second_store = TbProjectStore(config_path)

            first_store.save_project(self.make_project("游戏 A", tb_path))

            self.assertEqual(
                second_store.find_project("游戏 A"),
                self.make_project("游戏 A", tb_path),
            )

    def test_saving_same_name_updates_existing_project(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            temp_path = Path(tmp_dir)
            store = TbProjectStore(temp_path / "tb_projects.json")
            first_tb = temp_path / "first.xlsx"
            second_tb = temp_path / "second.xlsx"

            store.save_project(self.make_project("Project A", first_tb))
            store.save_project(
                self.make_project(
                    "project a",
                    second_tb,
                    sheet="Glossary",
                    source_column="C",
                    target_column="D",
                    start_row=3,
                )
            )

            projects = store.list_projects()
            self.assertEqual(len(projects), 1)
            self.assertEqual(projects[0].file_path, str(second_tb))
            self.assertEqual(projects[0].sheet, "Glossary")
            self.assertEqual(projects[0].start_row, 3)

    def test_delete_project_only_removes_saved_configuration(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            temp_path = Path(tmp_dir)
            store = TbProjectStore(temp_path / "tb_projects.json")
            tb_path = temp_path / "history.xlsx"
            tb_path.touch()
            store.save_project(self.make_project("游戏 A", tb_path))

            self.assertTrue(store.delete_project("游戏 A"))

            self.assertTrue(tb_path.is_file())
            self.assertEqual(store.list_projects(), ())

    def test_invalid_config_reports_a_readable_error(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            config_path = Path(tmp_dir) / "tb_projects.json"
            config_path.write_text("{invalid", encoding="utf-8")
            store = TbProjectStore(config_path)

            with self.assertRaisesRegex(ValueError, "无法读取 TB 项目配置"):
                store.list_projects()


if __name__ == "__main__":
    unittest.main()
