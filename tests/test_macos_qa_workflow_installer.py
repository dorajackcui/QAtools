from __future__ import annotations

import plistlib
import tempfile
import unittest
from pathlib import Path, PurePosixPath
from unittest.mock import call, patch

from scripts.install_macos_qa_workflow import (
    EXCEL_UTIS,
    NBSP_QUICK_ACTION,
    QUICK_ACTION_BUNDLE_NAME,
    build_argument_parser,
    build_quick_action_bundle,
    project_venv_python,
    sync_project_environment,
)


class MacosQaWorkflowInstallerTests(unittest.TestCase):
    def test_default_install_updates_both_finder_actions(self) -> None:
        args = build_argument_parser().parse_args([])

        self.assertEqual(args.action, "all")
        self.assertIsNone(args.python_executable)
        self.assertFalse(args.skip_environment_sync)

    def test_default_runtime_is_the_repository_virtualenv(self) -> None:
        self.assertEqual(
            project_venv_python(Path("/Users/example/tagExactor")),
            Path("/Users/example/tagExactor/.venv/bin/python"),
        )

    def test_environment_sync_installs_project_then_builds_full_gui(self) -> None:
        python_executable = Path("/Users/example/tagExactor/.venv/bin/python")
        project_root = Path("/Users/example/tagExactor")
        launcher = project_root / "toolshub_gui.py"

        with patch("scripts.install_macos_qa_workflow.subprocess.run") as run:
            sync_project_environment(
                python_executable=python_executable,
                project_root=project_root,
                launcher=launcher,
            )

        self.assertEqual(
            run.call_args_list,
            [
                call(
                    [
                        str(python_executable),
                        "-m",
                        "pip",
                        "install",
                        "-e",
                        str(project_root),
                    ],
                    check=True,
                ),
                call(
                    [str(python_executable), str(launcher), "--smoke-test"],
                    check=True,
                ),
            ],
        )

    def test_bundle_targets_excel_files_and_passes_input_as_argument(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            destination = Path(tmp_dir)
            bundle_path = build_quick_action_bundle(
                destination,
                python_executable=PurePosixPath("/opt/example/python3"),
                launcher=PurePosixPath(
                    "/Users/example/tagExactor/toolshub_gui.py"
                ),
            )

            self.assertEqual(bundle_path.name, QUICK_ACTION_BUNDLE_NAME)
            with (bundle_path / "Contents" / "Info.plist").open("rb") as info_file:
                info = plistlib.load(info_file)
            with (
                bundle_path / "Contents" / "Resources" / "document.wflow"
            ).open("rb") as document_file:
                document = plistlib.load(document_file)

            service = info["NSServices"][0]
            self.assertEqual(
                service["NSMenuItem"]["default"],
                "ABC · QA workflow",
            )
            self.assertEqual(tuple(service["NSSendFileTypes"]), EXCEL_UTIS)

            action_parameters = document["actions"][0]["action"]["ActionParameters"]
            self.assertEqual(action_parameters["inputMethod"], 1)
            command = action_parameters["COMMAND_STRING"]
            self.assertIn("--qa-workflow \"$qa_file\"", command)
            self.assertIn("/opt/example/python3", command)
            self.assertIn(
                "/Users/example/tagExactor/toolshub_gui.py",
                command,
            )

    def test_nbsp_bundle_runs_the_nbsp_restore_launcher_action(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            destination = Path(tmp_dir)
            bundle_path = build_quick_action_bundle(
                destination,
                python_executable=Path("/opt/example/python3"),
                launcher=Path("/Users/example/tagExactor/toolshub_gui.py"),
                spec=NBSP_QUICK_ACTION,
            )

            self.assertEqual(bundle_path.name, "NBSP restore.workflow")
            with (bundle_path / "Contents" / "Info.plist").open("rb") as info_file:
                info = plistlib.load(info_file)
            with (
                bundle_path / "Contents" / "Resources" / "document.wflow"
            ).open("rb") as document_file:
                document = plistlib.load(document_file)

            service = info["NSServices"][0]
            self.assertEqual(
                service["NSMenuItem"]["default"],
                "ABC · NBSP restore",
            )
            self.assertEqual(
                info["CFBundleIdentifier"],
                "com.tagexactor.services.nbsp-restore",
            )
            command = document["actions"][0]["action"]["ActionParameters"][
                "COMMAND_STRING"
            ]
            self.assertIn("--nbsp-restore \"$qa_file\"", command)


if __name__ == "__main__":
    unittest.main()
