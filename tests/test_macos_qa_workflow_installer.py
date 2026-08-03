from __future__ import annotations

import plistlib
import tempfile
import unittest
from pathlib import Path, PurePosixPath

from scripts.install_macos_qa_workflow import (
    EXCEL_UTIS,
    QUICK_ACTION_BUNDLE_NAME,
    build_quick_action_bundle,
)


class MacosQaWorkflowInstallerTests(unittest.TestCase):
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
            self.assertEqual(service["NSMenuItem"]["default"], "QA workflow")
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


if __name__ == "__main__":
    unittest.main()
