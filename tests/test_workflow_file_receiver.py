from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path

from tools.workflow.file_receiver import (
    FRENCH_NBSP_RESTORE_ACTION,
    WorkflowFileReceiver,
    normalize_workflow_input_file,
    send_tool_input_file,
    send_workflow_input_file,
)


class WorkflowFileReceiverTests(unittest.TestCase):
    def test_normalize_accepts_supported_excel_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            workbook_path = Path(tmp_dir) / "QA Input.XLSX"
            workbook_path.touch()

            normalized = normalize_workflow_input_file(workbook_path)

            self.assertEqual(normalized, workbook_path.absolute())

    def test_normalize_rejects_non_excel_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp_dir:
            text_path = Path(tmp_dir) / "notes.txt"
            text_path.touch()

            with self.assertRaisesRegex(ValueError, r"\.xlsx"):
                normalize_workflow_input_file(text_path)

    @unittest.skipIf(os.name == "nt", "Finder file forwarding requires fcntl")
    def test_receiver_accepts_forwarded_file_from_local_inbox(self) -> None:
        with tempfile.TemporaryDirectory(dir="/tmp") as tmp_dir:
            temp_path = Path(tmp_dir)
            workbook_path = temp_path / "input file.xlsx"
            workbook_path.touch()
            receiver_directory = temp_path / "receiver"
            receiver = WorkflowFileReceiver(
                receiver_directory=receiver_directory
            )

            try:
                self.assertTrue(receiver.start())
                self.assertTrue(
                    send_workflow_input_file(
                        workbook_path,
                        receiver_directory=receiver_directory,
                    )
                )

                received_paths = receiver.pop_pending_paths()
                self.assertEqual(received_paths, (str(workbook_path.absolute()),))
            finally:
                receiver.close()

    @unittest.skipIf(os.name == "nt", "Finder file forwarding requires fcntl")
    def test_second_receiver_does_not_replace_active_receiver(self) -> None:
        with tempfile.TemporaryDirectory(dir="/tmp") as tmp_dir:
            receiver_directory = Path(tmp_dir) / "receiver"
            first_receiver = WorkflowFileReceiver(
                receiver_directory=receiver_directory
            )
            second_receiver = WorkflowFileReceiver(
                receiver_directory=receiver_directory
            )

            try:
                self.assertTrue(first_receiver.start())
                self.assertFalse(second_receiver.start())
            finally:
                second_receiver.close()
                first_receiver.close()

    def test_receiver_preserves_nbsp_restore_action(self) -> None:
        with tempfile.TemporaryDirectory(dir="/tmp") as tmp_dir:
            temp_path = Path(tmp_dir)
            workbook_path = temp_path / "input.xlsx"
            workbook_path.touch()
            receiver_directory = temp_path / "receiver"
            receiver = WorkflowFileReceiver(
                receiver_directory=receiver_directory
            )

            try:
                self.assertTrue(receiver.start())
                self.assertTrue(
                    send_tool_input_file(
                        FRENCH_NBSP_RESTORE_ACTION,
                        workbook_path,
                        receiver_directory=receiver_directory,
                    )
                )

                requests = receiver.pop_pending_requests()
                self.assertEqual(len(requests), 1)
                self.assertEqual(
                    requests[0].action,
                    FRENCH_NBSP_RESTORE_ACTION,
                )
                self.assertEqual(
                    requests[0].file_path,
                    str(workbook_path.absolute()),
                )
            finally:
                receiver.close()


if __name__ == "__main__":
    unittest.main()
