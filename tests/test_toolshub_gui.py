from __future__ import annotations

import tkinter as tk
import unittest

from toolshub_gui import ToolshubApp


class ToolshubLayoutTests(unittest.TestCase):
    def test_initial_window_is_not_smaller_than_requested_content(self) -> None:
        try:
            root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk display is unavailable: {exc}")

        try:
            ToolshubApp(root)
            root.update()

            content = root.winfo_children()[0]
            self.assertGreaterEqual(root.winfo_width(), content.winfo_reqwidth())
            self.assertGreaterEqual(root.winfo_height(), content.winfo_reqheight())
        finally:
            root.destroy()


if __name__ == "__main__":
    unittest.main()
