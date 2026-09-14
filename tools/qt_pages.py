"""Page registry and stable class imports. Implement pages in their tool packages."""

from tools.qt_settings_page import SettingsPage
from phraseloom.qt_page import PhraseLoomPage
from tools.french_nbsp_restorer.qt_page import FrenchNbspPage
from tools.xbench_report_transformer.qt_page import XbenchPage
from tools.excel_batcher.qt_page import ExcelBatcherPage
from tools.excel_merger.qt_page import ExcelMergerPage
from tools.workflow.qt_page import WorkflowPage
from tools.content_sync.qt_page import ContentSyncPage
from tools.file_collector.qt_page import FileCollectorPage
from tools.excel_utilities_pages import ColumnToolsPage, CompatibilityPage, DeepReplacePage, UntranslatedStatsPage


PAGE_FACTORIES = {
    "content_sync": ContentSyncPage,
    "column_tools": ColumnToolsPage,
    "compatibility": CompatibilityPage,
    "deep_replace": DeepReplacePage,
    "file_collector": FileCollectorPage,
    "untranslated_stats": UntranslatedStatsPage,
    "workflow": WorkflowPage,
    "phraseloom": PhraseLoomPage,
    "french_nbsp": FrenchNbspPage,
    "excel_batcher": ExcelBatcherPage,
    "excel_merger": ExcelMergerPage,
    "xbench_report": XbenchPage,
    "settings": SettingsPage,
}


__all__ = [
    "ExcelBatcherPage",
    "ExcelMergerPage",
    "FrenchNbspPage",
    "PAGE_FACTORIES",
    "PhraseLoomPage",
    "SettingsPage",
    "WorkflowPage",
    "XbenchPage",
]
