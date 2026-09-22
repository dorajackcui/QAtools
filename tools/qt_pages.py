"""Lazy page registry and stable class imports; implementations live in tool packages."""

from functools import partial


def _load_class(name: str):
    # Explicit imports also let PyInstaller discover every optional page.
    if name == "SettingsPage":
        from tools.qt_settings_page import SettingsPage
        return SettingsPage
    if name == "PhraseLoomPage":
        from phraseloom.qt_page import PhraseLoomPage
        return PhraseLoomPage
    if name == "FrenchNbspPage":
        from tools.french_nbsp_restorer.qt_page import FrenchNbspPage
        return FrenchNbspPage
    if name == "XbenchPage":
        from tools.xbench_report_transformer.qt_page import XbenchPage
        return XbenchPage
    if name == "ExcelBatcherPage":
        from tools.excel_batcher.qt_page import ExcelBatcherPage
        return ExcelBatcherPage
    if name == "ExcelMergerPage":
        from tools.excel_merger.qt_page import ExcelMergerPage
        return ExcelMergerPage
    if name == "WorkflowPage":
        from tools.workflow.qt_page import WorkflowPage
        return WorkflowPage
    if name == "ContentSyncPage":
        from tools.content_sync.qt_page import ContentSyncPage
        return ContentSyncPage
    if name == "FileCollectorPage":
        from tools.file_collector.qt_page import FileCollectorPage
        return FileCollectorPage
    if name in {"ColumnToolsPage", "CompatibilityPage", "DeepReplacePage", "UntranslatedStatsPage"}:
        from tools.excel_utilities_pages import (
            ColumnToolsPage, CompatibilityPage, DeepReplacePage, UntranslatedStatsPage,
        )
        return {
            "ColumnToolsPage": ColumnToolsPage, "CompatibilityPage": CompatibilityPage,
            "DeepReplacePage": DeepReplacePage, "UntranslatedStatsPage": UntranslatedStatsPage,
        }[name]
    raise AttributeError(f"module {__name__!r} has no attribute {name!r}")


def __getattr__(name: str):
    page_class = _load_class(name)
    globals()[name] = page_class
    return page_class


def _create_page(class_name: str, *args, **kwargs):
    return _load_class(class_name)(*args, **kwargs)


PAGE_FACTORIES = {
    key: partial(_create_page, class_name)
    for key, class_name in (
        ("content_sync", "ContentSyncPage"),
        ("column_tools", "ColumnToolsPage"),
        ("compatibility", "CompatibilityPage"),
        ("deep_replace", "DeepReplacePage"),
        ("file_collector", "FileCollectorPage"),
        ("untranslated_stats", "UntranslatedStatsPage"),
        ("workflow", "WorkflowPage"),
        ("phraseloom", "PhraseLoomPage"),
        ("french_nbsp", "FrenchNbspPage"),
        ("excel_batcher", "ExcelBatcherPage"),
        ("excel_merger", "ExcelMergerPage"),
        ("xbench_report", "XbenchPage"),
        ("settings", "SettingsPage"),
    )
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
