class PhraseLoomError(Exception):
    """Base class for user-facing PhraseLoom errors."""


class ConfigError(PhraseLoomError):
    """Raised when command or workflow configuration is invalid."""


class WorkbookFormatError(PhraseLoomError):
    """Raised when an input workbook does not match an expected schema."""


class ColumnNotFoundError(WorkbookFormatError):
    """Raised when a requested workbook column is missing."""

    def __init__(self, column: str | int | None, available_columns: list[str]) -> None:
        available = ", ".join(available_columns) if available_columns else "(none)"
        super().__init__(
            f"Column {column!r} not found in header row.\n"
            f"Available columns: {available}"
        )
