class PhraseLoomError(Exception):
    """Base class for user-facing PhraseLoom errors."""


class ConfigError(PhraseLoomError):
    """Raised when command or workflow configuration is invalid."""


class WorkbookFormatError(PhraseLoomError):
    """Raised when an input workbook does not match an expected schema."""


class ColumnNotFoundError(WorkbookFormatError):
    """Raised when a requested workbook column is missing."""


class TranslationUnitLoadError(WorkbookFormatError):
    """Raised when translated units cannot be loaded from a workbook."""


class WorkflowError(PhraseLoomError):
    """Raised when a workflow cannot complete with valid inputs."""
