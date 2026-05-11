# PhraseLoom Internal Tool Architecture Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Convert PhraseLoom into a stable internal Python package and CLI while preserving the current Excel workflow behavior.

**Architecture:** Use a strangler-style first refactor: move the existing scripts into package-internal compatibility modules, then expose responsibility-based package modules (`workflow`, `template_engine`, `excel_io`, `interactive`, `cli`, `entity_cluster`). This creates the public architecture without rewriting the already-tested algorithms in the same step.

**Tech Stack:** Python 3.11+, `openpyxl`, `unittest`, `argparse`, standard-library packaging through `pyproject.toml`.

---

## File Structure

Create:

- `pyproject.toml`: package metadata, dependency declaration, console entry point.
- `phraseloom/__init__.py`: package version and package marker.
- `phraseloom/workbook_schema.py`: centralized sheet and column names.
- `phraseloom/errors.py`: structured exception types for future workflow and CLI errors.
- `phraseloom/models.py`: exported model classes from the current implementation modules.
- `phraseloom/template_engine.py`: exported template parsing and application API.
- `phraseloom/workflow.py`: exported workflow API and public aliases.
- `phraseloom/excel_io.py`: exported Excel/path helpers and compatibility I/O functions.
- `phraseloom/interactive.py`: exported interactive prompt flow and prompt helpers.
- `phraseloom/cli.py`: package CLI entry point.
- `phraseloom/entity_cluster.py`: exported entity cluster API.
- `tests/test_template_workflow.py`: migrated workflow tests.
- `tests/test_entity_cluster.py`: migrated entity cluster tests.

Move:

- `template_demo.py` -> `phraseloom/_template_workflow.py`
- `entity_cluster_probe.py` -> `phraseloom/_entity_cluster_probe.py`

Replace:

- `template_demo.py`: compatibility shim for old direct script usage and existing imports.
- `entity_cluster_probe.py`: compatibility shim for old direct script usage and existing imports.

Modify:

- `README.md`: development setup, package CLI usage, compatibility command notes.

---

### Task 1: Add Package Scaffold And Tooling

**Files:**
- Create: `pyproject.toml`
- Create: `phraseloom/__init__.py`
- Create: `phraseloom/workbook_schema.py`
- Create: `phraseloom/errors.py`

- [ ] **Step 1: Add packaging metadata**

Create `pyproject.toml` with:

```toml
[build-system]
requires = ["setuptools>=68"]
build-backend = "setuptools.build_meta"

[project]
name = "phraseloom"
version = "0.1.0"
description = "Internal Excel localization workflow tooling."
readme = "README.md"
requires-python = ">=3.11"
dependencies = [
  "openpyxl>=3.1",
]

[project.scripts]
phraseloom = "phraseloom.cli:main"

[tool.setuptools.packages.find]
include = ["phraseloom*"]
```

- [ ] **Step 2: Add package marker**

Create `phraseloom/__init__.py` with:

```python
"""PhraseLoom package."""

__version__ = "0.1.0"
```

- [ ] **Step 3: Add workbook schema constants**

Create `phraseloom/workbook_schema.py` with constants for the existing workbook contract:

```python
SCHEMA_VERSION = "1.0"

SUMMARY_SHEET = "summary"
TRANSLATION_UNITS_SHEET = "translation_units"
TO_TRANSLATE_SHEET = "to_translate"
PREFILLED_UNITS_SHEET = "prefilled_units"
SOURCE_MAP_SHEET = "source_map"
FILLED_WORKBOOK_SHEET = "filled_workbook"
QA_REPORT_SHEET = "qa_report"
TM_PAIRS_SHEET = "tm_pairs"
TM_SOURCE_MAP_SHEET = "tm_source_map"
ENTITY_CLUSTERS_SHEET = "entity_clusters"

UNIT_TYPE_COLUMN = "unit_type"
SOURCE_UNIT_COLUMN = "source_unit"
TARGET_UNIT_COLUMN = "target_unit"
UNIT_ID_COLUMN = "unit_id"
```

- [ ] **Step 4: Add structured exceptions**

Create `phraseloom/errors.py` with:

```python
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
```

- [ ] **Step 5: Verify file presence**

Run: `rg --files phraseloom pyproject.toml`

Expected: shows the four new package files and `pyproject.toml`.

- [ ] **Step 6: Commit**

Run:

```bash
git add pyproject.toml phraseloom/__init__.py phraseloom/workbook_schema.py phraseloom/errors.py
git commit -m "Add PhraseLoom package scaffold"
```

---

### Task 2: Move Existing Implementations Behind Package Internals

**Files:**
- Move: `template_demo.py` -> `phraseloom/_template_workflow.py`
- Move: `entity_cluster_probe.py` -> `phraseloom/_entity_cluster_probe.py`
- Create: `template_demo.py`
- Create: `entity_cluster_probe.py`

- [ ] **Step 1: Move the current scripts**

Run:

```bash
git mv template_demo.py phraseloom/_template_workflow.py
git mv entity_cluster_probe.py phraseloom/_entity_cluster_probe.py
```

Expected: `git status --short` shows two renames.

- [ ] **Step 2: Add the `template_demo.py` compatibility shim**

Replace `template_demo.py` with:

```python
from __future__ import annotations

from phraseloom.cli import main
from phraseloom.excel_io import (
    _default_extract_output_path,
    _default_fill_output_path,
    _default_tm_output_path,
    _default_to_translate_output_path,
    _default_work_dir,
)
from phraseloom.interactive import _user_path
from phraseloom.template_engine import (
    apply_target_template,
    infer_target_template,
    parse_template,
)
from phraseloom.workflow import (
    fill_target_column_workbook,
    generate_tm_pairs,
    generate_workbook,
)

__all__ = [
    "apply_target_template",
    "fill_target_column_workbook",
    "generate_tm_pairs",
    "generate_workbook",
    "infer_target_template",
    "main",
    "parse_template",
    "_default_extract_output_path",
    "_default_fill_output_path",
    "_default_tm_output_path",
    "_default_to_translate_output_path",
    "_default_work_dir",
    "_user_path",
]


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 3: Add the `entity_cluster_probe.py` compatibility shim**

Replace `entity_cluster_probe.py` with:

```python
from __future__ import annotations

from phraseloom.entity_cluster import (
    EntityCluster,
    EntityOccurrence,
    find_entity_clusters,
    generate_entity_cluster_workbook,
    main,
)

__all__ = [
    "EntityCluster",
    "EntityOccurrence",
    "find_entity_clusters",
    "generate_entity_cluster_workbook",
    "main",
]


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 4: Verify moved files are still present**

Run: `rg --files phraseloom template_demo.py entity_cluster_probe.py`

Expected: shows both package-internal implementation files and both top-level shims.

- [ ] **Step 5: Commit**

Run:

```bash
git add phraseloom/_template_workflow.py phraseloom/_entity_cluster_probe.py template_demo.py entity_cluster_probe.py
git commit -m "Move script implementations into package internals"
```

---

### Task 3: Add Responsibility-Based Package Facades

**Files:**
- Create: `phraseloom/models.py`
- Create: `phraseloom/template_engine.py`
- Create: `phraseloom/workflow.py`
- Create: `phraseloom/excel_io.py`
- Create: `phraseloom/interactive.py`
- Create: `phraseloom/cli.py`
- Create: `phraseloom/entity_cluster.py`

- [ ] **Step 1: Add model exports**

Create `phraseloom/models.py` with:

```python
from __future__ import annotations

from ._entity_cluster_probe import EntityCluster, EntityOccurrence
from ._template_workflow import RowItem, TemplateMatch, TranslationUnit

__all__ = [
    "EntityCluster",
    "EntityOccurrence",
    "RowItem",
    "TemplateMatch",
    "TranslationUnit",
]
```

- [ ] **Step 2: Add template engine exports**

Create `phraseloom/template_engine.py` with:

```python
from __future__ import annotations

from ._template_workflow import (
    PLACEHOLDER_RE,
    VAR_RE,
    apply_target_template,
    infer_target_template,
    parse_template,
)

__all__ = [
    "PLACEHOLDER_RE",
    "VAR_RE",
    "apply_target_template",
    "infer_target_template",
    "parse_template",
]
```

- [ ] **Step 3: Add workflow exports and stable aliases**

Create `phraseloom/workflow.py` with:

```python
from __future__ import annotations

from ._template_workflow import (
    fill_target_column_workbook,
    generate_tm_pairs,
    generate_workbook,
)

extract_tm_pairs = generate_tm_pairs
prepare_translation = generate_workbook
fill_translation = fill_target_column_workbook

__all__ = [
    "extract_tm_pairs",
    "fill_target_column_workbook",
    "fill_translation",
    "generate_tm_pairs",
    "generate_workbook",
    "prepare_translation",
]
```

- [ ] **Step 4: Add Excel/path exports**

Create `phraseloom/excel_io.py` with:

```python
from __future__ import annotations

from ._template_workflow import (
    _default_extract_output_path,
    _default_fill_output_path,
    _default_legacy_output_path,
    _default_tm_output_path,
    _default_to_translate_output_path,
    _default_work_dir,
    _load_translated_units,
    _read_headers,
    _read_source_rows,
    _resolve_column,
    _write_output_workbook,
    _write_target_column_workbook,
    _write_tm_workbook,
    _write_to_translate_workbook,
)

__all__ = [
    "_default_extract_output_path",
    "_default_fill_output_path",
    "_default_legacy_output_path",
    "_default_tm_output_path",
    "_default_to_translate_output_path",
    "_default_work_dir",
    "_load_translated_units",
    "_read_headers",
    "_read_source_rows",
    "_resolve_column",
    "_write_output_workbook",
    "_write_target_column_workbook",
    "_write_tm_workbook",
    "_write_to_translate_workbook",
]
```

- [ ] **Step 5: Add interactive exports**

Create `phraseloom/interactive.py` with:

```python
from __future__ import annotations

from ._template_workflow import (
    _normalize_optional_column,
    _prompt_int,
    _prompt_text,
    _prompt_yes_no,
    _user_path,
    run_interactive,
)

__all__ = [
    "run_interactive",
    "_normalize_optional_column",
    "_prompt_int",
    "_prompt_text",
    "_prompt_yes_no",
    "_user_path",
]
```

- [ ] **Step 6: Add CLI entry point**

Create `phraseloom/cli.py` with:

```python
from __future__ import annotations

from ._template_workflow import main

__all__ = ["main"]


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 7: Add entity cluster exports**

Create `phraseloom/entity_cluster.py` with:

```python
from __future__ import annotations

from ._entity_cluster_probe import (
    EntityCluster,
    EntityOccurrence,
    find_entity_clusters,
    generate_entity_cluster_workbook,
    main,
)

__all__ = [
    "EntityCluster",
    "EntityOccurrence",
    "find_entity_clusters",
    "generate_entity_cluster_workbook",
    "main",
]
```

- [ ] **Step 8: Verify facade imports by static search**

Run:

```bash
rg -n "from \._template_workflow|from \._entity_cluster_probe" phraseloom
```

Expected: facade modules import from package-internal compatibility modules.

- [ ] **Step 9: Commit**

Run:

```bash
git add phraseloom/models.py phraseloom/template_engine.py phraseloom/workflow.py phraseloom/excel_io.py phraseloom/interactive.py phraseloom/cli.py phraseloom/entity_cluster.py
git commit -m "Add package responsibility facades"
```

---

### Task 4: Migrate Tests Into The Package Layout

**Files:**
- Move: `test_template_demo.py` -> `tests/test_template_workflow.py`
- Move: `test_entity_cluster_probe.py` -> `tests/test_entity_cluster.py`
- Modify: `tests/test_template_workflow.py`
- Modify: `tests/test_entity_cluster.py`

- [ ] **Step 1: Move test files**

Run:

```bash
New-Item -ItemType Directory -Force tests
git mv test_template_demo.py tests/test_template_workflow.py
git mv test_entity_cluster_probe.py tests/test_entity_cluster.py
```

Expected: `git status --short` shows both test files renamed.

- [ ] **Step 2: Update workflow test imports**

In `tests/test_template_workflow.py`, replace imports from `template_demo` as follows:

```python
from phraseloom.template_engine import apply_target_template, infer_target_template, parse_template
from phraseloom.workflow import fill_target_column_workbook, generate_tm_pairs, generate_workbook
from phraseloom.excel_io import (
    _default_extract_output_path,
    _default_fill_output_path,
    _default_tm_output_path,
    _default_to_translate_output_path,
    _default_work_dir,
)
from phraseloom.interactive import _user_path
from phraseloom.cli import main
```

Apply the replacements at each test import site so each test imports from the module matching the responsibility it exercises.

- [ ] **Step 3: Update entity cluster test imports**

In `tests/test_entity_cluster.py`, replace:

```python
from entity_cluster_probe import find_entity_clusters
```

with:

```python
from phraseloom.entity_cluster import find_entity_clusters
```

- [ ] **Step 4: Add compatibility shim tests**

Append two small tests to `tests/test_template_workflow.py`:

```python
class CompatibilityShimTests(unittest.TestCase):
    def test_template_demo_shim_exports_existing_workflow_api(self):
        from template_demo import generate_workbook, main, parse_template

        self.assertTrue(callable(generate_workbook))
        self.assertTrue(callable(main))
        self.assertEqual(parse_template("VIP10 Pack").template, "VIP{num1} Pack")

    def test_entity_cluster_probe_shim_exports_existing_api(self):
        from entity_cluster_probe import find_entity_clusters

        self.assertTrue(callable(find_entity_clusters))
```

- [ ] **Step 5: Verify test file imports**

Run:

```bash
rg -n "from template_demo import|from entity_cluster_probe import" tests
```

Expected: only the compatibility shim tests import the top-level modules.

- [ ] **Step 6: Commit**

Run:

```bash
git add tests/test_template_workflow.py tests/test_entity_cluster.py
git commit -m "Migrate tests to package imports"
```

---

### Task 5: Update Documentation For Internal Tool Usage

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Add development setup section**

Add a section near the top of `README.md`:

```markdown
## 开发环境

推荐使用虚拟环境安装为可编辑包：

```bash
python -m venv .venv
.venv\Scripts\python -m pip install -e .
```

运行测试：

```bash
python -m unittest discover -v
```
```

- [ ] **Step 2: Update command examples**

Keep existing `python3 template_demo.py ...` examples as compatibility notes, and add package CLI equivalents:

```bash
phraseloom tm-extract '/path/to/tm.xlsx' --source-col en --target-col fr
phraseloom extract '/path/to/source.xlsx' --source-col '英語' --target-col - --tm '/path/to/tm_l10n/tm_tm_pairs.xlsx' --no-existing-targets
phraseloom fill '/path/to/source.xlsx' --templates '/path/to/source_l10n/source_to_translate.xlsx' --source-col '英語' --target-col '法语' --mode target-column
```

- [ ] **Step 3: Add architecture note**

Add a short note:

```markdown
## 项目结构

核心代码位于 `phraseloom/` 包中。`template_demo.py` 和 `entity_cluster_probe.py` 保留为兼容入口；新代码应优先从 `phraseloom.workflow`、`phraseloom.template_engine`、`phraseloom.excel_io`、`phraseloom.entity_cluster` 等模块导入。
```

- [ ] **Step 4: Commit**

Run:

```bash
git add README.md
git commit -m "Document package development workflow"
```

---

### Task 6: Verify And Finish

**Files:**
- Inspect all changed files.

- [ ] **Step 1: Run whitespace check**

Run: `git -c core.excludesfile= diff --check`

Expected: no output and exit code 0.

- [ ] **Step 2: Run test command if Python is available**

Run: `python -m unittest discover -v`

Expected when Python and dependencies are installed: all tests pass.

If `python` is unavailable in this environment, run:

```bash
py -3 -m unittest discover -v
```

Expected when the launcher has Python installed: all tests pass.

If neither Python command is available, record the exact command failure in the final report.

- [ ] **Step 3: Inspect git status**

Run: `git -c core.excludesfile= status --short`

Expected: no unstaged or staged changes after the final implementation commit.

- [ ] **Step 4: Inspect recent commits**

Run: `git -c core.excludesfile= log --oneline -5`

Expected: shows commits for the plan and each implementation checkpoint.
