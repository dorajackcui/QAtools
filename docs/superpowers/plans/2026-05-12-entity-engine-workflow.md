# Entity Engine Workflow Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the first independent entity workflow that splits a preprocessed todo workbook, extracts reusable entity structures, prefills from entity TM, fills approved entity rows, and merges them back with non-entity rows.

**Architecture:** Add a focused `phraseloom.entity_workflow` module for workbook-level orchestration and strategy-driven entity extraction. Keep entity discovery pluggable through a small strategy interface, with the first strategy wrapping `find_entity_clusters()` from `_entity_cluster_probe.py`. Extend CLI with separate `entity-split`, `entity-prefill`, `entity-fill`, and `entity-merge` commands without touching `tag_engine` or `template_engine`.

**Tech Stack:** Python, openpyxl, unittest, existing PhraseLoom workbook schema and entity cluster probe.

---

## File Structure

- Create `phraseloom/entity_workflow.py`: pure entity workflow orchestration, workbook reading/writing helpers, default cluster strategy, split/prefill/fill/merge public APIs.
- Modify `phraseloom/workbook_schema.py`: add sheet and column constants for entity workbooks.
- Modify `phraseloom/cli.py`: add entity workflow subcommands that call `entity_workflow`.
- Modify `phraseloom/entity_cluster.py`: optionally export workflow helpers only if needed after the workflow module is stable.
- Create `tests/test_entity_workflow.py`: red/green workflow tests covering split, prefill/fill, missing-entity blocking, and merge.

## Tasks

### Task 1: Split A Todo Workbook Into Entity And Non-Entity Files

**Files:**
- Create: `tests/test_entity_workflow.py`
- Create: `phraseloom/entity_workflow.py`
- Modify: `phraseloom/workbook_schema.py`

- [ ] **Step 1: Write the failing split test**

Add a test that creates a `to_translate` workbook with two reusable entity rows and one unrelated row. Call `split_entity_workbook(..., min_group_size=2)`. Assert the entity workbook contains the two reusable rows, `entity_structures`, `entity_terms`, and `entity_source_map`; assert the non-entity workbook contains only the unrelated row.

- [ ] **Step 2: Run the split test to verify RED**

Run: `py -3 -m unittest tests.test_entity_workflow.EntityWorkflowTests.test_split_creates_parallel_entity_and_non_entity_workbooks -v`

Expected: FAIL because `phraseloom.entity_workflow` does not exist.

- [ ] **Step 3: Add schema constants and minimal split implementation**

Add constants for `entity_structures`, `entity_terms`, `entity_source_map`, `original_index`, `structure_id`, `source_structure`, `target_structure`, `term_id`, `source_entity`, `target_entity`, `entities_json`, `preview_target`, and `fill_status`.

Implement `split_entity_workbook(input_path, entity_output_path, non_entity_output_path, min_group_size=3)` in `phraseloom/entity_workflow.py`. It reads `to_translate`, runs `ClusterProbeStrategy`, maps cluster row numbers back to todo rows, writes the two workbooks, and preserves non-`to_translate` sheets in the non-entity workbook.

- [ ] **Step 4: Run the split test to verify GREEN**

Run: `py -3 -m unittest tests.test_entity_workflow.EntityWorkflowTests.test_split_creates_parallel_entity_and_non_entity_workbooks -v`

Expected: PASS.

### Task 2: Prefill Entity Structures And Terms From Entity TM

**Files:**
- Modify: `tests/test_entity_workflow.py`
- Modify: `phraseloom/entity_workflow.py`

- [ ] **Step 1: Write the failing prefill test**

Create a target entity workbook with blank `target_structure` and `target_entity`, and a TM entity workbook with matching completed rows. Call `prefill_entity_workbook(...)`. Assert matching structures and terms are copied into the target workbook.

- [ ] **Step 2: Run the prefill test to verify RED**

Run: `py -3 -m unittest tests.test_entity_workflow.EntityWorkflowTests.test_prefill_copies_matching_structure_and_terms -v`

Expected: FAIL because `prefill_entity_workbook` is missing.

- [ ] **Step 3: Implement prefill**

Implement `prefill_entity_workbook(entity_input_path, tm_path, output_path)`. It loads TM `entity_structures` by `source_structure` and TM `entity_terms` by `source_entity`. A single matching target value is copied. Multiple distinct target values leave the target blank and append an `ambiguous_*` warning.

- [ ] **Step 4: Run the prefill test to verify GREEN**

Run: `py -3 -m unittest tests.test_entity_workflow.EntityWorkflowTests.test_prefill_copies_matching_structure_and_terms -v`

Expected: PASS.

### Task 3: Fill Completed Entity Rows Back Into The Entity-Related Todo Sheet

**Files:**
- Modify: `tests/test_entity_workflow.py`
- Modify: `phraseloom/entity_workflow.py`

- [ ] **Step 1: Write the failing fill test**

Use an entity workbook whose structure and terms are `ready`. Call `fill_entity_workbook(...)`. Assert `entity_source_map.preview_target` and the entity workbook `to_translate.target_unit` are populated with the target structure after replacing `{entityN}` placeholders with translated terms.

- [ ] **Step 2: Write the missing-term blocking test**

Use an entity workbook with a ready target structure but one missing `target_entity`. Call `fill_entity_workbook(...)`. Assert the matching todo row remains blank and `fill_status` is `missing_entity_translation`.

- [ ] **Step 3: Run fill tests to verify RED**

Run: `py -3 -m unittest tests.test_entity_workflow.EntityWorkflowTests.test_fill_writes_ready_entity_targets tests.test_entity_workflow.EntityWorkflowTests.test_fill_blocks_when_entity_translation_is_missing -v`

Expected: FAIL because `fill_entity_workbook` is missing.

- [ ] **Step 4: Implement fill**

Implement `fill_entity_workbook(entity_input_path, output_path)`. It enforces structure status, structure target, term status, term target, and todo row consistency before writing a target. It writes precise `fill_status` and `warning` values for blocked rows.

- [ ] **Step 5: Run fill tests to verify GREEN**

Run: `py -3 -m unittest tests.test_entity_workflow.EntityWorkflowTests.test_fill_writes_ready_entity_targets tests.test_entity_workflow.EntityWorkflowTests.test_fill_blocks_when_entity_translation_is_missing -v`

Expected: PASS.

### Task 4: Merge Entity And Non-Entity Workbooks

**Files:**
- Modify: `tests/test_entity_workflow.py`
- Modify: `phraseloom/entity_workflow.py`

- [ ] **Step 1: Write the failing merge test**

Use a filled entity-related workbook and a non-entity workbook. Call `merge_entity_workbooks(...)`. Assert the merged `to_translate` sheet restores original order and contains both entity-filled target rows and untouched non-entity rows.

- [ ] **Step 2: Run merge test to verify RED**

Run: `py -3 -m unittest tests.test_entity_workflow.EntityWorkflowTests.test_merge_restores_original_order_and_targets -v`

Expected: FAIL because `merge_entity_workbooks` is missing.

- [ ] **Step 3: Implement merge**

Implement `merge_entity_workbooks(entity_path, non_entity_path, output_path)`. It verifies unique `original_index`, required `unit_id`, required `source_unit`, no index gaps, and writes a merged `to_translate` sheet without `original_index`. It copies support sheets from the non-entity workbook.

- [ ] **Step 4: Run merge test to verify GREEN**

Run: `py -3 -m unittest tests.test_entity_workflow.EntityWorkflowTests.test_merge_restores_original_order_and_targets -v`

Expected: PASS.

### Task 5: Add CLI Commands And Full Verification

**Files:**
- Modify: `phraseloom/cli.py`
- Modify: `tests/test_entity_workflow.py`

- [ ] **Step 1: Write CLI dispatch tests**

Add lightweight tests for `_dispatch(["entity-split", ...])`, `_dispatch(["entity-prefill", ...])`, `_dispatch(["entity-fill", ...])`, and `_dispatch(["entity-merge", ...])` using temp workbooks.

- [ ] **Step 2: Run CLI tests to verify RED**

Run: `py -3 -m unittest tests.test_entity_workflow.EntityWorkflowCliTests -v`

Expected: FAIL because the commands are not registered.

- [ ] **Step 3: Register CLI commands**

Add `_main_entity_split`, `_main_entity_prefill`, `_main_entity_fill`, and `_main_entity_merge` to `phraseloom/cli.py`. Use explicit output arguments with sensible defaults beside the input workbook.

- [ ] **Step 4: Run targeted tests**

Run: `py -3 -m unittest tests.test_entity_workflow -v`

Expected: PASS.

- [ ] **Step 5: Run full verification**

Run: `py -3 -m unittest discover -v`

Expected: PASS with no errors.

## Self-Review

Spec coverage: the plan covers split, independent entity/non-entity files, TM prefill, final entity fill, merge, strategy extensibility, and CLI access. It preserves the constraint that `tag_engine` and `template_engine` are untouched.

Placeholder scan: no task relies on an unspecified future step; each task names exact files, functions, and verification commands.

Type consistency: public API names are `split_entity_workbook`, `prefill_entity_workbook`, `fill_entity_workbook`, and `merge_entity_workbooks`; the same names are used across tests, implementation, and CLI tasks.
