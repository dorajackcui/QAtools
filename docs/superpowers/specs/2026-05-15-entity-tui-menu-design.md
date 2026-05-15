# Entity TUI Menu Design

## Purpose

Add an interactive entity workflow menu that mirrors the existing template
workflow menu. The goal is to make the recommended four-step entity workflow
discoverable without requiring users to remember subcommand names or exact
argument order.

The existing non-interactive entity subcommands remain supported. This change
adds a guided terminal entry point on top of the current workflow functions.

## Scope

The TUI covers only the recommended four-step entity workflow:

```text
1. Build entity memory from TM reusable units
2. Prepare source entity pack
3. Fill completed entity pack
4. Merge filled entity pack back to translator todo
```

The advanced/debug entity commands stay available as direct CLI subcommands, but
they are not shown in the interactive entity menu.

## Entry Points

The existing top-level interactive command gains an entity option:

```text
phraseloom
  4) Entity workflow
```

Two direct entry points open the entity menu immediately:

```text
phraseloom entity
phraseloom entity-interactive
```

These direct entry points are interactive-only aliases. The existing
non-interactive commands keep their current names:

```text
phraseloom entity-tm ...
phraseloom entity-prepare ...
phraseloom entity-fill-pack ...
phraseloom entity-merge-pack ...
```

## Menu Flow

The entity menu uses the same prompt style as the existing template menu:

```text
Entity Workflow

1) Build entity memory from TM reusable units
2) Prepare source entity pack
3) Fill completed entity pack
4) Merge filled entity pack back to translator todo
b) Back
q) Quit
```

`b` returns to the top-level menu when the entity menu was reached from
`phraseloom`. When the menu was reached through `phraseloom entity` or
`phraseloom entity-interactive`, `b` exits successfully.

`q` always exits successfully after printing the same short quit message used by
the existing interactive flow.

## Prompt Behavior

Step 1 asks for:

```text
TM reusable units path
Output entity memory workbook
Minimum variants for a reusable entity structure
```

The output prompt defaults to `default_entity_memory_output_path(input_path)`.
The minimum group size defaults to `3`.

Step 2 asks for:

```text
Translator todo path
Entity memory path (- for none)
Output entity pack workbook
Minimum variants for a reusable entity structure
```

The entity memory prompt is optional. `-`, empty, `none`, `no`, and `skip` mean
no entity memory workbook. The output prompt defaults to
`default_entity_pack_output_path(input_path)`. The minimum group size defaults
to `3`.

Step 3 asks for:

```text
Source entity pack path
Output filled entity pack workbook
```

The output prompt defaults to `default_entity_filled_pack_output_path(input)`.
The TUI does not expose the `--in-place` mode; users who need in-place updates
can continue using the direct CLI command.

Step 4 asks for:

```text
Filled source entity pack path
Output merged translator todo workbook
```

The output prompt defaults to `default_entity_merged_todo_output_path(input)`.

All path prompts continue to accept copied shell quotes through the existing
`_user_path` helper.

## Implementation Shape

`phraseloom.interactive` owns the TUI functions because the current template
interactive flow already lives there. The new functions should call the same
entity workflow functions used by `phraseloom.cli`:

```text
extract_entity_memory_workbook
prepare_entity_pack_workbook
fill_entity_pack_workbook
merge_entity_pack_workbook
```

`phraseloom.cli._dispatch` routes `entity` and `entity-interactive` to the new
entity interactive function. It also keeps routing `phraseloom` and
`phraseloom interactive` to the existing top-level interactive menu, now with
the added entity option.

## Output And Errors

Each TUI step prints the same stats as the matching non-interactive command.
This keeps the CLI and TUI results consistent and avoids a second reporting
format.

Business errors should continue to flow through `PhraseLoomError` and the
existing `main()` error handling. Input validation for integer prompts continues
to use `_prompt_int`.

## Tests

Add focused tests for:

```text
Top-level interactive menu shows Entity workflow
Top-level option 4 enters the entity menu
phraseloom entity opens the entity menu directly
phraseloom entity-interactive opens the entity menu directly
Entity menu step 1 creates an entity memory workbook
Entity menu step 2 creates an entity pack with optional memory prefill
Entity menu step 3 creates a filled entity pack
Entity menu step 4 creates a merged translator todo
```

The workbook tests should reuse the existing entity workflow test helpers where
possible.

## Out Of Scope

This change does not add a full end-to-end wizard that runs all four steps in
one session. The workflow has a necessary human editing break between preparing
and filling the entity pack.

This change does not add an advanced/debug submenu. Low-level commands remain
available through direct CLI invocation and top-level `--help`.

This change does not alter workbook schemas, default output paths, or existing
direct subcommand behavior.
