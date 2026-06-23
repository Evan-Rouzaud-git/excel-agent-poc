# Architecture and Pipeline

This add-in is built around controlled execution. The model can propose intent, but the application decides whether the plan is valid, safe, and executable. That is the key engineering decision in the repository.

## Technical Intent

- Build a local-first Excel agent that is useful in a real workbook.
- Prevent the LLM from writing directly to Excel.
- Convert workbook state into a structured context snapshot.
- Execute only a fixed set of Office.js macros.
- Surface the result through logs, artifacts, and confirmations.

## Why Control Matters in an Excel Agent

Spreadsheets are mutable state. A bad plan can delete data, overwrite formulas, or create a misleading chart. For that reason, the code treats planning as data processing, not as code generation.

The planner may be wrong, incomplete, or overconfident. The executor therefore needs:

- a strong input contract
- schema validation
- deterministic macro dispatch
- human confirmation for risky steps
- test coverage that does not depend on a live Excel session

## High-Level System Flow

1. `getWorkbookContext` reads the workbook and produces a snapshot.
2. `planWithOllama` or `mockPlanner` proposes a plan from that snapshot and the user prompt.
3. `normalizePlan` resolves workbook references and prepares the plan for execution.
4. `sanitizePlan` removes unsupported fields and adds deterministic confirmation metadata.
5. `preSchemaRepair` and optional formula repair patch legacy or fragile plan details.
6. `validatePlan` and `validatePlanInvariants` verify the final contract.
7. `applyConfirmationsToPlan` and `autoAnswerConfirmations` apply runtime decisions.
8. `executePlan` runs the macros and returns logs, artifacts, and status.

`runAgentPipeline` in [src/taskpane/agent/pipeline/runAgentPipeline.ts](../src/taskpane/agent/pipeline/runAgentPipeline.ts) ties those steps together and returns both the final executed plan and the execution result.

## Workbook Context Snapshot

`getWorkbookContext` in [src/taskpane/context/getWorkbookContext.ts](../src/taskpane/context/getWorkbookContext.ts) collects:

- workbook metadata
- active sheet and current selection
- `selectionInBlockId` and `nearestBlockId`
- per-sheet used ranges
- tables, blocks, and charts
- capability and limitation notes
- capped previews and value bounds

The snapshot is intentionally bounded with limits such as `maxSheets`, `maxUsedRows`, `maxUsedCols`, and `maxPreviewRows`. That keeps the planner input small enough to be practical and avoids serializing the entire workbook.

When a sheet has no table but the used range looks table-like, the snapshot can promote it to a synthetic range block. That gives the planner a usable target without pretending the workbook model is perfect.

## Planning Contract: Strict JSON Plan

The planner is asked to return a single JSON object with a strict shape. [src/taskpane/agent/planner/systemPrompt.ts](../src/taskpane/agent/planner/systemPrompt.ts) constrains the allowed macro names and tells the model not to emit free-form code or extra metadata.

The runtime still re-validates everything:

- `parseJson` and `extractStrictJSONObject` strip noise and recover JSON.
- `canonicalizePlan` normalizes legacy shapes and operator formats.
- `validatePlan` in [src/taskpane/agent/planSchema.ts](../src/taskpane/agent/planSchema.ts) uses AJV.
- `validatePlanInvariants` adds a small set of business rules beyond schema shape.

A subtle but important detail: the schema currently tolerates `artifacts`, but the planner prompt forbids them and `sanitizePlan` strips them. The executor, not the planner, owns artifacts.

## Plan Lifecycle: raw -> normalize -> sanitize -> repair -> validate -> execute

### Raw

The planner output is treated as untrusted text. It can include markdown fences, extra prose, invalid keys, or partial JSON. The planner path handles that with parsing and a single repair retry when needed.

### Normalize

`normalizePlan` in [src/taskpane/agent/planner/normalizePlan.ts](../src/taskpane/agent/planner/normalizePlan.ts) binds the plan to the actual workbook context:

- resolves `blockRef` values to real blocks
- maps table names to workbook blocks when possible
- merges adjacent `table_view` steps
- rewrites references to newly created columns through `$lastAddedColumn`
- carries created headers across steps so later steps can still find them
- injects safe defaults for some formatting and output cases

This step exists because a planner output that is syntactically valid can still be semantically misaligned with the workbook.

### Sanitize

`sanitizePlan` in [src/taskpane/agent/planner/sanitizePlan.ts](../src/taskpane/agent/planner/sanitizePlan.ts) removes unsupported keys and keeps only known params for each macro. It also:

- normalizes `table_view` `select`, `filter`, `sort`, `rename`, and `dest`
- rewrites aliased headers when a join or a previous step created alternate names
- injects deterministic confirmation templates for `validate_data`
- resolves ambiguous join keys into explicit confirmation prompts
- drops planner-owned artifacts

### Repair

Repair is split into two practical layers:

- `preSchemaRepair` in [src/taskpane/agent/planner/ollamaPlanner.ts](../src/taskpane/agent/planner/ollamaPlanner.ts) fills missing step ids, coerces legacy `source` and `output` shapes, and performs a small amount of deterministic cleanup.
- `repairPlanWriteFormulas` in [src/taskpane/agent/planRepairer.ts](../src/taskpane/agent/planRepairer.ts) is an optional formula-specific hook that can patch `write_formula` steps using workbook headers and the user prompt.

There is also a small intent-based repair path for `table_view` when a newly created column should be selected in the next step. That prevents a common failure mode: the plan creates the column and then immediately stops referring to it.

### Validate

The final plan must pass:

- AJV schema validation
- plan invariants
- any additional runtime checks in the planner pipeline

If validation fails, the plan is not executed. If the planner could not produce a safe plan, the pipeline returns a failure state instead of guessing.

### Execute

`executePlan` in [src/taskpane/agent/executor.ts](../src/taskpane/agent/executor.ts) is deterministic:

- it canonicalizes the final plan
- it refuses invalid input
- it dispatches each step to a known macro
- it records logs and artifacts
- it stops when a step requests confirmation

## Deterministic Execution Model

The executor does not expose a generic scripting surface. It calls a known macro by name and passes a structured parameter object. That means the model cannot invent a new operation at runtime.

The supported macro surface is intentionally small:

- `table_view`
- `write_formula`
- `apply_format`
- `create_chart`
- `join_tables`
- `validate_data`
- `place_output`
- `summarize_actions`

A useful side effect of this design is that every workbook mutation is reviewable in code. If a macro changes behavior, the change is visible in a single function rather than hidden in prompt text.

## Controlled Macro Surface

### `table_view`

Projects, filters, sorts, and renames a block into a new table or an in-place view. It requires a selection when it creates a separate output and rejects incompatible combinations such as in-place projection with selected columns.

### `write_formula`

Writes a formula into a controlled column position. It supports structured references for table targets and A1 translation for range targets. In the fake host, it also precomputes values so downstream steps and tests can reason about the new column.

### `apply_format`

Applies a constrained formatting preset such as `corporate_blue`. The executor also uses it as a safe default on newly created table outputs so the demo looks coherent without manual formatting steps.

### `create_chart`

Resolves the source block, maps x/y columns by header or index, builds the source range from the workbook bounds, and places the chart relative to the source or on a new sheet.

### `join_tables`

Joins two workbook sources with explicit keys, selection policy, conflict handling, and output placement. It can ask for confirmation when the key choice is ambiguous.

### `validate_data`

Audits a source block for missing values, duplicates, and bad types. It can create an Issues sheet, highlight source cells, persist the issues reference in a named item, and then request deterministic yes/no confirmations for follow-up fixes. The key point is that the audit and the fix phase are separate.

### `place_output` and `summarize_actions`

These are helper macros for controlled placement and a final summary. They keep the core executor logic narrow.

## Human Confirmations and Runtime Control

Human input is part of the runtime, not a UI afterthought.

- `sanitizePlan` injects the `validate_data` confirmation templates based on the requested checks.
- `applyConfirmationsToPlan` rewrites the plan from user answers.
- `autoAnswerConfirmations` can resolve safe confirmations in demo or evaluation mode.
- `taskpane.ts` renders the confirmation buttons and resubmits the plan when the user picks an answer.
- `executePlan` returns `need_user_confirmation` when it cannot safely continue.

This is what makes the agent usable in a spreadsheet context: it can stop before a destructive action instead of trying to be clever.

## Logs, Artifacts, and Observability

Each execution returns structured output:

- timestamped logs
- artifacts created during the run
- status and error information
- requested confirmations when execution pauses

The planner pipeline also logs stage hashes such as `raw`, `canonical`, `sanitized`, and `validated`. That makes a run easier to debug without exposing the workbook contents in a free-form way.

The `validate_data` macro additionally persists the Issues table reference in the workbook named item `__validate_data_last_issues_ref`, so the follow-up apply pass can find the same output again.

## Testing Strategy

The test suite is built around a fake Excel host in [tests/mocks/fakeExcel.ts](../tests/mocks/fakeExcel.ts). It simulates enough of Office.js to exercise the pipeline without a live Excel session:

- workbook and worksheet objects
- ranges and writes
- tables and charts
- named items
- load/sync semantics

That fake host is important because many bugs in this kind of add-in are shape bugs, not pure logic bugs.

Fixtures in [tests/fixtures/workbooks.ts](../tests/fixtures/workbooks.ts) provide repeatable workbook snapshots for sales, join, and view scenarios. The evaluation runner in [tests/demoEvalRunner.ts](../tests/demoEvalRunner.ts) can run the same prompts in mock mode or against Ollama.

Covered behaviors include:

- workbook context extraction
- plan normalization and sanitization
- schema and invariant validation
- confirmation handling
- formula repair
- create_chart range resolution
- validate_data audit/apply flows
- end-to-end pipeline consistency

## Design Trade-offs

- The snapshot is bounded, which keeps planning practical but makes it incomplete.
- The macro surface is narrow, which reduces flexibility but improves safety.
- Normalization and sanitization sometimes rewrite the plan, which is less pure but much more robust.
- Optional Ollama keeps the demo local-first, but planner quality still depends on the model and prompt.
- The fake Excel host makes tests stable, but some Office.js edge cases still need real Excel validation.

## Known Limitations

- This is a POC, not a production platform.
- The workbook snapshot is not a full semantic model of Excel.
- Planner quality still depends on prompt quality and workbook context quality.
- Some behaviors are host-specific and may differ slightly between fake Excel and real Excel.
- The execution model is intentionally constrained, so it will not solve arbitrary spreadsheet tasks.

## Why These Choices Matter

The repository is stronger because it is easy to reason about:

- the model proposes intent
- the runtime validates that intent
- the executor only runs known macros
- confirmations exist for risky actions
- logs and artifacts make the run inspectable
- tests can run without Excel

That is the core engineering story in the repository.
