# Feature Specification: Calculated Cell

**Feature Branch**: `001-calculated-cell`  
**Created**: 2025-12-01  
**Status**: Draft  
**Input**: User description: "Add calculated cell support: allow Jinja templates to declare computed cells evaluated during ExcelXML generation"

## User Scenarios & Testing *(mandatory)*

### User Story 1 - Basic calculated cell evaluation (Priority: P1)

As a template author, I want to define a calculated cell in Jinja so the engine evaluates an expression during ExcelXML generation and writes its computed value to the generated Excel file.

**Why this priority**: This is the minimal feature that makes the new capability usable — templates that need computed values can set them without external pre-processing.

**Independent Test**: Provide a Jinja template that declares an Excel cell with a calculated attribute (e.g., `<cell formula="={{ var1 + var2 }}" type="Number" />`), run the engine, and verify the output Excel file contains the computed numerical value (not an unresolved formula text) in the expected cell.

**Acceptance Scenarios**:

1. **Given** a cell definition with a supported formula and known inputs, **When** the template is rendered via engine, **Then** the generated Excel cell value equals the computed result (e.g., numbers are summed correctly).
2. **Given** a cell formula that uses type functions (date/number formatting), **When** formula is rendered, **Then** output is typed correctly for Excel (Number vs Date vs Text) and reflected by the Excel file.

---

### User Story 2 - Excel formula pass-through and function support (Priority: P2)

As a template author, I want to use common Excel formula syntax and standard functions (SUM, AVERAGE, CONCAT, etc.) in either of two modes:

- Server-side evaluation mode (the engine evaluates and writes the computed value). This gives deterministic results and avoids Excel recalculation differences.
- Pass-through mode (the engine emits the formula text for Excel to compute). This is optional and should be controllable with an attribute.

**Why this priority**: Broader compatibility with existing Excel templates and functional parity improves adoption.

**Independent Test**: Create templates that use typical Excel functions in both modes, then verify server-side evaluation values match expected results and/or verify pass-through formula text is present in generated Excel files.

**Acceptance Scenarios**:

1. **Given** a template using SUM on a set of values, **When** server-side evaluation is enabled, **Then** the generated Excel cell shows a numeric sum equal to expected value.
2. **Given** the pass-through mode, **When** the generated file is opened by Excel and recalculation occurs, **Then** the formula cell displays the same value as server-side evaluation in typical cases.

---

### User Story 3 - Safety, caching and large dataset performance (Priority: P3)

As a service implementer, I want formula evaluation to be safe, predictable, and performant so the engine won't execute arbitrary code, and large spreadsheets with many computed cells do not cause OOM or unacceptable delays.

**Why this priority**: It is critical for security and scalability but can be added after basic evaluation capability is implemented.

**Independent Test**: Use large test data (e.g., 2k rows with a calculated cell each) and verify that:

- The engine completes generation without using excessive memory (no OOM on typical CI or local dev machine).
- P95 latency for generating outputs is within the specified target (see performance goals). For performance tests include a repeatable micro-benchmark for common scenarios.

**Acceptance Scenarios**:

1. **Given** 2k rows with a simple arithmetic formula, **When** the engine runs, **Then** the process finishes without OOM and meets P95 latency goal.
2. **Given** a template with a formula that references other cells, **When** a formula causes a circular reference, **Then** engine detects the circular reference and raises a deterministic, actionable error.

---

### Edge Cases

- Circular references in formulas — must be detected and reported as an actionable error.
- Invalid or unsupported functions — engine should report an informative error and optionally fallback to pass-through mode if configured.
- References to other sheets, ranges, or named ranges — support a minimal, documented subset and refuse unsupported patterns.
- Mixed data types — behavior should mirror Excel semantics where practical; explicit conversion rules must be documented.

## Requirements *(mandatory)*

### Functional Requirements

- **FR-001**: Templates MUST be able to declare a calculated cell using attributes like `formula` in the `<cell>` tag. Example: `<cell formula="={{ order_total + tax }}" type="Number" />`.
- **FR-002**: The engine MUST evaluate formulas using a safe evaluator (no direct Python eval). The evaluator supports a defined subset of operations and functions. By default the engine performs server-side evaluation (deterministic). A `pass_through="true"` attribute on the cell allows emitting an Excel formula string so the Excel client can compute values on open. (Choice: Option C selected.)
- **FR-003**: If a formula cannot be resolved (unsupported function / circular ref / missing data), the engine MUST emit a clear error and either:
  - Fail the generation (default), or
  - Fall back to emitting an Excel formula text if `pass_through` is set to true for the cell.
- **FR-004**: The calculated output MUST set the Excel cell type (Number, Date, Text) as inferred from the evaluation result or as explicitly specified by the template.
- **FR-005**: Provide a FormulaEvaluator module that parses and evaluates formula expressions safely; public API must be limited to the inputs and known functions.
- **FR-005**: Provide a FormulaEvaluator module that parses and evaluates formula expressions safely; public API must be limited to the inputs and known functions.
- **FR-006**: The engine MUST optionally support caching of repeated formula results when the same expression is evaluated multiple times to improve performance.
- **FR-006**: The engine MUST optionally support caching of repeated formula results when the same expression is evaluated multiple times to improve performance (parse-time caching implemented).
- **FR-007**: The feature MUST include unit tests (per function), integration tests (converting templates with formulas to Excel), and contract tests for expected behaviors.
- **FR-008**: Performance tests: For documents with up to 1k calculated cells, P95 generation time MUST be less than 2s on a standard development machine (document how P95 is measured — see plan.md for details).
- **FR-009**: For security, formula evaluation MUST be sandboxed: no access to the filesystem, network, or arbitrary executables.
- **FR-010**: The engine MUST support A1-style cell references and ranges (`A1`, `A1:C1`, `Sheet1!A1:B3`) in server-side evaluation for already-set cells.

### Key Entities *(include if feature involves data and processing)*

- **CalculatedCellDefinition**: Representation of a template cell with a formula attribute, including attributes: raw formula string, type hint (optional), pass_through flag, caching hint, and cell location.
- **FormulaEvaluator**: A module that parses and evaluates formula expressions safely; includes an AST for supported functions and a mapping of supported functions to implementations.
- **EvaluationContext**: The set of variables and data that are visible to formula evaluation (e.g., row variables, named variables in template rendering environment).
- **FormulaCache**: Caching store for repeated expressions; stores keyed by expression + canonicalized arguments and context to prevent incorrect reuse.

## Success Criteria *(mandatory)*

### Measurable Outcomes

- **SC-001 (Correctness)**: 100% of reference formula test cases (50+ diverse formulas: arithmetic, SUM, AVERAGE, CONCAT, DATE) return expected results under server-side and/or pass-through mode in unit tests.
- **SC-002 (Performance)**: Generation of a document with 1k basic computed cells completes with P95 latency < 2s on CI or a standard dev machine as documented by the plan's benchmark harness.
- **SC-003 (Safety)**: No insecure operations are possible through formulas; arithmetic, string, date functions allowed; filesystem/network access must be prevented (verified via contract tests that attempt to invoke disallowed functions and expect a failure).
- **SC-004 (Developer Experience)**: Public API and template syntax documented, with quickstart example added to README; enabling the feature requires no more than adding `formula` attribute to cells for simple cases.

### Constitution Compliance (Required)

All specs MUST include the entry items below: Code Quality (linters/type-checks used), Testing plan (unit/integration/contract tests), API compatibility notes if public API changes, and Performance goals/benchmarks if the feature affects runtime/resource usage.

**Assumptions**:

- We're adding a new `formula` attribute to `<cell>` that optionally supports `pass_through` and `type` attributes. This adds no breaking changes to existing template format unless optional attributes are malformed.
- The engine will include a safe formula evaluator implemented in a sandboxed environment (either via a lightweight interpreter implemented in-project or an independent safe evaluator dependency). The evaluator will not permit arbitrary code execution or access to the filesystem/network.
- Default behavior is server-side evaluation; pass-through mode is opt-in by setting `pass_through="true"` on the cell.

**[NEEDS CLARIFICATION: Q1] Formula evaluation vs pass-through**
**Context:** This feature can either (A) always evaluate formulas server-side and write computed values; or (B) emit formulas to Excel and let Excel compute; or (C) allow both modes with the default being server-side evaluation.

**Suggested Answers**:

| Option | Answer | Implications |
|--------|--------|--------------|
| A      | Server-side evaluation only | Deterministic results; implementation requires a safe evaluator and performance consideration; reduces need to rely on Excel recalculation; easier to write tests. |
| B      | Emit formulas only, let Excel calculate | Simpler to implement; more dependence on Excel variants; reduces testing determinism; formulas in output may not be computed until Excel recalculates. |
| C      | Both: default server-side with explicit `pass_through=true` option | Best of both worlds; more implementation complexity (two modes) but adds flexibility for templates and user preference. |

**Your choice**: _[Wait for user response — pick A/B/C or Custom]_ 

**Next steps**: Once the Q1 selection is confirmed, update FR-002 to require the correct mode(s), and adjust the plan to add evaluator (for A/C) or pass-through handling (for B/C).
# Feature Specification: [FEATURE NAME]

**Feature Branch**: `[###-feature-name]`  
**Created**: [DATE]  
**Status**: Draft  
**Input**: User description: "$ARGUMENTS"

## User Scenarios & Testing *(mandatory)*

<!--
  IMPORTANT: User stories should be PRIORITIZED as user journeys ordered by importance.
  Each user story/journey must be INDEPENDENTLY TESTABLE - meaning if you implement just ONE of them,
  you should still have a viable MVP (Minimum Viable Product) that delivers value.
  
  Assign priorities (P1, P2, P3, etc.) to each story, where P1 is the most critical.
  Think of each story as a standalone slice of functionality that can be:
  - Developed independently
  - Tested independently
  - Deployed independently
  - Demonstrated to users independently
-->

# Feature Specification: Calculated Cell

[Describe this user journey in plain language]

**Why this priority**: [Explain the value and why it has this priority level]

**Independent Test**: [Describe how this can be tested independently - e.g., "Can be fully tested by [specific action] and delivers [specific value]"]

**Acceptance Scenarios**:

1. **Given** [initial state], **When** [action], **Then** [expected outcome]
2. **Given** [initial state], **When** [action], **Then** [expected outcome]


### User Story 2 - [Brief Title] (Priority: P2)

[Describe this user journey in plain language]

**Why this priority**: [Explain the value and why it has this priority level]

**Independent Test**: [Describe how this can be tested independently]

**Acceptance Scenarios**:


### User Story 3 - [Brief Title] (Priority: P3)

[Describe this user journey in plain language]

**Why this priority**: [Explain the value and why it has this priority level]

**Independent Test**: [Describe how this can be tested independently]

**Acceptance Scenarios**:

1. **Given** [initial state], **When** [action], **Then** [expected outcome]


[Add more user stories as needed, each with an assigned priority]

<!--
  ACTION REQUIRED: The content in this section represents placeholders.
  Fill them out with the right edge cases.


  Fill them out with the right functional requirements.


### Key Entities *(include if feature involves data)*
## Success Criteria *(mandatory)*

<!--
  ACTION REQUIRED: Define measurable success criteria.
  These must be technology-agnostic and measurable.

### Constitution Compliance (Required)
