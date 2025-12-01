---
description: "Tasks to implement the Calculated Cell feature"
---

# Tasks: Calculated Cell

**Input**: spec.md, plan.md

## Phase 1: Setup (Shared Infrastructure)

- [ ] T001 Create project structure per implementation plan (src modules for evaluator and tests)
- [ ] T002 [P] Add pre-commit config for Black, isort, flake8
- [ ] T003 [P] Add mypy config and baseline for public APIs
- [ ] T004 Configure CI to run tests, lints, type-checks, and capture coverage
- [ ] T005 Add pytest-benchmark harness and document baseline performance metrics


- [x] T006 Implement safe formula parsing and AST structure `src/jinja_excel_template/formula_eval/`
- [x] T007 [P] Implement `FormulaEvaluator` API: parse(), evaluate(context) => result
- [x] T008 [P] Implement `EvaluationContext` and `FormulaCache` modules (basic cache present)
- [x] T009 Add contract tests for invalid formulas, circular references, and pass-through behavior

## Phase 3: User Story 1 - Basic Evaluation (P1)

- [x] T010 Unit tests: parse and evaluate simple arithmetic expressions
- [x] T011 Integration test: render a Jinja template that includes `formula` in cell and verify Excel output cell value
- [ ] T012 Add API docs and READMEs with basic example

## Phase 4: User Story 2 - Excel formula and function support (P2)

- [x] T013 Add function implementations for SUM, AVERAGE, CONCAT, DATE, ROUND, INT, ABS, MAX, MIN as a minimum subset
- [x] T014 Add tests for function correctness and edge-case handling
- [x] T015 Pass-through mode: add `pass_through` attribute handling and tests

## A1 & Cross-Sheet Ranges

- [x] T016 Add support for A1 ranges and cross-sheet references using `RANGE` and `GET` helpers
- [x] T017 Add tests for range handling (SUM(A1:C1), SUM(Sheet1!A1:B1))

## Phase 5: User Story 3 - Safety & Performance (P3)

- [x] T016 Add caching and optimization for repeated formula evaluation (parse-time cache via lru_cache implemented)
- [x] T017 Add benchmarking CI job and performance-related tests for P95 latency (pytest-benchmark + workflow added)
- [x] T018 Add sandboxing for evaluator to prevent filesystem/network access (safe AST visitor; disallows imports and attrs)
- [ ] T019 Add migration notes and a compatibility check for any public API changes

## Phase N: Polish & Cross-Cutting Concerns

- [ ] T020 Documentation: Add examples, quickstart, and troubleshooting sections in README and docs/.
- [ ] T021 Refactor to improve test coverage and fix any uncovered edge cases
- [ ] T022 Add additional unit tests to cover all new functions and AST operations
 - [x] T020 Documentation: Add examples, quickstart, and troubleshooting sections in README and docs/.
 - [x] T021 Refactor to improve test coverage and fix any uncovered edge cases
 - [x] T022 Add additional unit tests to cover all new functions and AST operations
 - [x] T023 Add CI benchmark threshold check script and baseline/thresholds
