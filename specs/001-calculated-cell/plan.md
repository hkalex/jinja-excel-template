# Implementation Plan: Calculated Cell

**Branch**: `001-calculated-cell` | **Date**: 2025-12-01 | **Spec**: spec.md
**Input**: Feature specification from `/specs/001-calculated-cell/spec.md`

## Summary

Add calculated cell support for Jinja templates used with the Jinja Excel Template Engine. This includes a safe formula evaluator for common functions, pass-through mode to emit Excel formulas, caching for repeated expressions, and performance/benchmark validation to ensure the library remains efficient for large spreadsheets.

## Technical Context

**Language/Version**: Python 3.11  
**Primary Dependencies**: Keep minimal: rely on standard Python, optional safe formula interpreter or in-house evaluator (final choice made during Phase 1).  
**Storage**: Not applicable (in-memory processing).  
**Testing**: pytest for unit/integration tests, pytest-benchmark or similar for micro-benchmarks.  
**Target Platform**: Cross-platform library consumers (Linux, macOS, Windows).  
**Performance Goals**: 1k calculated cells P95 generation time < 2s on a modern dev machine.  
**Constraints**: Sandbox evaluator — no filesystem or network access; composition of complex Excel features (pivot, advanced formulas) out-of-scope for initial implementation.

## Constitution Check

The plan includes the required entries to pass the Constitution Check:

- Code Quality: pre-commit with Black/isort; Flake8 and mypy for type checks; run on CI.
- Testing Standards: Unit tests and integration tests cover functionality; contract tests for error handling and pass-through; benchmark harness and performance tests added.
- API Stability/UX: New `formula` attribute on `<cell>` is optional and backward-compatible; old templates remain unchanged.
- Performance: Benchmarks defined and baseline for P95; plan includes caching strategy for repeated expressions.
- Documentation: Update README with quickstart example and `docs/` with detailed API usage.

## Project Structure

``text
specs/001-calculated-cell
├── spec.md
├── plan.md
├── checklists/requirements.md
└── tasks.md
```

## Complexity Tracking

The feature will avoid adding heavy third-party dependencies by implementing a lightweight evaluator for common functions in the first iteration and optionally adding a dependency in a follow-up if necessary.
