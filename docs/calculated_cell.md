# Calculated Cell guide

This document explains how to use the Calculated Cell feature in Jinja Excel Template Engine.

## Overview

Add a `formula` attribute to any `<cell>` to evaluate a computed value during Excel generation. By default, the engine evaluates the expression server-side (deterministic). You may opt to emit a pass-through Excel formula for the client to compute by setting `pass_through="true"`.

Supported features:

- Basic arithmetic expressions (e.g., `1 + 2 * 3`).
- Built-in functions: `SUM`, `AVERAGE`, `CONCAT`, `ROUND`, `INT`, `ABS`, `MAX`, `MIN`, `DATE`.
- A1-style references and ranges: `A1`, `SUM(A1:C10)`.
- Cross-sheet ranges: `SUM(Sheet1!A1:Sheet1!A10)` (use `Sheet1!A1:A10` syntax).

## Server-side evaluation

Example (simple expression):

```xml
<cell type="Number" formula="1 + 2" />
```

Example (SUM of a range):

```xml
<row>
  <cell type="Number" value="1" />
  <cell type="Number" value="2" />
  <cell type="Number" value="3" />
</row>
<row>
  <cell type="Number" formula="SUM(A1:C1)" /> <!-- server-side evaluated value -> 6 -->
</row>
```

## Pass-through (client-side Excel formula)

To emit a formula to Excel rather than calculating it server-side, set `pass_through="true"` on a cell:

```xml
<cell type="Number" formula="SUM([1,2,3])" pass_through="true" />
```

This will write an Excel formula string into the cell (e.g., `=SUM(1,2,3)`), causing Excel to compute it when the file is opened.

## A1 and cross-sheet references

Server-side evaluation supports simple A1 references and ranges using current sheet's values that have already been generated (top-down). Cross-sheet references are supported via `SheetName!A1` or `SheetName!A1:A5` and are resolved when those sheets have values available.

Note: Server-side evaluation is deterministic only for references to cells that have already been written; references to future cells or circular references will cause an error and must be handled via pass-through.

## Limitations

- Some complex Excel functions and pivot/range operations are not yet implemented; pass-through will help for advanced Excel-specific behavior.
- Circular references are detected and reported as errors during server-side evaluation.
- For performance and memory considerations, please use `pass_through` for extremely large spreadsheets or heavy formulas when server-side evaluation is impractical.

## Examples and tooling

We include integration tests and a benchmark harness in `tests/`. To run the example in README, use:

```bash
python -m pytest tests/calculated_cell_test.py -q
```
