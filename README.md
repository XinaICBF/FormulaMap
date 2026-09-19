
# FormulaMap

FormulaMap reads Excel formulas without replacing them with cached values, extracts cell and range references, and builds a directed dependency graph.

The graph direction is **formula cell -> referenced cell**. This makes `dependencies_of()` useful for finding the inputs needed to redraw a formula's related data, while `dependents_of()` finds formulas affected by an input.

## Quick start

```python
from formula_map import CellReference, DependencyGraph, WorkbookLoader

loaded = WorkbookLoader().load("data/local/model.xlsx")
try:
	dependencies = DependencyGraph.from_workbook(loaded)
	cell = CellReference("GHG", "D10")
	print(dependencies.dependencies_of(cell))
finally:
	loaded.close()
```

Install the project in the repository virtual environment with:

```bash
.venv/bin/pip install -e .
```

Run the tests with:

```bash
.venv/bin/pytest
```

Files under `data/local/` are intentionally ignored by git because they may contain sensitive workbooks.

## Current scope

- `.xlsx`, `.xlsm`, `.xltx`, and `.xltm` workbooks
- A1-style cell and rectangular range references
- Local and cross-worksheet references, including quoted worksheet names
- Cell-level graph expansion for rectangular ranges

Structured references, whole-row/whole-column references, external workbook links, and Excel calculation are intentionally left for later stages. `openpyxl` reads formulas but does not calculate them.
