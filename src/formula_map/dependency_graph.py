"""Build a cell-level dependency graph from an opened workbook."""

from pathlib import Path

import networkx as nx

from .loader import LoadedWorkbook
from .parser import CellReference, FormulaParser


class DependencyGraph:
    """Directed graph where each formula cell points to its input cells."""

    def __init__(self, graph: nx.DiGraph, *, path: Path | None = None):
        self.graph = graph
        self.path = path

    @classmethod
    def from_workbook(
        cls,
        loaded: LoadedWorkbook,
        *,
        parser: FormulaParser | None = None,
    ) -> "DependencyGraph":
        parser = parser or FormulaParser()
        graph = nx.DiGraph()
        workbook = loaded.workbook

        for worksheet in workbook.worksheets:
            for row in worksheet.iter_rows():
                for cell in row:
                    if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                        continue
                    source = CellReference(worksheet.title, cell.coordinate)
                    graph.add_node(source, formula=cell.value)
                    for reference in parser.parse(cell.value, worksheet.title):
                        for dependency in reference.cells():
                            graph.add_edge(source, dependency)

        return cls(graph, path=loaded.path)

    def dependencies_of(self, cell: CellReference) -> tuple[CellReference, ...]:
        return tuple(self.graph.successors(cell))

    def dependents_of(self, cell: CellReference) -> tuple[CellReference, ...]:
        return tuple(self.graph.predecessors(cell))

    @property
    def formula_cells(self) -> tuple[CellReference, ...]:
        return tuple(
            node for node, data in self.graph.nodes(data=True) if "formula" in data
        )