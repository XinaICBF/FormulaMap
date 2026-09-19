from openpyxl import Workbook

from formula_map.dependency_graph import DependencyGraph
from formula_map.formula_parser import CellReference, FormulaParser
from formula_map.workbook_loader import WorkbookLoader


def test_parser_resolves_local_and_cross_sheet_ranges():
    parser = FormulaParser()

    references = parser.parse("=SUM(A1:B2)+'Inputs'!$C$3", "Model")

    assert references[0].sheet == "Model"
    assert references[0].start == "A1"
    assert references[0].end == "B2"
    assert references[1].sheet == "Inputs"
    assert references[1].cells() == (CellReference("Inputs", "C3"),)


def test_dependency_graph_keeps_formula_and_expands_ranges(tmp_path):
    path = tmp_path / "model.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Model"
    worksheet["A1"] = 1
    worksheet["A2"] = 2
    worksheet["B1"] = "=SUM(A1:A2)"
    workbook.save(path)
    workbook.close()

    loaded = WorkbookLoader().load(path)
    try:
        graph = DependencyGraph.from_workbook(loaded)
    finally:
        loaded.close()

    source = CellReference("Model", "B1")
    assert graph.graph.nodes[source]["formula"] == "=SUM(A1:A2)"
    assert set(graph.dependencies_of(source)) == {
        CellReference("Model", "A1"),
        CellReference("Model", "A2"),
    }