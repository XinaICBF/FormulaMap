"""Command-line entry point for inspecting an Excel dependency graph."""

import argparse
from pathlib import Path
import sys

if __package__:
    from .dependency_graph import DependencyGraph
    from .formula_parser import CellReference
    from .workbook_loader import WorkbookLoader
else:
    # Allow `python src/formula_map/main.py` from a source checkout.
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from formula_map.dependency_graph import DependencyGraph
    from formula_map.formula_parser import CellReference
    from formula_map.workbook_loader import WorkbookLoader


def _default_workbook() -> Path:
    candidates = sorted(Path("data/local").glob("*.xlsx"))
    if not candidates:
        raise FileNotFoundError("No .xlsx workbook found in data/local")
    return candidates[0]


def _parse_cell(value: str) -> CellReference:
    try:
        sheet, coordinate = value.rsplit("!", maxsplit=1)
    except ValueError as error:
        raise argparse.ArgumentTypeError(
            "cell must use the format SHEET!A1"
        ) from error
    if not sheet or not coordinate:
        raise argparse.ArgumentTypeError("cell must use the format SHEET!A1")
    return CellReference(sheet.strip("'"), coordinate.upper())


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Load an Excel workbook and inspect its formula dependencies."
    )
    parser.add_argument(
        "workbook",
        nargs="?",
        type=Path,
        help="workbook path; defaults to the first .xlsx file in data/local",
    )
    parser.add_argument(
        "--cell",
        type=_parse_cell,
        help="show dependencies and dependents for a cell, e.g. GHG!D10",
    )
    parser.add_argument(
        "--max-items",
        type=int,
        default=20,
        help="maximum dependency addresses to print (default: 20)",
    )
    return parser


def inspect_workbook(path: Path, target: CellReference | None, max_items: int) -> int:
    if max_items < 0:
        raise ValueError("max-items must be non-negative")

    loaded = WorkbookLoader().load(path)
    try:
        graph = DependencyGraph.from_workbook(loaded)
        print(f"workbook={path}")
        print(f"sheets={len(loaded.workbook.sheetnames)}")
        print(f"formula_nodes={len(graph.formula_cells)}")
        print(f"total_nodes={graph.graph.number_of_nodes()}")
        print(f"edges={graph.graph.number_of_edges()}")

        if target is not None:
            dependencies = graph.dependencies_of(target)
            dependents = graph.dependents_of(target)
            print(f"target={target.sheet}!{target.coordinate}")
            print(f"dependency_count={len(dependencies)}")
            print(f"dependent_count={len(dependents)}")
            print("dependencies=" + ", ".join(
                f"{cell.sheet}!{cell.coordinate}" for cell in dependencies[:max_items]
            ))
            print("dependents=" + ", ".join(
                f"{cell.sheet}!{cell.coordinate}" for cell in dependents[:max_items]
            ))
    finally:
        loaded.close()
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        path = args.workbook or _default_workbook()
        return inspect_workbook(path, args.cell, args.max_items)
    except (FileNotFoundError, ValueError) as error:
        parser.error(str(error))
    return 2


if __name__ == "__main__":
    raise SystemExit(main())