"""Parse Excel cell and range references from formulas."""

from dataclasses import dataclass
import re

from openpyxl.formula import Tokenizer
from openpyxl.utils.cell import range_boundaries


_CELL = r"\$?[A-Z]{1,3}\$?\d+"
_REFERENCE = re.compile(
    rf"^(?:(?P<sheet>'(?:[^']|'')+'|[^!]+)!)?"
    rf"(?P<start>{_CELL})(?::(?P<end>{_CELL}))?$",
    re.IGNORECASE,
)


@dataclass(frozen=True, order=True)
class CellReference:
    """A single worksheet cell address."""

    sheet: str
    coordinate: str


@dataclass(frozen=True)
class FormulaReference:
    """A cell or rectangular range referenced by a formula."""

    sheet: str
    start: str
    end: str
    raw: str

    @property
    def is_single_cell(self) -> bool:
        return self.start == self.end

    def cells(self) -> tuple[CellReference, ...]:
        min_col, min_row, max_col, max_row = range_boundaries(
            f"{self.start}:{self.end}"
        )
        return tuple(
            CellReference(self.sheet, f"{column}{row}")
            for row in range(min_row, max_row + 1)
            for column in _column_names(min_col, max_col)
        )


def _column_names(min_col: int, max_col: int):
    for value in range(min_col, max_col + 1):
        column = ""
        number = value
        while number:
            number, remainder = divmod(number - 1, 26)
            column = chr(65 + remainder) + column
        yield column


def _unquote_sheet(sheet: str) -> str:
    if sheet.startswith("'") and sheet.endswith("'"):
        return sheet[1:-1].replace("''", "'")
    return sheet


class FormulaParser:
    """Extract worksheet cell references from an Excel formula."""

    def parse(self, formula: str, current_sheet: str) -> tuple[FormulaReference, ...]:
        if not isinstance(formula, str) or not formula.startswith("="):
            return ()

        references: list[FormulaReference] = []
        for token in Tokenizer(formula).items:
            if token.type != "OPERAND" or token.subtype != "RANGE":
                continue
            for raw_reference in token.value.split(","):
                match = _REFERENCE.fullmatch(raw_reference.strip())
                if match is None:
                    continue
                sheet = match.group("sheet")
                sheet = _unquote_sheet(sheet) if sheet else current_sheet
                start = match.group("start").replace("$", "").upper()
                end = (match.group("end") or start).replace("$", "").upper()
                references.append(
                    FormulaReference(
                        sheet=sheet,
                        start=start,
                        end=end,
                        raw=raw_reference.strip(),
                    )
                )
        return tuple(references)