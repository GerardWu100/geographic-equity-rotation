"""Convert the project's source workbook to CSV without changing cell data.

The converter reads the Office Open XML files stored inside the workbook. It
preserves row order, column positions, blank cells, text, and the exact numeric
text stored by Excel. The workbook contains no formulas, so formula evaluation
is intentionally outside this script's scope.
"""

from __future__ import annotations

import csv
import re
import xml.etree.ElementTree as ElementTree
from pathlib import Path
from zipfile import ZipFile


PROJECT_ROOT = Path(__file__).resolve().parent.parent
INPUT_PATH = PROJECT_ROOT / "data" / "raw" / "Index_ETF_Data.xlsx"
OUTPUT_PATH = PROJECT_ROOT / "data" / "processed" / "Index_ETF_Data.csv"

SPREADSHEET_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
XML_NAMESPACE = f"{{{SPREADSHEET_NAMESPACE}}}"
WORKSHEET_PATH = "xl/worksheets/sheet1.xml"
SHARED_STRINGS_PATH = "xl/sharedStrings.xml"
CELL_REFERENCE_PATTERN = re.compile(r"^([A-Z]+)([0-9]+)$")


def column_letters_to_index(column_letters: str) -> int:
    """Convert Excel column letters to a zero-based column index.

    Parameters
    ----------
    column_letters
        Excel column label such as ``A``, ``V``, or ``AA``.

    Returns
    -------
    int
        Zero-based column position.
    """

    index = 0
    for letter in column_letters:
        index = index * 26 + (ord(letter) - ord("A") + 1)
    return index - 1


def parse_cell_reference(cell_reference: str) -> tuple[int, int]:
    """Parse an Excel cell reference into zero-based row and column indexes.

    Parameters
    ----------
    cell_reference
        Cell address in A1 notation, such as ``V6938``.

    Returns
    -------
    tuple[int, int]
        Row index followed by column index, both zero-based.

    Raises
    ------
    ValueError
        If the reference is not valid A1 notation.
    """

    match = CELL_REFERENCE_PATTERN.fullmatch(cell_reference)
    if match is None:
        raise ValueError(f"Invalid cell reference: {cell_reference}")

    column_letters, row_number = match.groups()
    return int(row_number) - 1, column_letters_to_index(column_letters)


def read_shared_strings(workbook: ZipFile) -> list[str]:
    """Read Excel's shared-string table in its original order.

    Parameters
    ----------
    workbook
        Open Excel workbook archive.

    Returns
    -------
    list[str]
        Text value for each shared-string index. Rich-text runs are joined in
        display order without changing their text.
    """

    root = ElementTree.parse(workbook.open(SHARED_STRINGS_PATH)).getroot()
    shared_strings: list[str] = []

    for string_item in root.findall(f"{XML_NAMESPACE}si"):
        text_parts = [
            text_node.text or ""
            for text_node in string_item.iter(f"{XML_NAMESPACE}t")
        ]
        shared_strings.append("".join(text_parts))

    return shared_strings


def read_worksheet_values(
    workbook: ZipFile,
    shared_strings: list[str],
) -> list[list[str]]:
    """Read the first worksheet while preserving its rectangular dimensions.

    Parameters
    ----------
    workbook
        Open Excel workbook archive.
    shared_strings
        Shared strings indexed in the same order as Excel stores them.

    Returns
    -------
    list[list[str]]
        Rectangular worksheet matrix. Missing Excel cells remain empty strings,
        and stored numeric text is copied exactly.

    Raises
    ------
    ValueError
        If the worksheet has no declared dimension or contains an unsupported
        cell type.
    """

    worksheet_root = ElementTree.parse(workbook.open(WORKSHEET_PATH)).getroot()
    dimension = worksheet_root.find(f"{XML_NAMESPACE}dimension")
    if dimension is None or dimension.get("ref") is None:
        raise ValueError("The worksheet does not declare its used range.")

    final_reference = dimension.get("ref", "").split(":")[-1]
    final_row_index, final_column_index = parse_cell_reference(final_reference)
    rows = [
        [""] * (final_column_index + 1)
        for _ in range(final_row_index + 1)
    ]

    for cell in worksheet_root.iter(f"{XML_NAMESPACE}c"):
        cell_reference = cell.get("r")
        if cell_reference is None:
            raise ValueError("Encountered a worksheet cell without a reference.")

        row_index, column_index = parse_cell_reference(cell_reference)
        value_node = cell.find(f"{XML_NAMESPACE}v")
        stored_value = "" if value_node is None else value_node.text or ""
        cell_type = cell.get("t")

        if cell_type == "s":
            value = shared_strings[int(stored_value)]
        elif cell_type in (None, "n", "str", "e"):
            value = stored_value
        elif cell_type == "b":
            value = "TRUE" if stored_value == "1" else "FALSE"
        elif cell_type == "inlineStr":
            value = "".join(
                text_node.text or ""
                for text_node in cell.iter(f"{XML_NAMESPACE}t")
            )
        else:
            raise ValueError(
                f"Unsupported cell type {cell_type!r} at {cell_reference}."
            )

        rows[row_index][column_index] = value

    return rows


def write_csv(rows: list[list[str]], output_path: Path) -> None:
    """Write worksheet rows as a UTF-8, comma-delimited CSV file.

    Parameters
    ----------
    rows
        Rectangular matrix of worksheet values.
    output_path
        Destination CSV path.

    Returns
    -------
    None
        The function writes the CSV file to disk.
    """

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8", newline="") as csv_file:
        writer = csv.writer(csv_file, dialect="excel", lineterminator="\n")
        writer.writerows(rows)


def verify_csv(rows: list[list[str]], output_path: Path) -> None:
    """Verify that the CSV exactly matches the extracted worksheet matrix.

    Parameters
    ----------
    rows
        Worksheet matrix used to create the CSV.
    output_path
        CSV file to read back and compare.

    Returns
    -------
    None
        The function returns after a successful cell-by-cell comparison.

    Raises
    ------
    ValueError
        If any row, column, blank position, or populated value differs.
    """

    with output_path.open("r", encoding="utf-8", newline="") as csv_file:
        csv_rows = list(csv.reader(csv_file, dialect="excel"))

    if csv_rows != rows:
        raise ValueError("CSV verification failed: output differs from the workbook.")


def main() -> None:
    """Convert and verify the project's source workbook.

    Returns
    -------
    None
        The function writes and verifies ``data/processed/Index_ETF_Data.csv``.
    """

    with ZipFile(INPUT_PATH) as workbook:
        shared_strings = read_shared_strings(workbook)
        rows = read_worksheet_values(workbook, shared_strings)

    write_csv(rows, OUTPUT_PATH)
    verify_csv(rows, OUTPUT_PATH)

    populated_cells = sum(value != "" for row in rows for value in row)
    print(
        f"Wrote {OUTPUT_PATH.relative_to(PROJECT_ROOT)}: "
        f"{len(rows)} rows x {len(rows[0])} columns, "
        f"{populated_cells} populated cells verified."
    )


if __name__ == "__main__":
    main()
