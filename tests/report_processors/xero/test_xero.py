from pathlib import Path
from typing import Tuple, List

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from correct_hours.report_processors.xero import XeroReportProcessor, ROW_OFFSET
from correct_hours.utils import get_col_number, get_col_name

HOUR_START_ROW = 6


def assert_other_cells_should_be_the_same(
        old_sheet: Worksheet,
        new_sheet: Worksheet,
        skip_cells: List[str]
) -> None:
    for row_idx, row in enumerate(new_sheet.iter_rows(min_row=ROW_OFFSET, values_only=True)):
        row_number = row_idx + ROW_OFFSET
        if row_number == new_sheet.max_row:
            # skip bottom row
            continue
        for col_number in range(get_col_number("A"), get_col_number("N") + 1):
            col_name = get_col_name(col_number)
            cell_name = f"{col_name}{row_number}"
            if cell_name in skip_cells:
                continue
            old_value = old_sheet[cell_name].value
            new_value = new_sheet[cell_name].value
            assert old_value == new_value, (
                f"{cell_name} should have the same values in both sheets: {old_value} == {new_value}"
            )


def assert_corrected_hours(
        old_sheet: Worksheet,
        new_sheet: Worksheet,
        corrected_cells: List[Tuple[str, int, int]]
) -> None:
    for corrected_cell in corrected_cells:
        cell_name, expected_old_value, expected_new_value = corrected_cell
        old_value = old_sheet[cell_name].value
        new_value = new_sheet[cell_name].value
        assert old_value != new_value
        assert old_value == expected_old_value
        assert new_value == expected_new_value


def assert_cell_values(new_sheet: Worksheet, cells: List[Tuple[str, str]]) -> None:
    for cell in cells:
        cell_name, expected_value = cell
        assert new_sheet[cell_name].value == expected_value


def test_process_file() -> None:
    hours_workbook = load_workbook("tests/data/xero-report.xlsx")
    rates_workbook = load_workbook("tests/data/rates.xlsx")
    processor = XeroReportProcessor(hours_workbook, rates_workbook)
    processor.process()
    Path(f"tests/data/output").mkdir(parents=True, exist_ok=True)
    hours_workbook.save(filename="tests/data/output/copy-xero-report.xlsx")
    old_sheet = hours_workbook["Timesheet Details"]
    new_sheet = hours_workbook["Timesheet Details Copy"]

    # Assert corrected hours cells
    corrected_hours = [
        ("I8", 2, 0),
        ("I7", 2, 0),
        ("I6", 2, 0),
        ("H8", 2, 0),
        ("H7", 7, 6),
        ("K13", 8, 6),
        ("I16", 2, 0),
        ("I15", 2, 0),
        ("I14", 2, 0),
        ("H16", 2, 0),
        ("H15", 7, 6),
    ]
    assert_corrected_hours(old_sheet, new_sheet, corrected_hours)
    # Other cells should be the same
    skip_cells = [corrected_hour[0] for corrected_hour in corrected_hours]
    assert_other_cells_should_be_the_same(
        old_sheet,
        new_sheet,
        skip_cells
    )

    # Assert "new total" column
    for row_number in range(ROW_OFFSET, new_sheet.max_row):
        assert_cell_values(new_sheet, [
            (f"O{row_number}", f"=SUM(G{row_number}:M{row_number})"),
        ])

    # Asert "New total per day" and "Days worked" columns
    rows_with_total = [
        8,
        9,
        10,
        12,
        13,
        16,
    ]
    previous_row_with_total = 6
    for row_number in rows_with_total:
        assert_cell_values(new_sheet, [
            (f"P{row_number}", f"=SUM(O{previous_row_with_total}:O{row_number})"),
            (f"Q{row_number}", f"=P{row_number}/7.6"),
        ])
        previous_row_with_total = row_number + 1

    # Assert "Rates" column
    expected_rates = [
        "25.52",
        "36.24",
        "48.32",
        "48.32",
        "25.52",
        "25.52",
        "36.24",
        "25.52",
    ]
    for rate_idx, rate in enumerate(expected_rates):
        row_number = ROW_OFFSET + rate_idx
        assert_cell_values(new_sheet, [
            (f"R{row_number}", f"=O{row_number}*{rate}")
        ])

