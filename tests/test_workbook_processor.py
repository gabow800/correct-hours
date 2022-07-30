from openpyxl import load_workbook

from correct_hours.workbook_processor import WorkbookProcessor


def test_process_file():
    workbook = load_workbook("data/timesheet-david.xlsx")
    processor = WorkbookProcessor(workbook)
    processor.process()
    old_sheet = workbook.get_sheet_by_name("Timesheet Details")
    new_sheet = workbook.get_sheet_by_name("Timesheet Details Copy")
    # Old total
    assert new_sheet["O8"].value == 40
    # New total
    assert new_sheet["P8"].value == "=SUM(G6:M8)"  # should sum up to 38 hours
    # Days worked
    assert new_sheet["Q8"].value == "=P8/7.6"  # should be 5 days
    # Corrected hours
    assert new_sheet["J6"].value != old_sheet["J6"].value
    assert old_sheet["J6"].value == 8
    assert new_sheet["J6"].value == 6
    # Other hours should stay the same
    assert new_sheet["G6"].value == old_sheet["G6"].value
    assert new_sheet["G7"].value == old_sheet["G7"].value
    assert new_sheet["G8"].value == old_sheet["G8"].value

    assert new_sheet["H6"].value == old_sheet["H6"].value
    assert new_sheet["H7"].value == old_sheet["H7"].value
    assert new_sheet["H8"].value == old_sheet["H8"].value

    assert new_sheet["I6"].value == old_sheet["I6"].value
    assert new_sheet["I7"].value == old_sheet["I7"].value
    assert new_sheet["I8"].value == old_sheet["I8"].value

    assert new_sheet["J7"].value == old_sheet["J7"].value
    assert new_sheet["J8"].value == old_sheet["J8"].value

    assert new_sheet["K6"].value == old_sheet["K6"].value
    assert new_sheet["K7"].value == old_sheet["K7"].value
    assert new_sheet["K8"].value == old_sheet["K8"].value

    assert new_sheet["L6"].value == old_sheet["L6"].value
    assert new_sheet["L7"].value == old_sheet["L7"].value
    assert new_sheet["L8"].value == old_sheet["L8"].value

    assert new_sheet["M6"].value == old_sheet["M6"].value
    assert new_sheet["M7"].value == old_sheet["M7"].value
    assert new_sheet["M8"].value == old_sheet["M8"].value

