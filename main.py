from openpyxl import load_workbook, Workbook

workbook = load_workbook(filename="timesheets/callum.xlsx")
original_sheet = workbook.active

new_sheet = workbook.copy_worksheet(from_worksheet=original_sheet)

def add_up_hours(current_week_ending, row_start):
    rows_added = 0
    total_hours = 0
    for row in new_sheet.iter_rows(min_row=row_start, min_col=1, max_col=14, values_only=True):
        week_ending = row[0]
        if week_ending != current_week_ending:
            break
        total = row[13]
        total_hours += total
        rows_added += 1
    print(f"Total hours for {current_week_ending} is {total_hours} within {rows_added} rows")
    return rows_added, total_hours


def correct_hours(row_start, row_offset, overtime):
    print("Hours need to be corrected")
    time_left = overtime
    for col_idx in range(13, 7, -1):
        for row_idx in range((row_start + row_offset) - 1, row_start - 1, -1):
            value = new_sheet.cell(row_idx, col_idx).value
            if value <= 0:
                continue
            corrected_value = value - time_left
            if corrected_value < 0:
                time_left = corrected_value * -1
                corrected_value = 0
                new_sheet.cell(row_idx, col_idx, corrected_value)
            else:
                new_sheet.cell(row_idx, col_idx, corrected_value)
                return


row_count = 0
for idx, row in enumerate(new_sheet.iter_rows(min_row=6, min_col=1, max_col=17, values_only=True)):
    if idx < row_count:
        continue
    week_ending = row[0]
    row_idx = idx + 6
    rows_added, total_hours = add_up_hours(week_ending, row_idx)
    new_sheet.cell(row_idx, 15, total_hours)
    if total_hours > 38:
        overtime = total_hours - 38
        correct_hours(row_idx, rows_added, overtime)
    row_count += rows_added

workbook.save(filename="timesheets/callum2.xlsx")