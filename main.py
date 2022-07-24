from openpyxl import load_workbook, Workbook

workbook = load_workbook(filename="timesheets/callum.xlsx")
original_sheet = workbook.active

new_sheet = workbook.copy_worksheet(from_worksheet=original_sheet)
workbook.save(filename="timesheets/callum2.xlsx")

current_work_ending = None


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


row_count = 0
for idx, row in enumerate(new_sheet.iter_rows(min_row=6, min_col=1, max_col=14, values_only=True)):
    #print(f"{idx}")
    if idx < row_count:
        continue
    week_ending = row[0]
    rows_added, total_hours = add_up_hours(week_ending, idx + 6)

    row_count += rows_added
