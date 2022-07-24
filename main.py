from workbook_processor import WorkbookProcessor

filepath = "input/callum.xlsx"
output_folder = "output"
processor = WorkbookProcessor(filepath, output_folder)
processor.process()
