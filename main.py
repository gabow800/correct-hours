import sys
from pathlib import Path

from workbook_processor import WorkbookProcessor

input_folder = sys.argv[1]
# create output folder
Path(f"{input_folder}/output").mkdir(parents=True, exist_ok=True)
files = Path(input_folder).glob('*')
for f in files:
    if f.is_file():
        filepath = f.absolute()
        if not str.startswith(f.name, "~"):
            print(f"Processing file {filepath}...")
            processor = WorkbookProcessor(filepath)
            processor.process()
            new_file_name = processor.get_new_file_name()
            print(f"Finished processing file. Created file {new_file_name}.")
