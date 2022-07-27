from pathlib import Path
from workbook_processor import WorkbookProcessor
from argparse import ArgumentParser


parser = ArgumentParser()
parser.add_argument("directory", help="Location of Excel files", type=str)

args = parser.parse_args()
directory = args.directory

# create output folder
Path(f"{directory}/output").mkdir(parents=True, exist_ok=True)
files = Path(directory).glob('*')
for f in files:
    if f.is_file():
        if not str.startswith(f.name, "~"):
            filepath = f.absolute()
            print(f"Processing file {filepath}...")
            processor = WorkbookProcessor(filepath)
            processor.process()
            new_file_name = processor.get_new_file_name()
            print(f"Finished processing file. Created file {new_file_name}.")
