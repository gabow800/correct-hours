# Correct hours script

## Instructions

Install [Python 3](https://www.python.org/downloads/windows/) using the Windows installer.

After installing, make sure you can run this command: 
```bash
python3 --version
```

Download this project as zip file `Code > Download ZIP`. 

Unzip file.

Go to the unzipped location in the command line:
```bash
cd C:\Users\user\Downloads\correct-hours 
```

Install project dependencies
```bash
pip3 install -r requirements.txt
```

Run project with the location of your Excel files:

```bash
python3 main.py -d C:\Users\user\Downloads\excel-files
```

The project will generate a folder `output` in the same location with a copy of the files corrected.