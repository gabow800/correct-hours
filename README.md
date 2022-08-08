# Correct hours script

## Install script

---

Install **Python 3** for your operating system, for example you can use the [Windows installer](https://www.python.org/downloads/windows/).

> **Note:** The commands below use **python3**. If your system doesn't have this command then use **python** instead.

After installing, make sure you can run this command: 
```bash
python3 --version
```

Install the script:
```bash
pip3 install correct_hours
```

After installing the script, make sure you can run this command:
```bash
python3 -m correct_hours            
```

You should get the following message after running the command above. Even it's an error message, this means that
you have installed the script successfully.
```bash
usage: __main__.py [-h] directory
__main__.py: error: the following arguments are required: directory
```

## Using the script

---

First make sure you have the file `rates.xlsx` located in the same directory where your reports live. You can download 
and use [this file](./examples/xero/rates.xlsx) as an example.

Run script and pass the location of your report files, for example:

```bash
python3 -m correct_hours C:\Users\user\Downloads\reports
```

The project will generate a folder `output/` in the same location with a copy of the files corrected.

## Upgrading the script

---

Whenever there is an update of the script, you can run the following command to get the latest changes:

```bash
pip3 install correct_hours --upgrade
```

After upgrading, you should get a message indicating the new version of the script, for example:

```bash
Successfully installed correct-hours-0.1.5
```

If there is no message indicating the new version, then it means you have the latest version.