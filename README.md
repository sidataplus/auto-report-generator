# auto-report-generator

Automatically generate filled-in reports in PDF
with data from MSSQL Server using Mail Merge & Python

## Requirements

- python 3.9
- poetry ^1.1 (`pip install poetry`) [Documentation](https://python-poetry.org/docs/)
- MSSQL Server ODBC 17 on [macOS](https://docs.microsoft.com/en-us/sql/connect/odbc/linux-mac/install-microsoft-odbc-driver-sql-server-macos?view=sql-server-ver15), [Linux](https://docs.microsoft.com/en-us/sql/connect/odbc/linux-mac/installing-the-microsoft-odbc-driver-for-sql-server?view=sql-server-ver15), [Windows](https://docs.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver15)

## How-to use

### with Poetry (recommended)

```
cd [path-of-the-repository-root-directory]
poetry env use [path-of-python3.9-executable-file]
poetry install
poetry run python src/gen-report.py \
  --host         [database host server]\
  --db           [database name]\
  --username     [username]\
  --password     [password]\
  --query        [SQL Query]\
  --template     [template filename.docx]\
  --output       [filename without extension, see remark 3]\
  --pdf_password [password for the generated pdf file]
```

### without Poetry

```
cd [path-of-the-repository-root-directory]
[path-of-python3.9-executable-file] -m pip install -r requirements.txt
[path-of-python3.9-executable-file] python src/gen-report.py
  [argument the same as using Peotry]
```

### Remark
1. Please replace contents between `[...]` appropriately.
2. If you are using windows, please use `"` instead of `'`
   to quote arguments.
3. The argument for `--output` will be parsed to python function `strftime`,
   which allow user to timestamp, e.g. `%H:%M:%S`, in the PDF filename. See
   https://docs.python.org/3/library/datetime.html#strftime-strptime-behavior
   for the complete list of available format.

## References

- [Populating MS Word Templates with Python by Chris
  Moffitt](https://pbpython.com/python-word-template.html)
