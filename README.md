# auto-report-generator

Automatically generate filled-in reports in PDF with data from MSSQL Server using Mail Merge & Python

## Requirements

- python 3.9
- poetry ^1.1 (`pip install poetry`) [Documentation](https://python-poetry.org/docs/)
- MSSQL Server ODBC 17 on [macOS](https://docs.microsoft.com/en-us/sql/connect/odbc/linux-mac/install-microsoft-odbc-driver-sql-server-macos?view=sql-server-ver15), [Linux](https://docs.microsoft.com/en-us/sql/connect/odbc/linux-mac/installing-the-microsoft-odbc-driver-for-sql-server?view=sql-server-ver15), [Windows](https://docs.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server?view=sql-server-ver15)

## How-to use

### with Poetry

1. Install required python packages with commands `poetry install`

2. Run script

```
poetry run python src/gen-report.py \
    --host [database host server] --db [database name] \
    --username [username] --password [password] --query [SQL Query] \
    --template [template filename.docx] --output [filename without extension] \
    --pdf_password [password for the generated pdf file]
```

replace contents in [...] appropriately

### without Poetry

1. Install required python packages with commands `pip install -r requirements.txt`
2. Run script

```
   python src/gen-report.py \
       --host [database host server] --db [database name] \
       --username [username] --password [password] --query [SQL Query] \
       --template [template filename.docx] --output [filename without extension] \
       --pdf_password [password for the generated pdf file]
```

## References

- [Populating MS Word Templates with Python by Chris Moffitt](https://pbpython.com/python-word-template.html)
