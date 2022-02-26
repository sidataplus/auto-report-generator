import argparse
import sqlalchemy as sql
from sqlalchemy.engine import URL
import pandas as pd
from mailmerge import MailMerge
from docx2pdf import convert
from pikepdf import Pdf, Encryption, Permissions
from datetime import datetime


parser = argparse.ArgumentParser()
for arg_name, default_value in [
        ("driver", "ODBC Driver 17 for SQL Server"),
        ("host", None), ("port", "1433"), ("db", None),
        ("username", None), ("password", None), ("query", None),
        ("template", "template/template.docx"),
        ("output", "output_%Y-%m-%d_%H-%M-%S"), ("pdf_password", None)]:
    parser.add_argument(f"--{arg_name}",
                        action="store", type=str,
                        required=bool(default_value),
                        default=default_value)
args = parser.parse_args()

# Create database connection engine
connection_url = URL.create(
    "mssql+pyodbc", host=args.host, port=args.port, database=args.db,
    username=args.username, password=args.password,
    query={"driver": args.driver})

engine = sql.create_engine(connection_url)

# Parse potential timestamp in the string
filename = datetime.now().strftime(args.output)

# Query data from the connected database
df = pd.read_sql_query(args.query, engine).astype("string").to_dict("records")

# Read template file with MailMerge tags
document = MailMerge(args.template)

# Initialize the pdf file
pdf = Pdf.new()

print("processing... ")
for index, row in enumerate(df):
    # Fill in MailMerge tags with data fields from database row by row
    document.merge_templates([row], separator="page_break")
    # Save filled-in file as .docx
    document.write(f"out/{filename}_tmp.docx")
    # Convert .docx to .pdf
    convert(f"out/{filename}_tmp.docx", f"out/{filename}_tmp.pdf")
    pdf.pages.extend(Pdf.open(f"out/{filename}_tmp.pdf").pages)
    # Report progress
    print("[" + index + "]", end="")

# Encrypt and save the pdf
pdf.save("out/{filename}.pdf",
         encryption=Encryption(user=args.pdf_password,
                               owner=args.pdf_password,
                               allow=Permissions(extract=False)))

print(f"{filename}.pdf successfully generated! -- " +
      datetime.now().strftime("%Y-%m-%dT%H-%M-%S%z"))
