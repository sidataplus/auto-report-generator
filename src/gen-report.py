import argparse
from datetime import datetime
import os

import sqlalchemy as sql
from sqlalchemy.engine import URL
import pandas as pd

from mailmerge import MailMerge
from docxcompose.composer import Composer
from docx import Document
from docx2pdf import convert
from pikepdf import Pdf, Encryption, Permissions

parser = argparse.ArgumentParser()
for arg_name, kwargs in [
        ["driver",   {"default":  "ODBC Driver 17 for SQL Server"}],
        ["host",     {"required": True}],
        ["port",     {"default":  "1433"}],
        ["db",       {"required": True}],
        ["username", {"required": True}],
        ["password", {"required": True}],
        ["query",    {"required": True}],
        ["template", {"default":  "template/template.docx"}],
        ["output",   {"default":  "output_%Y-%m-%d_%H-%M-%S"}],
        ["pdf_password", {}]]:
    parser.add_argument(f"--{arg_name}", action="store", type=str, **kwargs)
args = parser.parse_args()

# Create database connection engine
print("connecting to the database ...")
engine = sql.create_engine(URL.create(
    "mssql+pyodbc", host=args.host, port=args.port, database=args.db,
    username=args.username, password=args.password,
    query={"driver": args.driver}))


# Query data from the connected database
print("querying the database ...")
df = pd.read_sql_query(args.query, engine).astype("string").to_dict("records")

# Parse potential timestamp in the string
filename = datetime.now().strftime(args.output)

# Create directory if not exists
os.makedirs("out/tmp", exist_ok=True)

print("filling each row on the template ...")
for index, row in enumerate(df):
    # Read template file with MailMerge tags
    # (need to be in for loop to prevent caching)
    document = MailMerge(args.template)
    # Fill in MailMerge tags with data fields from database
    document.merge(**row)
    # Save filled-in file as .docx
    document.write(f"out/tmp/{index}.docx")
    # Report progress
    print("[" + str(index) + "]", end="")

print("merging ...")
composer = Composer(Document("out/tmp/0.docx"))
for index, _ in enumerate(df):
    if index > 0:
        composer.append(Document(f"out/tmp/index.docx"))
    print("[" + str(index) + "]", end="")
composer.save("out/{filename}_tmp.docx")

# Convert .docx to .pdf
print("covert docx to pdf ...")
convert("out/{filename}_tmp.docx", f"out/{filename}_tmp.pdf")

# Encrypt and save the pdf
print("clean up tmp files ...")
Pdf.open(f"out/{filename}_tmp.pdf").save(
    r"out/{filename}.pdf",
    encryption=Encryption(user=args.pdf_password,
                          owner=args.pdf_password,
                          allow=Permissions(extract=False)))

# Initialize the pdf file
print("clean up tmp files ...")
os.removedirs("out/tmp")
os.remove(f"out/{filename}_tmp.docx")
os.remove(f"out/{filename}_tmp.pdf")

print(f"{filename}.pdf successfully generated! -- " +
      datetime.now().strftime("%Y-%m-%dT%H-%M-%S%z"))
