import argparse
import sqlalchemy as sql
import pandas as pd
from mailmerge import MailMerge
from docx2pdf import convert
from pikepdf import Pdf, Encryption, Permissions
from datetime import datetime


parser = argparse.ArgumentParser()
parser.add_argument('--driver', action='store', type=str, default='ODBC+Driver+17+for+SQL+Server')
parser.add_argument('--db_host',  action='store', type=str, required=True)
parser.add_argument('--port', action='store', type=str, default='1433')
parser.add_argument('--db',  action='store', type=str, required=True)
parser.add_argument('--username',  action='store', type=str, required=True)
parser.add_argument('--password',  action='store', type=str, required=True)
parser.add_argument('--query',  action='store', type=str, required=True)
parser.add_argument('--template', action='store', type=str, default='../template/template.docx')
parser.add_argument('--output', action='store', type=str, default='output')
parser.add_argument('--pdf_password', action='store', type=str)
args = parser.parse_args()

# Create database connection engine
engine = sql.create_engine(
    f'mssql+pyodbc://{args.username}:{args.password}@{args.db_host}:{args.port}/{args.db}?driver={args.driver}', echo=True)

# Query data from the connected database
df = pd.read_sql(args.query, engine)

# Read template file with MailMerge tags
document = MailMerge(args.template)

# Fill in MailMerge tags with data fields from database row by row
document.merge_templates(df.to_dict('records'), separator='page_brake')

# Save filled-in file as .docx
document.write(f'../out/docx/{args.output}.docx')

# Convert .docx to .pdf
convert(f'../out/docx/{args.output}.docx', f'../out/docx/{args.output}.pdf')

# Secure the pdf
with Pdf.open(f'../out/docx/{args.output}.pdf') as pdf:
	pdf.save(
	    f'../out/pdf/{args.output}.pdf',
	    encryption=Encryption(user=args.pdf_password, allow=Permissions(extract=False))
	)

print(f'{args.output}.pdf successfully generated! – {datetime.now().strftime("%Y-%m-%dT%H-%M-%S%z")}')
