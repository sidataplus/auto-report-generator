import argparse
import sqlalchemy as sql
from sqlalchemy.engine import URL
import pandas as pd
from mailmerge import MailMerge
from docx2pdf import convert
from pikepdf import Pdf, Encryption, Permissions
from datetime import datetime


parser = argparse.ArgumentParser()
parser.add_argument('--driver', action='store', type=str, default='ODBC Driver 17 for SQL Server')
parser.add_argument('--host',  action='store', type=str, required=True)
parser.add_argument('--port', action='store', type=str, default='1433')
parser.add_argument('--db',  action='store', type=str, required=True)
parser.add_argument('--username',  action='store', type=str, required=True)
parser.add_argument('--password',  action='store', type=str, required=True)
parser.add_argument('--query',  action='store', type=str, required=True)
parser.add_argument('--template', action='store', type=str, default='template/template.docx')
parser.add_argument('--output', action='store', type=str, default='output')
parser.add_argument('--pdf_password', action='store', type=str)
args = parser.parse_args()

# Create database connection engine
connection_url = URL.create(
	'mssql+pyodbc',
	username=args.username,
	password=args.password,
	host=args.host,
	port=args.port,
	database=args.db,
	query={
		'driver': 'ODBC Driver 17 for SQL Server'
	}
)

engine = sql.create_engine(connection_url)

# Query data from the connected database
df = pd.read_sql_query(args.query, engine).astype('string')

# Read template file with MailMerge tags
document = MailMerge(args.template)

# Fill in MailMerge tags with data fields from database row by row
document.merge_templates(df.to_dict('records'), separator='page_break')

# Save filled-in file as .docx
document.write(f'out/docx/{args.output}.docx')

# Convert .docx to .pdf
convert(f'out/docx/{args.output}.docx', f'out/pdf/{args.output}.pdf')

# Encrypt the pdf
with Pdf.open(f'out/pdf/{args.output}.pdf', allow_overwriting_input=True) as pdf:
	pdf.save(
	    f'out/pdf/{args.output}.pdf',
	    encryption=Encryption(user=args.pdf_password, owner=args.pdf_password, allow=Permissions(extract=False))
	)

print(f'{args.output}.pdf successfully generated! – {datetime.now().strftime("%Y-%m-%dT%H-%M-%S%z")}')
