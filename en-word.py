import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
import traceback

doc = DocxTemplate("CLEARANCE_REPORT.docx")
my_name = "Frank Andrade"
my_phone = "(123) 456-789"
my_email = "frank@gmail.com"
my_address = "123 Main Street, NY"
today_date = datetime.today().strftime("%d %b, %Y")

my_context = {
    
    'my_name': my_name,
    'my_phone': my_phone,
    'my_email': my_email,
    'my_address': my_address,
    'TODAY': today_date
}

df = pd.read_csv('CUSTOM_DUTY_DETAILS_&_DEMURRAGE_CHARGES.csv')

for index, row in df.iterrows():
    context = {
        'PO': row['PO'],
        'DOPO': '',  # Add the appropriate value for DOPO
        'UC': row['Ultimate Consignee'],
        'S/A': '',  # Add the appropriate value for S/A
        'BL No.': row['BL / AWB No.'],
        'BL Date': row['BL / AWB Date.'],
        'BE No.': row['BE No.'],
        'BE Date': row['BE Date'],
        'Late Fine': row.get('Late Fine', 'N/A'),
        'Interest': row.get('Interest Charges', 'N/A'),
        'Demurrage': row.get('Demurrage Charges', 'N/A')
    }

    context.update(my_context)

   
    doc.render(context)
    doc.save(f"generated_doc_{index}.docx")
   
