import pandas as pd
import glob
from fpdf import FPDF
import os

filepaths = glob.glob('invoices\\*.xlsx')

for filepath in filepaths:
    file_name_and_date = os.path.basename(filepath)[:-5].split('-')

    
    pdf = FPDF(orientation='P', format='A4', unit='mm')
    pdf.add_page()

    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=50, h=8, txt=f'Invoice mr. {file_name_and_date[0]}', ln=1)

    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=50, h=8, txt=f'Date {file_name_and_date[1]}', ln=1)

   
    if '~$' in filepath:
        continue
    else:
        df= pd.read_excel(filepath, sheet_name='Sheet 1', engine='openpyxl')

    # add table header
    columns = list(df.columns)
    columns = [column.replace('_', ' ').title() for column in columns]
    pdf.set_font(family='Times', style='B', size=14)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=str(columns[0]), border=1)
    pdf.cell(w=70, h=8, txt=str(columns[1]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[2]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[3]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns[4]), border=1, ln=1)

    # add table row
    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=8)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)

    pdf.output(f'invoices_pdf\\{file_name_and_date[0]}.pdf')