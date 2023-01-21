import pandas as pd
import glob
from fpdf import FPDF
import os

filepaths = glob.glob('invoices\\*.xlsx')

for filepath in filepaths:
    file_name_and_date = os.path.basename(filepath)[:-5].split('-')
    print()
    if '~$' in filepath:
        continue
    else:
        df= pd.read_excel(filepath, sheet_name='Sheet 1', engine='openpyxl')
    
    pdf = FPDF(orientation='P', format='A4', unit='mm')
    pdf.add_page()

    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=50, h=8, txt=f'Invoice mr. {file_name_and_date[0]}', ln=1)

    pdf.set_font(family='Times', style='B', size=16)
    pdf.cell(w=50, h=8, txt=f'Date {file_name_and_date[1]}')

    pdf.output(f'invoices_pdf\\{file_name_and_date[0]}.pdf')