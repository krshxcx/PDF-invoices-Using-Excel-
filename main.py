import pandas
import glob
from pathlib import Path
from fpdf import FPDF

filepaths = glob.glob('excel/*.xlsx')

for filepath in filepaths:
    pdf = FPDF(orientation='P',unit='mm',format='A4')
    pdf.add_page()
    filename = Path(filepath).stem
    invoices = filename.split('-')[0]
    date = Path(filepath).stem
    date = date.split('-')[1]
    
    pdf.set_font(family='Times',style='B',size= 20)
    pdf.cell(w=50,h=20,txt=f'invoices.{invoices}',ln=1)
    pdf.cell(w=20,h=10,txt=f'{date}',ln=1)
    
    df = pandas.read_excel(filepath,sheet_name='Sheet 1')
    col = list(df.columns)
    pdf.set_font(family='Times',style='B',size=10)
    pdf.cell(w=30,h=10,txt=col[0],border=1)
    pdf.cell(w=70,h=10,txt=col[1],border=1)
    pdf.cell(w=30,h=10,txt=col[2],border=1)
    pdf.cell(w=30,h=10,txt=col[3],border=1)
    pdf.cell(w=30,h=10,txt=col[4],border=1,ln=1)
    
    for index, row in df.iterrows():
        pdf.set_font(family='Times',style='I',size=15)
        pdf.cell(w=30,h=10,txt=str(row['product_id']),border=1)
        pdf.cell(w=70,h=10,txt=str(row['product_name']),border=1)
        pdf.cell(w=30,h=10,txt=str(row['amount_purchased']),border=1)
        pdf.cell(w=30,h=10,txt=str(row['price_per_unit']),border=1)
        pdf.cell(w=30,h=10,txt=str(row['total_price']),border=1,ln=1)
    
    total_sum = df['total_price'].sum()
    pdf.set_font(family='Times',style='B',size=10)
    pdf.cell(w=30,h=10,txt="",border=1)
    pdf.cell(w=70,h=10,txt="",border=1)
    pdf.cell(w=30,h=10,txt="",border=1)
    pdf.cell(w=30,h=10,txt="",border=1)
    pdf.cell(w=30,h=10,txt=str(total_sum),border=1,ln=1)
    
    pdf.set_font(family='Times',style='B',size= 20)
    pdf.cell(w=50,h=20,txt=f'yahhhh assignment 3',ln=1)
    pdf.image('image.png',w=40)

    pdf.output(f'pdfs/{filename}.pdf')
    