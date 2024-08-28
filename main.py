import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


filepaths = glob.glob("invoices/*xlsx")


for filepath in filepaths:
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    pdf.add_page()

    # Invoice number
    invoice_no = Path(filepath).stem.split('-')[0]
    invoice_date = filepath.split('-')[1].split('.')[:3]
    invoice_date = f"{invoice_date[2]}.{invoice_date[1]}.{invoice_date[0]}"

    # Set the header
    pdf.set_font(family="Times", style="B", size=20)
    pdf.cell(w=0, h=20, txt=f"Invoice: {invoice_no}", align="L", ln=1, border=0)
    pdf.line(10, 25, 200, 25)

    pdf.set_font(family="Times", style="IB", size=12)
    pdf.cell(w=0, h=12, txt=f"Date: {invoice_date}", align="L", ln=1, border=0)

    # Set table - head
    column_names = df.columns.tolist()
    column_names = [item.replace('_', ' ').title() for item in column_names]

    for name in column_names:
        pdf.set_font(family="Times", style="B", size=11)
        pdf.cell(w=35, h=12, txt=name, align="L", border=1)

    pdf.cell(w=0, h=12, align="L", ln=1)

    # Set table - body
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=8)
        pdf.cell(w=35, h=8, txt=str(row['product_id']), align="L", border=1)
        pdf.cell(w=35, h=8, txt=row['product_name'], align="L", border=1)
        pdf.cell(w=35, h=8, txt=str(row['amount_purchased']), align="L", border=1)
        pdf.cell(w=35, h=8, txt=str(row['price_per_unit']), align="L", border=1)
        pdf.cell(w=35, h=8, txt=str(row['total_price']), align="L", ln=1, border=1)

    # Count total price
    amount_columns = ['total_price']
    amount_counts = df[amount_columns].sum().sum()

    pdf.cell(w=140, h=8, txt="Total", align="L", border=1)
    pdf.cell(w=35, h=8, txt=str(amount_counts), align="L", ln=1, border=1)

    # Set the footer
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=0, h=20, txt=f"The total due amount is: {amount_counts} $", align="L", ln=1, border=0)
    pdf.line(10, 25, 200, 25)

    # Generate PDF
    pdf.output(f"PDFs/{invoice_no}.pdf")







