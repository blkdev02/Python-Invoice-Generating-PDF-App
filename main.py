import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx*")

print(filepaths)

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    
    # extract filename and invoice date 
    filename = Path(filepath).stem
    invoice_number, invoice_date = filename.split("-")

    # set the header
    pdf.set_font(family="Times", style="B", size=24)
    pdf.set_text_color(200, 100, 130)
    pdf.cell(w=50, h=12, txt=f"Invoice No.{invoice_number}", ln=1)

    pdf.set_font(family="Times", style="B", size=20)
    pdf.set_text_color(100, 150, 140)
    pdf.cell(w=50, h=12, txt=f"Date: {invoice_date}", ln=1)

    # read excel file
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    
    # add table header
    table_header = list(df.columns)
    table_header = [item.replace("_", " ") for item in table_header]
    pdf.set_font(family="Times", size=10, style="BU")
    pdf.set_text_color(100,80,80)
    pdf.cell(w=30, h=8, txt=table_header[0].capitalize(), border=1)
    pdf.cell(w=70, h=8, txt=table_header[1].title(),border=1)
    pdf.cell(w=34, h=8, txt=table_header[2].title(),border=1)
    pdf.cell(w=30, h=8, txt=table_header[3].title(), border=1)
    pdf.cell(w=30, h=8, txt=table_header[4].title(), border=1, ln=1)

 
    # add table contents
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=row["product_name"],border=1)
        pdf.cell(w=34, h=8, txt=str(row["amount_purchased"]),border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    

    # write the total sum row
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80,80,80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="",border=1)
    pdf.cell(w=34, h=8, txt="",border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # write the total sum in text 
    pdf.set_text_color(0,0,0)
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}", ln=1)

     # add company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=25, h=8, txt=f"PythonHow")
    pdf.image("images/pythonhow.png", w=11)




    # output the pdf files to disk
    pdf.output(f"invoice_pdf/Invoice-{invoice_number}.pdf")

