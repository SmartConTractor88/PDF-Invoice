import pandas as pd # for reading .xslx
#import openpyxl
import glob
from fpdf import FPDF
from pathlib import Path # to extract filepaths from the Excel files

filepaths = glob.glob("Invoices/*.xlsx")
print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1") # read Sheet 1 from each excel file
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem # extract the file name
    invoice_nr = filename.split("-")[0]
    invoice_date = filename.split("-")[1]
    # print(f"Invoice nr.{invoice_nr[0]}")

    pdf.set_font(family="Times", style="B", size=20)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=8, txt=f"Invoice nr.{invoice_nr}", align="C", ln=1, border=0)

    pdf.set_font(family="Times", style="B", size=16)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(w=0, h=18, txt=invoice_date, align="L", ln=1, border=1)

    pdf.output(f"PDFs/Invoice{invoice_nr}.pdf")

    #print(df)
