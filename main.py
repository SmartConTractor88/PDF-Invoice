import pandas as pd # for reading .xslx
#import openpyxl
import glob
from fpdf import FPDF
from pathlib import Path # to extract filepaths from the Excel files

filepaths = glob.glob("Invoices/*.xlsx")
print(filepaths)

for filepath in filepaths:
    
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
    pdf.cell(w=0, h=18, txt=invoice_date, align="L", ln=1, border=0)

    

    df = pd.read_excel(filepath, sheet_name="Sheet 1") # read Sheet 1 from each excel file

    # add a header
    columns = list(df.columns) # create a list of column headings
    columns = [item.replace("_", " ").title() for item in columns] # replace underscore with space, capitalize all first letters

    # Add headings for each column

    pdf.set_font(family="Times", size=12, style="B")
    pdf.set_text_color(0, 0, 0)

    pdf.cell(w=30, h=12, txt=columns[0], align="C", border=1)
    pdf.cell(w=50, h=12, txt=columns[1], align="C", border=1)
    pdf.cell(w=40, h=12, txt=columns[2], align="C", border=1)
    pdf.cell(w=30, h=12, txt=columns[3], align="C", border=1)
    pdf.cell(w=40, h=12, txt=columns[4], align="C", border=1, ln=1)

    for index, row in df.iterrows():

        pdf.set_font(family="Times", size=12)
        pdf.set_text_color(0, 0, 0)

        pdf.cell(w=30, h=12, txt=str(row["product_id"]), align="L", border=1)
        pdf.cell(w=50, h=12, txt=str(row["product_name"]), align="L", border=1)
        pdf.cell(w=40, h=12, txt=str(row["amount_purchased"]), align="L", border=1)
        pdf.cell(w=30, h=12, txt=str(row["price_per_unit"]), align="L", border=1)
        pdf.cell(w=40, h=12, txt=str(row["total_price"]), align="L", border=1, ln=1)

    # Add a rpw for totals

    pdf.set_font(family="Times", size=12)
    pdf.set_text_color(0, 0, 0)

    total_sum = df["total_price"].sum()

    pdf.cell(w=30, h=12, txt="", align="L", border=1)
    pdf.cell(w=50, h=12, txt="", align="L", border=1)
    pdf.cell(w=40, h=12, txt="", align="L", border=1)
    pdf.cell(w=30, h=12, txt="", align="L", border=1)
    pdf.cell(w=40, h=12, txt=str(total_sum), align="L", border=1, ln=1)

    pdf.set_font(family="Times", style="B", size=16)
    pdf.set_text_color(0, 0, 0)

    # Add a statement of totals

    total_statement = f"Total price is {total_sum} Eur."
    pdf.cell(w=0, h=12, txt=total_statement, align="L", border=0, ln=1)

    # Add company name and logo

    pdf.set_font(family="Times", style="B", size=24)
    pdf.set_text_color(0, 0, 0)

    pdf.cell(w=55, h=28, txt="Winner Company", align="L", border=0)
    pdf.image("logo_01.png", w=50)
    
    pdf.output(f"PDFs/Invoice{invoice_nr}.pdf")
    #print(df)
