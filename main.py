import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    # Read the Excel file
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Create PDF
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Extract filename and invoice number
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    # Set title
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=10, txt=f"Invoice nr. {invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=10, txt=f"Date: {date}", ln=1)
    # Add some space
    pdf.ln(10)

    # Add a header
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=60, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Add sum of price
    total_sum =df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Add total sum sentence
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price for your purchased is {total_sum} Euros", ln=1)

    # Add name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=30, h=8, txt="Bass player.com")  # Text
    pdf.cell(w=10, h=8)  # Empty cell for space (adjust w for more/less space)
    pdf.image("bassplayer.png", w=14)

    # Save PDF
    pdf.output(f"PDFs/{filename}.pdf")


