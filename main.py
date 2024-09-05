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
    invoice_nr = filename.split("-")[0]

    # Set title
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=10, txt=f"Invoice nr. {invoice_nr}", ln=True)

    # Add some space
    pdf.ln(10)

    # Set table header
    pdf.set_font(family="Times", size=12, style="B")
    col_widths = [30, 60, 40, 30, 25]  # Adjust column widths based on your data

    # Assuming df.columns are the table headers
    for header in df.columns:
        pdf.cell(w=col_widths[df.columns.get_loc(header)], h=10, txt=header, border=1)

    pdf.ln()

    # Set table body
    pdf.set_font(family="Times", size=10)

    # Iterate through each row and add data to the PDF
    for row in df.itertuples(index=False):
        for idx, cell in enumerate(row):
            pdf.cell(w=col_widths[idx], h=8, txt=str(cell), border=1)
        pdf.ln()

    # Save PDF
    pdf.output(f"PDFs/{filename}.pdf")
