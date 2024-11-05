import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

invoice_list = glob.glob("invoices/*.xlsx")
print(invoice_list)

for invoice in invoice_list:
    # Set up PDF file
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Set up variables for invoice name and date
    path = Path(invoice).stem
    invoice_number, invoice_date = path.split("-")

    # Set the title using the invoice file name
    pdf.set_font("Times", "B", 16)
    pdf.cell(50, 8, f"Invoice Number: {invoice_number}", ln=1)

    # Sets the date using the invoice file name
    pdf.cell(50, 8, f"Date:  {invoice_date}", ln=1)

    # Gather data frame
    df = pd.read_excel(invoice, "Sheet 1")

    # Set up header for table
    headings = list(df.columns)
    headings = [item.replace("_", " ").title() for item in headings]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(30, 8, headings[0], border=1)
    pdf.cell(70, 8, headings[1], border=1)
    pdf.cell(30, 8, headings[2], border=1)
    pdf.cell(30, 8, headings[3], border=1)
    pdf.cell(30, 8, headings[4], border=1, ln=1)

    # Set up the main content of the table
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.cell(30, 8, str(row["product_id"]), border=1)
        pdf.cell(70, 8, str(row["product_name"]), border=1)
        pdf.cell(30, 8, str(row["amount_purchased"]), border=1)
        pdf.cell(30, 8, str(row["price_per_unit"]), border=1)
        pdf.cell(30, 8, str(row["total_price"]), border=1, ln=1)

    # Set up the total
    pdf.set_font(family="Times", size=10)
    pdf.cell(30, 8, "", border=1)
    pdf.cell(70, 8, "", border=1)
    pdf.cell(30, 8, "", border=1)
    pdf.cell(30, 8, "", border=1)
    pdf.cell(30, 8, str(df["total_price"].sum()), border=1, ln=1)

    # Add total sum sentence
    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(30, 8, f"The total price is £{df['total_price'].sum()}", ln=1)

    # Add company name
    pdf.cell(30, 8, f"Company Name Ltd", ln=1)

    # Output the PDFs
    pdf.output(name=f"PDFs/{path}.pdf")
