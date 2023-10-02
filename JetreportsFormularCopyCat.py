import openpyxl
import re
import os

# Directory where your Excel files are located
excel_directory = "C:\\Users\\c.stegen\\Desktop"

# Regular expression to recognize Jet Reports formulas (=NL, =NP, =NF)
jet_reports_formula_pattern = re.compile(r'(?i)=\s*(?:NL|NP|NF)\s*\([^)]*\)')

# AsciiDoc content for the formulas
asciidoc_content = ""

# Set to store unique formulas
unique_formulas = set()

# Iterate through files in the directory
for filename in os.listdir(excel_directory):
    if filename.endswith(".xlsx") or filename.endswith(".xlsm"):  # Check for Excel file extensions
        excel_file = os.path.join(excel_directory, filename)

        try:
            # Open the Excel file
            wb = openpyxl.load_workbook(excel_file)

            # Get the category or purpose of the Excel file (you need to fill this in)
            excel_category = "Fill in the category or purpose here"

            # Add the Excel file name and category to AsciiDoc content
            asciidoc_content += f"= Excel File: {filename}\n"
            asciidoc_content += f"== Category/Purpose: {excel_category}\n\n"

            # Iterate through worksheets in the workbook
            for ws in wb.worksheets:
                # Search the worksheet for formulas and add them to the AsciiDoc content
                asciidoc_content += f"=== Sheet: {ws.title}\n\n"
                for row_num, row in enumerate(ws.iter_rows(values_only=True), start=1):
                    for col_num, cell in enumerate(row, start=1):
                        if isinstance(cell, str) and jet_reports_formula_pattern.search(cell):
                            formula = f"Row: {row_num}, Column: {col_num}\nFormula: {cell}\n\n"
                            if formula not in unique_formulas:
                                asciidoc_content += formula
                                unique_formulas.add(formula)

            # Close the Excel file
            wb.close()
        except Exception as e:
            print(f"Error processing {excel_file}: {e}")

# Save the AsciiDoc content to a file
with open('formulas.adoc', 'w') as asciidoc_file:
    asciidoc_file.write(asciidoc_content)

print("Done searching through all workbooks. Formulas saved in 'formulas.adoc'.")
