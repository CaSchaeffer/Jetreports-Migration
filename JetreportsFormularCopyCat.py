import openpyxl
import re
import os

# Directory where your Excel files are located
excel_directory = "C:\\Users\\c.stegen\\Desktop"

# Regulärer Ausdruck zum Erkennen von Jet Reports-Formeln (=NL, =NP, =NF)
jet_reports_formel_pattern = re.compile(r'(?i)=\s*(?:NL|NP|NF)\s*\([^)]*\)')

# Iterate through files in the directory
for filename in os.listdir(excel_directory):
    if filename.endswith(".xlsx") or filename.endswith(".xlsm"):  # Check for Excel file extensions
        excel_file = os.path.join(excel_directory, filename)

        try:
            # Öffnen Sie die Excel-Datei
            wb = openpyxl.load_workbook(excel_file)

            # Iterate through worksheets in the workbook
            for ws in wb.worksheets:
                # Durchsuchen Sie das Arbeitsblatt nach Formeln und schreiben Sie sie in eine Textdatei
                with open('jet_reports_formeln.txt', 'a') as text_datei:
                    for row_num, row in enumerate(ws.iter_rows(values_only=True), start=1):
                        for col_num, zelle in enumerate(row, start=1):
                            if isinstance(zelle, str) and jet_reports_formel_pattern.search(zelle):
                                text_datei.write(f"Formula found in Sheet: {ws.title}, Row: {row_num}, Column: {col_num}\n")
                                text_datei.write(f"Formula: {zelle}\n\n")

            # Schließen Sie die Excel-Datei
            wb.close()
        except Exception as e:
            print(f"Error processing {excel_file}: {e}")

print("Done searching through all workbooks.")
