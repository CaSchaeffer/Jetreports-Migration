import os
import fnmatch

# Das Verzeichnis, das du untersuchen möchtest
verzeichnis = 'P:\PowerBi'

# Funktion zur Zählung der XLSX- und XLSM-Dateien
def zaehle_xlsx_xlsm_dateien(verzeichnis):
    xlsx_anzahl = 0
    xlsm_anzahl = 0

    for ordnername, _, dateien in os.walk(verzeichnis):
        for datei in dateien:
            if fnmatch.fnmatch(datei, '*.xlsx'):
                xlsx_anzahl += 1
            elif fnmatch.fnmatch(datei, '*.xlsm'):
                xlsm_anzahl += 1

    return xlsx_anzahl, xlsm_anzahl

if __name__ == "__main__":
    xlsx_anzahl, xlsm_anzahl = zaehle_xlsx_xlsm_dateien(verzeichnis)
    print(f"Anzahl der XLSX-Dateien: {xlsx_anzahl}")
    print(f"Anzahl der XLSM-Dateien: {xlsm_anzahl}")
