import os
import shutil
import datetime

# Quellverzeichnis, das du untersuchen möchtest
quelle_verzeichnis = 'P:\JETREPORTS NAV'

# Zielverzeichnis, in das die Dateien kopiert werden sollen
ziel_verzeichnis = 'P:\PowerBi'

# Ziel-Jahr (2022)
ziel_jahr = 2022
global counter
counter = 0

def kopiere_dateien(src, dst):
    try:
        shutil.copy2(src, dst)
        print(f'Kopiere {src} nach {dst}')
    except Exception as e:
        print(f'Fehler beim Kopieren von {src} nach {dst}: {str(e)}')
        counter + 1

def untersuche_verzeichnis(verzeichnis):
    for ordnername, _, dateien in os.walk(verzeichnis):
        for datei in dateien:
            dateipfad = os.path.join(ordnername, datei)
            # Überprüfe, ob die Datei im Jahr 2023 geändert wurde
            dateiinfo = os.stat(dateipfad)
            modifikationsdatum = datetime.datetime.fromtimestamp(dateiinfo.st_mtime)
            if modifikationsdatum.year == ziel_jahr:
                ziel_dateipfad = os.path.join(ziel_verzeichnis, os.path.relpath(dateipfad, quelle_verzeichnis))
                kopiere_dateien(dateipfad, ziel_dateipfad)

if __name__ == "__main__":
    if not os.path.exists(ziel_verzeichnis):
        os.makedirs(ziel_verzeichnis)

    untersuche_verzeichnis(quelle_verzeichnis)
print("Fehleranzahl:",counter)
