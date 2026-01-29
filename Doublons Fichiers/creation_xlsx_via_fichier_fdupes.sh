#!/bin/bash
# le script prends dans le meme repertoire le fichier doublons.txt 
# généré avec la commande suivante
# fdupes -R /share/VCA/DGSRV/ -G 20000000 -S -l -t > doublons.txt
# pip install openpyxl
INPUT="doublons.txt"
OUTPUT="doublons.xlsx"
UNIT="Mo"   # octets | Ko | Mo

python3 <<EOF
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side
from openpyxl.utils import get_column_letter
from urllib.parse import quote
import os


# --- Variables Python depuis Bash ---
INPUT = "$INPUT"
OUTPUT = "$OUTPUT"
UNIT = "$UNIT"

# Lire et stocker tous les doublons
groups = []

def convert_size(size):
    if UNIT == "octets":
        return size
    elif UNIT == "Ko":
        return round(size / 1024, 2)
    elif UNIT == "Mo":
        return round(size / 1024 / 1024, 2)

current_size = None
paths = []

with open(INPUT, encoding="utf-8") as f:
    for line in f:
        line = line.strip()
        if line.endswith("bytes each:"):
            if current_size is not None:
                groups.append((current_size, list(paths)))
            current_size = int(line.split()[0])
            paths = []
        elif line.startswith("/"):
            path = line.replace("/share", r"\\\\data", 1)
            path = path.replace("/", "\\\\")
            path = os.path.normpath(path)
            paths.append(path)

# Ajouter le dernier groupe
if current_size is not None:
    groups.append((current_size, list(paths)))

# Trier par taille décroissante
groups.sort(key=lambda x: x[0], reverse=True)

# --- Création Excel ---
wb = Workbook()
ws = wb.active
ws.title = "Doublons"

# En-têtes
#ws.append([f"Taille ({UNIT})"] + [f"Fichier {i}" for i in range(1, 26)])
ws.append(
    [f"Taille ({UNIT})"] +
    [item
     for i in range(1, 26)
     for item in (f"DOSSIER {i}", f"FICHIER {i}")]
)

for cell in ws[1]:
    cell.font = Font(bold=True, underline="single")
ws.freeze_panes = "A2"
for size, paths in groups:
    row = [convert_size(size)]
    #for p in paths:
    #    row.append(p)  # Affiche le chemin complet du fichier
    ws.append(row)

    # Ajouter hyperliens vers le dossier contenant le fichier
    col = 2
    for p in paths:
        p = p.replace("\\\\", "/")
        p_norm = os.path.normpath(p)
        folder = quote(os.path.dirname(p_norm), safe="/:")
        file = os.path.basename(p_norm)
        parts = [x for x in p_norm.split(os.sep) if x]
        # Convertir pour Excel
        folder_link = folder
        cell = ws.cell(row=ws.max_row, column=col)
        cell.value = parts[3] 
        cell.font = Font(color="FF0000")
        cell = ws.cell(row=ws.max_row, column=col+1)
        cell.value = file  # Affiche le chemin complet du fichier
        cell.hyperlink = "file:" + folder_link
        cell.font = Font(color="0000FF", underline="single")
        col += 2

thin = Side(style="thin")

full_border = Border(
    left=thin,
    right=thin,
    top=thin,
    bottom=thin
)

for row in ws.iter_rows(
        min_row=1,
        max_row=ws.max_row,
        min_col=1,
        max_col=ws.max_column):
    for cell in row:
        cell.border = full_border

for col in range(1, ws.max_column + 1):
    max_length = 0
    col_letter = get_column_letter(col)

    for cell in ws[col_letter]:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))

    # marge de confort
    ws.column_dimensions[col_letter].width = max_length + 2

wb.save(OUTPUT)
print("Fichier généré :", OUTPUT)
EOF

