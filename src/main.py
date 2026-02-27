import os
import csv
from openpyxl import Workbook, load_workbook

INPUT_DIR = "input"
OUTPUT_DIR = "output"
OUTPUT_FILE = "merged.xlsx"

os.makedirs(OUTPUT_DIR, exist_ok=True)
output_path = os.path.join(OUTPUT_DIR, OUTPUT_FILE)

def merge_all_files_streaming():
    wb_out = Workbook()
    ws_out = wb_out.active

    files_found = False

    for filename in os.listdir(INPUT_DIR):
        path = os.path.join(INPUT_DIR, filename)

        # CSV
        if filename.lower().endswith(".csv"):
            files_found = True
            with open(path, newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                for row in reader:
                    ws_out.append(row) 

        # Excel
        elif filename.lower().endswith((".xls", ".xlsx")):
            files_found = True
            wb_in = load_workbook(path, read_only=True)
            sheet_in = wb_in.active
            for row in sheet_in.iter_rows(values_only=True):
                ws_out.append(list(row))  # convertit tuple en liste et écrit

        else:
            continue

    if not files_found:
        raise ValueError("Aucun fichier CSV ou Excel trouvé dans le dossier.")

    wb_out.save(output_path)
    print(f"Fusion terminée → {output_path}")
