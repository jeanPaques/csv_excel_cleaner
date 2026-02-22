# 📊 CSV/XLSX Cleaner & Merger

**Automate the cleaning and merging of CSV and Excel files. Save time and get consistent, ready-to-use spreadsheets.**

---

## 🛠️ Features
- Merge all CSV/XLS/XLSX files in a folder automatically.
- Standardize data:
  - Names → uppercase
  - Prices → formatted as float
  - Phone numbers → uniform international format
- Remove invalid rows (missing name or email)
- Remove duplicates
- Generate summary report:
  - Total number of clients
  - Total revenue (if price column present)

---

## 🚀 How to Use
1. Clone the repository:
   ```
   git clone https://github.com/jeanPaques/csv_excel_cleaner.git
   cd csv_excel_cleaner
   ```
2. Create and activate virtual environment:
   ```
   python -m venv venv
   # Windows
   .\venv\Scripts\activate
   # macOS/Linux
   source venv/bin/activate
   ```
3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
4. Put your CSV/XLS/XLSX files in the `/input` folder.
5. Run the script:
   ```
   python src/main.py
   ```
6. Your cleaned, merged master file will appear in the `/output` folder.

---

## 💡 Notes
- Works with CSV and Excel files only.
- `/output` folder is ignored in GitHub (`.gitignore`) to avoid uploading generated files.

---

## 📂 Project Structure
```
project-folder/
├─ input/          # place your CSV/XLSX files here
├─ output/         # cleaned/merged files will be saved here
├─ src/main.py    # main script
├─ requirements.txt
├─ .gitignore
├─ LICENSE
└─ README.md
```

---

## 🔗 Contact / Portfolio 
[jeanPaques](https://github.com/jeanPaques)