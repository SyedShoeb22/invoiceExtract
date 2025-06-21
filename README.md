# 🧾 Invoice Extractor

A Python script to extract **Invoice Numbers** and **Grand Total Amounts** from airline invoice PDFs and save the results to an Excel file.

---

## ✅ Features

- 📁 Reads all PDFs from a specified folder  
- 🔍 Extracts **Invoice Numbers** using regex patterns  
- 💰 Extracts **Grand Total / Amount Payable** (even from tabular formats)  
- 📊 Outputs the results to `invoice_output.xlsx`

---

## 📦 Requirements

Install dependencies using:

```bash
pip install -r requirements.txt

python extraction.py "D:\path\to\your\pdfs"

pip install pyinstaller

pyinstaller --onefile --noconsole extraction.py

extraction.exe "D:\path\to\your\pdfs"

