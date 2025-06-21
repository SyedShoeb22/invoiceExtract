import sys
import os
import fitz  # PyMuPDF
import re
import openpyxl

# ========== CHECK ARGUMENTS ==========
if len(sys.argv) < 2:
    print("❌ Please provide a folder path containing PDF files.")
    print("Usage: python extraction.py <path_to_pdf_folder>")
    sys.exit(1)

PDF_FOLDER = sys.argv[1]
if not os.path.isdir(PDF_FOLDER):
    print(f"❌ '{PDF_FOLDER}' is not a valid folder.")
    sys.exit(1)

OUTPUT_FILE = os.path.join(PDF_FOLDER, "invoice_output.xlsx")

# ========== EXTRACT INVOICE NUMBER ==========
def extract_invoice_no(text):
    lines = text.splitlines()
    invoice_no = "NOT FOUND"
    bad_values = {"invoice", "number", "issued", "summary", "makemytrip", "yourref", "details"}

    for i, line in enumerate(lines):
        # Skip lines like "Invoice Details"
        if line.strip().lower() == "invoice details":
            continue

        # Match patterns like: Invoice No.: 12345
        match = re.search(r"(Invoice\s*(No\.?|Number)?|Number)\s*[:\-]?\s*([A-Z0-9/\-]{6,})", line, re.IGNORECASE)
        if match:
            candidate = match.group(3).strip()
            if candidate.lower() not in bad_values:
                invoice_no = candidate
                break

        # Handle multi-line: line = 'Invoice No.', next = actual number
        if re.search(r"Invoice\s*(No\.?|Number)?", line, re.IGNORECASE) and i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            if re.match(r"^[A-Z0-9/\-]{6,}$", next_line) and next_line.lower() not in bad_values:
                invoice_no = next_line
                break

    return invoice_no


# ========== EXTRACT AMOUNT ==========
def extract_amount(text, page_dicts=None):
    import re

    currency_pattern = r"([\d,]+\.\d{1,2})"
    lines = text.splitlines()
    lines = [line.strip().replace("₹", "").replace("INR", "").replace("Rs.", "") for line in lines if line.strip()]
    keywords = [
        "Grand Total", "Amount Payable", "Total Amount",
        "Net Payable", "Total Due", "Balance Due", "Invoice Total", "Payable Amount"
    ]

    # ---------- 1. Try plain-text keyword match ----------
    for i, line in enumerate(lines):
        for keyword in keywords:
            if keyword.lower() in line.lower():
                values = re.findall(currency_pattern, line)
                if values:
                    try:
                        return str(float(values[-1].replace(",", "")))  # ✅ rightmost value
                    except:
                        continue
                # Try next 2 lines
                for j in range(i + 1, min(i + 3, len(lines))):
                    values = re.findall(currency_pattern, lines[j])
                    if values:
                        try:
                            return str(float(values[-1].replace(",", "")))  # ✅ rightmost
                        except:
                            continue

    # ---------- 2. Try layout-aware span-based matching ----------
    if page_dicts:
        for page in page_dicts:
            for block in page.get("blocks", []):
                for line in block.get("lines", []):
                    spans = line.get("spans", [])
                    if not spans:
                        continue

                    combined = " ".join(span["text"] for span in spans).lower()
                    if "grand total" in combined or "amount payable" in combined:
                        y = spans[0]["bbox"][1]
                        values = []
                        for span in spans:
                            if abs(span["bbox"][1] - y) < 2.0:
                                txt = span["text"]
                                if re.fullmatch(r"[\d,]+\.\d{2}", txt):
                                    try:
                                        values.append((float(txt.replace(",", "")), span["bbox"][0]))  # (value, x-pos)
                                    except:
                                        continue
                        if values:
                            return str(sorted(values, key=lambda x: x[1])[-1][0])  # ✅ rightmost span

    # ---------- 3. Try last 20 lines ----------
    last_lines = "\n".join(lines[-20:])
    try:
        tail_values = [float(v.replace(",", "")) for v in re.findall(currency_pattern, last_lines)]
        if tail_values:
            return str(max(tail_values))
    except:
        pass

    # ---------- 4. Fallback: entire document ----------
    try:
        all_values = [float(v.replace(",", "")) for v in re.findall(currency_pattern, text)]
        if all_values:
            return str(max(all_values))
    except:
        pass

    return "NOT FOUND"



# ========== MAIN PROCESS ==========
invoice_data = []

for filename in os.listdir(PDF_FOLDER):
    if not filename.lower().endswith(".pdf"):
        continue

    path = os.path.join(PDF_FOLDER, filename)
    doc = fitz.open(path)

    full_text = ""
    page_dicts = []
    for page in doc:
        full_text += page.get_text()
        page_dicts.append(page.get_text("dict"))

    invoice_no = extract_invoice_no(full_text)
    amount = extract_amount(full_text, page_dicts)

    invoice_data.append([filename, invoice_no, amount])
    print(f"✅ Processed: {filename} → Invoice: {invoice_no}, Amount: {amount}")

# ========== WRITE TO EXCEL ==========
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Invoice Data"
ws.append(["PDF File", "Invoice Number", "Amount"])

for row in invoice_data:
    ws.append(row)

wb.save(OUTPUT_FILE)
print(f"\n✅ Done. Data saved to '{OUTPUT_FILE}'")
