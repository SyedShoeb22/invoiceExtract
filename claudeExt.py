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

    # ✅ 1. Text-based approach for tabular "Grand Total" rows (more reliable)
    lines = text.splitlines()
    for line in lines:
        line_clean = line.strip().replace("₹", "").replace("INR", "").replace("Rs.", "")
        
        # Look for "Grand Total" lines with multiple amounts
        if "grand total" in line_clean.lower():
            # Find all decimal amounts in the line
            amounts = re.findall(r"[\d,]+\.\d{1,2}", line_clean)
            if len(amounts) >= 2:  # Multiple amounts = likely tabular
                try:
                    # Return the last (rightmost) amount
                    return str(float(amounts[-1].replace(",", "")))
                except:
                    pass

    # ✅ 2. Enhanced structured fallback for tabular data
    if page_dicts:
        for page in page_dicts:
            for block in page.get("blocks", []):
                for line in block.get("lines", []):
                    spans = line.get("spans", [])
                    if not spans:
                        continue

                    full_line_text = " ".join(span["text"] for span in spans).lower()

                    # Look for Grand Total, Total, Final Amount, etc.
                    total_keywords = ["grand total", "total amount", "final amount", "amount payable", "invoice total"]
                    
                    if any(keyword in full_line_text for keyword in total_keywords):
                        y_line = spans[0]["bbox"][1]
                        values_on_line = []

                        # Collect all numeric values on the same line
                        for span in spans:
                            if abs(span["bbox"][1] - y_line) < 3.0:  # Increased tolerance
                                txt = span["text"].strip()
                                # Match currency patterns: 12001.00, 1,234.56, etc.
                                if re.match(r"[\d,]+\.\d{1,2}$", txt):
                                    try:
                                        numeric_value = float(txt.replace(",", ""))
                                        x_position = span["bbox"][0]  # x-coordinate
                                        values_on_line.append((numeric_value, x_position))
                                    except:
                                        continue

                        if values_on_line:
                            # Return the rightmost (highest x-coordinate) value
                            rightmost_value = sorted(values_on_line, key=lambda x: x[1])[-1][0]
                            return str(rightmost_value)

    # ✅ 2. Original approach + enhanced patterns
    currency_pattern = r"([\d,]+\.\d{1,2})"
    lines = text.splitlines()
    lines = [line.strip().replace("₹", "").replace("INR", "").replace("Rs.", "") for line in lines if line.strip()]
    
    # Expanded keywords for better matching
    keywords = [
        "Grand Total", "Amount Payable", "Total Amount", "Balance Due", 
        "Invoice Total", "Total Invoice Value", "Payable", "Final Amount",
        "Net Amount", "Total Due", "Amount Due", "Total Payable"
    ]

    # Enhanced keyword matching
    for i, line in enumerate(lines):
        if any(k.lower() in line.lower() for k in keywords):
            # Find all amounts in current line
            values = re.findall(currency_pattern, line)
            if values:
                try:
                    # For tabular data, return the last value (rightmost)
                    return str(float(values[-1].replace(",", "")))
                except:
                    pass
            
            # Check next line if current line has no amounts
            if i + 1 < len(lines):
                next_values = re.findall(currency_pattern, lines[i + 1])
                if next_values:
                    try:
                        return str(float(next_values[-1].replace(",", "")))
                    except:
                        pass

    # ✅ 3. Look for lines with multiple amounts (tabular format)
    for line in lines:
        values = re.findall(currency_pattern, line)
        # If line has multiple amounts (3+) and contains "total"
        if len(values) >= 3 and "total" in line.lower():
            try:
                # Return the last amount (usually the final total)
                return str(float(values[-1].replace(",", "")))
            except:
                pass

    # ✅ 4. Broader pattern matching for various formats
    # Look for common amount indicators
    amount_indicators = [
        r"total[:\s]+([,\d]+\.\d{1,2})",
        r"amount[:\s]+([,\d]+\.\d{1,2})",
        r"payable[:\s]+([,\d]+\.\d{1,2})",
        r"due[:\s]+([,\d]+\.\d{1,2})"
    ]
    
    text_clean = text.replace("₹", "").replace("INR", "").replace("Rs.", "")
    for pattern in amount_indicators:
        matches = re.findall(pattern, text_clean, re.IGNORECASE)
        if matches:
            try:
                return str(float(matches[-1].replace(",", "")))
            except:
                pass

    # ✅ 5. Last 20 lines fallback (return largest reasonable amount)
    last_20 = "\n".join(lines[-20:])
    values = re.findall(currency_pattern, last_20)
    try:
        floats = [float(v.replace(",", "")) for v in values if float(v.replace(",", "")) > 10]  # Minimum 10 currency units
        if floats:
            return str(max(floats))  # Return largest amount
    except:
        pass

    return "NOT FOUND"


# ========== MAIN PROCESS ==========
invoice_data = []

for filename in os.listdir(PDF_FOLDER):
    if not filename.lower().endswith(".pdf"):
        continue

    path = os.path.join(PDF_FOLDER, filename)
    try:
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
        
        doc.close()
        
    except Exception as e:
        print(f"❌ Error processing {filename}: {str(e)}")
        invoice_data.append([filename, "ERROR", "ERROR"])

# ========== WRITE TO EXCEL ==========
try:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Invoice Data"
    ws.append(["PDF File", "Invoice Number", "Amount"])

    for row in invoice_data:
        ws.append(row)

    wb.save(OUTPUT_FILE)
    print(f"\n✅ Done. Data saved to '{OUTPUT_FILE}'")
    
except Exception as e:
    print(f"❌ Error saving Excel file: {str(e)}")