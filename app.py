import os
import pdfplumber
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from openpyxl import load_workbook

class QCFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Quotation to QCF Converter - Production Build")
        self.root.geometry("550x450")

        self.template_path = ""
        self.pdf_list = []

        # --- UI Elements ---
        tk.Label(root, text="Step 1: Select QCF Template", font=('Arial', 10, 'bold')).pack(pady=5)
        self.btn_template = tk.Button(root, text="Browse Excel Template", command=self.load_template)
        self.btn_template.pack()

        tk.Label(root, text="Step 2: Drag & Drop Quotations Below", font=('Arial', 10, 'bold')).pack(pady=15)
        self.drop_target = tk.Listbox(root, width=70, height=10)
        self.drop_target.pack(padx=20, pady=5)

        self.drop_target.drop_target_register(DND_FILES)
        self.drop_target.dnd_bind('<<Drop>>', self.handle_drop)

        self.btn_run = tk.Button(root, text="Process & Generate QCF",
                                 command=self.start_processing, bg="green", fg="white", font=('Arial', 12, 'bold'))
        self.btn_run.pack(pady=20)

    def load_template(self):
        self.template_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.template_path:
            self.btn_template.config(text=os.path.basename(self.template_path), fg="blue")

    def handle_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        for f in files:
            if f.endswith(".pdf"):
                self.pdf_list.append(f)
                self.drop_target.insert(tk.END, os.path.basename(f))

    def start_processing(self):
        if not self.template_path or not self.pdf_list:
            messagebox.showwarning("Missing Files", "Please provide the template and at least one PDF.")
            return

        output_path = filedialog.asksaveasfilename(
            title="Save Output As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="Consolidated_QCF_Result.xlsx",
        )
        if not output_path:
            return

        success_count = 0
        print(f"Opening template: {self.template_path}")
        wb = load_workbook(self.template_path)
        ws = wb.active

        for pdf_path in self.pdf_list:
            try:
                self.process_single_pdf(pdf_path, ws)
                success_count += 1
            except Exception as e:
                print(f"Error processing {os.path.basename(pdf_path)}: {e}")

        if success_count > 0:
            wb.save(output_path)
            messagebox.showinfo("Success", f"Data inserted!\nSaved as: {os.path.basename(output_path)}")
        else:
            messagebox.showerror("Error", "Failed to process the files.")

    def extract_pdf_data(self, pdf_path):
        """ Extracts data dynamically handling ANY vendor company """
        data = {
            "vendor": "Unknown Company",
            "ref_no": "N/A",
            "payment": "N/A",
            "delivery": "N/A",
            "items": [] # Format will be: (Description, Qty, UOM, Rate)
        }

        with pdfplumber.open(pdf_path) as pdf:
            # Combine all pages into one text block
            pages_text = []
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                if page_text:
                    pages_text.append(page_text)
            text = "\n".join(pages_text)
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            
            if lines:
                # DYNAMIC VENDOR NAME: Prefer a company-like line near the top
                company_hints = [
                    r"\bSDN\s*BHD\b",
                    r"\bS/B\b",
                    r"\bBHD\b",
                    r"\bCONSULTING\b",
                    r"\bSOLUTIONS?\b",
                    r"\bTECHNOLOGY\b",
                ]
                for line in lines[:12]:
                    upper = line.upper()
                    if "HTTP" in upper or "WWW." in upper:
                        continue
                    if "EMAIL" in upper or "TEL" in upper or "FAX" in upper:
                        continue
                    if re.search(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", line):
                        continue
                    if re.search(r"\b\d{1,2}:\d{2}\s*(AM|PM)?\b", upper):
                        continue
                    if any(re.search(h, upper) for h in company_hints):
                        data['vendor'] = line[:60]
                        break
                if data['vendor'] == "Unknown Company":
                    # Fallback to first clean line
                    for line in lines[:12]:
                        upper = line.upper()
                        if "HTTP" in upper or "WWW." in upper:
                            continue
                        if "EMAIL" in upper or "TEL" in upper or "FAX" in upper:
                            continue
                        data['vendor'] = line[:60]
                        break

            # Extract Terms
            # First pass: line-based label search to capture full reference
            def clean_ref(raw):
                val = raw.strip()
                # Cut off trailing labels if line has multiple fields
                val = re.split(r"\b(Date|Quotation Date|Valid Through|Payment Term|Attention|Attn)\b", val, flags=re.IGNORECASE)[0].strip()
                # Trim trailing punctuation
                val = val.rstrip(" .,:;")
                return val

            for line in lines:
                m = re.search(r"\bReference\s*Number\s*[:\-]\s*(.+)$", line, re.IGNORECASE)
                if m:
                    data['ref_no'] = clean_ref(m.group(1))
                    break
                m = re.search(r"\bRef(?:erence)?\s*(?:No\.?|#|Number)?\s*[:\-]\s*(.+)$", line, re.IGNORECASE)
                if m:
                    data['ref_no'] = clean_ref(m.group(1))
                    break
                m = re.search(r"\b(?:Quotation|Quote)\s*(?:No\.?|#|Ref\.?|Reference)?\s*[:\-]\s*(.+)$", line, re.IGNORECASE)
                if m:
                    data['ref_no'] = clean_ref(m.group(1))
                    break

            # Fallback: regex scan for labeled refs anywhere in text
            if data['ref_no'] == "N/A":
                ref_patterns = [
                    r"(?:Quotation|Quote)\s*(?:No\.?|#|Ref\.?|Reference)?\s*[:\-]?\s*([A-Z0-9][A-Z0-9\-\/&_.]+)",
                    r"\b(?:Ref\.?|Reference)\s*(?:No\.?|#|Number)?\s*[:\-]?\s*([A-Z0-9][A-Z0-9\-\/&_.]+)",
                    r"\bReference\s*Number\s*[:\-]\s*([A-Z0-9][A-Z0-9\-\/&_.]+)",
                ]
                ref_candidates = []
                for pat in ref_patterns:
                    ref_candidates.extend(re.findall(pat, text, re.IGNORECASE))
                if ref_candidates:
                    data['ref_no'] = max(ref_candidates, key=len)

            pay_match = re.search(r"(?:Payment|Terms?)[:\s]*(.*?(?:Days|COD|Cash|Advance|Month).*?)(?=\n|$)", text, re.IGNORECASE)
            if pay_match: data['payment'] = pay_match.group(1).strip()

            del_match = re.search(r"(?:Delivery|Lead\s*Time)[:\s]*(.*?(?:Weeks|Days|Months|ARO).*?)(?=\n|$)", text, re.IGNORECASE)
            if del_match: data['delivery'] = del_match.group(1).strip()

            # Extract line items (best-effort patterns)
            def parse_item_line(line):
                # Only accept lines that look like item starts (e.g., "1 ..." or "1. ...")
                if not re.match(r"^\s*\d+(?:\.\d+)?\s+", line):
                    return None

                # Pattern: description + qty+unit + unit price + amount
                m = re.search(
                    r"^(?:\d+(?:\.\d+)?\s+)?(.+?)\s+(\d+(?:\.\d+)?)([A-Za-z]+)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$",
                    line,
                )
                if m:
                    desc = m.group(1).strip().rstrip(".")
                    qty = float(m.group(2))
                    unit = m.group(3)
                    rate = float(m.group(4).replace(",", ""))
                    return desc, qty, unit, rate

                # Pattern: description + L/S + unit price + amount
                m = re.search(
                    r"^(?:\d+(?:\.\d+)?\s+)?(.+?)\s+L\/S\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$",
                    line,
                    re.IGNORECASE,
                )
                if m:
                    desc = m.group(1).strip().rstrip(".")
                    rate = float(m.group(2).replace(",", ""))
                    return desc, 1.0, "L/S", rate

                # Pattern: description + qty+unit + price (no amount)
                m = re.search(
                    r"^(?:\d+(?:\.\d+)?\s+)?(.+?)\s+(\d+(?:\.\d+)?)([A-Za-z]+)\s+([\d,]+\.\d{2})\s*$",
                    line,
                )
                if m:
                    desc = m.group(1).strip().rstrip(".")
                    qty = float(m.group(2))
                    unit = m.group(3)
                    rate = float(m.group(4).replace(",", ""))
                    return desc, qty, unit, rate

                return None

            in_items = False
            last_idx = None
            skip_next = False
            for i, line in enumerate(lines):
                if skip_next:
                    skip_next = False
                    continue

                upper = line.upper()
                if "NO" in upper and "DESCRIPTION" in upper and "QTY" in upper:
                    in_items = True
                    continue
                if "TERMS" in upper or "CONDITION" in upper or "TOTAL" in upper or "SST" in upper:
                    in_items = False
                    last_idx = None
                    continue

                if not in_items:
                    continue

                match = parse_item_line(line)
                if not match and i < len(lines) - 1:
                    # Some PDFs wrap the price/qty to the next line
                    combined = f"{line} {lines[i + 1]}"
                    match = parse_item_line(combined)
                    if match:
                        skip_next = True

                if match:
                    desc, qty, unit, rate = match
                    data['items'].append((desc, int(qty) if qty.is_integer() else qty, unit, rate))
                    last_idx = len(data['items']) - 1
                elif last_idx is not None:
                    # Append continuation lines to the last description
                    if not re.search(r"^\d+(?:\.\d+)?\s+", line):
                        # Stop appending when footer/terms content begins
                        if re.search(r"\b(TOTAL|SST|TERM|PAYMENT|DELIVERY|VALIDITY|WARRANTY|BANK|ACCOUNT|REFERENCE|QUOTE|QUOTATION)\b", upper):
                            in_items = False
                            last_idx = None
                            continue
                        prev = data['items'][last_idx]
                        data['items'][last_idx] = (prev[0] + " " + line.strip(), prev[1], prev[2], prev[3])

            # Table-based extraction (more reliable for structured quotations)
            if not data['items']:
                table_settings = {
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "intersection_tolerance": 5,
                    "snap_tolerance": 3,
                    "join_tolerance": 3,
                    "edge_min_length": 3,
                }

                def normalize(s):
                    return re.sub(r"\s+", " ", (s or "")).strip().lower()

                def parse_number(val):
                    if val is None:
                        return None
                    txt = normalize(val).replace("rm", "").replace(",", "").strip()
                    m = re.search(r"(\d+(?:\.\d{2})?)", txt)
                    return float(m.group(1)) if m else None

                last_idx = None
                for page in pdf.pages:
                    try:
                        tables = page.extract_tables(table_settings) or []
                    except Exception:
                        tables = []
                    for table in tables:
                        if not table or len(table) < 2:
                            continue
                        header = [normalize(c) for c in table[0]]
                        desc_idx = None
                        qty_idx = None
                        unit_price_idx = None

                        for i, h in enumerate(header):
                            if "description" in h or "product details" in h:
                                desc_idx = i
                            if h in ("qty", "quantity"):
                                qty_idx = i
                            if "unit price" in h or "price ex.sst" in h or "price ex" in h:
                                unit_price_idx = i
                            if "unit price (rm)" in h or h == "unit price":
                                unit_price_idx = i

                        if desc_idx is None:
                            continue

                        for row in table[1:]:
                            if not row or all((c or "").strip() == "" for c in row):
                                continue
                            # Some tables include a "No" column; skip if row is header-like
                            row_norm = [normalize(c) for c in row]
                            if any("description" in c for c in row_norm):
                                continue

                            desc = row[desc_idx] if desc_idx < len(row) else ""
                            qty_val = row[qty_idx] if (qty_idx is not None and qty_idx < len(row)) else None
                            price_val = row[unit_price_idx] if (unit_price_idx is not None and unit_price_idx < len(row)) else None

                            qty = parse_number(qty_val) or None
                            rate = parse_number(price_val) or None

                            if desc and (qty is not None or rate is not None):
                                data['items'].append((desc.strip().rstrip("."), int(qty or 1), "Unit", float(rate or 0)))
                                last_idx = len(data['items']) - 1
                            elif desc and last_idx is not None:
                                # Append continuation line to previous description
                                prev = data['items'][last_idx]
                                data['items'][last_idx] = (prev[0] + " " + desc.strip(), prev[1], prev[2], prev[3])

            # Dynamic Price Extraction (Fallback for unknown table structures)
            if not data['items']:
                price_matches = re.findall(r"(?:Grand\s*)?Total.*?([\d,]+\.\d{2})", text, re.IGNORECASE)
                total_price = float(price_matches[-1].replace(",", "")) if price_matches else 0.00
                # Note: Because reading random PDF tables is unreliable without AI,
                # we insert the Total as a single item line to guarantee the Excel file gets populated.
                # Format: (Description, QTY, UOM, Rate)
                data['items'].append(("Extracted Total (Please verify PDF for line items)", 1, "Lot", total_price))

        return data

    def get_vendor_columns(self, ws, extracted_vendor_name):
        """ Scans Row 6 to find the vendor. If empty, assigns the new vendor to that slot. """
        # Based on your Codex, columns are H(8), J(10), L(12)
        vendor_slots = [8, 10, 12]
        
        # 1. Check if vendor already exists in row 6
        for col in vendor_slots:
            cell_val = str(ws.cell(row=6, column=col).value or "").strip()
            if cell_val:
                # Partial match (e.g. "IPENET" matches "IPENET Solution Sdn Bhd")
                if extracted_vendor_name.lower() in cell_val.lower() or cell_val.lower() in extracted_vendor_name.lower():
                    return col, col + 1

        # 2. If vendor is new, find the first EMPTY slot and claim it
        for col in vendor_slots:
            if not ws.cell(row=6, column=col).value:
                print(f"Assigning new vendor '{extracted_vendor_name}' to Column {col}")
                ws.cell(row=6, column=col).value = extracted_vendor_name
                return col, col + 1

        # Fallback if all slots are full
        return 8, 9 

    def process_single_pdf(self, pdf_path, ws):
        # 1. Extract Data
        extracted = self.extract_pdf_data(pdf_path)
        print(f"Extracted from {os.path.basename(pdf_path)}:", extracted)

        # 2. Find Dynamic Columns
        rate_col, amt_col = self.get_vendor_columns(ws, extracted['vendor'])

        # 3. Ensure enough rows for items (insert before footer if needed)
        start_row = 8
        footer_start = 16
        items_count = len(extracted['items'])
        available_rows = footer_start - start_row
        rows_to_insert = max(0, items_count - available_rows)
        if rows_to_insert > 0:
            ws.insert_rows(footer_start, amount=rows_to_insert)
            footer_start += rows_to_insert

        # 4. Write Items based on Codex (Rows 8+, Columns C, F, G)
        for i, item in enumerate(extracted['items']):
            curr_row = start_row + i
            desc, qty, uom, rate = item
            
            ws.cell(row=curr_row, column=3).value = desc   # Column C (Description)
            ws.cell(row=curr_row, column=6).value = qty    # Column F (QTY)
            ws.cell(row=curr_row, column=7).value = uom    # Column G (UOM)
            
            ws.cell(row=curr_row, column=rate_col).value = rate         # Vendor Rate Col
            ws.cell(row=curr_row, column=amt_col).value = qty * rate    # Vendor Amount Col

        # 5. Write Footer Terms based on Codex (Rows 16, 17, 18)
        ws.cell(row=footer_start, column=rate_col).value = extracted['payment']
        ws.cell(row=footer_start + 1, column=rate_col).value = extracted['delivery']
        ws.cell(row=footer_start + 2, column=rate_col).value = extracted['ref_no']


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = QCFApp(root)
    root.mainloop()
