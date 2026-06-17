"""
Invoice Splitter & Matcher  V2
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Input : one PDF containing:
          Page 1  – summary table (Excel-style scan)
          Pages 2+ – invoices / credit-advices / receipts

Output per summary row:
  📁 Payment_[Company]_[Date]/
       summary.pdf             ← page 1 (the Excel table)
       VendorName.pdf          ← all pages belonging to this row
       ...

Matching logic:
  1. OCR page 1  → extract rows  (vendor, expected_amount)
  2. OCR pages 2+→ detect Credit Advice pages (amount + beneficiary)
  3. Match Credit Advice amounts to summary rows  → group anchors
  4. Forward-fill remaining pages into the nearest group
  5. Split PDF into named files per row
"""

import io
import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import fitz
import numpy as np
import pytesseract
from PIL import Image

try:
    import openpyxl
    _HAVE_OPENPYXL = True
except ImportError:
    _HAVE_OPENPYXL = False

# ── Tesseract path ────────────────────────────────────────────────────────────
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

APP_VERSION = "V2"
ZOOM        = 2
LANG        = "eng+tha"

NUM_RE        = re.compile(r"\b(\d{1,3}(?:,\d{3})*\.\d{2})\b")
CREDIT_ADV_RE = re.compile(r"credit\s+advice", re.IGNORECASE)
SUCCESS_RE    = re.compile(r"\bsuccess\b", re.IGNORECASE)
PMT_AMT_RE    = re.compile(
    r"payment\s+amount\s*[:\-]?\s*([\d,]+\.\d{2})", re.IGNORECASE)
AMOUNT_RE     = re.compile(
    r"(?<!\w)amount\s*[:\-]?\s*([\d,]+\.\d{2})", re.IGNORECASE)

_COMPANY_KEYWORDS = re.compile(
    r"\b(co\.?,?\s*ltd\.?|company|corporation|corp\.?|public|co\.,ltd|pte\.?|inc\.?|llc)\b",
    re.IGNORECASE,
)
_DATE_FORMATS = re.compile(r"\b(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})\b")


# ── Excel summary reader ──────────────────────────────────────────────────────
def _cell_to_float(val) -> float | None:
    if val is None:
        return None
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


def parse_summary_from_excel(excel_path: str) -> tuple[list[dict], str, str]:
    """Read summary rows from the active sheet of an Excel file.
    Returns (summary_rows, company_name, date_str).

    Expected format (CHILAT-style):
      Row 1: ... company_name ... date ...
      Row 2: headers (Detail / Amount)
      Rows 3+: item_no(B) | vendor(C) | desc(D) | ... | amount(F or H)
    """
    if not _HAVE_OPENPYXL:
        raise RuntimeError("openpyxl not installed. Run setup.bat option 1.")

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active

    all_rows = list(ws.iter_rows(values_only=True))

    # ── Company name & date: scan first 5 non-empty rows ─────────────────────
    import datetime as _dt
    company_name = ""
    date_str     = ""
    for row in all_rows[:10]:
        if not any(c is not None for c in row):
            continue
        for cell in row:
            if cell is None:
                continue
            if isinstance(cell, _dt.datetime):
                if not date_str:
                    date_str = cell.strftime("%d/%m/%Y")
                continue
            val = str(cell).strip()
            if not val:
                continue
            if not company_name:
                val_low = val.lower()
                if any(k in val_low for k in [
                    "co.,ltd", "co., ltd", "co,ltd", "holding", "limited",
                    "corporation", "corp", "pte", "inc", "llc", "company"
                ]):
                    company_name = val
            if not date_str:
                m = _DATE_FORMATS.search(val)
                if m:
                    date_str = m.group(1)
        if company_name and date_str:
            break

    # ── Find data rows: row where col B is a positive integer (item number) ───
    # Columns (0-based): B=1, C=2, D=3, E=4, F=5, G=6, H=7
    summary_rows = []
    for row in all_rows[1:]:          # skip row 1 (company header)
        if len(row) < 3:
            continue
        # col B must be a small positive integer (item number)
        item_no = _cell_to_float(row[1])
        if item_no is None or item_no <= 0 or item_no != int(item_no) or item_no > 999:
            continue
        # col C = vendor name
        vendor = str(row[2]).strip() if row[2] is not None else ""
        if not vendor:
            continue
        # amount: prefer col F (index 5), fall back to col H (index 7)
        amount = None
        for col_idx in [5, 7, 6, 4]:
            if len(row) > col_idx:
                v = _cell_to_float(row[col_idx])
                if v and v > 0:
                    amount = v
                    break
        if amount is None:
            continue
        summary_rows.append({
            "row_idx": int(item_no),
            "vendor":  vendor,
            "amount":  amount,
        })

    # Fallback company from file name
    if not company_name:
        company_name = os.path.splitext(os.path.basename(excel_path))[0]

    return summary_rows, company_name, date_str


# ── OCR helpers ───────────────────────────────────────────────────────────────
def ocr_page(doc: fitz.Document, page_idx: int) -> str:
    page = doc[page_idx]
    mat  = fitz.Matrix(ZOOM, ZOOM)
    pix  = page.get_pixmap(matrix=mat)
    img  = Image.open(io.BytesIO(pix.tobytes("png")))
    return pytesseract.image_to_string(img, lang=LANG, config="--psm 6")


def _ocr_header_cell(doc: fitz.Document) -> str:
    """OCR the top-left header cell of page 1 to find the company name."""
    page = doc[0]
    mat  = fitz.Matrix(4, 4)
    pix  = page.get_pixmap(matrix=mat)
    img  = Image.open(io.BytesIO(pix.tobytes("png")))
    w, h = img.size
    cell = img.crop((0, int(h * 0.01), int(w * 0.45), int(h * 0.055)))
    cell = cell.convert("L")
    arr  = np.array(cell)
    arr  = np.where(arr < 180, 0, 255).astype("uint8")
    cell = Image.fromarray(arr)
    return pytesseract.image_to_string(cell, lang="eng", config="--psm 6 --oem 3")


# ── Page classifier ───────────────────────────────────────────────────────────
def classify_page(text: str) -> dict:
    is_ca       = bool(CREDIT_ADV_RE.search(text)) or bool(SUCCESS_RE.search(text))
    pmt_amount  = None
    if is_ca:
        m = PMT_AMT_RE.search(text) or AMOUNT_RE.search(text)
        if m:
            pmt_amount = float(m.group(1).replace(",", ""))

    # beneficiary
    beneficiary = ""
    BEN_RE = re.compile(
        r"(?:beneficiary|benef\.?|ben\.?)\s*(?:name)?\s*[:\-]?\s*(.+)",
        re.IGNORECASE)
    bm = BEN_RE.search(text)
    if bm:
        beneficiary = bm.group(1).strip()[:80]

    # all amounts
    all_amounts = [float(x.replace(",", "")) for x in NUM_RE.findall(text)]

    # dates
    dates = _DATE_FORMATS.findall(text)

    # invoice / ref numbers
    INV_RE = re.compile(
        r"(?:invoice|inv\.?|ref\.?|reference)\s*(?:no\.?|number|#)?\s*[:\-]?\s*([A-Z0-9\-/]+)",
        re.IGNORECASE)
    inv_numbers = INV_RE.findall(text)

    return {
        "is_credit_advice": is_ca,
        "payment_amount":   pmt_amount,
        "beneficiary":      beneficiary,
        "all_amounts":      all_amounts,
        "dates":            dates,
        "raw_text":         text,
        "inv_numbers":      inv_numbers,
        "full_text":        text,
    }


# ── Vendor extraction ─────────────────────────────────────────────────────────
def _score_vendor_token(s: str) -> int:
    if not s:
        return 0
    alpha = sum(c.isalpha() for c in s)
    if alpha == 0:
        return 0
    vowels = sum(c in "aeiouAEIOU" for c in s)
    if vowels == 0 and len(s) > 3:
        return 0
    return alpha


def _extract_vendor(line: str) -> str:
    line = line.strip()
    if not line:
        return ""
    # split on pipes, take best segment
    segments = [s.strip() for s in re.split(r"\|", line)]

    def best_seg(segs):
        best, best_score = "", 0
        for seg in segs:
            # trim at first 2+ digit run
            seg = re.split(r"\d{2,}", seg)[0].strip(" -:.,|[]")
            score = _score_vendor_token(seg)
            if score > best_score:
                best, best_score = seg, score
        return best

    # Priority 1: text before first '['
    pre_bracket = re.split(r"\[", line)[0]
    pb_segs = [s.strip() for s in re.split(r"\|", pre_bracket)]
    candidate = best_seg(pb_segs)
    if candidate and _score_vendor_token(candidate) > 3:
        return candidate

    # Priority 2: full line segments
    return best_seg(segments)


def _best_vendor(lines: list, amount_idx: int,
                 all_amounts: list[float], target: float) -> str:
    # Look back up to 3 lines
    for offset in range(1, 4):
        idx = amount_idx - offset
        if idx < 0:
            break
        # stop if another CA amount on this line
        line_nums = [float(x.replace(",", "")) for x in NUM_RE.findall(lines[idx])]
        if any(abs(n - a) < 0.02 for n in line_nums for a in all_amounts if abs(a - target) > 0.02):
            break
        v = _extract_vendor(lines[idx])
        if v:
            return v
    # fallback: same line
    return _extract_vendor(lines[amount_idx])


# ── Summary parser ────────────────────────────────────────────────────────────
def parse_summary(text: str, ca_amounts: list[float],
                  prefilled_rows: list[dict] | None = None) -> list[dict]:
    # Excel mode: rows already parsed — return them as-is (order from Excel)
    if prefilled_rows is not None:
        return [dict(r) for r in prefilled_rows]

    # PDF mode: OCR text
    lines    = text.splitlines()
    rows     = []
    seen2: set[float] = set()
    all_set  = set(ca_amounts)

    for line_idx, line in enumerate(lines):
        for raw in NUM_RE.findall(line):
            amt = float(raw.replace(",", ""))
            if amt in all_set and amt not in seen2:
                seen2.add(amt)
                vendor = _best_vendor(lines, line_idx, list(all_set), amt)
                rows.append({
                    "row_idx": len(rows) + 1,
                    "vendor":  vendor,
                    "amount":  amt,
                })

    return rows


# ── Page assignment ───────────────────────────────────────────────────────────
def _vendor_in_text(vendor: str, text: str) -> bool:
    """Check if any significant word from vendor name appears in the OCR text."""
    text_lower = text.lower()
    # try full name first
    words = [w for w in re.split(r"[\s,.\-/]+", vendor) if len(w) > 3]
    if not words:
        return vendor.lower() in text_lower
    matches = sum(1 for w in words if w.lower() in text_lower)
    return matches >= max(1, len(words) // 2)


def assign_pages(summary_rows: list[dict],
                 page_infos: list[dict]) -> list[int | None]:
    n           = len(page_infos)
    assignments = [None] * n

    # Find rows that share the same amount (ambiguous)
    from collections import Counter
    amt_counts = Counter(round(r["amount"], 2) for r in summary_rows)

    # build map: abs_page_idx → row_idx  for CA anchors
    ca_map: dict[int, int] = {}
    for abs_idx, info in enumerate(page_infos):
        if info["is_credit_advice"] and info["payment_amount"] is not None:
            page_text = info.get("raw_text", "")
            candidates = [r for r in summary_rows
                          if abs(info["payment_amount"] - r["amount"]) < 0.02]
            if not candidates:
                continue
            if len(candidates) == 1:
                ca_map[abs_idx] = candidates[0]["row_idx"]
            else:
                # Same amount — try vendor name match in page text
                matched = [r for r in candidates if _vendor_in_text(r["vendor"], page_text)]
                if matched:
                    ca_map[abs_idx] = matched[0]["row_idx"]
                else:
                    # fallback: first unassigned candidate
                    assigned_rows = set(ca_map.values())
                    for r in candidates:
                        if r["row_idx"] not in assigned_rows:
                            ca_map[abs_idx] = r["row_idx"]
                            break

    if not ca_map:
        return assignments

    # forward-fill from first anchor
    first_anchor = min(ca_map.keys())
    current      = ca_map[first_anchor]
    assignments[first_anchor] = current

    for abs_idx in range(first_anchor + 1, n):
        if abs_idx in ca_map:
            current = ca_map[abs_idx]
        assignments[abs_idx] = current

    # pages before first anchor → first row
    for abs_idx in range(0, first_anchor):
        assignments[abs_idx] = ca_map[first_anchor]

    return assignments


# ── Group data extractor ──────────────────────────────────────────────────────
def extract_group_data(page_infos: list[dict], pages: list[int]) -> dict:
    ca_amount   = None
    beneficiary = ""
    all_amounts: list[float] = []
    dates: list[str]         = []
    inv_numbers: list[str]   = []

    for p in pages:
        info = page_infos[p]
        if info["is_credit_advice"] and info["payment_amount"] and ca_amount is None:
            ca_amount   = info["payment_amount"]
            beneficiary = info["beneficiary"]
        all_amounts.extend(info["all_amounts"])
        dates.extend(info["dates"])
        inv_numbers.extend(info["inv_numbers"])

    return {
        "ca_amount":   ca_amount,
        "beneficiary": beneficiary,
        "all_amounts": sorted(set(all_amounts)),
        "dates":       list(dict.fromkeys(dates)),
        "inv_numbers": list(dict.fromkeys(inv_numbers)),
    }


# ── Company + date extraction ─────────────────────────────────────────────────
def extract_company_and_date(summary_text: str,
                              page_infos: list[dict],
                              doc=None) -> tuple[str, str]:
    lines = summary_text.splitlines()

    # ── Company name ──────────────────────────────────────────────────────────
    company = ""

    def _clean_candidate(raw: str) -> str:
        cand = re.split(r"[\[\|]", raw)[0].strip(" -:.,")
        cand = re.sub(r"^[\s\-|:.,\[\]]+", "", cand)
        m = re.search(
            r"(co\.?,?\s*ltd\.?|limited|corporation|corp\.?|public\s+co|pte\.?|inc\.?|llc)",
            cand, re.IGNORECASE)
        if m:
            cand = cand[:m.end()].strip(" -:.,")
        return cand

    # Priority 1: OCR header cell (top-left of page 1)
    if doc is not None:
        try:
            header_text = _ocr_header_cell(doc)
            for line in header_text.splitlines():
                line_s = line.strip()
                if _COMPANY_KEYWORDS.search(line_s):
                    candidate = _clean_candidate(line_s)
                    if len(candidate) >= 5:
                        company = candidate[:80]
                        break
        except Exception:
            pass

    # Priority 2: full page 1 text
    if not company:
        for line in lines:
            line_s = line.strip()
            if not line_s:
                continue
            if _COMPANY_KEYWORDS.search(line_s):
                candidate = _clean_candidate(line_s)
                if len(candidate) >= 5:
                    company = candidate[:80]
                    break

    # Fallback: payer from Credit Advice
    if not company:
        PAYER_RE = re.compile(
            r"payer[/\s]*remitter\s+name\s*[:\-]?\s*(.+)", re.IGNORECASE)
        for info in page_infos:
            m = PAYER_RE.search(info.get("full_text", ""))
            if m:
                company = m.group(1).strip()[:80]
                break

    if not company:
        company = "Unknown_Company"

    # ── Date ─────────────────────────────────────────────────────────────────
    date_str = ""
    for line in lines:
        dates = _DATE_FORMATS.findall(line)
        if dates:
            date_str = dates[0]
            break

    if not date_str:
        for info in page_infos:
            if info["is_credit_advice"] and info["dates"]:
                date_str = info["dates"][0]
                break

    if not date_str:
        from datetime import date as _d
        date_str = _d.today().strftime("%d/%m/%Y")

    # Sanitise for folder name
    company_safe = re.sub(r"[\\/:*?\"<>|()\[\],.]+", " ", company)
    company_safe = re.sub(r"\s+", "_", company_safe.strip()).strip("_")
    date_safe    = re.sub(r"[\/\\\-\.]", "-", date_str)

    return company_safe, date_safe


# ── Excel → PDF converter ─────────────────────────────────────────────────────
def _excel_to_pdf(excel_path: str, pdf_path: str):
    """Export the first Excel sheet directly to PDF using Office COM."""
    import subprocess, tempfile

    abs_excel = os.path.abspath(excel_path)
    abs_pdf   = os.path.abspath(pdf_path)

    # Ensure output directory exists
    os.makedirs(os.path.dirname(abs_pdf), exist_ok=True)

    # Remove any pre-existing PDF so Excel won't fail on a locked file
    if os.path.exists(abs_pdf):
        try:
            os.remove(abs_pdf)
        except Exception:
            pass

    # Pass paths via env vars (avoids all escaping issues with spaces/parens).
    # Open read-only — works even when the file is already open in Excel.
    # Export at workbook level — more reliable than worksheet level.
    script = (
        "$ErrorActionPreference = 'Stop'\n"
        "$ep = $env:_XL_SRC\n"
        "$pp = $env:_XL_DST\n"
        "$xl = New-Object -ComObject Excel.Application\n"
        "$xl.Visible = $false\n"
        "$xl.DisplayAlerts = $false\n"
        "try {\n"
        "    $wb = $xl.Workbooks.Open($ep, 0, $true)\n"
        "    $wb.ExportAsFixedFormat(0, $pp)\n"
        "    $wb.Close($false)\n"
        "} finally {\n"
        "    $xl.Quit()\n"
        "    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null\n"
        "}\n"
    )

    ps1 = tempfile.NamedTemporaryFile(suffix=".ps1", mode="w",
                                       encoding="utf-16", delete=False)
    ps1name = ps1.name
    ps1.write(script)
    ps1.close()

    env = os.environ.copy()
    env["_XL_SRC"] = abs_excel
    env["_XL_DST"] = abs_pdf

    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1name],
            capture_output=True, timeout=60, text=True, env=env)
        if not os.path.exists(abs_pdf):
            raise RuntimeError(f"summary.pdf not created.\nSTDERR: {result.stderr[:400]}")
    finally:
        try:
            os.remove(ps1name)
        except Exception:
            pass


# ── PDF splitter ──────────────────────────────────────────────────────────────
def split_pdf(doc: fitz.Document, assignments: list[int | None],
              summary_rows: list[dict], base_out_dir: str,
              page_infos: list[dict],
              summary_text: str = "",
              excel_company: str = "",
              excel_date: str = "",
              excel_path: str = "") -> tuple[list[dict], str]:
    """
    Split doc into named PDF files inside:
        Payment_[CompanyName]_[Date]/
    Returns (results, out_dir).
    """
    if excel_company:
        company_safe = re.sub(r"[\\/:*?\"<>|()\[\],. ]+", "_", excel_company).strip("_")
        date_safe    = re.sub(r"[\/\-\.]", "-", excel_date) if excel_date else \
                       __import__("datetime").date.today().strftime("%d-%m-%Y")
    else:
        company_safe, date_safe = extract_company_and_date(summary_text, page_infos, doc)
    root_folder = f"Payment_{company_safe}_{date_safe}"
    out_dir     = os.path.join(base_out_dir, root_folder)
    os.makedirs(out_dir, exist_ok=True)

    if excel_path and os.path.exists(excel_path):
        # Excel mode: convert first sheet to summary.pdf
        _excel_to_pdf(excel_path, os.path.join(out_dir, "summary.pdf"))
    else:
        # PDF mode: save page 1 as summary.pdf
        summary_pdf = fitz.open()
        summary_pdf.insert_pdf(doc, from_page=0, to_page=0)
        summary_pdf.save(os.path.join(out_dir, "summary.pdf"))
        summary_pdf.close()

    # Group pages by row
    row_pages: dict[int, list[int]] = {}
    for abs_idx, row_idx in enumerate(assignments):
        if row_idx is None:
            continue
        row_pages.setdefault(row_idx, []).append(abs_idx)

    results = []
    for row in summary_rows:
        ridx  = row["row_idx"]
        pages = row_pages.get(ridx, [])

        # Safe file name — vendor name only
        vendor_safe = re.sub(r"[\\/:*?\"<>|]+", "_", row["vendor"]).strip("_. ")
        file_name   = vendor_safe if vendor_safe else f"Row{ridx:02d}"

        # Save PDF directly in out_dir
        pdf_path = None
        if pages:
            sorted_pages = sorted(pages)
            out_pdf      = fitz.open()
            block_start  = sorted_pages[0]
            block_end    = sorted_pages[0]
            for pg in sorted_pages[1:]:
                if pg == block_end + 1:
                    block_end = pg
                else:
                    out_pdf.insert_pdf(doc, from_page=block_start, to_page=block_end)
                    block_start = block_end = pg
            out_pdf.insert_pdf(doc, from_page=block_start, to_page=block_end)
            pdf_path = os.path.join(out_dir, f"{file_name}.pdf")
            out_pdf.save(pdf_path)
            out_pdf.close()

        # Analyse
        group    = extract_group_data(page_infos, pages)
        expected = row["amount"]
        found    = group["ca_amount"]

        if found is None:
            match_status = "NO PAYMENT FOUND"
        elif abs(found - expected) < 0.02:
            match_status = "MATCH"
        else:
            match_status = f"MISMATCH (expected {expected:,.2f} / found {found:,.2f})"

        results.append({
            "row_idx":      ridx,
            "vendor":       row["vendor"],
            "expected_amt": expected,
            "ca_amount":    found,
            "beneficiary":  group["beneficiary"],
            "all_amounts":  group["all_amounts"],
            "dates":        group["dates"],
            "inv_numbers":  group["inv_numbers"],
            "pages":        [p + 1 for p in pages],
            "page_count":   len(pages),
            "folder":       out_dir,
            "pdf_path":     pdf_path,
            "match_status": match_status,
        })

    return results, out_dir


# ── GUI App ───────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Invoice Splitter & Matcher  {APP_VERSION}")
        self.resizable(False, False)
        self.configure(bg="#F0F4F8")
        self._pdf_paths    = []
        self._excel_path   = None
        self._summary_text = ""
        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - self.winfo_width())  // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _build_ui(self):
        tk.Frame(self, bg="#1F4E79", height=4).pack(fill="x")
        hdr = tk.Frame(self, bg="#1F4E79")
        hdr.pack(fill="x")
        tk.Label(hdr, text="  Invoice Splitter & Matcher",
                 font=("Calibri", 16, "bold"), bg="#1F4E79", fg="white",
                 pady=14).pack(side="left")
        tk.Label(hdr, text=APP_VERSION,
                 font=("Calibri", 9), bg="#1F4E79", fg="#90CAF9",
                 pady=14).pack(side="left")

        body = tk.Frame(self, bg="#F0F4F8", padx=28, pady=20)
        body.pack(fill="both")

        # PDF row
        tk.Label(body, text="PDF Files:", font=("Calibri", 11, "bold"),
                 bg="#F0F4F8").grid(row=0, column=0, sticky="w", pady=6)
        self._pdf_lbl = tk.Label(body, text="No files selected",
                                  font=("Calibri", 10), bg="#F0F4F8", fg="#666",
                                  width=46, anchor="w")
        self._pdf_lbl.grid(row=0, column=1, padx=8)
        pdf_btn_frame = tk.Frame(body, bg="#F0F4F8")
        pdf_btn_frame.grid(row=0, column=2)
        tk.Button(pdf_btn_frame, text="Browse…", command=self._pick_pdf,
                  font=("Calibri", 10), bg="#1F4E79", fg="white",
                  relief="flat", padx=10).pack(side="left")
        tk.Button(pdf_btn_frame, text="✕", command=self._clear_pdf,
                  font=("Calibri", 10), bg="#888", fg="white",
                  relief="flat", padx=6).pack(side="left", padx=(4, 0))

        # Excel row (optional)
        tk.Label(body, text="Excel Summary\n(optional):", font=("Calibri", 11, "bold"),
                 bg="#F0F4F8").grid(row=1, column=0, sticky="w", pady=6)
        self._excel_lbl = tk.Label(body, text="Not selected  (will use page 1 of PDF)",
                                    font=("Calibri", 10), bg="#F0F4F8", fg="#888",
                                    width=46, anchor="w")
        self._excel_lbl.grid(row=1, column=1, padx=8)
        xbtn_frame = tk.Frame(body, bg="#F0F4F8")
        xbtn_frame.grid(row=1, column=2)
        tk.Button(xbtn_frame, text="Browse…", command=self._pick_excel,
                  font=("Calibri", 10), bg="#2E7D32", fg="white",
                  relief="flat", padx=10).pack(side="left")
        tk.Button(xbtn_frame, text="✕", command=self._clear_excel,
                  font=("Calibri", 10), bg="#888", fg="white",
                  relief="flat", padx=6).pack(side="left", padx=(4, 0))

        # Output row
        tk.Label(body, text="Output Folder:", font=("Calibri", 11, "bold"),
                 bg="#F0F4F8").grid(row=2, column=0, sticky="w", pady=6)
        self._out_lbl = tk.Label(body, text="No folder selected",
                                  font=("Calibri", 10), bg="#F0F4F8", fg="#666",
                                  width=46, anchor="w")
        self._out_lbl.grid(row=2, column=1, padx=8)
        tk.Button(body, text="Browse…", command=self._pick_out,
                  font=("Calibri", 10), bg="#1F4E79", fg="white",
                  relief="flat", padx=10).grid(row=2, column=2)

        # Progress
        self._prog_var = tk.DoubleVar()
        self._prog_bar = ttk.Progressbar(body, variable=self._prog_var,
                                          maximum=100, length=460)
        self._prog_bar.grid(row=3, column=0, columnspan=3, pady=(18, 4))

        self._status_lbl = tk.Label(body, text="Ready.",
                                     font=("Calibri", 9), bg="#F0F4F8", fg="#555")
        self._status_lbl.grid(row=4, column=0, columnspan=3)

        # Run button
        self._btn = tk.Button(body, text="▶  Analyse & Split",
                               font=("Calibri", 13, "bold"),
                               bg="#1F4E79", fg="white", relief="flat",
                               padx=24, pady=10, cursor="hand2",
                               command=self._run)
        self._btn.grid(row=5, column=0, columnspan=3, pady=(20, 4))

    def _pick_pdf(self):
        paths = filedialog.askopenfilenames(
            title="Select PDF files (hold Ctrl for multiple)",
            filetypes=[("PDF files", "*.pdf")])
        if paths:
            self._pdf_paths = list(paths)
            n = len(self._pdf_paths)
            if n == 1:
                label = os.path.basename(self._pdf_paths[0])
            else:
                label = f"{n} files selected"
            self._pdf_lbl.config(text=label, fg="#1F4E79")

    def _clear_pdf(self):
        self._pdf_paths = []
        self._pdf_lbl.config(text="No files selected", fg="#666")

    def _pick_excel(self):
        p = filedialog.askopenfilename(
            title="Select Excel Summary",
            filetypes=[("Excel files", "*.xlsx *.xls")])
        if p:
            self._excel_path = p
            self._excel_lbl.config(text=os.path.basename(p), fg="#2E7D32")

    def _clear_excel(self):
        self._excel_path = None
        self._excel_lbl.config(text="Not selected  (will use page 1 of PDF)", fg="#888")

    def _pick_out(self):
        d = filedialog.askdirectory(title="Select output folder")
        if d:
            self._out_dir = d
            self._out_lbl.config(text=d, fg="#1F4E79")

    def _status(self, msg: str):
        self.after(0, lambda: self._status_lbl.config(text=msg))

    def _prog(self, val: float):
        self.after(0, lambda: self._prog_var.set(val))

    def _run(self):
        if not self._pdf_paths:
            messagebox.showwarning("Missing", "Please select at least one PDF file.")
            return
        if not hasattr(self, "_out_dir") or not self._out_dir:
            messagebox.showwarning("Missing", "Please select an output folder.")
            return
        self._btn.config(state="disabled")
        threading.Thread(
            target=self._worker,
            args=(list(self._pdf_paths), self._out_dir),
            daemon=True).start()

    def _worker(self, pdf_paths: list, out_dir: str):
        try:
            # Merge all selected PDFs into one document
            if len(pdf_paths) == 1:
                doc = fitz.open(pdf_paths[0])
            else:
                self._status(f"Merging {len(pdf_paths)} PDF files…")
                doc = fitz.open()
                for p in pdf_paths:
                    src = fitz.open(p)
                    doc.insert_pdf(src)
                    src.close()
            n   = doc.page_count
            self._status(f"Opened PDF — {n} pages")
            self._prog(0)

            excel_path = self._excel_path
            summary_text = ""

            if excel_path:
                # ── Excel mode: all PDF pages are invoices ──
                self._status("Reading summary from Excel…")
                summary_rows_pre, company_from_excel, date_from_excel = \
                    parse_summary_from_excel(excel_path)
                self._prog(5)

                excel_amounts = {round(r["amount"], 2) for r in summary_rows_pre}

                page_infos: list[dict] = []
                for i in range(0, n):
                    self._status(f"Scanning page {i+1} / {n}…")
                    text = ocr_page(doc, i)
                    info = classify_page(text)
                    # In Excel mode: any page containing a matching amount is an anchor
                    if not info["is_credit_advice"]:
                        for amt in info["all_amounts"]:
                            if round(amt, 2) in excel_amounts:
                                info["is_credit_advice"] = True
                                info["payment_amount"]   = amt
                                break
                    page_infos.append(info)
                    self._prog(5 + int(75 * (i + 1) / n))

                ca_amounts = [
                    info["payment_amount"]
                    for info in page_infos
                    if info["is_credit_advice"] and info["payment_amount"]
                ]
                if not ca_amounts:
                    self.after(0, lambda: messagebox.showerror(
                        "Error",
                        "No matching amounts found in PDF.\n"
                        "Make sure the PDF contains the payment amounts listed in the Excel."
                    ))
                    return

                summary_rows = parse_summary(summary_text, ca_amounts,
                                             prefilled_rows=summary_rows_pre)
                self._prog(82)
                self._excel_company = company_from_excel
                self._excel_date    = date_from_excel

            else:
                # ── PDF mode: page 1 = summary scan ──
                self._status("Reading summary page (page 1)…")
                summary_text = ocr_page(doc, 0)
                self._prog(5)

                page_infos: list[dict] = []
                for i in range(1, n):
                    self._status(f"Scanning page {i+1} / {n}…")
                    text = ocr_page(doc, i)
                    info = classify_page(text)
                    page_infos.append(info)
                    self._prog(5 + int(75 * i / max(n - 1, 1)))

                ca_amounts = [
                    info["payment_amount"]
                    for info in page_infos
                    if info["is_credit_advice"] and info["payment_amount"]
                ]
                if not ca_amounts:
                    self.after(0, lambda: messagebox.showerror(
                        "Error",
                        "No payment confirmation pages found.\n"
                        "Make sure pages 2+ contain Credit Advice or Success transfer pages."
                    ))
                    return

                self._status("Parsing summary rows…")
                summary_rows = parse_summary(summary_text, ca_amounts)
                self._prog(82)
                self._excel_company = ""
                self._excel_date    = ""

            # Step 4: Assign pages to rows
            self._status("Assigning pages to summary rows…")
            assignments = assign_pages(summary_rows, page_infos)
            self._prog(85)

            # Step 5: Show preview → user confirms before splitting
            self._summary_text = summary_text
            self.after(0, lambda: self._show_preview(
                doc, summary_rows, page_infos, assignments, out_dir))

        except Exception as exc:
            import traceback
            _msg = f"{exc}\n\n{traceback.format_exc()[-600:]}"
            self.after(0, lambda m=_msg: messagebox.showerror("Error", m))
            self.after(0, lambda: self._btn.config(state="normal"))

    # ── preview dialog ────────────────────────────────────────────────────────
    def _show_preview(self, doc, summary_rows, page_infos, assignments, out_dir):
        dlg = tk.Toplevel(self)
        dlg.title("Preview — confirm page assignments before splitting")
        dlg.resizable(True, True)
        dlg.grab_set()
        dlg.geometry("920x560")

        tk.Label(dlg,
                 text="  Review assignments — then click  ✔ Split  to create folders.",
                 font=("Calibri", 11, "bold"), bg="#1F4E79", fg="white",
                 anchor="w", pady=10).pack(fill="x")

        # summary legend
        leg = tk.Frame(dlg, bg="#EBF3FB", padx=14, pady=8)
        leg.pack(fill="x")
        tk.Label(leg, text="Summary rows found:", font=("Calibri", 10, "bold"),
                 bg="#EBF3FB").pack(anchor="w")
        for row in summary_rows:
            tk.Label(leg,
                     text=f"  Row {row['row_idx']:02d}:  {row['vendor'][:50]}"
                          f"   —  {row['amount']:,.2f} THB",
                     font=("Calibri", 10), bg="#EBF3FB").pack(anchor="w")

        # page table
        cols = ("page", "type", "amount", "assigned_row")
        tree = ttk.Treeview(dlg, columns=cols, show="headings", height=14)
        tree.heading("page",         text="Page")
        tree.heading("type",         text="Type")
        tree.heading("amount",       text="Amount")
        tree.heading("assigned_row", text="Assigned Row")
        tree.column("page",         width=60,  anchor="center")
        tree.column("type",         width=160, anchor="w")
        tree.column("amount",       width=130, anchor="e")
        tree.column("assigned_row", width=520, anchor="w")

        row_map = {r["row_idx"]: r for r in summary_rows}
        for abs_idx, row_idx in enumerate(assignments):
            info     = page_infos[abs_idx]
            pg_label = str(abs_idx + 2)  # 1-based, +1 for summary page
            pg_type  = "Credit Advice" if info["is_credit_advice"] else "Document"
            amt_s    = f"{info['payment_amount']:,.2f}" if info["payment_amount"] else ""
            if row_idx and row_idx in row_map:
                row_lbl = f"Row {row_idx:02d}: {row_map[row_idx]['vendor'][:50]}"
            else:
                row_lbl = "—"
            tree.insert("", "end", values=(pg_label, pg_type, amt_s, row_lbl))

        sb = ttk.Scrollbar(dlg, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True, padx=(14, 0), pady=10)
        sb.pack(side="left", fill="y", pady=10)

        btn_f = tk.Frame(dlg, bg="#F0F4F8", pady=12)
        btn_f.pack(fill="x")

        def do_split():
            dlg.destroy()
            threading.Thread(
                target=self._do_split,
                args=(doc, summary_rows, page_infos, assignments, out_dir),
                daemon=True).start()

        tk.Button(btn_f, text="✔  Split",
                  font=("Calibri", 12, "bold"),
                  bg="#1F4E79", fg="white", relief="flat",
                  padx=22, pady=8, cursor="hand2",
                  command=do_split).pack(side="left", padx=16)
        tk.Button(btn_f, text="Cancel",
                  font=("Calibri", 10), bg="#CCC", relief="flat",
                  padx=14, pady=8,
                  command=lambda: (dlg.destroy(),
                                   self._btn.config(state="normal"),
                                   self._status("Cancelled."))
                  ).pack(side="left")

    # ── actual split ──────────────────────────────────────────────────────────
    def _do_split(self, doc, summary_rows, page_infos, assignments, out_dir):
        try:
            self._status("Splitting PDF into files…")
            self._prog(88)
            results, payment_dir = split_pdf(
                doc, assignments, summary_rows, out_dir,
                page_infos, summary_text=self._summary_text,
                excel_company=getattr(self, "_excel_company", ""),
                excel_date=getattr(self, "_excel_date", ""),
                excel_path=getattr(self, "_excel_path", "") or "")
            self._prog(100)

            n_match    = sum(1 for r in results if r["match_status"] == "MATCH")
            n_mismatch = sum(1 for r in results if r["match_status"].startswith("MISMATCH"))
            n_missing  = sum(1 for r in results if r["match_status"].startswith("NO"))

            summary = (
                f"Split complete!\n\n"
                f"  Rows processed : {len(results)}\n"
                f"  ✔ MATCH        : {n_match}\n"
                f"  ⚠ MISMATCH     : {n_mismatch}\n"
                f"  ✗ MISSING      : {n_missing}\n\n"
                f"Saved to:\n{payment_dir}"
            )
            self._status(f"Done — {n_match}/{len(results)} rows matched.")
            self.after(0, lambda: messagebox.showinfo("Complete", summary))

        except Exception as exc:
            import traceback
            _msg = f"{exc}\n\n{traceback.format_exc()[-600:]}"
            self.after(0, lambda m=_msg: messagebox.showerror("Error", m))
        finally:
            self.after(0, lambda: self._btn.config(state="normal"))


if __name__ == "__main__":
    App().mainloop()
