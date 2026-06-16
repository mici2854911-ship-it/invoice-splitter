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
def parse_summary(text: str, ca_amounts: list[float]) -> list[dict]:
    lines    = text.splitlines()
    rows     = []
    seen: set[float] = set()
    all_set  = set(ca_amounts)

    for line_idx, line in enumerate(lines):
        for raw in NUM_RE.findall(line):
            amt = float(raw.replace(",", ""))
            if amt in all_set and amt not in seen:
                seen.add(amt)
                vendor = _best_vendor(lines, line_idx, list(all_set), amt)
                rows.append({
                    "row_idx": len(rows) + 1,
                    "vendor":  vendor,
                    "amount":  amt,
                })

    return rows


# ── Page assignment ───────────────────────────────────────────────────────────
def assign_pages(summary_rows: list[dict],
                 page_infos: list[dict]) -> list[int | None]:
    n           = len(page_infos)
    assignments = [None] * n

    # build map: abs_page_idx → row_idx  for CA anchors
    ca_map: dict[int, int] = {}
    for abs_idx, info in enumerate(page_infos):
        if info["is_credit_advice"] and info["payment_amount"] is not None:
            for row in summary_rows:
                if abs(info["payment_amount"] - row["amount"]) < 0.02:
                    ca_map[abs_idx] = row["row_idx"]
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


# ── PDF splitter ──────────────────────────────────────────────────────────────
def split_pdf(doc: fitz.Document, assignments: list[int | None],
              summary_rows: list[dict], base_out_dir: str,
              page_infos: list[dict],
              summary_text: str = "") -> tuple[list[dict], str]:
    """
    Split doc into named PDF files inside:
        Payment_[CompanyName]_[Date]/
    Returns (results, out_dir).
    """
    company_safe, date_safe = extract_company_and_date(summary_text, page_infos, doc)
    root_folder = f"Payment_{company_safe}_{date_safe}"
    out_dir     = os.path.join(base_out_dir, root_folder)
    os.makedirs(out_dir, exist_ok=True)

    # Save page 1 (summary table) as summary.pdf
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
        self._pdf_path     = None
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
        tk.Label(body, text="PDF File:", font=("Calibri", 11, "bold"),
                 bg="#F0F4F8").grid(row=0, column=0, sticky="w", pady=6)
        self._pdf_lbl = tk.Label(body, text="No file selected",
                                  font=("Calibri", 10), bg="#F0F4F8", fg="#666",
                                  width=46, anchor="w")
        self._pdf_lbl.grid(row=0, column=1, padx=8)
        tk.Button(body, text="Browse…", command=self._pick_pdf,
                  font=("Calibri", 10), bg="#1F4E79", fg="white",
                  relief="flat", padx=10).grid(row=0, column=2)

        # Output row
        tk.Label(body, text="Output Folder:", font=("Calibri", 11, "bold"),
                 bg="#F0F4F8").grid(row=1, column=0, sticky="w", pady=6)
        self._out_lbl = tk.Label(body, text="No folder selected",
                                  font=("Calibri", 10), bg="#F0F4F8", fg="#666",
                                  width=46, anchor="w")
        self._out_lbl.grid(row=1, column=1, padx=8)
        tk.Button(body, text="Browse…", command=self._pick_out,
                  font=("Calibri", 10), bg="#1F4E79", fg="white",
                  relief="flat", padx=10).grid(row=1, column=2)

        # Progress
        self._prog_var = tk.DoubleVar()
        self._prog_bar = ttk.Progressbar(body, variable=self._prog_var,
                                          maximum=100, length=460)
        self._prog_bar.grid(row=2, column=0, columnspan=3, pady=(18, 4))

        self._status_lbl = tk.Label(body, text="Ready.",
                                     font=("Calibri", 9), bg="#F0F4F8", fg="#555")
        self._status_lbl.grid(row=3, column=0, columnspan=3)

        # Run button
        self._btn = tk.Button(body, text="▶  Analyse & Split",
                               font=("Calibri", 13, "bold"),
                               bg="#1F4E79", fg="white", relief="flat",
                               padx=24, pady=10, cursor="hand2",
                               command=self._run)
        self._btn.grid(row=4, column=0, columnspan=3, pady=(20, 4))

    def _pick_pdf(self):
        p = filedialog.askopenfilename(
            title="Select PDF", filetypes=[("PDF files", "*.pdf")])
        if p:
            self._pdf_path = p
            self._pdf_lbl.config(text=os.path.basename(p), fg="#1F4E79")

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
        if not self._pdf_path:
            messagebox.showwarning("Missing", "Please select a PDF file.")
            return
        if not hasattr(self, "_out_dir") or not self._out_dir:
            messagebox.showwarning("Missing", "Please select an output folder.")
            return
        self._btn.config(state="disabled")
        threading.Thread(
            target=self._worker,
            args=(self._pdf_path, self._out_dir),
            daemon=True).start()

    def _worker(self, pdf_path: str, out_dir: str):
        try:
            doc = fitz.open(pdf_path)
            n   = doc.page_count
            self._status(f"Opened PDF — {n} pages")
            self._prog(0)

            # Step 1: OCR summary page
            self._status("Reading summary page (page 1)…")
            summary_text = ocr_page(doc, 0)
            self._prog(5)

            # Step 2: OCR all content pages
            page_infos: list[dict] = []
            for i in range(1, n):
                self._status(f"Scanning page {i+1} / {n}…")
                text = ocr_page(doc, i)
                info = classify_page(text)
                page_infos.append(info)
                self._prog(5 + int(75 * i / max(n - 1, 1)))

            # Step 3: Re-parse summary using confirmed CA amounts
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
            self.after(0, lambda: messagebox.showerror("Error",
                f"{exc}\n\n{traceback.format_exc()[-600:]}"))
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
                page_infos, summary_text=self._summary_text)
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
            self.after(0, lambda: messagebox.showerror(
                "Error", f"{exc}\n\n{traceback.format_exc()[-600:]}"))
        finally:
            self.after(0, lambda: self._btn.config(state="normal"))


if __name__ == "__main__":
    App().mainloop()
