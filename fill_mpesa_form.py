#!/usr/bin/env python3
"""
Fill the Safaricom "M-PESA Customer-to-Business Service Application Form" (2025)
including Page 2 (Terms & Conditions) and Page 3 (Account Opening Authorization Form)
from an Excel workbook with pristine spacing and alignment.

Usage (from the folder that holds the files):
    python fill_mpesa_form.py
Optional:
    python fill_mpesa_form.py --xlsx YourWorkbook.xlsx --pdf blank.pdf --out filled.pdf --open

One-time setup:
    pip install openpyxl pypdf reportlab
"""
import argparse
import datetime as dt
import io
import os
import re
import sys
from pathlib import Path

from openpyxl import load_workbook
from pypdf import PdfReader, PdfWriter
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

BASE = Path(__file__).resolve().parent
FONT = "Helvetica-Bold"

# ----------------------------------------------------------------------------
# POSITIONS  (PDF points, measured from the TOP-LEFT of each page)
# key: (page, x_start, x_end, row_bottom, label_font_size, text_size)
# ----------------------------------------------------------------------------
P1 = 6.0    # label font size on page 1
P2 = 9.0    # label font size on page 2
P3 = 10.0   # label font size on page 3
T1 = 8.0    # text size used on page 1
T2 = 10.0   # text size used on page 2
T3 = 9.5    # text size used on page 3

TEXT_FIELDS = {
    # ---- Page 1: Company profile
    "industry":              (1, 85.0, 589.0, 144.8, P1, T1),
    "physical_address":      (1, 153.0, 589.0, 159.0, P1, T1),
    "postal_address":        (1, 105.5, 588.0, 172.9, P1, T1),
    "telephone":             (1, 115.0, 586.0, 186.9, P1, T1),
    "office_email":          (1, 127.5, 586.0, 201.0, P1, T1),
    "vat_number":            (1, 99.5, 587.0, 214.9, P1, T1),
    "business_type":         (1, 115.0, 587.0, 228.9, P1, T1),
    "proposed_account_name": (1, 422.0, 586.0, 243.3, P1, T1),
    # ---- Reason for M-PESA
    "collection_reason":     (1, 220.5, 582.0, 298.4, P1, T1),
    "disbursement_reason":   (1, 220.5, 582.0, 312.3, P1, T1),
    # ---- Contact details
    "contact_name":          (1, 141.0, 581.0, 339.5, P1, T1),
    "contact_email":         (1, 106.5, 312.0, 353.5, P1, T1),
    "contact_tel":           (1, 385.0, 580.0, 353.5, P1, T1),
    "finance_name":          (1, 142.0, 580.0, 367.5, P1, T1),
    "finance_email":         (1, 107.0, 315.0, 381.6, P1, T1),
    "finance_tel":           (1, 389.5, 581.0, 381.6, P1, T1),
    # ---- Administrator
    "admin_name":            (1, 128.0, 311.0, 409.1, P1, T1),
    "admin_tel":             (1, 382.5, 579.0, 409.1, P1, T1),
    "admin_username":        (1, 91.0, 311.0, 423.1, P1, T1),
    "admin_email":           (1, 373.5, 580.0, 423.1, P1, T1),
    "admin_dob":             (1, 98.0, 311.0, 437.1, P1, T1),
    "admin_id":              (1, 358.5, 579.0, 437.1, P1, T1),
    # ---- Payment details: M-PESA
    "mpesa_name":            (1, 165.0, 575.0, 484.1, P1, T1),
    "mpesa_mobile":          (1, 166.5, 312.0, 498.1, P1, T1),
    "mpesa_id":              (1, 358.5, 574.0, 498.1, P1, T1),
    # ---- Payment details: Bank
    "bank_name":             (1, 106.0, 316.0, 531.5, P1, T1),
    "bank_branch":           (1, 350.5, 579.0, 531.5, P1, T1),
    "bank_account_name":     (1, 106.0, 317.0, 545.5, P1, T1),
    "bank_account_number":   (1, 382.5, 583.0, 545.5, P1, T1),
    "ussd_mobile":           (1, 242.0, 406.0, 586.2, P1, T1),
    "ussd_id":               (1, 445.0, 585.0, 586.2, P1, T1),
    # ---- Account declaration (page 1)
    "sign_day":              (1, 93.0, 312.0, 627.2, P1, T1),
    "sign_month":            (1, 343.0, 408.0, 627.2, P1, T1),
    "sign_year":             (1, 419.0, 445.0, 627.2, P1, T1),
    "sign_location":         (1, 475.0, 585.0, 627.2, P1, T1),
    "sig1_name":             (1, 151.5, 310.0, 641.3, P1, T1),
    "sig1_designation":      (1, 355.5, 446.0, 641.3, P1, T1),
    "sig2_name":             (1, 125.5, 309.0, 655.4, P1, T1),
    "sig2_designation":      (1, 352.0, 446.0, 655.4, P1, T1),

    # ---- Page 2: Terms & conditions acceptance
    "p2_company":            (2, 88.0, 478.0, 321.5, P2, T2),
    "p2_s1_name":            (2, 66.0, 306.0, 461.2, P2, T2),
    "p2_s1_date":            (2, 66.0, 306.0, 523.4, P2, T2),
    "p2_s2_name":            (2, 66.0, 306.0, 601.0, P2, T2),
    "p2_s2_address":         (2, 66.0, 306.0, 647.6, P2, T2),
    "p2_s2_date":            (2, 66.0, 306.0, 740.7, P2, T2),

    # ---- Page 3: M-PESA Account Opening Authorization Form (Fine-tuned baseline spacing)
    "auth_org_name":         (3, 183.0, 554.0, 156.0, P3, T3),
    "auth_reg_office":       (3, 173.0, 554.0, 170.0, P3, T3),
    "auth_day":              (3, 73.0,  132.0, 187.0, P3, T3),
    "auth_month":            (3, 167.0, 230.0, 187.0, P3, T3),
    "auth_year":             (3, 244.0, 282.0, 187.0, P3, T3),
    "auth_person_name":      (3, 119.0, 436.0, 326.5, P3, T3),
    "auth_purpose":          (3, 200.0, 555.0, 456.5, P3, T3),
    "auth_bank_name":        (3, 92.0,  293.0, 501.5, P3, T3),
    "auth_bank_branch":      (3, 367.0, 554.0, 501.5, P3, T3),
    "auth_account_name":     (3, 104.0, 292.0, 516.5, P3, T3),
    "auth_account_number":   (3, 377.0, 555.0, 516.5, P3, T3),
    "auth_nominated_mobile": (3, 407.0, 552.0, 551.5, P3, T3),
    "auth_s1_name":          (3, 64.0,  291.0, 649.5, P3, T3),
    "auth_s1_date":          (3, 58.0,  290.0, 664.5, P3, T3),
    "auth_s2_name":          (3, 326.0, 553.0, 649.5, P3, T3),
    "auth_s2_date":          (3, 318.0, 553.0, 664.5, P3, T3),
}

# Checkbox images: (page, x0, top, x1, bottom)
BOXES = {
    # Page 1
    "PayBill":                 (1, 77,  263, 84,  270),
    "Buy Goods":               (1, 188, 263, 195, 270),
    "Aggregated PayBill":      (1, 326, 263, 333, 270),
    "Business Till":           (1, 447, 263, 454, 270),
    "collection":              (1, 122, 290, 129, 297),
    "disbursement":            (1, 121, 307, 127, 314),
    "M-PESA":                  (1, 45,  462, 52,  469),
    "Bank Account":            (1, 45,  510, 52,  517),
    "ussd_yes":                (1, 273, 566, 280, 573),
    "ussd_no":                 (1, 308, 566, 315, 573),
    # Page 3 (Authorization Form)
    "auth_paybill_b2c":        (3, 36.1,  234.1, 49.5,  247.5),
    "auth_business_till":      (3, 296.1, 234.1, 309.5, 247.5),
    "auth_buy_goods":          (3, 36.1,  265.1, 49.5,  278.5),
    "auth_bulk_payment":       (3, 296.1, 264.1, 309.5, 277.5),
    "auth_paybill":            (3, 36.1,  290.1, 49.5,  303.5),
    "auth_sure_pay":           (3, 296.1, 290.1, 309.5, 303.5),
}

# "Trading for" comb boxes on Page 1: (x0, x1, row_bottom)
YEARS_BOX = (113.0, 144.0, 243.3)
MONTHS_BOX = (215.0, 246.0, 243.3)


# ----------------------------------------------------------------------------
# helpers
# ----------------------------------------------------------------------------
def ordinal(n: int) -> str:
    if 11 <= n % 100 <= 13:
        return f"{n}th"
    return f"{n}{ {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th') }"


def clean(v):
    """Turn any Excel cell value into a tidy string ('' if empty)."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, (dt.datetime, dt.date)):
        return v
    return str(v).strip()


def as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str) and v:
        for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d %B %Y", "%d %b %Y", "%d.%m.%Y"):
            try:
                return dt.datetime.strptime(v, fmt).date()
            except ValueError:
                pass
    return None


def fmt_date(v):
    d = as_date(v)
    return d.strftime("%d/%m/%Y") if d else (str(v).strip() if v not in ("", None) else "")


# ----------------------------------------------------------------------------
# drawing
# ----------------------------------------------------------------------------
class Overlay:
    def __init__(self, page_sizes):
        self.page_sizes = page_sizes
        self.total_pages = len(page_sizes)
        self.buf = io.BytesIO()
        self.c = canvas.Canvas(self.buf, pagesize=page_sizes[0])
        self.page = 1
        self.warnings = []

    @property
    def h(self):
        return self.page_sizes[self.page - 1][1]

    def goto(self, page):
        while self.page < page:
            self.c.showPage()
            self.page += 1
            self.c.setPageSize(self.page_sizes[self.page - 1])

    def text(self, page, key, value, x0, x1, bottom, label_size, size):
        if value in ("", None):
            return False
        self.goto(page)
        value = str(value)
        avail = x1 - x0
        s = size
        while stringWidth(value, FONT, s) > avail and s > 5.0:
            s -= 0.25
        if stringWidth(value, FONT, s) > avail:
            self.warnings.append(f"'{key}' is too long for its space on the form – shorten it.")
        lift = 1.0 if page == 1 else (3.5 if page == 2 else 2.5)
        baseline_top = bottom - 0.22 * label_size - lift
        self.c.setFont(FONT, s)
        self.c.setFillColorRGB(0, 0, 0.5)
        self.c.drawString(x0, self.h - baseline_top, value)
        return True

    def tick(self, name):
        page, x0, top, x1, bottom = BOXES[name]
        self.goto(page)
        pad = 1.2 if page == 1 else 2.2
        lw = 1.2 if page == 1 else 1.5
        c = self.c
        c.setLineWidth(lw)
        c.setStrokeColorRGB(0, 0, 0)
        c.line(x0 + pad, self.h - (top + pad), x1 - pad, self.h - (bottom - pad))
        c.line(x0 + pad, self.h - (bottom - pad), x1 - pad, self.h - (top + pad))

    def comb(self, box, number, label):
        """Write a 2-digit number into a two-cell box."""
        x0, x1, bottom = box
        digits = f"{number:02d}"
        cell = (x1 - x0) / 2
        self.goto(1)
        self.c.setFont(FONT, 7)
        for i, ch in enumerate(digits):
            cx = x0 + cell * i + cell / 2
            self.c.drawCentredString(cx, self.h - (bottom - 1.6), ch)

    def finish(self):
        while self.page < self.total_pages:
            self.c.showPage()
            self.page += 1
            self.c.setPageSize(self.page_sizes[self.page - 1])
        self.c.drawString(0, 0, " ")  # guarantees the last page is kept even if nothing was drawn on it
        self.c.save()
        self.buf.seek(0)
        return PdfReader(self.buf)


# ----------------------------------------------------------------------------
# READING YOUR WORKBOOK
# ----------------------------------------------------------------------------
AF = "Application Form"
TC = "Terms & Conditions"
AU = "Authorization form"

TEXT_CELLS = {
    # Sheet 1: Application Form
    "industry": (AF, "B6"), "physical_address": (AF, "B7"), "postal_address": (AF, "B8"),
    "telephone": (AF, "B9"), "office_email": (AF, "B10"), "vat_number": (AF, "B11"),
    "business_type": (AF, "B12"), "trading_years": (AF, "B13"), "trading_months": (AF, "D13"),
    "proposed_account_name": (AF, "F13"),
    "collection_reason": (AF, "D20"), "disbursement_reason": (AF, "D21"),
    "contact_name": (AF, "B24"), "contact_email": (AF, "B25"), "contact_tel": (AF, "E25"),
    "finance_name": (AF, "B26"), "finance_email": (AF, "B27"), "finance_tel": (AF, "E27"),
    "admin_name": (AF, "B30"), "admin_tel": (AF, "E30"), "admin_username": (AF, "B31"),
    "admin_email": (AF, "E31"), "admin_dob": (AF, "B32"), "admin_id": (AF, "E32"),
    "mpesa_name": (AF, "B36"), "mpesa_mobile": (AF, "B37"), "mpesa_id": (AF, "E37"),
    "bank_name": (AF, "B40"), "bank_branch": (AF, "E40"),
    "bank_account_name": (AF, "B41"), "bank_account_number": (AF, "E41"),
    "ussd_mobile": (AF, "B44"), "ussd_id": (AF, "E44"),
    "sign_day": (AF, "B48"), "sign_month_year": (AF, "D48"), "sign_location": (AF, "F48"),
    "sig1_name": (AF, "B49"), "sig1_designation": (AF, "D49"),
    "sig2_name": (AF, "B50"), "sig2_designation": (AF, "D50"),

    # Sheet 2: Terms & Conditions
    "company_name": (TC, "B8"),
    "t_s1_name": (TC, "B12"), "t_s1_date": (TC, "B14"),
    "t_s2_name": (TC, "B16"), "t_s2_address": (TC, "B17"), "t_s2_date": (TC, "B19"),

    # Sheet 3: Authorization form
    "auth_org_name": (AU, "C5"),
    "auth_reg_office": (AU, "C6"),
    "auth_day": (AU, "B7"),
    "auth_month": (AU, "D7"),
    "auth_year": (AU, "F7"),
    "auth_person_name": (AU, "C16"),
    "auth_purpose": (AU, "C20"),
    "auth_bank_name": (AU, "B23"),
    "auth_bank_branch": (AU, "D23"),
    "auth_account_name": (AU, "B24"),
    "auth_account_number": (AU, "D24"),
    "auth_nominated_mobile": (AU, "E26"),
    "auth_s1_name": (AU, "B31"),
    "auth_s1_date": (AU, "B32"),
    "auth_s2_name": (AU, "E31"),
    "auth_s2_date": (AU, "E32"),
}

TICK_CELLS = {
    # Sheet 1: Application Form
    "PayBill": (AF, "B16"),
    "Buy Goods": (AF, "D16"),
    "Aggregated PayBill": (AF, "F16"),
    "Business Till": (AF, "B17"),
    "collection": (AF, "B20"),
    "disbursement": (AF, "B21"),
    "M-PESA": (AF, "A35"),
    "Bank Account": (AF, "A39"),
    "ussd_yes": (AF, "C43"),
    "ussd_no": (AF, "E43"),

    # Sheet 3: Authorization form
    "auth_paybill_b2c": (AU, "A11"),
    "auth_business_till": (AU, "C11"),
    "auth_buy_goods": (AU, "A12"),
    "auth_bulk_payment": (AU, "C12"),
    "auth_paybill": (AU, "A13"),
    "auth_sure_pay": (AU, "C13"),
}


def is_ticked(v) -> bool:
    if isinstance(v, bool):
        return v
    return str(v).strip().lower() in ("\u2612", "\u2611", "\u2713", "\u2714", "x", "yes", "true", "1")


def normalize_sheet_name(name):
    return " ".join(str(name).strip().split()).casefold()


def load_workbook_data(xlsx: Path):
    wb = load_workbook(xlsx, data_only=True)
    sheets = {normalize_sheet_name(name): name for name in wb.sheetnames}
    missing = [name for name in (AF, TC, AU) if normalize_sheet_name(name) not in sheets]
    if missing:
        missing_names = ", ".join(f"'{n}'" for n in missing)
        sys.exit(f"Missing required sheets in {xlsx.name}: {missing_names}. Please don't rename the sheets.")

    values = {k: clean(wb[sheets[normalize_sheet_name(sh)]][addr].value) for k, (sh, addr) in TEXT_CELLS.items()}
    ticks = {k: is_ticked(wb[sheets[normalize_sheet_name(sh)]][addr].value) for k, (sh, addr) in TICK_CELLS.items()}
    return values, ticks


def find_workbook() -> Path:
    cands = [p for p in sorted(BASE.glob("*.xlsx")) if not p.name.startswith("~$")]
    for p in cands:
        try:
            wb = load_workbook(p, read_only=True)
            if any(normalize_sheet_name(AF) == normalize_sheet_name(name) for name in wb.sheetnames):
                return p
        except Exception:
            continue
    sys.exit("Could not find the Excel workbook (with an 'Application Form' sheet) in this folder.")


def month_year(v):
    """'November 2026' (text) or a real date -> ('November', '26')."""
    d = as_date(v)
    if d:
        return d.strftime("%B"), d.strftime("%y")
    txt = str(v).strip()
    if not txt:
        return "", ""
    m = re.search(r"(\d{4}|\d{2})\s*$", txt)
    year = m.group(1)[-2:] if m else ""
    month = re.sub(r"[\d/\-.,]+", " ", txt).split()
    return (month[0].title() if month else ""), year


def day_text(v):
    d = as_date(v) if not isinstance(v, str) or "/" in v or "-" in v else None
    if d:
        return ordinal(d.day)
    txt = str(v).strip()
    return ordinal(int(txt)) if txt.isdigit() else txt


def short_year(v):
    """Turn 2026 or 26 into '26' for the '20__' field."""
    d = as_date(v)
    if d:
        return d.strftime("%y")
    txt = str(v).strip()
    if not txt:
        return ""
    m = re.search(r"(\d{2})\s*$", txt)
    return m.group(1) if m else txt


def compose_auth_date(day_v, month_v, year_v):
    """Build DD/MM/YYYY from B7, D7, F7 if all three are provided."""
    if day_v == "" or month_v == "" or year_v == "":
        return ""
    try:
        d_num = int(re.sub(r"\D", "", str(day_v)))
        yr_str = short_year(year_v)
        yr_num = int(f"20{yr_str}") if len(yr_str) == 2 else int(yr_str)
        m_txt = str(month_v).strip()
        if m_txt.isdigit():
            m_num = int(m_txt)
        else:
            m_num = dt.datetime.strptime(m_txt[:3].title(), "%b").month
        return dt.date(yr_num, m_num, d_num).strftime("%d/%m/%Y")
    except Exception:
        return ""


# ----------------------------------------------------------------------------
# main logic
# ----------------------------------------------------------------------------
def build(values: dict, ticks: dict, blank_pdf: Path, out_pdf: Path):
    reader = PdfReader(str(blank_pdf))
    page_sizes = [(float(p.mediabox.width), float(p.mediabox.height)) for p in reader.pages]
    ov = Overlay(page_sizes)
    filled, notes = [], []

    def put(field_key, value, tag=None):
        spec = TEXT_FIELDS[field_key]
        if ov.text(spec[0], tag or field_key, value, *spec[1:]):
            filled.append(tag or field_key)

    # --- Page 1: plain text fields
    direct = [
        "industry", "physical_address", "postal_address", "telephone", "office_email",
        "vat_number", "business_type", "proposed_account_name",
        "collection_reason", "disbursement_reason",
        "contact_name", "contact_email", "contact_tel",
        "finance_name", "finance_email", "finance_tel",
        "admin_name", "admin_tel", "admin_username", "admin_email", "admin_id",
        "ussd_mobile", "ussd_id", "sign_location",
        "sig1_name", "sig1_designation", "sig2_name", "sig2_designation",
    ]
    for k in direct:
        put(k, values.get(k, ""))
    put("admin_dob", fmt_date(values.get("admin_dob", "")), "admin_dob")

    # --- Page 1: trading period
    for key, box_ in (("trading_years", YEARS_BOX), ("trading_months", MONTHS_BOX)):
        v = values.get(key, "")
        if v != "":
            try:
                n = int(float(v))
                if not 0 <= n <= 99:
                    raise ValueError
                ov.comb(box_, n, key)
                filled.append(key)
            except ValueError:
                notes.append(f"'{key}' must be a whole number from 0 to 99 - skipped.")

    # --- Page 1: ticks
    for name in ("PayBill", "Buy Goods", "Aggregated PayBill", "Business Till",
                 "collection", "disbursement", "M-PESA", "Bank Account", "ussd_yes", "ussd_no"):
        if ticks.get(name):
            ov.tick(name)
            filled.append(f"tick:{name}")

    # --- Page 1: sanity checks (warnings only)
    if not any(ticks[n] for n in ("PayBill", "Buy Goods", "Aggregated PayBill", "Business Till")):
        notes.append("No service type ticked on Page 1 (PayBill / Buy Goods / Aggregated PayBill / Business Till).")
    if values.get("collection_reason") and not ticks["collection"]:
        notes.append("A collection reason is written but 'Collection of funds' is not ticked.")
    if values.get("disbursement_reason") and not ticks["disbursement"]:
        notes.append("A disbursement reason is written but 'Disbursement of funds' is not ticked.")
    if ticks["M-PESA"] and ticks["Bank Account"]:
        notes.append("Both M-PESA and Bank Account are ticked - the form says tick one.")
    if not ticks["M-PESA"] and not ticks["Bank Account"]:
        notes.append("Neither M-PESA nor Bank Account is ticked under Payment Details.")
    if ticks["ussd_yes"] and ticks["ussd_no"]:
        notes.append("Both Yes and No are ticked for the USSD nomination.")

    # --- Page 1: payment detail blocks follow the tick
    mp_keys = ("mpesa_name", "mpesa_mobile", "mpesa_id")
    bk_keys = ("bank_name", "bank_branch", "bank_account_name", "bank_account_number")
    if ticks["M-PESA"]:
        for k in mp_keys:
            put(k, values.get(k, ""))
    elif any(values.get(k) for k in mp_keys):
        notes.append("M-PESA details are filled but the M-PESA box is not ticked - they were skipped.")
    if ticks["Bank Account"]:
        for k in bk_keys:
            put(k, values.get(k, ""))
    elif any(values.get(k) for k in bk_keys):
        notes.append("Bank details are filled but the Bank Account box is not ticked - they were skipped.")

    # --- Page 1: declaration: "Signed this [14th] Day of: [November] 20[26]"
    if values.get("sign_day"):
        put("sign_day", day_text(values["sign_day"]))
    month, year = month_year(values.get("sign_month_year", ""))
    if month:
        put("sign_month", month)
    if year:
        put("sign_year", year)
    elif values.get("sign_month_year"):
        notes.append("Couldn't read the year in 'Month / Year' - write it like 'November 2026'.")

    # --- Page 2: Terms & Conditions
    company = values.get("company_name", "")
    for suffix in (" limited", " ltd", " ltd."):
        if company.lower().endswith(suffix):
            company = company[: -len(suffix)].rstrip()
            notes.append("'Limited' removed from the company name on page 2 because the form already prints it.")
            break
    put("p2_company", company, "company name (page 2)")
    s1 = values.get("t_s1_name") or values.get("sig1_name", "")
    s2 = values.get("t_s2_name") or values.get("sig2_name", "")
    put("p2_s1_name", s1, "signatory 1 name (page 2)")
    put("p2_s2_name", s2, "signatory 2 name (page 2)")
    put("p2_s2_address", values.get("t_s2_address", ""), "signatory 2 address (page 2)")
    for src, dst, label in (("t_s1_date", "p2_s1_date", "signatory 1 date (page 2)"),
                            ("t_s2_date", "p2_s2_date", "signatory 2 date (page 2)")):
        v = values.get(src, "")
        d = as_date(v)
        if d:
            put(dst, f"{d.day} {d.strftime('%B %Y')}", label)
        elif v:
            put(dst, str(v), label)

    # --- Page 3: M-PESA Account Opening Authorization Form
    if len(reader.pages) < 3:
        notes.append("The blank PDF has fewer than 3 pages - Authorization Form (page 3) was skipped.")
    else:
        put("auth_org_name", values.get("auth_org_name", ""), "organization name (page 3)")
        put("auth_reg_office", values.get("auth_reg_office", ""), "registered office (page 3)")

        if values.get("auth_day") != "":
            put("auth_day", day_text(values["auth_day"]), "meeting day (page 3)")
        if values.get("auth_month") != "":
            m_val = values["auth_month"]
            d_m = as_date(m_val)
            m_str = d_m.strftime("%B") if d_m else str(m_val).strip()
            put("auth_month", m_str, "meeting month (page 3)")
        if values.get("auth_year") != "":
            put("auth_year", short_year(values["auth_year"]), "meeting year (page 3)")

        auth_boxes = (
            "auth_paybill_b2c", "auth_business_till", "auth_buy_goods",
            "auth_bulk_payment", "auth_paybill", "auth_sure_pay",
        )
        for name in auth_boxes:
            if ticks.get(name):
                ov.tick(name)
                filled.append(f"tick:{name}")
        if not any(ticks[n] for n in auth_boxes):
            notes.append("No service type ticked on Page 3 (Authorization Form).")

        put("auth_person_name", values.get("auth_person_name", ""), "authorized person (page 3)")
        put("auth_purpose", values.get("auth_purpose", ""), "account purpose (page 3)")
        put("auth_bank_name", values.get("auth_bank_name", ""), "bank name (page 3)")
        put("auth_bank_branch", values.get("auth_bank_branch", ""), "bank branch (page 3)")
        put("auth_account_name", values.get("auth_account_name", ""), "bank account name (page 3)")
        put("auth_account_number", values.get("auth_account_number", ""), "bank account number (page 3)")
        put("auth_nominated_mobile", values.get("auth_nominated_mobile", ""), "nominated mobile (page 3)")

        # Signatories on Page 3 (falls back to Page 1 signatories & meeting date if left blank on Sheet 3)
        default_auth_date = compose_auth_date(
            values.get("auth_day", ""),
            values.get("auth_month", ""),
            values.get("auth_year", ""),
        )
        a_s1_name = values.get("auth_s1_name") or values.get("sig1_name", "")
        a_s2_name = values.get("auth_s2_name") or values.get("sig2_name", "")
        a_s1_date = fmt_date(values.get("auth_s1_date")) or default_auth_date
        a_s2_date = fmt_date(values.get("auth_s2_date")) or (default_auth_date if a_s2_name else "")

        put("auth_s1_name", a_s1_name, "signatory 1 name (page 3)")
        put("auth_s1_date", a_s1_date, "signatory 1 date (page 3)")
        put("auth_s2_name", a_s2_name, "signatory 2 name (page 3)")
        put("auth_s2_date", a_s2_date, "signatory 2 date (page 3)")

    # --- merge overlay onto the blank form
    overlay = ov.finish()
    writer = PdfWriter()
    for i, page in enumerate(reader.pages):
        new_page = writer.add_page(page)
        if i < len(overlay.pages):
            new_page.merge_page(overlay.pages[i])
    try:
        with open(out_pdf, "wb") as f:
            writer.write(f)
    except PermissionError:
        out_pdf = out_pdf.with_name(out_pdf.stem + dt.datetime.now().strftime("_%H%M%S") + ".pdf")
        with open(out_pdf, "wb") as f:
            writer.write(f)
        notes.append("The previous PDF was open, so a new file name was used.")
    return out_pdf, filled, ov.warnings + notes


def find_blank_pdf() -> Path:
    for p in sorted(BASE.glob("*.pdf")):
        n = p.name.lower()
        if "filled" in n or "authorization" in n:
            continue
        if "mpesa" in n or "m-pesa" in n:
            return p
    sys.exit("Could not find the blank M-PESA application form PDF in this folder.")


def main():
    ap = argparse.ArgumentParser(description="Fill the M-PESA C2B application form (Pages 1-3) from Excel.")
    ap.add_argument("--xlsx", default=None, help="Workbook (auto-detected if omitted)")
    ap.add_argument("--pdf", default=None, help="Blank form (auto-detected if omitted)")
    ap.add_argument("--out", default=str(BASE / "MPESA_C2B_Application_FILLED.pdf"))
    ap.add_argument("--open", action="store_true", help="Open the PDF when finished (Windows)")
    a = ap.parse_args()

    xlsx = Path(a.xlsx) if a.xlsx else find_workbook()
    if not xlsx.exists():
        sys.exit(f"Excel file not found: {xlsx}")
    blank = Path(a.pdf) if a.pdf else find_blank_pdf()

    values, ticks = load_workbook_data(xlsx)
    out, filled, notes = build(values, ticks, blank, Path(a.out))

    print(f"\nDone. Filled {len(filled)} items -> {out}")
    for n in notes:
        print("  ! " + n)
    if a.open and sys.platform.startswith("win"):
        os.startfile(out)  # noqa


if __name__ == "__main__":
    main()