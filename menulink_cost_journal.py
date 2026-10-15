#!/usr/bin/env python3
"""
MenuLink 2.05 Cost Journal → Categorized Excel Breakdown
=========================================================
Parses a MenuLink Cost Journal (Receiving by Vendor) PDF and generates
a formatted .xlsx workbook — same structure as the GFS PO breakdown.

Key difference from GFS script: category codes are EMBEDDED in the 2.05
report (e.g., "01 COS-MEAT", "ECOLAB CHEMICALS"), so no keyword guessing.

Usage:
    python menulink_cost_journal.py <input_pdf> [output_xlsx]

Requirements:
    pip install pdfplumber openpyxl
"""

import re
import sys
import os
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

import pdfplumber

# Emoji in console output: avoid UnicodeEncodeError on Windows code pages.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================================
# DATA STRUCTURES
# ============================================================================

@dataclass
class LineItem:
    description: str
    category_raw: str       # Category string exactly as it appears in the PDF
    qty: float
    units: str
    unit_cost: float
    extension: float        # Negative for credits/returns
    vendor: str = ""
    invoice: str = ""

    @property
    def is_credit(self):
        return self.extension < 0


# ============================================================================
# CATEGORY MAPPING
# ============================================================================
# Maps raw category strings from PDF → internal bucket keys.
# Order matters for the KNOWN_CATEGORIES list (longest-match first).

CATEGORY_MAP = {
    # ----- F&B -----
    "01 COS-MEAT":          "meat_seafood",
    "03 COS-POULTRY":       "poultry",
    "04 COS-PRODUCE":       "produce_frozen",
    "05 COS-DAIRY":         "dairy_eggs",
    "07 COS-BAKERY":        "dry_goods",
    "08 COS-DRESSING":      "sauces",
    "09 COS-DESSERT":       "desserts",
    "10 COS-SOUP":          "prepared",
    "11 COS-SHORTENING":    "oils",
    "12 COS-SALSA":         "sauces",       # merged into sauces bucket
    "13 COS-OTHER FOOD":    "other_food",
    "14 COS-N/A BEVERAGE":  "beverages",
    # ----- Alcohol -----
    "15 COS-LIQUOR":        "liquor",
    "16 COS-BEER-BT":       "beer_bottle",
    "16 COS-BEER-KG":       "beer_keg",
    "17 COS-WINE":          "wine",
    "18 COS-BEER KEG DEPOSIT": "keg_deposit",
    "18 COS-BEER KEG":          "keg_deposit",   # pdfplumber wraps "DEPOSIT" to next line
    "BEER KEG RETURNS":     "keg_deposit",
    # ----- Non-F&B -----
    "SUPPLIES / FACILITIES": "misc",
    "TO/GO-IJW":            "togo",         # must be before plain "TO/GO"
    "TO/GO":                "togo",
    "ECOLAB CHEMICALS":     "chemicals",
    "PLATEWARE":            "plateware",
    "GLASSWARE":            "glassware",
    "FLATWARE":             "flatware",
    # ----- Excluded expenses (tracked but not in F&B) -----
    "FUEL SURCHARGE":       "non_food",
    "OTHER MAINTENANCE":    "non_food",
    "SALES TAX":            "non_food",
}

# Sorted longest-first so "TO/GO-IJW" matches before "TO/GO", etc.
KNOWN_CATEGORIES = sorted(CATEGORY_MAP.keys(), key=len, reverse=True)

# ---- Display config for each sub-bucket ----

FB_SUBCATEGORIES = [
    ("meat_seafood",   "MEAT & SEAFOOD",                  "#1F6B2E", "#E2EFDA"),
    ("poultry",        "POULTRY",                          "#375623", "#EBF1DE"),
    ("dairy_eggs",     "DAIRY & EGGS",                     "#4E7B2A", "#F0F7E8"),
    ("produce_frozen", "PRODUCE & FROZEN VEG",             "#1F6B2E", "#E2EFDA"),
    ("prepared",       "PREPARED FOODS & SOUPS",           "#375623", "#EBF1DE"),
    ("sauces",         "SAUCES, DRESSINGS & CONDIMENTS",   "#1F6B2E", "#E2EFDA"),
    ("other_food",     "OTHER FOOD",                       "#4E7B2A", "#F0F7E8"),
    ("dry_goods",      "DRY GOODS & BAKERY",               "#375623", "#EBF1DE"),
    ("desserts",       "DESSERTS & BAKING",                "#1F6B2E", "#E2EFDA"),
    ("oils",           "OILS & SHORTENING",                "#4E7B2A", "#F0F7E8"),
    ("beverages",      "N/A BEVERAGES",                    "#375623", "#EBF1DE"),
]

BEER_LIQUOR_SUBCATEGORIES = [
    ("liquor",      "LIQUOR & SPIRITS",      "#7B3F00", "#FDF0E0"),
    ("beer_bottle", "BEER — BOTTLE / CAN",   "#7B3F00", "#FDF0E0"),
    ("beer_keg",    "BEER — DRAFT / KEG",    "#7B3F00", "#FDF0E0"),
    ("wine",        "WINE",                  "#7B3F00", "#FDF0E0"),
    ("keg_deposit", "KEG DEPOSITS / RETURNS","#808080", "#F2F2F2"),
]

MISC_SUBCATEGORIES = [
    ("misc",      "SUPPLIES / FACILITIES", "#833C00", "#FCE4D6"),
    ("plateware", "PLATEWARE",             "#833C00", "#FCE4D6"),
    ("glassware", "GLASSWARE",             "#833C00", "#FCE4D6"),
    ("flatware",  "FLATWARE",              "#833C00", "#FCE4D6"),
]


# ============================================================================
# PDF PARSING
# ============================================================================

# Tail pattern: qty  units  $unit_cost  $extension (handles credits in parens)
TAIL_RE = re.compile(
    r"^([-\d.]+)\s+(.+?)\s+\$([\d,]+\.\d{2})\s+(\(?\$[\d,]+\.\d{2}\)?)$"
)

SKIP_LINES = {
    "Description", "Category", "Amount", "Received", "Purchase",
    "Units", "Cost/Unit", "Extension", "COST DISTRIBUTION TOTALS",
}


def clean_desc(text: str) -> str:
    """Strip X-/T-/Y- prefixes and extra whitespace."""
    return re.sub(r"^[XTY]-\s*", "", text.strip())


def parse_ext(s: str) -> float:
    """Parse an extension string like '$1,234.56' or '($23.83)' → float."""
    negative = s.startswith("(")
    val = float(s.replace("(", "").replace(")", "").replace("$", "").replace(",", ""))
    return -val if negative else val


def parse_pdf(pdf_path: str):
    """
    Returns (items, metadata, cost_distribution).
      items             — list[LineItem]
      metadata          — dict with store/period/week/date_range/invoice_total
      cost_distribution — dict {category_label: total} from the summary page
    """
    metadata = {
        "store": os.environ.get("GFS_STORE_NAME", "Store"),
        "period": "",
        "week": "",
        "date_range": "",
        "invoice_total": 0.0,
    }
    cost_distribution = {}
    items = []

    with pdfplumber.open(pdf_path) as pdf:
        all_lines = []
        for page in pdf.pages:
            text = page.extract_text() or ""
            all_lines.extend(text.split("\n"))

    # ---- metadata pass ----
    for line in all_lines[:40]:
        m = re.search(r"Period:\s*(\d+)\s+Week:\s*(\d+)", line)
        if m:
            metadata["period"] = m.group(1)
            metadata["week"] = m.group(2)
        m = re.search(r"(\d{1,2}/\d{1,2}/\d{4})\s*[-–]\s*(\d{1,2}/\d{1,2}/\d{4})", line)
        if m:
            metadata["date_range"] = f"{m.group(1)} – {m.group(2)}"
        m = re.search(r"(C\d+[-–].+?)\s+Fiscal Year:", line)
        if m:
            metadata["store"] = m.group(1).strip()

    # ---- cost distribution pass ----
    in_dist = False
    for line in all_lines:
        if "COST DISTRIBUTION TOTALS" in line:
            in_dist = True
            continue
        if "Invoice Total" in line:
            m = re.search(r"Invoice Total\s+\$([\d,]+\.\d{2})", line)
            if m:
                metadata["invoice_total"] = float(m.group(1).replace(",", ""))
        if in_dist:
            m = re.search(r"G/L Account:\s+(.+?)\s+([-]?[\d,]+\.\d{2})$", line)
            if m:
                cost_distribution[m.group(1).strip()] = float(m.group(2).replace(",", ""))

    # ---- line item pass ----
    current_vendor = ""
    current_invoice = ""
    desc_buffer = []

    # Pre-compile a pattern that finds any known category in a line
    cat_alt = "|".join(re.escape(c) for c in KNOWN_CATEGORIES)
    CAT_RE = re.compile(r"(" + cat_alt + r")\s+([-\d].+)$")

    for line in all_lines:
        stripped = line.strip()

        # ---- vendor / invoice tracking ----
        m = re.match(r"Vendor:\s*\d+\s+(.+?)\s+\d+\s*$", stripped)
        if m:
            current_vendor = m.group(1).strip()
            desc_buffer.clear()
            continue

        m = re.match(r"Invoice:\s*(\S+)", stripped)
        if m:
            current_invoice = m.group(1)
            desc_buffer.clear()
            continue

        # ---- skip header / footer / admin lines ----
        if not stripped:
            continue
        if any(stripped.startswith(s) for s in (
            "***", "End:", "Memo:", "Site:", "V 21", "Copyright",
            "Description", "Amount\nReceived",
            "2.05 - Cost Journal",
        )) or re.match(r"C\d{4}.*Fiscal", stripped):
            continue
        if stripped in SKIP_LINES:
            continue
        if re.match(r"^\d{5}\.\d{3}", stripped):   # G/L account numbers
            continue
        if re.match(r"G/L Account:", stripped):
            continue
        if "COST DISTRIBUTION TOTALS" in stripped:
            continue
        if re.match(r"Invoice Total", stripped):
            continue
        if re.match(r"\d{1,2}/\d{1,2}/\d{4}", stripped):
            continue

        # ---- try to match a category anchor ----
        cat_m = CAT_RE.search(stripped)
        if cat_m:
            cat_raw = cat_m.group(1)
            tail = cat_m.group(2).strip()
            pre_cat = stripped[: cat_m.start()].strip()

            # Build full description
            desc_parts = [p for p in desc_buffer if p] + ([pre_cat] if pre_cat else [])
            full_desc = clean_desc(" ".join(desc_parts))
            desc_buffer.clear()

            # Parse tail: qty units $unit_cost $extension
            tail_m = TAIL_RE.match(tail)
            if tail_m and cat_raw in CATEGORY_MAP:
                qty = float(tail_m.group(1))
                units = tail_m.group(2).strip()
                unit_cost = float(tail_m.group(3).replace(",", ""))
                extension = parse_ext(tail_m.group(4))

                if full_desc:
                    items.append(LineItem(
                        description=full_desc,
                        category_raw=cat_raw,
                        qty=qty,
                        units=units,
                        unit_cost=unit_cost,
                        extension=extension,
                        vendor=current_vendor,
                        invoice=current_invoice,
                    ))
            continue

        # ---- possible description continuation ----
        # Reject lines that clearly aren't description text
        if "$" in stripped:
            # One more attempt: find qty+cost at the end, then scan the
            # remaining prefix for a known category.  This handles cases like
            #   "MISC-OTHER MAINTENANCE OTHER 1.00 DR $100.00 $100.00"
            # where the category string appears inside the description.
            fb_m = re.search(
                r"([-\d.]+)\s+(\S+(?:\s+\S+){0,3}?)\s+\$([\d,]+\.\d{2})\s+(\(?\$[\d,]+\.\d{2}\)?)$",
                stripped
            )
            if fb_m:
                prefix = stripped[: fb_m.start()].strip()
                # Find the rightmost known category in the prefix
                best_cat, best_pos = None, -1
                for cat_str in KNOWN_CATEGORIES:
                    pos = prefix.rfind(cat_str)
                    if pos >= 0 and pos > best_pos:
                        best_cat, best_pos = cat_str, pos
                if best_cat and best_cat in CATEGORY_MAP:
                    pre_cat = prefix[:best_pos].strip()
                    desc_parts = [p for p in desc_buffer if p] + ([pre_cat] if pre_cat else [])
                    full_desc = clean_desc(" ".join(desc_parts))
                    desc_buffer.clear()
                    qty = float(fb_m.group(1))
                    units = fb_m.group(2).strip()
                    unit_cost = float(fb_m.group(3).replace(",", ""))
                    extension = parse_ext(fb_m.group(4))
                    if full_desc and best_cat in CATEGORY_MAP:
                        items.append(LineItem(
                            description=full_desc,
                            category_raw=best_cat,
                            qty=qty,
                            units=units,
                            unit_cost=unit_cost,
                            extension=extension,
                            vendor=current_vendor,
                            invoice=current_invoice,
                        ))
                    continue
            desc_buffer.clear()
            continue
        if stripped.startswith("Invoice Date:") or stripped.startswith("Delivered"):
            desc_buffer.clear()
            continue

        desc_buffer.append(stripped)
        if len(desc_buffer) > 4:
            desc_buffer.pop(0)

    return items, metadata, cost_distribution


# ============================================================================
# CATEGORIZATION
# ============================================================================

def categorize_items(items: list) -> dict:
    """Returns {bucket_key: [LineItem, ...]}."""
    all_buckets = list(dict.fromkeys(CATEGORY_MAP.values()))  # preserve order, unique
    categories = {k: [] for k in all_buckets}

    for item in items:
        bucket = CATEGORY_MAP.get(item.category_raw)
        if bucket:
            categories[bucket].append(item)
        else:
            print(f"  ⚠ UNMAPPED: {item.category_raw!r}  —  {item.description}")

    return categories


# ============================================================================
# EXCEL HELPERS  (same style as gfs_breakdown.py)
# ============================================================================

THIN_BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)


def hex_(c): return c.lstrip("#")


def title_row(ws, row, text, bg, font_size=12):
    ws.merge_cells(f"A{row}:G{row}")
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(name="Arial", size=font_size, bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color=hex_(bg), fill_type="solid")
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
    ws.row_dimensions[row].height = 20
    return row + 1


def col_headers(ws, row):
    headers = ["#", "Description", "Vendor", "Qty / Units", "Unit Cost", "Total", "Credit?"]
    widths_ref = ["A", "B", "C", "D", "E", "F", "G"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = Font(name="Arial", size=9, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="595959", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER
    ws.row_dimensions[row].height = 16
    return row + 1


def data_row(ws, row, idx, item: LineItem, alt_bg):
    shade = alt_bg if idx % 2 == 0 else "FFFFFF"
    vals = [
        idx,
        item.description,
        item.vendor,
        f"{item.qty} {item.units}",
        item.unit_cost,
        item.extension,
        "CREDIT" if item.is_credit else "",
    ]
    fmts = [None, None, None, None, "$#,##0.00", "$#,##0.00", None]
    aligns = ["center", "left", "left", "left", "right", "right", "center"]

    for col, (v, fmt, align) in enumerate(zip(vals, fmts, aligns), 1):
        cell = ws.cell(row=row, column=col, value=v)
        cell.font = Font(
            name="Arial", size=9,
            color="CC0000" if item.is_credit else "000000"
        )
        cell.fill = PatternFill(start_color=hex_(shade) if shade != "FFFFFF" else "FFFFFF",
                                fill_type="solid")
        cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=False)
        cell.border = THIN_BORDER
        if fmt:
            cell.number_format = fmt
    ws.row_dimensions[row].height = 14
    return row + 1


def subtotal_row(ws, row, label, data_start, data_end, bg):
    """Write a subtotal row that sums column F."""
    ws.merge_cells(f"A{row}:E{row}")
    cell_label = ws.cell(row=row, column=1, value=label)
    cell_label.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    cell_label.fill = PatternFill(start_color=hex_(bg), fill_type="solid")
    cell_label.alignment = Alignment(horizontal="right", vertical="center")
    cell_label.border = THIN_BORDER

    if data_end >= data_start:
        formula = f"=SUM(F{data_start}:F{data_end})"
    else:
        formula = "=$0.00"

    cell_f = ws.cell(row=row, column=6, value=formula)
    cell_f.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    cell_f.fill = PatternFill(start_color=hex_(bg), fill_type="solid")
    cell_f.alignment = Alignment(horizontal="right", vertical="center")
    cell_f.number_format = "$#,##0.00"
    cell_f.border = THIN_BORDER

    # Blank G
    cell_g = ws.cell(row=row, column=7, value="")
    cell_g.fill = PatternFill(start_color=hex_(bg), fill_type="solid")
    cell_g.border = THIN_BORDER

    ws.row_dimensions[row].height = 16
    return row + 1


def grand_total_row(ws, row, label, subtotal_rows_f, bg):
    ws.merge_cells(f"A{row}:E{row}")
    cell = ws.cell(row=row, column=1, value=label)
    cell.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color=hex_(bg), fill_type="solid")
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell.border = THIN_BORDER

    formula = "=" + "+".join(f"F{r}" for r in subtotal_rows_f)
    cell_f = ws.cell(row=row, column=6, value=formula)
    cell_f.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    cell_f.fill = PatternFill(start_color=hex_(bg), fill_type="solid")
    cell_f.alignment = Alignment(horizontal="right", vertical="center")
    cell_f.number_format = "$#,##0.00"
    cell_f.border = THIN_BORDER

    cell_g = ws.cell(row=row, column=7, value="")
    cell_g.fill = PatternFill(start_color=hex_(bg), fill_type="solid")
    cell_g.border = THIN_BORDER

    ws.row_dimensions[row].height = 18
    return row + 1


def set_widths(ws):
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 11
    ws.column_dimensions["F"].width = 13
    ws.column_dimensions["G"].width = 8


def build_detail_sheet(wb, sheet_title, tab_color, sheet_label,
                       subcategories, categories, bg_title):
    """
    Build a detail sheet containing one or more sub-category sections.
    Returns (ws, grand_total_row_number).
    """
    ws = wb.create_sheet(sheet_title)
    ws.sheet_properties.tabColor = tab_color.lstrip("#")
    set_widths(ws)

    row = title_row(ws, 1, sheet_label, bg_title, font_size=12)
    subtotal_row_nums = []

    for cat_key, display_name, hdr_bg, alt_bg in subcategories:
        items_in_cat = categories.get(cat_key, [])
        if not items_in_cat:
            continue
        row = title_row(ws, row, display_name, hdr_bg, font_size=10)
        row = col_headers(ws, row)
        data_start = row
        for idx, item in enumerate(items_in_cat, 1):
            row = data_row(ws, row, idx, item, alt_bg)
        data_end = row - 1
        row = subtotal_row(ws, row, display_name, data_start, data_end, hdr_bg)
        subtotal_row_nums.append(row - 1)
        row += 1  # blank spacer

    if subtotal_row_nums:
        gt_row = row
        grand_total_row(ws, gt_row, f"TOTAL {sheet_title}", subtotal_row_nums, bg_title)
        total_row_num = gt_row
    else:
        # Empty sheet placeholder
        ws.cell(row=row, column=2, value=f"No {sheet_title} items in this report").font = Font(
            name="Arial", size=10, italic=True, color="999999"
        )
        total_row_num = row

    ws.freeze_panes = "A3"
    return ws, total_row_num


# ============================================================================
# WORKBOOK BUILDER
# ============================================================================

def build_workbook(categories: dict, metadata: dict, cost_distribution: dict):
    wb = Workbook()
    wb.remove(wb.active)  # We'll add SUMMARY first at position 0 at the end

    store = metadata.get("store", "Store")
    period = metadata.get("period", "?")
    week = metadata.get("week", "?")
    dr = metadata.get("date_range", "")
    pw_label = f"P{period}W{week}  |  {dr}"

    # ===== F&B SHEET =====
    ws_fb, fb_total_row = build_detail_sheet(
        wb,
        sheet_title="F&B",
        tab_color="#1F6B2E",
        sheet_label=f"FOOD & BEVERAGE  —  {store}  |  {pw_label}",
        subcategories=FB_SUBCATEGORIES,
        categories=categories,
        bg_title="#1F6B2E",
    )

    # ===== BEER & LIQUOR SHEET =====
    ws_beer, beer_total_row = build_detail_sheet(
        wb,
        sheet_title="BEER & LIQUOR",
        tab_color="#7B3F00",
        sheet_label=f"BEER & LIQUOR  —  {store}  |  {pw_label}",
        subcategories=BEER_LIQUOR_SUBCATEGORIES,
        categories=categories,
        bg_title="#7B3F00",
    )

    # ===== TO-GO SHEET =====
    ws_togo, togo_total_row = build_detail_sheet(
        wb,
        sheet_title="TOGO SUPPLIES",
        tab_color="#1F4E79",
        sheet_label=f"TO-GO SUPPLIES  —  {store}  |  {pw_label}",
        subcategories=[("togo", "TO-GO SUPPLIES", "#1F4E79", "#DEEAF1")],
        categories=categories,
        bg_title="#1F4E79",
    )

    # ===== MISC SUPPLIES SHEET =====
    ws_misc, misc_total_row = build_detail_sheet(
        wb,
        sheet_title="MISC SUPPLIES",
        tab_color="#833C00",
        sheet_label=f"MISC SUPPLIES  —  {store}  |  {pw_label}",
        subcategories=MISC_SUBCATEGORIES,
        categories=categories,
        bg_title="#833C00",
    )

    # ===== CHEMICALS SHEET =====
    ws_chem, chem_total_row = build_detail_sheet(
        wb,
        sheet_title="CHEMICALS",
        tab_color="#4B0082",
        sheet_label=f"CHEMICALS / ECOLAB  —  {store}  |  {pw_label}",
        subcategories=[("chemicals", "ECOLAB CHEMICALS", "#4B0082", "#EDE7F6")],
        categories=categories,
        bg_title="#4B0082",
    )

    # ===== SUMMARY SHEET =====
    ws_sum = wb.create_sheet("SUMMARY", 0)
    ws_sum.sheet_properties.tabColor = "C00000"
    ws_sum.column_dimensions["A"].width = 18
    ws_sum.column_dimensions["B"].width = 36
    ws_sum.column_dimensions["C"].width = 13
    ws_sum.column_dimensions["D"].width = 15
    ws_sum.column_dimensions["E"].width = 13

    r = 1
    ws_sum.merge_cells(f"A{r}:E{r}")
    c = ws_sum.cell(row=r, column=1,
                    value=f"{store.upper()}  —  COST JOURNAL  |  {pw_label}")
    c.font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    c.fill = PatternFill(start_color="C00000", fill_type="solid")
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws_sum.row_dimensions[r].height = 24
    r += 1

    ws_sum.merge_cells(f"A{r}:E{r}")
    c = ws_sum.cell(row=r, column=1, value="EXPENSE CATEGORY SUMMARY")
    c.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    c.fill = PatternFill(start_color="404040", fill_type="solid")
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws_sum.row_dimensions[r].height = 18
    r += 1

    # Column headers
    for col, hdr in enumerate(["Category", "Description", "# Items", "Total Cost", "% of Order"], 1):
        cell = ws_sum.cell(row=r, column=col, value=hdr)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="595959", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER
    r += 1

    # Row definitions: label, description, bg, item_count_formula, cost_formula
    # item_count uses COUNTA of column B on the detail sheet (non-header rows)
    cat_rows = [
        ("F&B",           "Food & Beverage Purchases",
         "E2EFDA", f"='F&B'!F{fb_total_row}"),
        ("Beer & Liquor",  "Beer, Liquor, Wine & Seltzer",
         "FDF0E0", f"='BEER & LIQUOR'!F{beer_total_row}"),
        ("To-Go Supplies", "To-Go Packaging & Containers",
         "DEEAF1", f"='TOGO SUPPLIES'!F{togo_total_row}"),
        ("Misc Supplies",  "Supplies, Plateware, Glassware, Flatware",
         "FCE4D6", f"='MISC SUPPLIES'!F{misc_total_row}"),
        ("Chemicals",      "Ecolab / Cleaning Chemicals",
         "EDE7F6", f"='CHEMICALS'!F{chem_total_row}"),
    ]

    data_start_r = r
    for label, desc, bg, cost_f in cat_rows:
        for col in range(1, 6):
            cell = ws_sum.cell(row=r, column=col)
            cell.fill = PatternFill(start_color=hex_(bg), fill_type="solid")
            cell.border = THIN_BORDER

        ws_sum.cell(row=r, column=1, value=label).font = Font(name="Arial", size=10, bold=True)
        ws_sum.cell(row=r, column=1).alignment = Alignment(horizontal="center")

        ws_sum.cell(row=r, column=2, value=desc).font = Font(name="Arial", size=10)
        ws_sum.cell(row=r, column=2).alignment = Alignment(horizontal="left")

        # Item count — just a static count for now; cost formula is the live link
        ws_sum.cell(row=r, column=3, value="").font = Font(name="Arial", size=10)
        ws_sum.cell(row=r, column=3).alignment = Alignment(horizontal="center")

        ws_sum.cell(row=r, column=4, value=cost_f).font = Font(name="Arial", size=10)
        ws_sum.cell(row=r, column=4).alignment = Alignment(horizontal="right")
        ws_sum.cell(row=r, column=4).number_format = "$#,##0.00"

        grand_r = data_start_r + len(cat_rows)
        ws_sum.cell(row=r, column=5,
                    value=f"=IFERROR(D{r}/D{grand_r},\"\")").font = Font(name="Arial", size=10)
        ws_sum.cell(row=r, column=5).alignment = Alignment(horizontal="center")
        ws_sum.cell(row=r, column=5).number_format = "0.0%"

        ws_sum.row_dimensions[r].height = 16
        r += 1

    # Grand total row
    grand_r = r
    for col in range(1, 6):
        cell = ws_sum.cell(row=grand_r, column=col)
        cell.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="C00000", fill_type="solid")
        cell.border = THIN_BORDER
    ws_sum.cell(row=grand_r, column=1, value="GRAND TOTAL").alignment = Alignment(horizontal="left")
    ws_sum.cell(row=grand_r, column=4,
                value=f"=SUM(D{data_start_r}:D{grand_r-1})").number_format = "$#,##0.00"
    ws_sum.cell(row=grand_r, column=4).alignment = Alignment(horizontal="right")
    ws_sum.cell(row=grand_r, column=5, value="100.0%").alignment = Alignment(horizontal="center")
    ws_sum.row_dimensions[grand_r].height = 20
    r += 2

    # ---- Cost Distribution reconciliation block (from the PDF's own totals page) ----
    if cost_distribution:
        r += 1
        ws_sum.merge_cells(f"A{r}:E{r}")
        c = ws_sum.cell(row=r, column=1, value="MENULINK COST DISTRIBUTION TOTALS  (source: PDF summary page)")
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill(start_color="404040", fill_type="solid")
        c.alignment = Alignment(horizontal="left")
        ws_sum.row_dimensions[r].height = 16
        r += 1

        for cat_label, total in cost_distribution.items():
            ws_sum.cell(row=r, column=1, value=cat_label).font = Font(name="Arial", size=9)
            ws_sum.cell(row=r, column=4, value=total).font = Font(name="Arial", size=9)
            ws_sum.cell(row=r, column=4).number_format = "$#,##0.00"
            ws_sum.cell(row=r, column=4).alignment = Alignment(horizontal="right")
            ws_sum.row_dimensions[r].height = 14
            r += 1

        # Invoice total from PDF
        if metadata["invoice_total"]:
            ws_sum.cell(row=r, column=1, value="Invoice Total (from PDF)").font = Font(
                name="Arial", size=9, bold=True)
            ws_sum.cell(row=r, column=4, value=metadata["invoice_total"]).font = Font(
                name="Arial", size=9, bold=True)
            ws_sum.cell(row=r, column=4).number_format = "$#,##0.00"
            ws_sum.cell(row=r, column=4).alignment = Alignment(horizontal="right")

    ws_sum.freeze_panes = "A4"
    return wb


# ============================================================================
# CONSOLE SUMMARY
# ============================================================================

def print_summary(categories, metadata, cost_distribution):
    store = metadata.get("store", "Store")
    period = metadata.get("period", "?")
    week = metadata.get("week", "?")
    dr = metadata.get("date_range", "")

    print(f"\n{'='*60}")
    print(f"   {store}  —  P{period}W{week}  |  {dr}")
    print(f"{'='*60}")

    fb_keys = [k for k, *_ in FB_SUBCATEGORIES]
    fb_total = sum(sum(i.extension for i in categories.get(k, [])) for k in fb_keys)
    beer_keys = [k for k, *_ in BEER_LIQUOR_SUBCATEGORIES]
    beer_total = sum(sum(i.extension for i in categories.get(k, [])) for k in beer_keys)

    print(f"\n   {'Category':<35} {'Items':>6} {'Cost':>12}")
    print(f"   {'-'*35} {'-'*6} {'-'*12}")

    for cat_key, display_name, *_ in FB_SUBCATEGORIES:
        items = categories.get(cat_key, [])
        if items:
            total = sum(i.extension for i in items)
            print(f"   {display_name:<35} {len(items):>6} ${total:>10,.2f}")
    print(f"   {'F&B TOTAL':<35} {'':>6} ${fb_total:>10,.2f}")
    print()

    for cat_key, display_name, *_ in BEER_LIQUOR_SUBCATEGORIES:
        items = categories.get(cat_key, [])
        if items:
            total = sum(i.extension for i in items)
            print(f"   {display_name:<35} {len(items):>6} ${total:>10,.2f}")
    if any(categories.get(k) for k in beer_keys):
        print(f"   {'BEER & LIQUOR TOTAL':<35} {'':>6} ${beer_total:>10,.2f}")
    print()

    for key, label in [
        ("togo", "TO-GO SUPPLIES"),
        ("misc", "MISC SUPPLIES"),
        ("plateware", "  PLATEWARE"),
        ("glassware", "  GLASSWARE"),
        ("flatware", "  FLATWARE"),
        ("chemicals", "CHEMICALS (ECOLAB)"),
    ]:
        items = categories.get(key, [])
        if items:
            total = sum(i.extension for i in items)
            print(f"   {label:<35} {len(items):>6} ${total:>10,.2f}")

    non_food_items = categories.get("non_food", [])
    if non_food_items:
        nf_total = sum(i.extension for i in non_food_items)
        print(f"\n   {'NON-FOOD (tax/surcharge/maint)':<35} {len(non_food_items):>6} ${nf_total:>10,.2f}")

    togo_total = sum(i.extension for i in categories.get("togo", []))
    misc_total = sum(
        sum(i.extension for i in categories.get(k, []))
        for k in ["misc", "plateware", "glassware", "flatware"]
    )
    chem_total = sum(i.extension for i in categories.get("chemicals", []))

    grand = fb_total + beer_total + togo_total + misc_total + chem_total
    print(f"   {'-'*35} {'-'*6} {'-'*12}")
    print(f"   {'GRAND TOTAL (excl. non-food)':<35} {'':>6} ${grand:>10,.2f}")

    if metadata["invoice_total"]:
        print(f"\n   Invoice Total (from PDF):  ${metadata['invoice_total']:>10,.2f}")
        nf = sum(i.extension for i in non_food_items) if non_food_items else 0
        print(f"   Our total + non-food:      ${grand + nf:>10,.2f}")


# ============================================================================
# MAIN
# ============================================================================

def main():
    if len(sys.argv) < 2:
        print("Usage: python menulink_cost_journal.py <input_pdf> [output_xlsx]")
        sys.exit(1)

    pdf_path = sys.argv[1]
    if not os.path.exists(pdf_path):
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)

    output_path = sys.argv[2] if len(sys.argv) >= 3 else \
        Path(pdf_path).stem + "_Breakdown.xlsx"

    print(f"📄 Parsing: {pdf_path}")
    items, metadata, cost_distribution = parse_pdf(pdf_path)

    pw = f"P{metadata['period']}W{metadata['week']}" if metadata['period'] else ""
    print(f"   Found {len(items)} line items  |  {pw}  |  {metadata['date_range']}")
    print(f"   Invoice total (PDF): ${metadata['invoice_total']:,.2f}")

    print("\n📦 Categorizing...")
    categories = categorize_items(items)

    print_summary(categories, metadata, cost_distribution)

    print("\n📊 Building workbook...")
    wb = build_workbook(categories, metadata, cost_distribution)
    wb.save(output_path)
    print(f"✅ Saved: {output_path}")
    print(f"   Sheets: {[ws.title for ws in wb.worksheets]}")


if __name__ == "__main__":
    main()
