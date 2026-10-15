#!/usr/bin/env python3
"""
GFS Purchase Order → Categorized Excel Breakdown
=================================================
Parses a Gordon Food Service PO PDF and generates a formatted .xlsx workbook
with a category breakdown (F&B, paper/to-go, chemicals, produce, beer/liquor, ...).

Usage:
    python gfs_breakdown.py <input_pdf> [output_xlsx]

    If output_xlsx is omitted, saves as "<input_pdf_stem>_Breakdown.xlsx"

Requirements:
    pip install pdfplumber openpyxl

Built for a restaurant's weekly GFS truck-ordering workflow.
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
    qty: int
    description: str
    unit_cost: float
    extension: float
    item_id: str = ""
    product_num: str = ""


@dataclass
class Category:
    name: str
    items: list = field(default_factory=list)

    @property
    def total_cases(self) -> int:
        return sum(i.qty for i in self.items)

    @property
    def total_cost(self) -> float:
        return round(sum(i.extension for i in self.items), 2)


# ============================================================================
# CATEGORY RULES
# ============================================================================
# Each rule is (category_key, list_of_keyword_matches).
# Order matters — first match wins. More specific rules go first.

CATEGORY_RULES = [
    # --- NON-F&B CATEGORIES (checked first) ---
    ("chemicals", [
        "DETERGENT", "CLEANER", "GREASE EXPRESS", "SANITIZER", "SOAP,",
        "DEGREASER", "RINSE, ADDITIVE",
    ]),
    ("togo", [
        "LID, 7 OZ AND 12 OZ", "LID, 7OZ AND 12OZ",
        "BAG, T-SHIRT", "BAG, SM PAPER TILE", "BAG, CHIP W/LOGO",
        "CONTAINER, TILE LOGO", "CUP, TILE LOGO",
        "UTENSILS, MEAL KIT", "UTENSIL, MEAL KIT",
        "CONTAINER, 9X9", "CONTAINER, 8X8 FOAM",
        "CONTAINER, 12OZ", "CONTAINER, 7OZ",
        "BAG, FLATWARE", "CARRIER, BEV",
        "LINER, CHIP TRAY",
        "CONTAINER, 7 OZ", "CONTAINER, 12 OZ", "CONTAINER, 6X6",
        "CUP, 12 OZ PLAS", "STRAW,", "LID, KID'S CUP",
        "BAG, LRG PAPER TILE", "BOWL PLAS", "CONTAINER, PLAS", "LID, PLAS",
    ]),
    ("misc", [
        "CUP, PORTION", "LID, PORTION",
        "TOWEL, PAPER", "TISSUE, TOILET",
        "GLOVE,", "LID, 12-24", "LID, FLAT FOR",
        "LINER, INSERT", "LINER, TRASH",
        "NAPKIN,", "TRAY,", "PVC FILM",
        "PAPER, PRINTER",
        "BAG, PORTION", "TOWELETTE,", "SUGAR,", "KID MENU,",
    ]),

    # --- F&B SUB-CATEGORIES ---
    ("poultry", [
        "CKN,", "CHICKEN,", "CHICKEN ",
    ]),
    ("meat_seafood", [
        "BEEF,", "BACON,", "STEAK,", "SAUSAGE,", "SALMON", "SHRIMP",
        "RIBS,", "BURGER", "PORK",
    ]),
    ("dairy_eggs", [
        "CHEESE,", "QUESO,", "BUTTERMILK", "SOUR CREAM", "MILK,",
        "EGG,", "EGGS,", "CREAM CHEESE",
    ]),
    ("produce_frozen", [
        "CORN,", "POTATOES,", "BEANS,", "STRAWBERR", "LETTUCE",
        "TOMATO", "ONION", "PEPPER,", "AVOCADO", "CELERY", "CARROT",
        "BROCCOLI", "SPINACH", "MUSHROOM", "CUCUMBER", "JALAPENO",
        "FRIES,", "ASPARAGUS", "LEMONS", "LIMES", "ORANGES", "CILANTRO",
    ]),
    ("beer_liquor", [
        "KEG,", "DRAFT,", "BEER,",
        "VODKA", "TEQUILA", "WHISKEY", "BOURBON", "RUM,", "GIN,",
        "SCOTCH", "MEZCAL", "BRANDY", "TRIPLE SEC", "KAHLUA",
        "WINE,", "WINE ", "CHARDONNAY", "MERLOT", "CABERNET",
        "SELTZER,", "HARD SELTZER", "WHITE CLAW", "TRULY,",
    ]),
    ("beverages", [
        "BIB,", "JUICE,", "WATER,", "TEA,", "LEMON, SOUR",
        "GRENADINE", "SYRUP,", "AGAVE", "COFFEE",
    ]),
    ("oils", [
        "OIL,", "SHORTENING,", "PAN SPRAY",
    ]),
    ("desserts", [
        "DOUGH, CHOC", "DOUGH, COOKIE", "CAKE,", "ICE CREAM",
        "COATING, CHOC", "CROUTONS", "FUDGE SAUCE", "BROWNIE",
        "PIE,", "COBBLER", "CARAMEL TOPPING",
    ]),
    ("sauces", [
        "DRESS,", "DRESSING,", "SAUCE,", "SALSA", "SEASONING,",
        "MUSTARD,", "KETCHUP,", "MAYONNAISE", "SPREAD,", "SALT,",
        "HONEY MUSTARD", "RANCH", "BBQ",
        "SWEET-N-SOUR", "PUREE,",
    ]),
    ("prepared", [
        "EGGROLLS", "EGGROLL", "MAC & CHEESE", "MACARONI",
        "BREAD, GARLIC", "PIZZA,", "SOUP,", "PICKLES,",
    ]),
    ("dry_goods", [
        "TORT,", "TORTILLA,", "BUN,", "PASTA,", "RICE,",
        "BATTER MIX", "BASE, MARINARA", "CHIPS,", "CHIP,",
        "FLOUR,", "BREAD", "CRACKER",
    ]),
]

# Display names and sheet colors for each category group
FB_SUBCATEGORIES = [
    ("meat_seafood",  "MEAT & SEAFOOD",                  "#1F6B2E", "#E2EFDA"),
    ("poultry",       "POULTRY",                          "#375623", "#EBF1DE"),
    ("dairy_eggs",    "DAIRY & EGGS",                     "#4E7B2A", "#F0F7E8"),
    ("produce_frozen","PRODUCE & FROZEN VEG",             "#1F6B2E", "#E2EFDA"),
    ("prepared",      "PREPARED FOODS & SOUPS",           "#375623", "#EBF1DE"),
    ("sauces",        "SAUCES, DRESSINGS & CONDIMENTS",   "#1F6B2E", "#E2EFDA"),
    ("dry_goods",     "DRY GOODS, BAKERY & STARCHES",     "#4E7B2A", "#F0F7E8"),
    ("desserts",      "DESSERTS & BAKING",                "#1F6B2E", "#E2EFDA"),
    ("oils",          "OILS & COOKING FATS",              "#1F6B2E", "#E2EFDA"),
    ("beverages",     "BEVERAGES",                        "#1F6B2E", "#E2EFDA"),
    ("beer_liquor",   "BEER & LIQUOR",                    "#7B3F00", "#FDF0E0"),
]


# ============================================================================
# PDF PARSING
# ============================================================================

def parse_pdf(pdf_path: str) -> tuple:
    """Parse a GFS PO PDF and return (line_items, metadata)."""
    items = []
    metadata = {"po_number": "", "date": "", "store": os.environ.get("GFS_STORE_NAME", "Store")}

    with pdfplumber.open(pdf_path) as pdf:
        all_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text += text + "\n"

    lines = all_text.split("\n")

    # Extract metadata — handles both receiving report and direct PO formats
    for line in lines[:25]:
        # Format A: "Purchase Order: 6439  3/1/2026"
        po_match = re.search(r"Purchase Order:\s*(\d+)", line)
        if po_match:
            metadata["po_number"] = po_match.group(1)
        # Format B: "No.6446 Date: 3/11/2026"
        no_match = re.search(r"No\.\s*(\d{4,})", line)
        if no_match and not metadata["po_number"]:
            metadata["po_number"] = no_match.group(1)
        date_match = re.search(r"(?:Date:|Date PO Sent)\s*(\d{1,2}/\d{1,2}/\d{4})", line)
        if date_match and not metadata["date"]:
            metadata["date"] = date_match.group(1)
        store_match = re.match(r"(C\d+.+?)\s+\d+\.\d+\s*-\s*Purchase Orders", line)
        if store_match:
            metadata["store"] = store_match.group(1).strip()

    # Two item patterns:
    # Format A (receiving report): qty  units  product#  item_id  desc  $cost  $ext
    item_pattern_a = re.compile(
        r"^(\d+\.\d+)\s+"
        r"(.+?)\s+"
        r"(\d{5,6})\s+"
        r"(\d{4,5})\s+"
        r"(.+?)\s+"
        r"\$([\d,.]+)\s+"
        r"\$([\d,.]+)$"
    )
    # Format B (direct PO): item#  qty  units  code  desc  $cost  $ext
    item_pattern_b = re.compile(
        r"^\d+\s+"
        r"(\d+\.\d+)\s+"
        r"(.+?)\s+"
        r"(\d{5,6})\s+"
        r"(.+?)\s+"
        r"\$([\d,.]+)\s+"
        r"\$([\d,.]+)$"
    )

    SKIP_STARTS = ("V ", "Vendor", "Quantity", "Item ", "P.O.",
                   "Purchase Order", "VENDOR", "Ship to", "No.", "Telephone",
                   "Fax", "Attn", "Account", "Memo", "Delivery", "Date PO",
                   "Vendor\'s", "Call In", "Total ", "***")

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = item_pattern_a.match(line)
        fmt_b = False
        if not m:
            m = item_pattern_b.match(line)
            fmt_b = bool(m)

        if m:
            if fmt_b:
                qty_raw     = float(m.group(1))
                product_num = m.group(3)
                item_id     = ""
                desc        = m.group(4).strip()
                unit_cost   = float(m.group(5).replace(",", ""))
                extension   = float(m.group(6).replace(",", ""))
            else:
                qty_raw     = float(m.group(1))
                product_num = m.group(3)
                item_id     = m.group(4)
                desc        = m.group(5).strip()
                unit_cost   = float(m.group(6).replace(",", ""))
                extension   = float(m.group(7).replace(",", ""))

            while i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                if (next_line
                    and not item_pattern_a.match(next_line)
                    and not item_pattern_b.match(next_line)
                    and not any(next_line.startswith(s) for s in SKIP_STARTS)
                    and not re.match(r"C\d{4}", next_line)
                    and "$" not in next_line):
                    desc += " " + next_line
                    i += 1
                else:
                    break

            desc = re.sub(r"^[XTY]-\s*", "", desc.strip())

            items.append(LineItem(
                qty=int(qty_raw) if qty_raw == int(qty_raw) else int(round(qty_raw)),
                description=desc,
                unit_cost=unit_cost,
                extension=extension,
                item_id=item_id,
                product_num=product_num,
            ))
        i += 1

    return items, metadata


# ============================================================================
# CATEGORIZATION
# ============================================================================

def categorize_items(items: list) -> dict:
    """Sort line items into categories. Returns dict of category_key -> Category."""
    categories = {}
    all_keys = [r[0] for r in CATEGORY_RULES]
    for key in all_keys:
        categories[key] = Category(name=key)
    categories["uncategorized"] = Category(name="uncategorized")

    for item in items:
        d = item.description.upper()
        matched = False
        for cat_key, keywords in CATEGORY_RULES:
            if any(kw in d for kw in keywords):
                categories[cat_key].items.append(item)
                matched = True
                break
        if not matched:
            categories["uncategorized"].items.append(item)
            print(f"  ⚠ UNCATEGORIZED: {item.description} (${item.extension:.2f})")

    return categories


# ============================================================================
# EXCEL FORMATTING HELPERS
# ============================================================================

THIN_BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)


def hex_to_argb(hex_color: str) -> str:
    """Convert '#RRGGBB' to 'RRGGBB' for openpyxl."""
    return hex_color.lstrip("#")


def write_title_row(ws, row, text, bg_color, font_size=12):
    """Write a merged title row across A:E."""
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(name="Arial", size=font_size, bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color=hex_to_argb(bg_color), fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    for col in range(1, 6):
        ws.cell(row=row, column=col).border = THIN_BORDER
    return row + 1


def write_col_headers(ws, row, headers=None):
    """Write column header row."""
    if headers is None:
        headers = ["Item #", "Description", "Qty (CS)", "Unit Cost", "Extension"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col, value=header)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="595959", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER
    return row + 1


def write_data_row(ws, row, item_num, item, alt_bg):
    """Write a single data row with alternating background colors."""
    is_odd = (item_num % 2 == 1)
    bg = hex_to_argb(alt_bg) if is_odd else "FFFFFF"

    values = [str(item_num), item.description, item.qty, item.unit_cost, item.extension]
    aligns = ["center", "left", "right", "right", "right"]
    formats = [None, None, None, "#,##0.00", "#,##0.00"]

    for col, (val, align, fmt) in enumerate(zip(values, aligns, formats), 1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = Font(name="Arial", size=10, color="000000")
        cell.fill = PatternFill(start_color=bg, fill_type="solid")
        cell.alignment = Alignment(horizontal=align, vertical="center")
        cell.border = THIN_BORDER
        if fmt:
            cell.number_format = fmt
    return row + 1


def write_subtotal_row(ws, row, label, data_start, data_end, bg_color):
    """Write a category subtotal row with SUM formulas."""
    for col in range(1, 3):
        cell = ws.cell(row=row, column=col)
        if col == 1:
            cell.value = f"  {label}  SUBTOTAL"
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=hex_to_argb(bg_color), fill_type="solid")
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.border = THIN_BORDER

    cell_c = ws.cell(row=row, column=3)
    cell_c.value = f"=SUM(C{data_start}:C{data_end})"
    cell_c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    cell_c.fill = PatternFill(start_color=hex_to_argb(bg_color), fill_type="solid")
    cell_c.alignment = Alignment(horizontal="right", vertical="center")
    cell_c.border = THIN_BORDER

    cell_d = ws.cell(row=row, column=4)
    cell_d.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    cell_d.fill = PatternFill(start_color=hex_to_argb(bg_color), fill_type="solid")
    cell_d.border = THIN_BORDER

    cell_e = ws.cell(row=row, column=5)
    cell_e.value = f"=SUM(E{data_start}:E{data_end})"
    cell_e.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    cell_e.fill = PatternFill(start_color=hex_to_argb(bg_color), fill_type="solid")
    cell_e.alignment = Alignment(horizontal="right", vertical="center")
    cell_e.border = THIN_BORDER

    return row + 1


def write_grand_total_row(ws, row, label, subtotal_rows_c, subtotal_rows_e, bg_color):
    """Write the grand total row referencing subtotal cells."""
    for col in range(1, 3):
        cell = ws.cell(row=row, column=col)
        if col == 1:
            cell.value = label
        cell.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=hex_to_argb(bg_color), fill_type="solid")
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = THIN_BORDER

    formula_c = "=" + "+".join(f"C{r}" for r in subtotal_rows_c)
    cell_c = ws.cell(row=row, column=3, value=formula_c)
    cell_c.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    cell_c.fill = PatternFill(start_color=hex_to_argb(bg_color), fill_type="solid")
    cell_c.alignment = Alignment(horizontal="right", vertical="center")
    cell_c.border = THIN_BORDER

    cell_d = ws.cell(row=row, column=4)
    cell_d.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    cell_d.fill = PatternFill(start_color=hex_to_argb(bg_color), fill_type="solid")
    cell_d.border = THIN_BORDER

    formula_e = "=" + "+".join(f"E{r}" for r in subtotal_rows_e)
    cell_e = ws.cell(row=row, column=5, value=formula_e)
    cell_e.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
    cell_e.fill = PatternFill(start_color=hex_to_argb(bg_color), fill_type="solid")
    cell_e.alignment = Alignment(horizontal="right", vertical="center")
    cell_e.border = THIN_BORDER
    cell_e.number_format = "$#,##0.00"

    return row + 1


def set_col_widths(ws, widths):
    """Set column widths. widths = {col_letter: width}."""
    for letter, width in widths.items():
        ws.column_dimensions[letter].width = width


# ============================================================================
# WORKBOOK BUILDER
# ============================================================================

def build_workbook(categories, metadata):
    """Build the full formatted workbook."""
    wb = Workbook()

    po = metadata.get("po_number", "????")
    date = metadata.get("date", "")
    store = metadata.get("store", "Store")
    store_short = store.split("-")[0].strip() if "-" in store else store

    col_widths = {"A": 10, "B": 45, "C": 10, "D": 12, "E": 14}

    # ===== F&B SHEET =====
    ws_fb = wb.active
    ws_fb.title = "F&B"
    ws_fb.sheet_properties.tabColor = "1F6B2E"
    set_col_widths(ws_fb, col_widths)

    row = write_title_row(
        ws_fb, 1,
        f"FOOD & BEVERAGE  —  GFS PO #{po}  |  {store_short}  |  Delivery {date}",
        "#1F6B2E"
    )

    subtotal_rows = []

    for cat_key, display_name, header_bg, alt_bg in FB_SUBCATEGORIES:
        cat = categories.get(cat_key)
        if not cat or not cat.items:
            continue

        row = write_title_row(ws_fb, row, display_name, header_bg, font_size=10)
        row = write_col_headers(ws_fb, row)
        data_start = row
        for idx, item in enumerate(cat.items, 1):
            row = write_data_row(ws_fb, row, idx, item, alt_bg)
        data_end = row - 1
        row = write_subtotal_row(ws_fb, row, display_name, data_start, data_end, header_bg)
        subtotal_rows.append(row - 1)
        row += 1

    fb_total_row = row
    write_grand_total_row(ws_fb, row, "TOTAL F&B", subtotal_rows, subtotal_rows, "#1F6B2E")
    ws_fb.freeze_panes = "A3"

    # ===== TOGO SUPPLIES SHEET =====
    ws_togo = wb.create_sheet("TOGO SUPPLIES")
    ws_togo.sheet_properties.tabColor = "1F4E79"
    set_col_widths(ws_togo, col_widths)

    row = write_title_row(
        ws_togo, 1,
        f"TO-GO SUPPLIES  —  GFS PO #{po}  |  {store_short}  |  Delivery {date}",
        "#1F4E79"
    )
    row = write_col_headers(ws_togo, row)
    togo_data_start = row
    for idx, item in enumerate(categories["togo"].items, 1):
        row = write_data_row(ws_togo, row, idx, item, "#DEEAF1")
    togo_data_end = row - 1
    togo_total_row = row
    write_subtotal_row(ws_togo, row, "TOTAL TOGO SUPPLIES", togo_data_start, togo_data_end, "#1F4E79")
    ws_togo.freeze_panes = "A3"

    # ===== MISC SUPPLIES SHEET =====
    ws_misc = wb.create_sheet("MISC SUPPLIES")
    ws_misc.sheet_properties.tabColor = "833C00"
    set_col_widths(ws_misc, col_widths)

    row = write_title_row(
        ws_misc, 1,
        f"MISC SUPPLIES  —  GFS PO #{po}  |  {store_short}  |  Delivery {date}",
        "#833C00"
    )
    row = write_col_headers(ws_misc, row)
    misc_data_start = row
    for idx, item in enumerate(categories["misc"].items, 1):
        row = write_data_row(ws_misc, row, idx, item, "#FCE4D6")
    misc_data_end = row - 1
    misc_total_row = row
    write_subtotal_row(ws_misc, row, "TOTAL MISC SUPPLIES", misc_data_start, misc_data_end, "#833C00")
    ws_misc.freeze_panes = "A3"

    # ===== CHEMICALS SHEET =====
    ws_chem = wb.create_sheet("CHEMICALS")
    ws_chem.sheet_properties.tabColor = "4B0082"
    set_col_widths(ws_chem, col_widths)

    row = write_title_row(
        ws_chem, 1,
        f"CHEMICALS / CLEANING  —  GFS PO #{po}  |  {store_short}  |  Delivery {date}",
        "#4B0082"
    )
    row = write_col_headers(ws_chem, row)
    chem_data_start = row
    for idx, item in enumerate(categories["chemicals"].items, 1):
        row = write_data_row(ws_chem, row, idx, item, "#EDE7F6")
    chem_data_end = row - 1
    chem_total_row = row
    write_subtotal_row(ws_chem, row, "TOTAL CHEMICALS", chem_data_start, chem_data_end, "#4B0082")
    ws_chem.freeze_panes = "A3"

    # ===== BEER & LIQUOR SHEET =====
    ws_beer = wb.create_sheet("BEER & LIQUOR")
    ws_beer.sheet_properties.tabColor = "7B3F00"
    set_col_widths(ws_beer, col_widths)

    row = write_title_row(
        ws_beer, 1,
        f"BEER & LIQUOR  —  GFS PO #{po}  |  {store_short}  |  Delivery {date}",
        "#7B3F00"
    )
    row = write_col_headers(ws_beer, row)
    beer_data_start = row
    beer_items = categories.get("beer_liquor", Category(name="beer_liquor")).items
    for idx, item in enumerate(beer_items, 1):
        row = write_data_row(ws_beer, row, idx, item, "#FDF0E0")
    beer_data_end = row - 1
    beer_total_row = row
    if not beer_items:
        ws_beer.cell(row=row, column=2, value="No beer/liquor items in this order").font = Font(
            name="Arial", size=10, italic=True, color="999999"
        )
        row += 1
    write_subtotal_row(ws_beer, row if beer_items else beer_total_row,
                       "TOTAL BEER & LIQUOR", beer_data_start,
                       beer_data_end if beer_items else beer_data_start, "#7B3F00")
    beer_total_row = row if beer_items else beer_total_row
    ws_beer.freeze_panes = "A3"

    # ===== SUMMARY SHEET (insert at position 0) =====
    ws_sum = wb.create_sheet("SUMMARY", 0)
    ws_sum.sheet_properties.tabColor = "C00000"
    set_col_widths(ws_sum, {"A": 16, "B": 38, "C": 12, "D": 14, "E": 12})

    r = write_title_row(
        ws_sum, 1,
        f"{store_short.upper()}  —  GFS PO #{po}  |  Delivery {date}",
        "#C00000", font_size=14
    )
    r = write_title_row(ws_sum, r, "EXPENSE CATEGORY SUMMARY", "#404040", font_size=11)

    sum_headers = ["Category", "Description", "# Cases", "Total Cost", "% of Order"]
    for col, header in enumerate(sum_headers, 1):
        cell = ws_sum.cell(row=r, column=col, value=header)
        cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="404040", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER
    r += 1

    cat_rows = [
        ("F&B",           "Food & Beverage Purchases",
         "E2EFDA",
         f"='F&B'!C{fb_total_row}",         f"='F&B'!E{fb_total_row}"),
        ("Beer & Liquor",  "Beer, Liquor, Wine & Seltzer",
         "FDF0E0",
         f"='BEER & LIQUOR'!C{beer_total_row}", f"='BEER & LIQUOR'!E{beer_total_row}"),
        ("To-Go Supplies", "To-Go Packaging & Containers",
         "DEEAF1",
         f"='TOGO SUPPLIES'!C{togo_total_row}", f"='TOGO SUPPLIES'!E{togo_total_row}"),
        ("Misc Supplies",  "Operational Supplies & Paper Goods",
         "FCE4D6",
         f"='MISC SUPPLIES'!C{misc_total_row}", f"='MISC SUPPLIES'!E{misc_total_row}"),
        ("Chemicals",      "Cleaning & Sanitation Chemicals",
         "EDE7F6",
         f"='CHEMICALS'!C{chem_total_row}",     f"='CHEMICALS'!E{chem_total_row}"),
    ]

    data_start_row = r
    for label, desc, bg, cases_formula, cost_formula in cat_rows:
        ws_sum.cell(row=r, column=1, value=label).font = Font(name="Arial", size=10, bold=True)
        ws_sum.cell(row=r, column=1).fill = PatternFill(start_color=bg, fill_type="solid")
        ws_sum.cell(row=r, column=1).alignment = Alignment(horizontal="center")
        ws_sum.cell(row=r, column=1).border = THIN_BORDER

        ws_sum.cell(row=r, column=2, value=desc).font = Font(name="Arial", size=10)
        ws_sum.cell(row=r, column=2).fill = PatternFill(start_color=bg, fill_type="solid")
        ws_sum.cell(row=r, column=2).alignment = Alignment(horizontal="left")
        ws_sum.cell(row=r, column=2).border = THIN_BORDER

        ws_sum.cell(row=r, column=3, value=cases_formula).font = Font(name="Arial", size=10)
        ws_sum.cell(row=r, column=3).fill = PatternFill(start_color=bg, fill_type="solid")
        ws_sum.cell(row=r, column=3).alignment = Alignment(horizontal="center")
        ws_sum.cell(row=r, column=3).border = THIN_BORDER

        ws_sum.cell(row=r, column=4, value=cost_formula).font = Font(name="Arial", size=10)
        ws_sum.cell(row=r, column=4).fill = PatternFill(start_color=bg, fill_type="solid")
        ws_sum.cell(row=r, column=4).alignment = Alignment(horizontal="center")
        ws_sum.cell(row=r, column=4).number_format = "$#,##0.00"
        ws_sum.cell(row=r, column=4).border = THIN_BORDER

        ws_sum.cell(row=r, column=5, value=f"=D{r}/D{data_start_row + len(cat_rows)}").font = Font(name="Arial", size=10)
        ws_sum.cell(row=r, column=5).fill = PatternFill(start_color=bg, fill_type="solid")
        ws_sum.cell(row=r, column=5).alignment = Alignment(horizontal="center")
        ws_sum.cell(row=r, column=5).number_format = "0.0%"
        ws_sum.cell(row=r, column=5).border = THIN_BORDER
        r += 1

    grand_row = r
    for col in range(1, 6):
        cell = ws_sum.cell(row=grand_row, column=col)
        cell.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="C00000", fill_type="solid")
        cell.border = THIN_BORDER

    ws_sum.cell(row=grand_row, column=1, value=f"GRAND TOTAL — GFS PO #{po}")
    ws_sum.cell(row=grand_row, column=1).alignment = Alignment(horizontal="left")
    ws_sum.cell(row=grand_row, column=3, value=f"=SUM(C{data_start_row}:C{grand_row-1})")
    ws_sum.cell(row=grand_row, column=3).alignment = Alignment(horizontal="right")
    ws_sum.cell(row=grand_row, column=4, value=f"=SUM(D{data_start_row}:D{grand_row-1})")
    ws_sum.cell(row=grand_row, column=4).alignment = Alignment(horizontal="right")
    ws_sum.cell(row=grand_row, column=4).number_format = "$#,##0.00"
    ws_sum.cell(row=grand_row, column=5, value="100.0%")
    ws_sum.cell(row=grand_row, column=5).alignment = Alignment(horizontal="right")

    ws_sum.freeze_panes = "A4"

    return wb


# ============================================================================
# MAIN
# ============================================================================

def main():
    if len(sys.argv) < 2:
        print("Usage: python gfs_breakdown.py <input_pdf> [output_xlsx]")
        print("       Parses a GFS Purchase Order PDF into a categorized Excel workbook.")
        sys.exit(1)

    pdf_path = sys.argv[1]
    if not os.path.exists(pdf_path):
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)

    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        stem = Path(pdf_path).stem
        output_path = f"{stem}_Breakdown.xlsx"

    print(f"📄 Parsing: {pdf_path}")
    items, metadata = parse_pdf(pdf_path)
    print(f"   Found {len(items)} line items | PO #{metadata['po_number']} | {metadata['date']}")

    total = sum(i.extension for i in items)
    print(f"   Raw total: ${total:,.2f}")

    print(f"\n📦 Categorizing...")
    categories = categorize_items(items)

    fb_keys = [k for k, _, _, _ in FB_SUBCATEGORIES]
    fb_total = sum(categories[k].total_cost for k in fb_keys if k in categories)
    fb_cases = sum(categories[k].total_cases for k in fb_keys if k in categories)

    print(f"\n{'='*60}")
    print(f"   CATEGORY SUMMARY — PO #{metadata['po_number']}")
    print(f"{'='*60}")
    print(f"   {'Category':<35} {'Cases':>6} {'Cost':>12}")
    print(f"   {'-'*35} {'-'*6} {'-'*12}")

    for cat_key, display_name, _, _ in FB_SUBCATEGORIES:
        cat = categories.get(cat_key)
        if cat and cat.items:
            print(f"   {display_name:<35} {cat.total_cases:>6} ${cat.total_cost:>10,.2f}")
    print(f"   {'F&B TOTAL':<35} {fb_cases:>6} ${fb_total:>10,.2f}")
    print()

    for key, label in [("togo", "TO-GO SUPPLIES"), ("misc", "MISC SUPPLIES"), ("chemicals", "CHEMICALS")]:
        cat = categories[key]
        print(f"   {label:<35} {cat.total_cases:>6} ${cat.total_cost:>10,.2f}")

    grand_cases = fb_cases + categories["togo"].total_cases + categories["misc"].total_cases + categories["chemicals"].total_cases
    grand_cost = fb_total + categories["togo"].total_cost + categories["misc"].total_cost + categories["chemicals"].total_cost
    print(f"   {'-'*35} {'-'*6} {'-'*12}")
    print(f"   {'GRAND TOTAL':<35} {grand_cases:>6} ${grand_cost:>10,.2f}")

    if categories["uncategorized"].items:
        print(f"\n   ⚠ {len(categories['uncategorized'].items)} uncategorized items:")
        for item in categories["uncategorized"].items:
            print(f"     - {item.description} ({item.qty} cs, ${item.extension:.2f})")

    print(f"\n📊 Building workbook...")
    wb = build_workbook(categories, metadata)

    wb.save(output_path)
    print(f"✅ Saved: {output_path}")
    print(f"   Sheets: {[ws.title for ws in wb.worksheets]}")


if __name__ == "__main__":
    main()
