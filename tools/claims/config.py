"""
config.py  -  ALL SETTINGS LIVE HERE.
Change values in this file only; you normally never need to touch the other files.
"""
import os

VERSION = "v0.7.1"

# ---------------------------------------------------------------------------
# Folders / files
# ---------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# The tool tries to open this file automatically on start (you can also browse).
DEFAULT_SALES_FILE = os.path.join(BASE_DIR, "input", "Sales_Data.xls")
# Reports are saved here. Existing files are never overwritten.
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")

# ---------------------------------------------------------------------------
# Where to find things in the sales file
# ---------------------------------------------------------------------------
# The sheet used is the first one whose name contains this text (not case sensitive).
SALES_SHEET_KEYWORD = "Pack Wise"
COL_CODE = "Distr. SAP Code"        # column heading holding the distributor code
COL_NAME = "Distr. Name"            # column heading holding the distributor name
COL_ADDRESS = "Distr. Address"      # pack columns start right after this column
COL_PACKS_END = "Grand Total Raw Cases Vol."   # pack columns stop just before this one

# ---------------------------------------------------------------------------
# Rows shown on the report, in this order.
#   ("Text shown in the 'Pack' column", ["column heading(s) in the sales file"])
# If several headings are listed, their cases are ADDED together.
# ---------------------------------------------------------------------------
PACKS = [
    ("250ML SSRB",            ["250ML (SSRB CSD)"]),
    ("1000ML Pet",            ["1000ML PET"]),
    ("1500ML Pet",            ["1500ML PET"]),
    ("300ML Sting",           ["300ML PET STING"]),
    ("500ML Pet Sting",       ["500ML PET STING"]),
    ("RB Slice 240 ML",       ["240ML SLICE-SSRB"]),
    ("RB Sting 240 ML",       ["240ML STING-SSRB"]),
    ("2000 ML (PET)",         ["2000ML PET (4-B)", "2000ML PET (6-B)"]),   # 4-B + 6-B added together
    ("A Fina 500 ML",         ["500ML AQUA FINA"]),
    ("A Fina 1500 ML",        ["1500ML AQUA FINA"]),
    ("200ML Slice Tetra Pack", ["200ML TP SLICE(24-B)"]),
    ("345ML PET",             ["345ML PET"]),
    ("500ML Pet",             ["500ML PET"]),
]
# Packs in the sales file that are NOT listed above are left off the report,
# but any cases they hold are listed on the "Check" sheet so nothing disappears silently.

# ---------------------------------------------------------------------------
# Header block of every report sheet.
# Month, Year, Unit, Route, Area Incharge and Freight Rate are all typed in the window (they start empty).
# ---------------------------------------------------------------------------
REPORT_TITLE = "Variable Freight Charges"

# Signature boxes at the bottom (row 1: left, middle, right / row 2: left, middle, right)
SIGNATURES_ROW_1 = ["Distributor", "SDM/SDO/SDS", "Unit Manager"]
SIGNATURES_ROW_2 = ["S & D Manager", "Head Of Sale", "HOD (MR&D)"]

# ---------------------------------------------------------------------------
# Sheet names:  "<Distributor> Freight Charges <Mon-yy>"
# Excel allows only 31 characters. If too long, "Freight Charges" is shortened
# to SHORT_WORD, and only if still too long the distributor name is trimmed.
# ---------------------------------------------------------------------------
FULL_WORD = "Freight Charges"
SHORT_WORD = "F.C"
MAX_SHEET_NAME = 31

# ---------------------------------------------------------------------------
# Printing
# ---------------------------------------------------------------------------
PAPER_SIZE_A4 = 9
ORIENTATION = "portrait"             # the report is narrow, so portrait fits well
MARGIN_LEFT_RIGHT = 0.4              # inches
MARGIN_TOP_BOTTOM = 0.75
FONT_NAME = "Arial"


# ===========================================================================
# WHOLESALE CLAIM  (added in v0.4)
# ===========================================================================
# The tool tries to open this file automatically when the Wholesale window opens (you can also browse).
WS_DEFAULT_FILE = os.path.join(BASE_DIR, "input", "Wholesale_data.xlsx")

# Column headings to look for in the wholesale data file.
# The sheet used is the first one that has the "dist_code" heading in its first 30 rows.
WS_COLS = {
    "dist_code":      "Dist_Code",
    "dist_name":      "Dist_Name",
    "unit":           "Unit",
    "sdm":            "SDM/SDO/SDS",
    "outlet_code":    "Outlet_Code",
    "outlet_name":    "Outlet_Name",
    "outlet_address": "Outlet_Address",
    "invoice":        "INVOICE_NUMBER",
    "order_date":     "ORDER_DATE",        # BOOKING date - this decides if a sale is inside the claim dates
    "delivery_date":  "DELIVERY_DATE",     # only used for the "delivered next month" warning
    "pack":           "PKG_NAME",
    "brand":          "Brand NAME",
    "cases":          "Sales_Raw Cases",
    "promo_cases":    "Promo_Raw_Cases",   # optional column; only checked, never claimed
}

WS_REPORT_TITLE = "Whole Seller Discount Claim"

# How pack / brand names from the data are WRITTEN on the report (keys must be CAPITAL letters).
# Anything not listed here is written as it appears in the data.
# The ORDER of WS_BRAND_LABELS is also the order brands are written in, e.g. "Pepsi, 7up, Dew".
WS_PACK_LABELS = {"1000 ML PET": "1000 ML (PET)", "1500 ML PET": "1500 ML (PET)"}
WS_BRAND_LABELS = {"PEPSI": "Pepsi", "7-UP": "7up", "M.DEW": "Dew", "MIRINDA": "Mirinda", "7UP MINT": "7up Mint"}

# Packs that are ALWAYS offered in the "Pack" list of the conditions (added to the packs found in the data file;
# a pack that is in both is shown once). Names should be written as in the data file (PKG_NAME).
# All brands named in WS_BRAND_LABELS above are likewise always offered in the "Brand" list, added to the brands found in the data.
WS_EXTRA_PACKS = ["500 ML PET", "1000 ML PET", "1500 ML PET", "2000 ML PET"]

WS_SHOW_ADDRESS = True    # True: outlet name + address in the "Outlet Name & Address" cell; False: name only

# Signature labels at the bottom (line 1: left, middle, right / line 2: left, middle, right)
WS_SIGNATURES_ROW_1 = ["Distributor", "SDM/O/S", "Unit Manager"]
WS_SIGNATURES_ROW_2 = ["S&D Manager", "C.C.O", "HOD (MR&D)"]


# ===========================================================================
# DISCOUNT SUMMARY  "[Unit] Distribution Wise Discount Amount [Month-Year]"  (added in v0.5)
# ===========================================================================
DS_SHEET_NAME = "Credit Note"            # name of the summary sheet (as in your sample)
DS_TITLE = "Total Amount Summary"        # title row of the sheet
DS_FILE_WORD = "Distribution Wise Discount Amount"   # file name = "<Unit> <this text> <Month-Year>.xlsx"

# Column headings, in order, copied from your sample (A to N).  Spelling is exactly as in the sample.
DS_HEADINGS = ["Sr. #", "Code", "Distributor Name",
               "Credit Note\n(Card Discount)\nGTM", "Non GTM\nAmount", "(HTH)\nAmount",
               "Priodic\nAmount", "Promo Claim Amount", "Area Coverage\nFreight Charges Amount",
               "Whole Sale\nAmount", "Special Approval Amount", "Total Amount", "Unit Name", "Claim Month"]

# Rows are sorted by distributor code ("code") or by name ("name").
DS_SORT = "code"

# How the PDFs are recognised (text printed under the page title).
DS_PDF_CARD_KEY = "(Card Discount)"
DS_PDF_HTH_KEY = "(HTH)"


# ===========================================================================
# DISTRIBUTION WISE COST  "[Unit] Distribution Wise Cost [Month-Year]"  (added in v0.6)
# ===========================================================================
CO_FILE_WORD = "Distribution Wise Cost"          # file name = "<Unit> <this text> <Month-Year>.xlsx"
CO_SHEET_NAME = "Credit Note"                    # name of the cost sheet (as in your sample)
CO_COL_RAW_CASES = COL_PACKS_END                 # column of the sales file that holds each distributor's Raw Cases
CO_HEADINGS_EXTRA = ["Raw Cases", "Per Unit Cost"]   # two columns added after "Total Amount" (M and N)
# Labels of the Sampling / unit block under the table (7 lines, in this order, as in your sample)
CO_BLOCK_LABELS = ["Sampling Cases", "Sampling Price", "Sampling Amount", "Claims Amount",
                   "Total Amount", "Raw Cases", "Per Case Cost"]


# ===========================================================================
# MONTHLY WHOLESALE CLAIMS DETAIL  "[Unit] Distribution Wise Monthly Wholesale Claims Detail [Month-Year]"  (added in v0.7)
# ===========================================================================
WCD_FILE_WORD = "Distribution Wise Monthly Wholesale Claims Detail"   # file name = "<Unit> <this text> <Month-Year>.xlsx"
WCD_SHEET_NAME = "Dic. Cs"                                           # name of the report sheet (as in your sample)
WCD_TITLE = "Distributor Wise Whole Sale Discount Claims Summary"    # title row of the sheet
WCD_HEAD_CODE = "Code"
WCD_HEAD_NAME = "Distributon\nName"                                  # spelling exactly as in the sample

# One block of 7 columns per pack, left to right, in this order.
#   (heading above the block,
#    [column heading(s) in the SALES file]   -> Primary Lifting (several are ADDED together),
#    [pack name(s) in the WHOLESALE CLAIM sheets]   -> Variable Disc. W.S,
#    [package name(s) printed in the HTH / Card Discount PDFs]  -> Committed Discount (card + HTH added))
# Names are compared ignoring capitals, spaces and brackets, so "1000 ML (PET)" = "1000 ML PET".
# "1500 ML PET AF" (Aqua Fina) is a DIFFERENT package and is NOT counted in "1500 ML PET".
WCD_PACKS = [
    ("1500 ML (PET)", ["1500ML PET"],                          ["1500 ML (PET)"], ["1500 ML PET"]),
    ("1000 ML (PET)", ["1000ML PET"],                          ["1000 ML (PET)"], ["1000 ML PET"]),
    ("2000 ML (PET)", ["2000ML PET (4-B)", "2000ML PET (6-B)"], ["2000 ML (PET)"], ["2000 ML PET"]),
]

# The 7 column headings inside every pack block (as in the sample).
WCD_HEADINGS = ["Primary Lifting", "Carry Forward Stock", "Variable Disc. W.S (Raw CS)", "% age of\nW.S Claim",
                "Committed Discount\n(Raw CS)", "Total Discounted\nCases", "% age of Primary Lifting"]

WCD_RED_LIMIT = 0.5          # "% age of W.S Claim" ABOVE this (50%) gets a red background
WCD_COLOR_RED = "FF0000"
WCD_COLOR_BLUE = "00B0F0"    # Committed Discount that is the SUM of the Card Discount figure and the HTH figure
