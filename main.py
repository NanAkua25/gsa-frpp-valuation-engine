import pandas as pd
import numpy as np
import os

print("🔥 CRE MODEL STARTED")

# =========================
# 1. LOAD DATASET
# =========================

file_path = "data/frpp_public_dataset_fy24_07022025 (1).xlsx"

print("\n1️⃣ Checking file...")
print("File exists:", os.path.exists(file_path))

if not os.path.exists(file_path):
    print("❌ File not found")
    exit()

print("2️⃣ Loading dataset...")

try:
    xls = pd.ExcelFile(file_path)
    df = pd.read_excel(xls, sheet_name=0)

    print("3️⃣ Dataset loaded successfully")
    print("Shape:", df.shape)

except Exception as e:
    print("❌ ERROR loading file:")
    print(e)
    exit()

df.columns = df.columns.str.strip()

print("\nColumns detected:")
print(df.columns)

# =========================
# 2. FIND PROPERTY USE COLUMN
# =========================

use_col = None

for col in df.columns:
    if "use" in col.lower() and "property" in col.lower():
        use_col = col
        break

if use_col is None:
    print("❌ No property use column found")
    exit()

print("\nUsing column:", use_col)

# =========================
# 3. FILTER OFFICE ASSETS
# =========================

office_df = df[df[use_col].astype(str).str.contains("Office", na=False)]

print("Office assets found:", len(office_df))

if len(office_df) == 0:
    print("❌ No office assets found")
    exit()

asset = office_df.iloc[0]

print("\nSelected Asset:")
print(asset)

# =========================
# 4. SAFE SQUARE FOOTAGE EXTRACTION
# =========================

print("\n🔍 Detecting square footage column...")

sqft_col = None

priority_cols = [
    "Gross Square Feet",
    "Gross Square Feet (Buildings)",
    "Building Gross Square Feet",
    "Gross Sq Ft",
    "GSF"
]

for col in priority_cols:
    if col in df.columns:
        sqft_col = col
        break

if sqft_col is None:
    for col in df.columns:
        c = col.lower()
        if "square" in c and ("feet" in c or "foot" in c):
            sqft_col = col
            break

print("Detected sqft column:", sqft_col)

sqft = None

if sqft_col is not None:
    try:
        sqft = float(asset[sqft_col])
    except:
        sqft = np.nan

if sqft is None or pd.isna(sqft) or sqft <= 0:
    print("⚠️ Invalid or missing SqFt → using fallback 50,000")
    sqft = 50000

print("Building Size USED:", sqft)

# =========================
# 5. RENT ROLL MODEL
# =========================

rent_roll = pd.DataFrame({
    "Unit": [100, 200, 300, 400],
    "Tenant": ["Agency A", "Agency B", "Agency C", "Vacant"],
    "SqFt": [
        sqft * 0.40,
        sqft * 0.35,
        sqft * 0.15,
        sqft * 0.10
    ],
    "RentPerSF": [2.10, 2.00, 1.80, 0.00]
})

rent_roll["MonthlyRent"] = rent_roll["SqFt"] * rent_roll["RentPerSF"]

print("\nRent Roll:")
print(rent_roll)

# =========================
# 6. PRO FORMA MODEL
# =========================

growth = 0.03
vacancy = 0.05

base_gpr = rent_roll["MonthlyRent"].sum() * 12

rows = []

for year in range(1, 6):

    revenue = base_gpr * (1 + growth) ** (year - 1)
    egi = revenue * (1 - vacancy)

    expenses = 130000 * (1.02 ** (year - 1))
    noi = egi - expenses

    rows.append([year, revenue, egi, expenses, noi])

pro_forma = pd.DataFrame(rows, columns=[
    "Year", "Revenue", "EGI", "Expenses", "NOI"
])

print("\nPro Forma:")
print(pro_forma)

# =========================
# 7. FINANCING
# =========================

loan = 3850000
equity = 1650000
rate = 0.065

annual_debt = loan * rate

pro_forma["DebtService"] = annual_debt
pro_forma["CashFlow"] = pro_forma["NOI"] - annual_debt
pro_forma["CoC_Return"] = pro_forma["CashFlow"] / equity

# =========================
# 8. EXPORT EXCEL
# =========================

output_file = "CRE_Model_Output.xlsx"

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    asset.to_frame().T.to_excel(writer, sheet_name="Asset", index=False)
    rent_roll.to_excel(writer, sheet_name="Rent_Roll", index=False)
    pro_forma.to_excel(writer, sheet_name="Pro_Forma", index=False)

print("\n✅ DONE! Excel file created:", output_file)