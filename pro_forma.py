"""
GSA FRPP Valuation Engine

Author: Nana Darko

Description:
Automates commercial real estate underwriting using the
FY2024 Federal Real Property Profile (FRPP) dataset.

Features
--------
- Loads FRPP dataset
- Filters office properties
- Detects square footage automatically
- Handles missing building size
- Generates rent roll
- Creates 5-year pro forma
- Calculates NOI
- Calculates Debt Service
- Calculates DSCR
- Calculates Cash Flow
- Calculates Cash-on-Cash Return
- Exports results to Excel
"""

import os
import pandas as pd
import numpy as np


# =====================================================
# CONFIGURATION
# =====================================================

DATA_FILE = "data/frpp_public_dataset_fy24_07022025 (1).xlsx"

OUTPUT_FOLDER = "output"

OUTPUT_FILE = os.path.join(
    OUTPUT_FOLDER,
    "CRE_Model_Output.xlsx"
)


# =====================================================
# LOAD DATASET
# =====================================================

def load_dataset():

    print("🔥 CRE MODEL STARTED")

    print("\n1️⃣ Checking file...")

    print("File exists:", os.path.exists(DATA_FILE))

    if not os.path.exists(DATA_FILE):
        raise FileNotFoundError(
            f"Dataset not found:\n{DATA_FILE}"
        )

    print("\n2️⃣ Loading dataset...")

    xls = pd.ExcelFile(DATA_FILE)

    df = pd.read_excel(
        xls,
        sheet_name=0
    )

    df.columns = df.columns.str.strip()

    print("✅ Dataset Loaded")

    print(f"Rows: {len(df):,}")

    print(f"Columns: {len(df.columns)}")

    return df


# =====================================================
# FIND PROPERTY USE COLUMN
# =====================================================

def get_property_use_column(df):

    for col in df.columns:

        if (
            "property" in col.lower()
            and
            "use" in col.lower()
        ):

            return col

    raise Exception(
        "Property Use column not found."
    )


# =====================================================
# FILTER OFFICE ASSETS
# =====================================================

def filter_office_assets(df):

    use_col = get_property_use_column(df)

    office_df = df[
        df[use_col]
        .astype(str)
        .str.contains(
            "Office",
            case=False,
            na=False
        )
    ]

    print("\nOffice Assets Found:", f"{len(office_df):,}")

    if len(office_df) == 0:
        raise Exception(
            "No office assets found."
        )

    return office_df


# =====================================================
# FIND SQUARE FOOTAGE
# =====================================================

def detect_sqft(asset, df):

    priority_cols = [

        "Gross Square Feet",

        "Gross Square Feet (Buildings)",

        "Building Gross Square Feet",

        "Square Feet (Buildings)",

        "Gross Sq Ft",

        "GSF"

    ]

    sqft_col = None

    for col in priority_cols:

        if col in df.columns:

            sqft_col = col

            break

    if sqft_col is None:

        for col in df.columns:

            c = col.lower()

            if (
                "square" in c
                and
                ("feet" in c or "foot" in c)
            ):

                sqft_col = col

                break

    print("\nDetected Square Footage Column:")

    print(sqft_col)

    sqft = np.nan

    if sqft_col:

        try:

            sqft = float(asset[sqft_col])

        except:

            sqft = np.nan

    if pd.isna(sqft) or sqft <= 0:

        print(
            "⚠️ Invalid or missing SqFt → using fallback 50,000"
        )

        sqft = 50000

    return sqft
# =====================================================
# BUILD RENT ROLL
# =====================================================

def build_rent_roll(sqft):

    print("\nGenerating Rent Roll...")

    rent_roll = pd.DataFrame({

        "Unit": [100, 200, 300, 400],

        "Tenant": [

            "Agency A",

            "Agency B",

            "Agency C",

            "Vacant"

        ],

        "SqFt": [

            sqft * 0.40,

            sqft * 0.35,

            sqft * 0.15,

            sqft * 0.10

        ],

        "RentPerSF": [

            2.10,

            2.00,

            1.80,

            0.00

        ]

    })

    rent_roll["MonthlyRent"] = (

        rent_roll["SqFt"]

        * rent_roll["RentPerSF"]

    )

    print(rent_roll)

    return rent_roll


# =====================================================
# BUILD 5-YEAR PRO FORMA
# =====================================================

def build_pro_forma(rent_roll):

    print("\nBuilding 5-Year Pro Forma...")

    growth = 0.03

    vacancy = 0.05

    operating_expense = 130000

    base_gpr = (

        rent_roll["MonthlyRent"].sum()

        * 12

    )

    rows = []

    for year in range(1, 6):

        revenue = (

            base_gpr

            * (1 + growth) ** (year - 1)

        )

        egi = revenue * (1 - vacancy)

        expenses = (

            operating_expense

            * (1.02 ** (year - 1))

        )

        noi = egi - expenses

        rows.append([

            year,

            revenue,

            egi,

            expenses,

            noi

        ])

    pro_forma = pd.DataFrame(

        rows,

        columns=[

            "Year",

            "Revenue",

            "EGI",

            "Expenses",

            "NOI"

        ]

    )

    print(pro_forma)

    return pro_forma


# =====================================================
# FINANCING
# =====================================================

def apply_financing(pro_forma):

    print("\nApplying Financing...")

    loan = 3_850_000

    equity = 1_650_000

    interest_rate = 0.065

    annual_debt = loan * interest_rate

    pro_forma["DebtService"] = annual_debt

    pro_forma["CashFlow"] = (

        pro_forma["NOI"]

        - annual_debt

    )

    pro_forma["CoC_Return"] = (

        pro_forma["CashFlow"]

        / equity

    )

    pro_forma["DSCR"] = (

        pro_forma["NOI"]

        / annual_debt

    )

    print("\nDebt Service Coverage Ratio")

    print(

        pro_forma[

            ["Year", "DSCR"]

        ]

    )

    return pro_forma, annual_debt


# =====================================================
# CREATE SUMMARY
# =====================================================

def create_summary(

    df,

    office_df,

    sqft,

    pro_forma,

    annual_debt

):

    summary = pd.DataFrame({

        "Metric": [

            "Dataset Records",

            "Office Assets",

            "Building Size (SF)",

            "Year 1 NOI",

            "Annual Debt Service",

            "Year 1 DSCR"

        ],

        "Value": [

            len(df),

            len(office_df),

            sqft,

            pro_forma.loc[0, "NOI"],

            annual_debt,

            pro_forma.loc[0, "DSCR"]

        ]

    })

    return summary
# =====================================================
# EXPORT TO EXCEL
# =====================================================

def export_to_excel(

    asset,

    rent_roll,

    pro_forma,

    summary

):

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    with pd.ExcelWriter(

        OUTPUT_FILE,

        engine="openpyxl"

    ) as writer:

        summary.to_excel(

            writer,

            sheet_name="Summary",

            index=False

        )

        asset.to_frame().T.to_excel(

            writer,

            sheet_name="Asset",

            index=False

        )

        rent_roll.to_excel(

            writer,

            sheet_name="Rent_Roll",

            index=False

        )

        pro_forma.to_excel(

            writer,

            sheet_name="Pro_Forma",

            index=False

        )

    print("\n✅ Excel workbook created successfully!")

    print(OUTPUT_FILE)


# =====================================================
# PRINT SUMMARY
# =====================================================

def print_summary(

    df,

    office_df,

    sqft,

    pro_forma,

    annual_debt

):

    print("\n======================================")

    print("      CRE VALUATION SUMMARY")

    print("======================================")

    print(f"Dataset Records : {len(df):,}")

    print(f"Office Assets   : {len(office_df):,}")

    print(f"Building Size   : {sqft:,.0f} SF")

    print(f"Year 1 NOI      : ${pro_forma.loc[0,'NOI']:,.0f}")

    print(f"Debt Service    : ${annual_debt:,.0f}")

    print(f"Year 1 DSCR     : {pro_forma.loc[0,'DSCR']:.2f}")

    if pro_forma.loc[0, "DSCR"] < 1.25:

        print("⚠️ Status        : Flag for Review")

    else:

        print("✅ Status        : Healthy")

    print("======================================")


# =====================================================
# MAIN
# =====================================================

def main():

    df = load_dataset()

    office_df = filter_office_assets(df)

    asset = office_df.iloc[0]

    print("\nSelected Asset")

    print(asset)

    sqft = detect_sqft(asset, df)

    print(f"\nBuilding Size Used: {sqft:,.0f} SF")

    rent_roll = build_rent_roll(sqft)

    pro_forma = build_pro_forma(rent_roll)

    pro_forma, annual_debt = apply_financing(pro_forma)

    summary = create_summary(

        df,

        office_df,

        sqft,

        pro_forma,

        annual_debt

    )

    export_to_excel(

        asset,

        rent_roll,

        pro_forma,

        summary

    )

    print_summary(

        df,

        office_df,

        sqft,

        pro_forma,

        annual_debt

    )


# =====================================================
# RUN PROGRAM
# =====================================================

if __name__ == "__main__":

    main()