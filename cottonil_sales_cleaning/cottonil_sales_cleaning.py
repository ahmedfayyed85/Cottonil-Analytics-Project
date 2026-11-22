
import zipfile
import re
import pandas as pd
import numpy as np

from pathlib import Path
from bs4 import BeautifulSoup
from collections import Counter


# ==========================================
# 1. PROJECT SETTINGS
# ==========================================

BASE_DIR = Path(__file__).parent

ZIP_FILE = BASE_DIR / "مبيعات.zip"

OUTPUT_FILE = BASE_DIR / "Cottonil_Sales_Cleaned.xlsx"

COLUMNS = [
    "Date",
    "Permission_Number",
    "Barcode",
    "Product_Code",
    "Product_Name",
    "Price",
    "Quantity",
    "Total_Value"
]


# ==========================================
# 2. HELPER FUNCTIONS
# ==========================================

def clean_text(value):

    if value is None:
        return None

    value = str(value)

    # Remove extra spaces
    value = re.sub(r"\s+", " ", value)

    value = value.strip()

    # Convert empty strings to NaN
    if value in ["", "-", "--", "—"]:
        return np.nan

    return value


def clean_numeric(value):

    if pd.isna(value):
        return np.nan

    value = str(value).strip()

    # Remove commas and spaces
    value = value.replace(",", "")
    value = value.replace(" ", "")

    # Convert Arabic numerals if present
    arabic_digits = "٠١٢٣٤٥٦٧٨٩"
    english_digits = "0123456789"

    translation = str.maketrans(arabic_digits, english_digits)
    value = value.translate(translation)

    # Keep only numeric characters and decimal point
    value = re.sub(r"[^0-9.\-]", "", value)

    try:
        return float(value)
    except ValueError:
        return np.nan


def natural_sort_key(filename):

    return [
        int(part) if part.isdigit() else part.lower()
        for part in re.split(r"(\d+)", filename)
    ]


# ==========================================
# 3. EXTRACT DATA FROM ALL HTML FILES
# ==========================================

print("=" * 50)
print("COTTONIL SALES DATA EXTRACTION")
print("=" * 50)

all_records = []

skipped_rows = []

with zipfile.ZipFile(ZIP_FILE, "r") as zip_ref:

    html_files = [
        name for name in zip_ref.namelist()
        if name.lower().endswith((".html", ".htm"))
    ]

    html_files = sorted(html_files, key=natural_sort_key)

    print(f"\nTotal HTML files: {len(html_files)}")

    for file_number, file_name in enumerate(html_files, start=1):

        try:

            # Read HTML using Windows-1256 encoding
            html_content = zip_ref.read(file_name).decode(
                "cp1256",
                errors="replace"
            )

            soup = BeautifulSoup(
                html_content,
                "html.parser"
            )

            # Extract all table rows
            table_rows = soup.find_all("tr")

            for row in table_rows:

                cells = row.find_all(
                    "td",
                    recursive=False
                )

                # Expected structure:
                # 1 spacer cell + 8 data cells
                if len(cells) != 9:
                    continue

                values = [
                    clean_text(cell.get_text(" ", strip=True))
                    for cell in cells
                ]

                # Remove first empty spacer cell
                values = values[1:]

                if len(values) != 8:
                    continue

                date_value = values[0]

                # Identify actual sales records
                # Ignore report headers and footers
                if pd.isna(date_value):
                    continue

                if not re.fullmatch(
                    r"\d{1,2}/\d{1,2}/\d{2,4}",
                    str(date_value)
                ):
                    continue

                record = dict(
                    zip(COLUMNS, values)
                )

                record["Source_File"] = file_name

                all_records.append(record)

            print(
                f"Processed {file_number}/{len(html_files)}: "
                f"{file_name}"
            )

        except Exception as e:

            skipped_rows.append({
                "Source_File": file_name,
                "Error": str(e)
            })

            print(f"Error processing {file_name}: {e}")


# ==========================================
# 4. CREATE DATAFRAME
# ==========================================

df = pd.DataFrame(all_records)

if df.empty:
    raise ValueError(
        "No sales records were extracted. "
        "Please check the HTML structure."
    )

print("\nExtraction completed!")

print("Total extracted records:", len(df))

print("\nOriginal Data:")
print(df.head())


# ==========================================
# 5. DATA CLEANING
# ==========================================

print("\n" + "=" * 50)
print("DATA CLEANING")
print("=" * 50)


# 5.1 Clean text columns

text_columns = [
    "Date",
    "Permission_Number",
    "Barcode",
    "Product_Code",
    "Product_Name"
]

for col in text_columns:

    df[col] = df[col].apply(clean_text)


# 5.2 Convert numerical columns

numeric_columns = [
    "Price",
    "Quantity",
    "Total_Value"
]

for col in numeric_columns:

    df[col] = df[col].apply(clean_numeric)


# ==========================================
# 6. HANDLE MISSING VALUES
# ==========================================

print("\nMissing Values Before Cleaning:")

print(df.isnull().sum())

# Remove rows missing essential information
essential_columns = [
    "Date",
    "Product_Name",
    "Price",
    "Quantity",
    "Total_Value"
]

before_missing = len(df)

df = df.dropna(
    subset=essential_columns
).copy()

removed_missing = before_missing - len(df)

print("\nRows removed due to missing essential data:",
      removed_missing)


# ==========================================
# 7. DATA VALIDATION
# ==========================================

# Check negative prices
negative_prices = (df["Price"] < 0).sum()

# Check negative quantities
negative_quantities = (df["Quantity"] < 0).sum()

# Check negative sales values
negative_values = (df["Total_Value"] < 0).sum()

# Check zero quantities
zero_quantities = (df["Quantity"] == 0).sum()

print("\nData Validation:")

print("Negative Prices:", negative_prices)

print("Negative Quantities:", negative_quantities)

print("Negative Sales Values:", negative_values)

print("Zero Quantities:", zero_quantities)


# ==========================================
# 8. CHECK DUPLICATES
# ==========================================

# Do not automatically delete duplicates.
# Two legitimate sales lines can have identical values.

duplicate_columns = [
    "Date",
    "Permission_Number",
    "Barcode",
    "Product_Code",
    "Product_Name",
    "Price",
    "Quantity",
    "Total_Value"
]

duplicates = df[
    df.duplicated(
        subset=duplicate_columns,
        keep=False
    )
].copy()

duplicate_count = df.duplicated(
    subset=duplicate_columns
).sum()

print("\nDuplicate Analysis:")

print("Duplicate Records:", duplicate_count)


# ==========================================
# 9. CHECK SALES CALCULATIONS
# ==========================================

# Expected value = Price * Quantity

df["Calculated_Value"] = (
    df["Price"] * df["Quantity"]
)

df["Value_Difference"] = (
    df["Total_Value"] - df["Calculated_Value"]
).round(2)

# Allow a small rounding difference
df["Value_Check"] = np.where(
    df["Value_Difference"].abs() <= 0.01,
    "Match",
    "Mismatch"
)

mismatch_count = (
    df["Value_Check"] == "Mismatch"
).sum()

print("\nSales Calculation Validation:")

print("Mismatched Records:", mismatch_count)


# ==========================================
# 10. CHECK DATE FORMAT
# ==========================================

# Preserve original date as text.
# The source contains shortened years such as 14/02/20.
# Do not assume that 20 means 2020 or 2024.

df["Date_Year_Length"] = df["Date"].apply(
    lambda x: len(str(x).split("/")[-1])
)

short_year_count = (
    df["Date_Year_Length"] == 2
).sum()

print("\nDate Validation:")

print("Dates with two-digit years:", short_year_count)


# ==========================================
# 11. FINAL DATASET
# ==========================================

# Keep source traceability and validation columns.

final_columns = [
    "Date",
    "Permission_Number",
    "Barcode",
    "Product_Code",
    "Product_Name",
    "Price",
    "Quantity",
    "Total_Value",
    "Calculated_Value",
    "Value_Difference",
    "Value_Check",
    "Source_File"
]

cleaned_df = df[final_columns].copy()


# ==========================================
# 12. DATA QUALITY REPORT
# ==========================================

quality_data = []

quality_data.append({
    "Metric": "Total Extracted Records",
    "Value": len(all_records)
})

quality_data.append({
    "Metric": "Final Clean Records",
    "Value": len(cleaned_df)
})

quality_data.append({
    "Metric": "Rows Removed - Missing Essential Values",
    "Value": removed_missing
})

quality_data.append({
    "Metric": "Duplicate Records",
    "Value": duplicate_count
})

quality_data.append({
    "Metric": "Negative Prices",
    "Value": negative_prices
})

quality_data.append({
    "Metric": "Negative Quantities",
    "Value": negative_quantities
})

quality_data.append({
    "Metric": "Negative Sales Values",
    "Value": negative_values
})

quality_data.append({
    "Metric": "Zero Quantities",
    "Value": zero_quantities
})

quality_data.append({
    "Metric": "Sales Value Mismatches",
    "Value": mismatch_count
})

quality_data.append({
    "Metric": "Dates with Two-Digit Years",
    "Value": short_year_count
})

quality_data.append({
    "Metric": "HTML Files Processed",
    "Value": len(html_files)
})

quality_data.append({
    "Metric": "Files with Processing Errors",
    "Value": len(skipped_rows)
})

quality_df = pd.DataFrame(quality_data)


# ==========================================
# 13. MISSING VALUES REPORT
# ==========================================

missing_report = pd.DataFrame({
    "Column": cleaned_df.columns,
    "Missing_Count": [
        cleaned_df[col].isnull().sum()
        for col in cleaned_df.columns
    ],
    "Missing_Percentage": [
        round(
            cleaned_df[col].isnull().mean() * 100,
            2
        )
        for col in cleaned_df.columns
    ]
})


# ==========================================
# 14. PRODUCT SALES SUMMARY
# ==========================================

product_summary = (
    cleaned_df.groupby(
        "Product_Name",
        dropna=False
    )
    .agg(
        Total_Quantity=("Quantity", "sum"),
        Total_Sales=("Total_Value", "sum"),
        Average_Price=("Price", "mean"),
        Number_of_Records=("Product_Name", "count")
    )
    .reset_index()
)

product_summary = product_summary.sort_values(
    by="Total_Sales",
    ascending=False
)


# ==========================================
# 15. EXPORT TO EXCEL
# ==========================================

print("\n" + "=" * 50)
print("EXPORTING DATA TO EXCEL")
print("=" * 50)

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    cleaned_df.to_excel(
        writer,
        sheet_name="Clean_Data",
        index=False
    )

    quality_df.to_excel(
        writer,
        sheet_name="Data_Quality",
        index=False
    )

    missing_report.to_excel(
        writer,
        sheet_name="Missing_Values",
        index=False
    )

    duplicates.to_excel(
        writer,
        sheet_name="Duplicates",
        index=False
    )

    product_summary.to_excel(
        writer,
        sheet_name="Product_Summary",
        index=False
    )

    pd.DataFrame(
        skipped_rows
    ).to_excel(
        writer,
        sheet_name="Processing_Errors",
        index=False
    )

    # Excel formatting
    for sheet in writer.sheets.values():

        sheet.freeze_panes = "A2"

        sheet.auto_filter.ref = sheet.dimensions

        for column_cells in sheet.columns:

            max_length = max(
                len(str(cell.value or ""))
                for cell in column_cells
            )

            column_letter = column_cells[0].column_letter

            sheet.column_dimensions[
                column_letter
            ].width = min(max_length + 3, 40)


# ==========================================
# 16. FINAL RESULTS
# ==========================================

print("\n" + "=" * 50)

print("PROJECT COMPLETED SUCCESSFULLY!")

print("=" * 50)

print("Final Records:", len(cleaned_df))

print("Total Sales:",
      round(cleaned_df["Total_Value"].sum(), 2))

print("Total Quantity:",
      round(cleaned_df["Quantity"].sum(), 2))

print("Unique Products:",
      cleaned_df["Product_Name"].nunique())

print("\nExcel File Saved At:")

print(OUTPUT_FILE)

print("\nAll Done!")
