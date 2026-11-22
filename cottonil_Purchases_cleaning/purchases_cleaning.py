
import zipfile
import re
import pandas as pd
import numpy as np

from pathlib import Path
from bs4 import BeautifulSoup


# ==================================================
# 1. PROJECT SETTINGS
# ==================================================

BASE_DIR = Path(__file__).parent

ZIP_FILE = BASE_DIR / "المشتريات.zip"

OUTPUT_FILE = BASE_DIR / "Cottonil_Purchases_Cleaned.xlsx"


# ==================================================
# 2. HELPER FUNCTIONS
# ==================================================

def clean_text(value):

    if value is None or pd.isna(value):
        return np.nan

    value = str(value).strip()

    if value.lower() in ["nan", "none", "null", "nat", ""]:
        return np.nan

    value = re.sub(r"\s+", " ", value)

    return value


def clean_numeric(value):

    if value is None or pd.isna(value):
        return np.nan

    value = str(value).strip()

    arabic_digits = "٠١٢٣٤٥٦٧٨٩"
    english_digits = "0123456789"

    translation = str.maketrans(
        arabic_digits,
        english_digits
    )

    value = value.translate(translation)

    value = value.replace(",", "").replace(" ", "")

    if value.lower() in ["nan", "none", "null", "", "--", "-"]:
        return np.nan

    value = re.sub(r"[^0-9.\-]", "", value)

    try:
        return float(value)

    except ValueError:
        return np.nan


def get_page_number(filename):

    match = re.search(
        r"Page(\d+)",
        filename,
        re.IGNORECASE
    )

    if match:
        return int(match.group(1))

    return 1


def get_row_values(row):

    return [
        cell.get_text(" ", strip=True)
        for cell in row.find_all(
            ["td", "th"],
            recursive=False
        )
    ]


# ==================================================
# 3. EXTRACT PURCHASE DATA
# ==================================================

print("=" * 60)
print("COTTONIL PURCHASE DATA EXTRACTION")
print("=" * 60)

all_records = []

processing_errors = []

with zipfile.ZipFile(ZIP_FILE, "r") as zip_ref:

    html_files = [
        name for name in zip_ref.namelist()
        if name.lower().endswith((".html", ".htm"))
    ]

    # Sort pages numerically
    html_files = sorted(
        html_files,
        key=get_page_number
    )

    print("\nTotal HTML Files:", len(html_files))

    # Metadata from the previous page
    current_metadata = {}

    pages_with_metadata = 0
    pages_without_metadata = 0

    for file_number, file_name in enumerate(
        html_files,
        start=1
    ):

        try:

            html_content = zip_ref.read(
                file_name
            ).decode(
                "cp1256",
                errors="replace"
            )

            soup = BeautifulSoup(
                html_content,
                "lxml"
            )

            rows = soup.find_all("tr")

            page_metadata = {}

            # --------------------------------------
            # Extract Purchase Document Information
            # --------------------------------------

            for row in rows:

                values = get_row_values(row)

                if len(values) >= 12 and "رقم إذن" in values:

                    page_metadata = {

                        "Permission_Number": values[3],

                        "Purchase_Date": values[6],

                        "Supplier": values[10],

                        "Supplier_Code": values[11]

                    }

                    break

            # Update metadata when a new document starts
            if page_metadata:

                current_metadata = page_metadata.copy()

                pages_with_metadata += 1

                metadata_inherited = False

            else:

                pages_without_metadata += 1

                metadata_inherited = True

            # --------------------------------------
            # Extract Additional Information
            # --------------------------------------

            purchase_type = np.nan

            warehouse = np.nan

            storekeeper = np.nan

            for row in rows:

                values = get_row_values(row)

                if len(values) == 4 and "الشراء" in values:

                    purchase_type = clean_text(values[3])

                if len(values) == 7 and "المخزن" in values:

                    warehouse = clean_text(values[3])

                    storekeeper = clean_text(values[6])

            # --------------------------------------
            # Extract Product Rows
            # --------------------------------------

            for row in rows:

                values = get_row_values(row)

                # Valid product rows contain 11 or 12 cells
                if len(values) not in [11, 12]:
                    continue

                barcode = values[1]

                product_code = values[2]

                product_name = values[3]

                department = values[4]

                size = values[5]

                color = values[6]

                purchase_price = values[7]

                retail_price = values[8]

                discount = values[9]

                quantity = values[10]

                # Internal Code is optional
                if len(values) == 12:

                    internal_code = values[11]

                else:

                    internal_code = np.nan

                # Validate product row
                if not clean_text(product_name) == clean_text(product_name):
                    continue

                if pd.isna(clean_text(product_name)):
                    continue

                if pd.isna(clean_numeric(purchase_price)):
                    continue

                if pd.isna(clean_numeric(quantity)):
                    continue

                record = {

                    **current_metadata,

                    "Purchase_Type": purchase_type,

                    "Barcode": barcode,

                    "Product_Code": product_code,

                    "Product_Name": product_name,

                    "Department": department,

                    "Size": size,

                    "Color": color,

                    "Purchase_Price": purchase_price,

                    "Retail_Price": retail_price,

                    "Discount": discount,

                    "Quantity": quantity,

                    "Internal_Code": internal_code,

                    "Warehouse": warehouse,

                    "Storekeeper": storekeeper,

                    "Source_File": file_name,

                    "Page_Number": get_page_number(file_name),

                    "Metadata_Inherited": metadata_inherited

                }

                all_records.append(record)

            print(
                f"Processed {file_number}/{len(html_files)}: "
                f"{file_name}"
            )

        except Exception as e:

            processing_errors.append({

                "Source_File": file_name,

                "Error": str(e)

            })

            print(
                f"Error processing {file_name}: {e}"
            )


# ==================================================
# 4. CREATE DATAFRAME
# ==================================================

df = pd.DataFrame(all_records)

if df.empty:

    raise ValueError(
        "No purchase records were extracted. "
        "Please check the HTML structure."
    )

print("\nExtraction Completed!")

print("Total Extracted Records:", len(df))

print("\nFirst 5 Records:")

print(df.head())


# ==================================================
# 5. CLEAN TEXT COLUMNS
# ==================================================

print("\n" + "=" * 60)
print("TEXT DATA CLEANING")
print("=" * 60)

text_columns = [

    "Permission_Number",

    "Purchase_Date",

    "Supplier",

    "Supplier_Code",

    "Purchase_Type",

    "Barcode",

    "Product_Code",

    "Product_Name",

    "Department",

    "Size",

    "Color",

    "Internal_Code",

    "Warehouse",

    "Storekeeper"

]

for col in text_columns:

    df[col] = df[col].apply(clean_text)


# ==================================================
# 6. FIX IDENTIFIER COLUMNS
# ==================================================

identifier_columns = [

    "Permission_Number",

    "Supplier_Code",

    "Barcode",

    "Product_Code",

    "Internal_Code"

]

for col in identifier_columns:

    df[col] = df[col].astype("string")

    df[col] = df[col].replace(

        ["nan", "None", "null", ""],

        pd.NA

    )


# ==================================================
# 7. CLEAN NUMERIC COLUMNS
# ==================================================

numeric_columns = [

    "Purchase_Price",

    "Retail_Price",

    "Discount",

    "Quantity"

]

for col in numeric_columns:

    df[col] = df[col].apply(clean_numeric)


# ==================================================
# 8. DATE CLEANING
# ==================================================

# Preserve original date before conversion

df["Purchase_Date_Original"] = df["Purchase_Date"]

df["Purchase_Date"] = pd.to_datetime(

    df["Purchase_Date"],

    format="%Y/%m/%d",

    errors="coerce"

)


# ==================================================
# 9. HANDLE MISSING VALUES
# ==================================================

print("\nMissing Values Before Cleaning:")

print(df.isnull().sum())

# Do not automatically delete incomplete purchase records.
# Keep them for investigation.

# Replace textual missing values in all text fields
for col in text_columns:

    df[col] = df[col].replace(

        ["nan", "None", "null", ""],

        pd.NA

    )


# ==================================================
# 10. DATA VALIDATION
# ==================================================

print("\n" + "=" * 60)
print("DATA VALIDATION")
print("=" * 60)

negative_purchase_prices = (

    df["Purchase_Price"] < 0

).sum()

negative_retail_prices = (

    df["Retail_Price"] < 0

).sum()

negative_quantities = (

    df["Quantity"] < 0

).sum()

zero_quantities = (

    df["Quantity"] == 0

).sum()

missing_dates = df["Purchase_Date"].isnull().sum()

missing_permission_numbers = (

    df["Permission_Number"].isnull()

).sum()

print(
    "Negative Purchase Prices:",
    negative_purchase_prices
)

print(
    "Negative Retail Prices:",
    negative_retail_prices
)

print(
    "Negative Quantities:",
    negative_quantities
)

print(
    "Zero Quantities:",
    zero_quantities
)

print(
    "Invalid or Missing Dates:",
    missing_dates
)

print(
    "Missing Permission Numbers:",
    missing_permission_numbers
)


# ==================================================
# 11. CALCULATE PURCHASE VALUES
# ==================================================

# Estimated gross purchase value before discount.
# Discount is kept separate because its exact calculation
# rule is not established by the source report.

df["Gross_Purchase_Value"] = (

    df["Purchase_Price"] * df["Quantity"]

).round(2)


# ==================================================
# 12. DUPLICATE ANALYSIS
# ==================================================

duplicate_columns = [

    "Permission_Number",

    "Purchase_Date",

    "Supplier",

    "Barcode",

    "Product_Code",

    "Product_Name",

    "Department",

    "Size",

    "Color",

    "Purchase_Price",

    "Retail_Price",

    "Discount",

    "Quantity"

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

print("\nDuplicate Records:", duplicate_count)

# Do not delete duplicates automatically.


# ==================================================
# 13. MISSING VALUES REPORT
# ==================================================

missing_report = pd.DataFrame({

    "Column": df.columns,

    "Missing_Count": [

        df[col].isnull().sum()

        for col in df.columns

    ],

    "Missing_Percentage": [

        round(

            df[col].isnull().mean() * 100,

            2

        )

        for col in df.columns

    ]

})


# ==================================================
# 14. PRODUCT PURCHASE SUMMARY
# ==================================================

product_summary = (

    df.groupby(

        ["Product_Name", "Department"],

        dropna=False

    )

    .agg(

        Total_Quantity=(

            "Quantity",

            "sum"

        ),

        Total_Purchase_Value=(

            "Gross_Purchase_Value",

            "sum"

        ),

        Average_Purchase_Price=(

            "Purchase_Price",

            "mean"

        ),

        Average_Retail_Price=(

            "Retail_Price",

            "mean"

        ),

        Number_Of_Records=(

            "Product_Name",

            "count"

        )

    )

    .reset_index()

)

product_summary = product_summary.sort_values(

    by="Total_Purchase_Value",

    ascending=False

)


# ==================================================
# 15. SUPPLIER SUMMARY
# ==================================================

supplier_summary = (

    df.groupby(

        "Supplier",

        dropna=False

    )

    .agg(

        Total_Quantity=(

            "Quantity",

            "sum"

        ),

        Total_Purchase_Value=(

            "Gross_Purchase_Value",

            "sum"

        ),

        Average_Purchase_Price=(

            "Purchase_Price",

            "mean"

        ),

        Number_Of_Records=(

            "Product_Name",

            "count"

        ),

        Number_Of_Purchase_Documents=(

            "Permission_Number",

            "nunique"

        )

    )

    .reset_index()

)

supplier_summary = supplier_summary.sort_values(

    by="Total_Purchase_Value",

    ascending=False

)


# ==================================================
# 16. PURCHASE DOCUMENT SUMMARY
# ==================================================

purchase_summary = (

    df.groupby(

        [

            "Permission_Number",

            "Purchase_Date",

            "Supplier"

        ],

        dropna=False

    )

    .agg(

        Total_Quantity=(

            "Quantity",

            "sum"

        ),

        Total_Purchase_Value=(

            "Gross_Purchase_Value",

            "sum"

        ),

        Number_Of_Products=(

            "Product_Name",

            "count"

        )

    )

    .reset_index()

)

purchase_summary = purchase_summary.sort_values(

    by="Purchase_Date"

)


# ==================================================
# 17. DATA QUALITY REPORT
# ==================================================

quality_data = [

    {

        "Metric": "Total HTML Files",

        "Value": len(html_files)

    },

    {

        "Metric": "Pages With Metadata",

        "Value": pages_with_metadata

    },

    {

        "Metric": "Pages Without Metadata",

        "Value": pages_without_metadata

    },

    {

        "Metric": "Total Extracted Records",

        "Value": len(all_records)

    },

    {

        "Metric": "Final Records",

        "Value": len(df)

    },

    {

        "Metric": "Duplicate Records",

        "Value": duplicate_count

    },

    {

        "Metric": "Missing Dates",

        "Value": missing_dates

    },

    {

        "Metric": "Missing Permission Numbers",

        "Value": missing_permission_numbers

    },

    {

        "Metric": "Negative Purchase Prices",

        "Value": negative_purchase_prices

    },

    {

        "Metric": "Negative Retail Prices",

        "Value": negative_retail_prices

    },

    {

        "Metric": "Negative Quantities",

        "Value": negative_quantities

    },

    {

        "Metric": "Zero Quantities",

        "Value": zero_quantities

    },

    {

        "Metric": "Processing Errors",

        "Value": len(processing_errors)

    }

]

quality_df = pd.DataFrame(quality_data)


# ==================================================
# 18. FINAL DATASET
# ==================================================

final_columns = [

    "Purchase_Date",

    "Permission_Number",

    "Supplier",

    "Supplier_Code",

    "Purchase_Type",

    "Barcode",

    "Product_Code",

    "Product_Name",

    "Department",

    "Size",

    "Color",

    "Purchase_Price",

    "Retail_Price",

    "Discount",

    "Quantity",

    "Gross_Purchase_Value",

    "Internal_Code",

    "Warehouse",

    "Storekeeper",

    "Source_File",

    "Page_Number",

    "Metadata_Inherited"

]

cleaned_df = df[final_columns].copy()


# ==================================================
# 19. EXPORT TO EXCEL
# ==================================================

print("\n" + "=" * 60)
print("EXPORTING TO EXCEL")
print("=" * 60)

with pd.ExcelWriter(

    OUTPUT_FILE,

    engine="openpyxl",

    date_format="yyyy-mm-dd"

) as writer:

    # --------------------------------------
    # Main Data
    # --------------------------------------

    cleaned_df.to_excel(

        writer,

        sheet_name="Clean_Data",

        index=False

    )

    # --------------------------------------
    # Data Quality
    # --------------------------------------

    quality_df.to_excel(

        writer,

        sheet_name="Data_Quality",

        index=False

    )

    # --------------------------------------
    # Missing Values
    # --------------------------------------

    missing_report.to_excel(

        writer,

        sheet_name="Missing_Values",

        index=False

    )

    # --------------------------------------
    # Duplicates
    # --------------------------------------

    duplicates.to_excel(

        writer,

        sheet_name="Duplicates",

        index=False

    )

    # --------------------------------------
    # Product Summary
    # --------------------------------------

    product_summary.to_excel(

        writer,

        sheet_name="Product_Summary",

        index=False

    )

    # --------------------------------------
    # Supplier Summary
    # --------------------------------------

    supplier_summary.to_excel(

        writer,

        sheet_name="Supplier_Summary",

        index=False

    )

    # --------------------------------------
    # Purchase Summary
    # --------------------------------------

    purchase_summary.to_excel(

        writer,

        sheet_name="Purchase_Summary",

        index=False

    )

    # --------------------------------------
    # Processing Errors
    # --------------------------------------

    pd.DataFrame(

        processing_errors,

        columns=["Source_File", "Error"]

    ).to_excel(

        writer,

        sheet_name="Processing_Errors",

        index=False

    )

    # --------------------------------------
    # Excel Formatting
    # --------------------------------------

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

            ].width = min(

                max_length + 3,

                40

            )

    # --------------------------------------
    # Identifier Formatting
    # --------------------------------------

    ws = writer.sheets["Clean_Data"]

    for col_name in identifier_columns:

        col_index = cleaned_df.columns.get_loc(

            col_name

        ) + 1

        for row in range(2, ws.max_row + 1):

            cell = ws.cell(

                row=row,

                column=col_index

            )

            cell.number_format = "@"

    # --------------------------------------
    # Numeric Formatting
    # --------------------------------------

    numeric_excel_columns = [

        "Purchase_Price",

        "Retail_Price",

        "Discount",

        "Gross_Purchase_Value"

    ]

    for col_name in numeric_excel_columns:

        col_index = cleaned_df.columns.get_loc(

            col_name

        ) + 1

        for row in range(2, ws.max_row + 1):

            ws.cell(

                row=row,

                column=col_index

            ).number_format = '#,##0.00'

    # Quantity formatting

    quantity_index = (

        cleaned_df.columns.get_loc("Quantity") + 1

    )

    for row in range(2, ws.max_row + 1):

        ws.cell(

            row=row,

            column=quantity_index

        ).number_format = '#,##0.##'


# ==================================================
# 20. FINAL RESULTS
# ==================================================

print("\n" + "=" * 60)

print("PURCHASE DATA CLEANING COMPLETED!")

print("=" * 60)

print("\nTotal Records:", len(cleaned_df))

print(

    "Total Quantity:",

    round(cleaned_df["Quantity"].sum(), 2)

)

print(

    "Estimated Gross Purchase Value:",

    round(

        cleaned_df["Gross_Purchase_Value"].sum(),

        2

    )

)

print(

    "Unique Products:",

    cleaned_df["Product_Name"].nunique()

)

print(

    "Unique Suppliers:",

    cleaned_df["Supplier"].nunique()

)

print(

    "Unique Purchase Documents:",

    cleaned_df["Permission_Number"].nunique()

)

print("\nExcel File Saved At:")

print(OUTPUT_FILE)

print("\nAll Done!")
