
import zipfile
import re
import pandas as pd

from pathlib import Path
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment


# =====================================================
# 1. إعدادات المشروع
# =====================================================

BASE_DIR = Path(__file__).resolve().parent

ZIP_FILE = BASE_DIR / "supplier.zip"

OUTPUT_FILE = BASE_DIR / "Cottonil_Suppliers_PowerBI_Ready.xlsx"

UNKNOWN_VALUE = "غير معروف"


# =====================================================
# 2. دوال تنظيف البيانات
# =====================================================

def natural_sort_key(value):

    return [
        int(part) if part.isdigit() else part.lower()
        for part in re.split(r"(\d+)", value)
    ]


def clean_text(value):

    if value is None:
        return None

    value = str(value).strip()

    if value.lower() in ["", "nan", "none", "null"]:
        return None

    if value in ["--", "-", "—"]:
        return None

    # إزالة الرموز والمسافات المخفية
    value = re.sub(
        r"[\u200e\u200f\u202a-\u202e\u00a0]",
        " ",
        value
    )

    # إزالة المسافات الزائدة
    value = " ".join(value.split())

    return value if value else None


def clean_supplier(value):

    value = clean_text(value)

    if value is None:
        return None

    # حذف عبارة للفترة وما بعدها من نهاية اسم المورد
    value = re.sub(
        r"\s*للفترة\b.*$",
        "",
        value
    )

    # تنظيف المسافات بعد الحذف
    value = " ".join(value.split())

    return value if value else None


def clean_number(value):

    value = clean_text(value)

    if value is None:
        return None

    value = value.replace(",", "")

    try:
        return float(value)

    except (ValueError, TypeError):
        return None


# =====================================================
# 3. التأكد من وجود ملف ZIP
# =====================================================

if not ZIP_FILE.exists():

    raise FileNotFoundError(
        f"لم يتم العثور على الملف: {ZIP_FILE}"
    )


# =====================================================
# 4. استخراج البيانات من ملفات HTML
# =====================================================

all_records = []

with zipfile.ZipFile(ZIP_FILE, "r") as zip_file:

    html_files = [
        name for name in zip_file.namelist()
        if name.lower().endswith((".html", ".htm"))
    ]

    html_files = sorted(
        html_files,
        key=natural_sort_key
    )

    print("Number of HTML files:", len(html_files))

    for file_number, file_name in enumerate(
        html_files,
        start=1
    ):

        print(
            f"Processing {file_number}/{len(html_files)}: {file_name}"
        )

        html_content = zip_file.read(file_name)

        # قراءة الملفات بترميز Windows-1256
        html_text = html_content.decode(
            "cp1256",
            errors="replace"
        )

        soup = BeautifulSoup(
            html_text,
            "lxml"
        )

        # ---------------------------------------------
        # 4.1 استخراج اسم الفرع
        # ---------------------------------------------

        store = None

        for table in soup.find_all("table"):

            table_text = table.get_text(" ", strip=True)

            if "قطونيل شيراتون" in table_text:

                store = "قطونيل شيراتون"
                break

        # ---------------------------------------------
        # 4.2 استخراج اسم المورد
        # ---------------------------------------------

        supplier = None

        for table in soup.find_all("table"):

            table_text = table.get_text(" ", strip=True)

            match = re.search(
                r"للمورد\s+(.+)",
                table_text
            )

            if match:

                supplier = clean_supplier(match.group(1))
                break

        # ---------------------------------------------
        # 4.3 استخراج بيانات المنتجات
        # ---------------------------------------------

        for table in soup.find_all("table"):

            for row in table.find_all("tr"):

                cells = row.find_all(["td", "th"])

                values = [
                    cell.get_text(" ", strip=True)
                    for cell in cells
                ]

                if len(values) != 12:
                    continue

                product_name = clean_text(values[1])

                price = clean_number(values[2])

                product_code = clean_text(values[3])

                if product_name is None:
                    continue

                if price is None:
                    continue

                record = {

                    "Store": store,

                    "Supplier": supplier,

                    "Product_Name": product_name,

                    "Price": price,

                    "Product_Code": product_code,

                    "Purchase_Quantity": clean_number(values[4]),

                    "Purchase_Return_Quantity": clean_number(values[5]),

                    "Sales_Quantity": clean_number(values[6]),

                    "Sales_Return_Quantity": clean_number(values[7]),

                    "Net_Sales_Quantity": clean_number(values[8]),

                    "Net_Sales_Value": clean_number(values[9]),

                    "Stock_Quantity": clean_number(values[10]),

                    "Stock_Value": clean_number(values[11])

                }

                all_records.append(record)


# =====================================================
# 5. تحويل البيانات إلى DataFrame
# =====================================================

df = pd.DataFrame(all_records)

if df.empty:

    raise ValueError(
        "لم يتم استخراج أي بيانات. راجع ملفات HTML."
    )

print("\nData extraction completed.")

print("Extracted Rows:", len(df))


# =====================================================
# 6. تنظيف الأعمدة النصية
# =====================================================

df["Store"] = df["Store"].apply(clean_text)

df["Supplier"] = df["Supplier"].apply(clean_supplier)

df["Product_Name"] = df["Product_Name"].apply(clean_text)


# =====================================================
# 7. معالجة القيم الفارغة
# =====================================================

text_columns = [
    "Store",
    "Supplier",
    "Product_Name"
]

for column in text_columns:

    df[column] = df[column].fillna(
        UNKNOWN_VALUE
    )


# =====================================================
# 8. تنظيف الأكواد
# =====================================================

df["Product_Code"] = (
    df["Product_Code"]
    .apply(clean_text)
    .astype("string")
)


# =====================================================
# 9. تنظيف الأعمدة الرقمية
# =====================================================

numeric_columns = [

    "Price",

    "Purchase_Quantity",

    "Purchase_Return_Quantity",

    "Sales_Quantity",

    "Sales_Return_Quantity",

    "Net_Sales_Quantity",

    "Net_Sales_Value",

    "Stock_Quantity",

    "Stock_Value"

]

for column in numeric_columns:

    df[column] = pd.to_numeric(
        df[column],
        errors="coerce"
    )


# =====================================================
# 10. التأكد من تنظيف أسماء الموردين
# =====================================================

df["Supplier"] = df["Supplier"].apply(clean_supplier)

print("\nSupplier Names After Cleaning:")

print(df["Supplier"].value_counts())


# =====================================================
# 11. ترتيب الأعمدة النهائية
# =====================================================

final_columns = [

    "Store",

    "Supplier",

    "Product_Name",

    "Product_Code",

    "Price",

    "Purchase_Quantity",

    "Purchase_Return_Quantity",

    "Sales_Quantity",

    "Sales_Return_Quantity",

    "Net_Sales_Quantity",

    "Net_Sales_Value",

    "Stock_Quantity",

    "Stock_Value"

]

df = df[final_columns]


# =====================================================
# 12. ترتيب البيانات
# =====================================================

df = df.sort_values(

    by=[
        "Supplier",
        "Product_Name",
        "Product_Code"
    ],

    na_position="last"

).reset_index(drop=True)


# =====================================================
# 13. فحص جودة البيانات
# =====================================================

print("\n" + "=" * 50)

print("FINAL DATA QUALITY CHECK")

print("=" * 50)

print("Total Rows:", len(df))

print("Total Columns:", len(df.columns))

print("\nMissing Values:")

print(df.isnull().sum())

print("\nSuppliers:")

print(df["Supplier"].value_counts())

print("\nUnique Products:")

print(df["Product_Name"].nunique())

print("\nTotal Purchase Quantity:")

print(df["Purchase_Quantity"].sum())

print("\nTotal Sales Quantity:")

print(df["Sales_Quantity"].sum())

print("\nTotal Net Sales Quantity:")

print(df["Net_Sales_Quantity"].sum())

print("\nTotal Net Sales Value:")

print(df["Net_Sales_Value"].sum())

print("\nTotal Stock Quantity:")

print(df["Stock_Quantity"].sum())

print("\nTotal Stock Value:")

print(df["Stock_Value"].sum())


# =====================================================
# 14. تصدير ملف Excel
# =====================================================

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    df.to_excel(
        writer,
        sheet_name="Suppliers",
        index=False
    )


# =====================================================
# 15. تنسيق ملف Excel
# =====================================================

workbook = load_workbook(OUTPUT_FILE)

worksheet = workbook["Suppliers"]

worksheet.freeze_panes = "A2"

worksheet.auto_filter.ref = worksheet.dimensions


# تنسيق العناوين
header_fill = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)

header_font = Font(
    bold=True,
    color="FFFFFF"
)

for cell in worksheet[1]:

    cell.fill = header_fill
    cell.font = header_font

    cell.alignment = Alignment(
        horizontal="center",
        vertical="center"
    )


# ضبط عرض الأعمدة
column_widths = {

    "A": 22,
    "B": 22,
    "C": 35,
    "D": 18,
    "E": 15,
    "F": 22,
    "G": 25,
    "H": 20,
    "I": 25,
    "J": 22,
    "K": 22,
    "L": 18,
    "M": 18

}

for column, width in column_widths.items():

    worksheet.column_dimensions[column].width = width


# تنسيق الأسعار والقيم
for column in ["E", "K", "M"]:

    for cell in worksheet[column][1:]:

        cell.number_format = '#,##0.00'


# تنسيق الكميات
for column in ["F", "G", "H", "I", "J", "L"]:

    for cell in worksheet[column][1:]:

        cell.number_format = '#,##0.00'


workbook.save(OUTPUT_FILE)


# =====================================================
# 16. النتيجة النهائية
# =====================================================

print("\n" + "=" * 50)

print("CLEANING COMPLETED SUCCESSFULLY")

print("=" * 50)

print("Output File:")

print(OUTPUT_FILE)

print("Rows:", len(df))

print("Columns:", len(df.columns))

print("\nSupplier data is ready for Power BI.")
