
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

ZIP_FILE = BASE_DIR / "Net sales.zip"

OUTPUT_FILE = BASE_DIR / "Cottonil_Net_Sales_PowerBI_Ready.xlsx"

UNKNOWN_VALUE = "غير معروف"

STORE_NAME = "قطونيل شيرتون"


# =====================================================
# 2. دوال مساعدة
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

    value = re.sub(
        r"[\u200e\u200f\u202a-\u202e\u00a0]",
        " ",
        value
    )

    value = " ".join(value.split())

    if value.lower() in ["", "nan", "none", "null"]:
        return None

    if value in ["--", "-", "—"]:
        return None

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


def clean_identifier(value):

    value = clean_text(value)

    if value is None:
        return None

    return value


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

        html_text = html_content.decode(
            "cp1256",
            errors="replace"
        )

        soup = BeautifulSoup(
            html_text,
            "lxml"
        )

        # ---------------------------------------------
        # قراءة صفوف الجداول
        # ---------------------------------------------

        previous_record = None

        for row in soup.find_all("tr"):

            cells = row.find_all(["td", "th"])

            values = [
                cell.get_text(" ", strip=True)
                for cell in cells
            ]

            if not values:
                continue

            # -----------------------------------------
            # التعامل مع الباركود الممتد
            # -----------------------------------------

            if len(values) == 2:

                first_value = clean_text(values[0])
                second_value = clean_identifier(values[1])

                if (
                    first_value is None
                    and second_value is not None
                    and second_value.isdigit()
                    and previous_record is not None
                ):

                    old_barcode = previous_record["Barcode"]

                    if old_barcode is None:

                        previous_record["Barcode"] = second_value

                    else:

                        previous_record["Barcode"] = (
                            old_barcode + second_value
                        )

                continue

            # -----------------------------------------
            # تجاهل صفوف العناوين
            # -----------------------------------------

            if "رقم الإذن" in values:

                continue

            # -----------------------------------------
            # تحديد شكل صف المنتج
            # -----------------------------------------

            if len(values) == 17:

                permission_number = values[0]
                product_code = values[2]
                barcode = values[3]
                product_name = values[5]

                price = values[7]
                quantity = values[8]
                discount = values[9]

                taxes = values[11]
                net_price = values[12]
                total_value = values[13]

                notes = values[14]
                size = values[15]
                color = values[16]

            elif len(values) == 16:

                permission_number = values[0]
                product_code = values[2]
                barcode = values[3]
                product_name = values[4]

                price = values[6]
                quantity = values[7]
                discount = values[8]

                taxes = values[10]
                net_price = values[11]
                total_value = values[12]

                notes = values[13]
                size = values[14]
                color = values[15]

            else:

                continue

            # -----------------------------------------
            # التحقق من بيانات المنتج
            # -----------------------------------------

            product_name = clean_text(product_name)

            price = clean_number(price)
            quantity = clean_number(quantity)

            if product_name is None:
                continue

            if price is None or quantity is None:
                continue

            # -----------------------------------------
            # إنشاء السجل
            # -----------------------------------------

            record = {

                "Store": STORE_NAME,

                "Transaction_Type": "مبيعات",

                "Permission_Number": clean_identifier(
                    permission_number
                ),

                "Product_Code": clean_identifier(
                    product_code
                ),

                "Barcode": clean_identifier(
                    barcode
                ),

                "Product_Name": product_name,

                "Price": price,

                "Quantity": quantity,

                "Discount": clean_number(discount),

                "Taxes": clean_number(taxes),

                "Net_Price": clean_number(net_price),

                "Total_Value": clean_number(total_value),

                "Notes": clean_text(notes),

                "Size": clean_text(size),

                "Color": clean_text(color)

            }

            all_records.append(record)

            previous_record = record


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

text_columns = [

    "Store",
    "Transaction_Type",
    "Product_Name",
    "Notes",
    "Size",
    "Color"

]

for column in text_columns:

    df[column] = df[column].apply(clean_text)


# =====================================================
# 7. معالجة القيم الفارغة
# =====================================================

unknown_columns = [

    "Transaction_Type",
    "Product_Name",
    "Notes",
    "Size",
    "Color"

]

for column in unknown_columns:

    df[column] = df[column].fillna(
        UNKNOWN_VALUE
    )


# تثبيت اسم الفرع في جميع الصفوف
df["Store"] = STORE_NAME


# =====================================================
# 8. تنظيف الأكواد والمعرفات
# =====================================================

identifier_columns = [

    "Permission_Number",
    "Product_Code",
    "Barcode"

]

for column in identifier_columns:

    df[column] = (
        df[column]
        .apply(clean_identifier)
        .astype("string")
    )


# =====================================================
# 9. تنظيف الأعمدة الرقمية
# =====================================================

numeric_columns = [

    "Price",
    "Quantity",
    "Discount",
    "Taxes",
    "Net_Price",
    "Total_Value"

]

for column in numeric_columns:

    df[column] = pd.to_numeric(
        df[column],
        errors="coerce"
    )


# =====================================================
# 10. ترتيب الأعمدة النهائية
# =====================================================

final_columns = [

    "Store",
    "Transaction_Type",
    "Permission_Number",
    "Product_Code",
    "Barcode",
    "Product_Name",
    "Price",
    "Quantity",
    "Discount",
    "Taxes",
    "Net_Price",
    "Total_Value",
    "Notes",
    "Size",
    "Color"

]

df = df[final_columns]


# =====================================================
# 11. ترتيب البيانات
# =====================================================

df = df.sort_values(

    by=[
        "Permission_Number",
        "Product_Name"
    ],

    na_position="last"

).reset_index(drop=True)


# =====================================================
# 12. فحص جودة البيانات
# =====================================================

print("\n" + "=" * 50)

print("FINAL DATA QUALITY CHECK")

print("=" * 50)

print("Total Rows:", len(df))

print("Total Columns:", len(df.columns))

print("\nMissing Values:")

print(df.isnull().sum())

print("\nStore Distribution:")

print(df["Store"].value_counts())

print("\nUnique Products:")

print(df["Product_Name"].nunique())

print("\nUnique Product Codes:")

print(df["Product_Code"].nunique())

print("\nUnique Barcodes:")

print(df["Barcode"].nunique())

print("\nTotal Quantity:")

print(df["Quantity"].sum())

print("\nTotal Discount:")

print(df["Discount"].sum())

print("\nTotal Taxes:")

print(df["Taxes"].sum())

print("\nTotal Net Sales Value:")

print(df["Total_Value"].sum())


# =====================================================
# 13. تصدير ملف Excel
# =====================================================

with pd.ExcelWriter(

    OUTPUT_FILE,

    engine="openpyxl"

) as writer:

    df.to_excel(

        writer,

        sheet_name="Net_Sales",

        index=False

    )


# =====================================================
# 14. تنسيق ملف Excel
# =====================================================

workbook = load_workbook(OUTPUT_FILE)

worksheet = workbook["Net_Sales"]

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
    "B": 20,
    "C": 18,
    "D": 18,
    "E": 20,
    "F": 35,
    "G": 15,
    "H": 15,
    "I": 15,
    "J": 15,
    "K": 15,
    "L": 18,
    "M": 25,
    "N": 15,
    "O": 18

}

for column, width in column_widths.items():

    worksheet.column_dimensions[column].width = width


# تنسيق الأسعار والقيم
for column in ["G", "I", "J", "K", "L"]:

    for cell in worksheet[column][1:]:

        cell.number_format = '#,##0.00'


# تنسيق الكمية
for cell in worksheet["H"][1:]:

    cell.number_format = '#,##0.00'


workbook.save(OUTPUT_FILE)


# =====================================================
# 15. النتيجة النهائية
# =====================================================

print("\n" + "=" * 50)

print("CLEANING COMPLETED SUCCESSFULLY")

print("=" * 50)

print("Output File:")

print(OUTPUT_FILE)

print("Rows:", len(df))

print("Columns:", len(df.columns))

print("Store:", STORE_NAME)

print("\nNet Sales data is ready for Power BI.")
