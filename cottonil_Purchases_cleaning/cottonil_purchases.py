
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

ZIP_FILE = BASE_DIR / "المشتريات.zip"
OUTPUT_FILE = BASE_DIR / "Cottonil_Purchases_PowerBI_Ready.xlsx"

UNKNOWN_VALUE = "غير معروف"


# =====================================================
# 2. دوال مساعدة
# =====================================================

def natural_sort_key(value):
    """
    ترتيب أسماء الملفات ترتيبًا طبيعيًا:
    Page2 قبل Page10
    """
    return [
        int(part) if part.isdigit() else part.lower()
        for part in re.split(r"(\d+)", value)
    ]


def clean_text(value):
    """
    تنظيف النصوص وتحويل القيم الفارغة إلى None.
    """
    if value is None:
        return None

    value = str(value).strip()

    if value.lower() in ["", "nan", "none", "null"]:
        return None

    if value in ["--", "-", "—"]:
        return None

    return value


def clean_number(value):
    """
    تحويل الأرقام إلى قيم رقمية.
    """
    value = clean_text(value)

    if value is None:
        return None

    value = value.replace(",", "")

    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def is_number(value):
    """
    التحقق من أن القيمة رقمية.
    """
    return clean_number(value) is not None


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

# بيانات إذن الشراء الحالية
current_metadata = {
    "Permission_Number": None,
    "Purchase_Date": None,
    "Supplier": None,
    "Supplier_Code": None,
    "Purchase_Type": None
}


with zipfile.ZipFile(ZIP_FILE, "r") as zip_file:

    html_files = [
        name for name in zip_file.namelist()
        if name.lower().endswith((".html", ".htm"))
    ]

    html_files = sorted(html_files, key=natural_sort_key)

    print(f"عدد ملفات HTML: {len(html_files)}")

    for file_number, file_name in enumerate(html_files, start=1):

        print(
            f"Processing {file_number}/{len(html_files)}: {file_name}"
        )

        html_content = zip_file.read(file_name)

        # الملفات محفوظة بترميز Windows-1256
        html_text = html_content.decode(
            "cp1256",
            errors="replace"
        )

        soup = BeautifulSoup(html_text, "lxml")

        for table in soup.find_all("table"):

            for row in table.find_all("tr"):

                cells = row.find_all(["td", "th"])

                values = [
                    cell.get_text(" ", strip=True)
                    for cell in cells
                ]

                if not values:
                    continue

                row_text = " ".join(values)

                # -----------------------------------------
                # 4.1 استخراج بيانات إذن الشراء
                # -----------------------------------------

                if "رقم إذن" in row_text:

                    if len(values) >= 12:

                        current_metadata = {
                            "Permission_Number": clean_text(values[3]),
                            "Purchase_Date": clean_text(values[6]),
                            "Supplier": clean_text(values[10]),
                            "Supplier_Code": clean_text(values[11]),
                            "Purchase_Type": None
                        }

                    continue

                # -----------------------------------------
                # 4.2 استخراج نوع الشراء
                # -----------------------------------------

                if len(values) >= 4 and "الشراء" in row_text:

                    current_metadata["Purchase_Type"] = clean_text(
                        values[3]
                    )

                    continue

                # -----------------------------------------
                # 4.3 استخراج تفاصيل المنتجات
                # -----------------------------------------

                # صف المنتج يتكون عادة من 11 أو 12 خلية
                if len(values) not in [11, 12]:
                    continue

                # الباركود موجود في الخلية الثانية
                barcode = clean_text(values[1])

                # نتأكد أن الصف يحتوي على بيانات منتج
                if not is_number(barcode):
                    continue

                product_name = clean_text(values[3])

                purchase_price = clean_number(values[7])
                retail_price = clean_number(values[8])
                discount = clean_number(values[9])
                quantity = clean_number(values[10])

                # تجاهل الصفوف التي لا تحتوي على بيانات منتج صحيحة
                if product_name is None:
                    continue

                if purchase_price is None or quantity is None:
                    continue

                # -----------------------------------------
                # 4.4 إضافة السجل
                # -----------------------------------------

                record = {
                    "Purchase_Date": current_metadata["Purchase_Date"],
                    "Permission_Number": current_metadata["Permission_Number"],
                    "Supplier": current_metadata["Supplier"],
                    "Supplier_Code": current_metadata["Supplier_Code"],
                    "Purchase_Type": current_metadata["Purchase_Type"],

                    "Barcode": barcode,
                    "Product_Name": product_name,

                    "Department": clean_text(values[4]),
                    "Size": clean_text(values[5]),

                    "Purchase_Price": purchase_price,
                    "Retail_Price": retail_price,
                    "Discount": discount,
                    "Quantity": quantity
                }

                all_records.append(record)


# =====================================================
# 5. تحويل البيانات إلى DataFrame
# =====================================================

df = pd.DataFrame(all_records)

if df.empty:
    raise ValueError(
        "لم يتم استخراج أي سجلات. راجع ملفات HTML وطريقة قراءة الجداول."
    )

print("\nتم استخراج البيانات بنجاح.")
print("عدد السجلات:", len(df))


# =====================================================
# 6. تنظيف النصوص
# =====================================================

text_columns = [
    "Supplier",
    "Purchase_Type",
    "Product_Name",
    "Department",
    "Size"
]

for column in text_columns:

    df[column] = df[column].apply(clean_text)


# =====================================================
# 7. معالجة القيم الفارغة في Department و Size
# =====================================================

df["Department"] = df["Department"].fillna(UNKNOWN_VALUE)

df["Size"] = df["Size"].fillna(UNKNOWN_VALUE)


# =====================================================
# 8. تنظيف التواريخ
# =====================================================

df["Purchase_Date"] = pd.to_datetime(
    df["Purchase_Date"],
    format="%Y/%m/%d",
    errors="coerce"
)


# =====================================================
# 9. تنظيف الأعمدة الرقمية
# =====================================================

numeric_columns = [
    "Purchase_Price",
    "Retail_Price",
    "Discount",
    "Quantity"
]

for column in numeric_columns:

    df[column] = pd.to_numeric(
        df[column],
        errors="coerce"
    )


# =====================================================
# 10. تجهيز أعمدة التعريف كـ Text
# =====================================================

identifier_columns = [
    "Permission_Number",
    "Supplier_Code",
    "Barcode"
]

for column in identifier_columns:

    df[column] = (
        df[column]
        .apply(clean_text)
        .astype("string")
    )


# =====================================================
# 11. حساب إجمالي قيمة الشراء
# =====================================================

df["Gross_Purchase_Value"] = (
    df["Purchase_Price"] * df["Quantity"]
).round(2)


# =====================================================
# 12. ترتيب الأعمدة النهائية
# =====================================================

final_columns = [
    "Purchase_Date",
    "Permission_Number",
    "Supplier",
    "Supplier_Code",
    "Purchase_Type",
    "Barcode",
    "Product_Name",
    "Department",
    "Size",
    "Purchase_Price",
    "Retail_Price",
    "Discount",
    "Quantity",
    "Gross_Purchase_Value"
]

df = df[final_columns]


# =====================================================
# 13. التأكد من أنواع البيانات
# =====================================================

# التاريخ
df["Purchase_Date"] = pd.to_datetime(
    df["Purchase_Date"],
    errors="coerce"
)

# الأرقام
for column in numeric_columns + ["Gross_Purchase_Value"]:

    df[column] = pd.to_numeric(
        df[column],
        errors="coerce"
    )


# =====================================================
# 14. ترتيب البيانات
# =====================================================

df = df.sort_values(
    by=["Purchase_Date", "Permission_Number"],
    na_position="last"
).reset_index(drop=True)


# =====================================================
# 15. إنشاء ملخص بسيط لجودة البيانات
# =====================================================

print("\n" + "=" * 50)
print("FINAL DATA QUALITY CHECK")
print("=" * 50)

print("Total Rows:", len(df))
print("Total Columns:", len(df.columns))

print("\nMissing Values:")
print(df.isnull().sum())

print("\nDepartment Distribution:")
print(df["Department"].value_counts(dropna=False))

print("\nSize Distribution:")
print(df["Size"].value_counts(dropna=False))

print("\nTotal Purchase Quantity:")
print(df["Quantity"].sum())

print("\nTotal Gross Purchase Value:")
print(round(df["Gross_Purchase_Value"].sum(), 2))


# =====================================================
# 16. تصدير ملف Excel
# =====================================================

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl",
    datetime_format="yyyy-mm-dd"
) as writer:

    df.to_excel(
        writer,
        sheet_name="Purchases",
        index=False
    )


# =====================================================
# 17. تنسيق ملف Excel
# =====================================================

workbook = load_workbook(OUTPUT_FILE)

worksheet = workbook["Purchases"]

# تثبيت الصف الأول
worksheet.freeze_panes = "A2"

# تفعيل الفلاتر
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
    "A": 15,
    "B": 18,
    "C": 25,
    "D": 15,
    "E": 18,
    "F": 18,
    "G": 35,
    "H": 18,
    "I": 18,
    "J": 18,
    "K": 18,
    "L": 15,
    "M": 15,
    "N": 22
}

for column, width in column_widths.items():

    worksheet.column_dimensions[column].width = width


# تنسيق التاريخ
for cell in worksheet["A"][1:]:

    cell.number_format = "yyyy-mm-dd"


# تنسيق الأسعار والإجماليات
for column in ["J", "K", "L", "N"]:

    for cell in worksheet[column][1:]:

        cell.number_format = '#,##0.00'


# تنسيق الكمية
for cell in worksheet["M"][1:]:

    cell.number_format = '#,##0.00'


workbook.save(OUTPUT_FILE)


# =====================================================
# 18. النتيجة النهائية
# =====================================================

print("\n" + "=" * 50)
print("CLEANING COMPLETED SUCCESSFULLY")
print("=" * 50)

print("Output File:")
print(OUTPUT_FILE)

print("Rows:", len(df))
print("Columns:", len(df.columns))

print("\nالملف جاهز للاستيراد في Power BI.")
