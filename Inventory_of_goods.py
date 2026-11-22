
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

ZIP_FILE = BASE_DIR / "Inventory of goods.zip"

OUTPUT_FILE = BASE_DIR / "Cottonil_Inventory_PowerBI_Ready.xlsx"

UNKNOWN_VALUE = "غير معروف"


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

    if value.lower() in ["", "nan", "none", "null"]:
        return None

    if value in ["--", "-", "—"]:
        return None

    return value


def clean_number(value):

    value = clean_text(value)

    if value is None:
        return None

    value = value.replace(",", "")

    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def is_numeric(value):

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

        # الملفات بترميز Windows-1256
        html_text = html_content.decode(
            "cp1256",
            errors="replace"
        )

        soup = BeautifulSoup(
            html_text,
            "lxml"
        )

        # بيانات الصفحة
        inventory_date = None
        department = None
        store = None

        # ---------------------------------------------
        # قراءة الجداول
        # ---------------------------------------------

        for table in soup.find_all("table"):

            rows = table.find_all("tr")

            if not rows:
                continue

            for row in rows:

                cells = row.find_all(["td", "th"])

                values = [
                    cell.get_text(" ", strip=True)
                    for cell in cells
                ]

                if not values:
                    continue

                row_text = " ".join(values)

                # -------------------------------------
                # 4.1 اسم الفرع
                # -------------------------------------

                if "قطونيل شيراتون" in row_text:

                    store = "قطونيل شيراتون"

                # -------------------------------------
                # 4.2 تاريخ الجرد والقسم
                # -------------------------------------

                if "طباعة الجرد على مستوى" in row_text:

                    for value in values:

                        if re.fullmatch(
                            r"\d{4}/\d{2}/\d{2}",
                            value
                        ):

                            inventory_date = value
                            break

                    # آخر خلية تحتوي على اسم القسم
                    section_value = clean_text(values[-1])

                    department = section_value

                    continue

                # -------------------------------------
                # 4.3 تجاهل صف العناوين
                # -------------------------------------

                if "الصنف" in values:

                    continue

                # -------------------------------------
                # 4.4 التحقق من صف المنتج
                # -------------------------------------

                # شكل صف المنتج في الملف:
                # 0: خلية فارغة
                # 1: اسم الصنف
                # 2: المقاس
                # 3: اللون
                # 4: المجموعة
                # 5: خلية فارغة مدمجة
                # 6: كود المنتج
                # 7: الباركود
                # 8: الرصيد

                if len(values) != 9:

                    continue

                product_name = clean_text(values[1])

                stock_quantity = clean_number(values[8])

                # التأكد من أن الصف يحتوي على اسم صنف ورصيد رقمي
                if product_name is None:

                    continue

                if stock_quantity is None:

                    continue

                # -------------------------------------
                # 4.5 استخراج بيانات المنتج
                # -------------------------------------

                record = {

                    "Inventory_Date": inventory_date,

                    "Store": store,

                    "Department": department,

                    "Product_Name": product_name,

                    "Size": clean_text(values[2]),

                    "Color": clean_text(values[3]),

                    "Product_Group": clean_text(values[4]),

                    "Product_Code": clean_text(values[6]),

                    "Barcode": clean_text(values[7]),

                    "Stock_Quantity": stock_quantity

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

text_columns = [

    "Store",

    "Department",

    "Product_Name",

    "Size",

    "Color",

    "Product_Group"

]

for column in text_columns:

    df[column] = df[column].apply(clean_text)


# =====================================================
# 7. معالجة القيم الفارغة
# =====================================================

# القسم
df["Department"] = df["Department"].fillna(
    UNKNOWN_VALUE
)

# المقاس
df["Size"] = df["Size"].fillna(
    UNKNOWN_VALUE
)

# اللون
df["Color"] = df["Color"].fillna(
    UNKNOWN_VALUE
)

# المجموعة
df["Product_Group"] = df["Product_Group"].fillna(
    UNKNOWN_VALUE
)

# الفرع
df["Store"] = df["Store"].fillna(
    UNKNOWN_VALUE
)


# =====================================================
# 8. تنظيف التاريخ
# =====================================================

df["Inventory_Date"] = pd.to_datetime(

    df["Inventory_Date"],

    format="%Y/%m/%d",

    errors="coerce"

)


# =====================================================
# 9. تنظيف الأعمدة الرقمية
# =====================================================

df["Stock_Quantity"] = pd.to_numeric(

    df["Stock_Quantity"],

    errors="coerce"

)


# =====================================================
# 10. تجهيز الأكواد كـ Text
# =====================================================

identifier_columns = [

    "Product_Code",

    "Barcode"

]

for column in identifier_columns:

    df[column] = (
        df[column]
        .apply(clean_text)
        .astype("string")
    )


# =====================================================
# 11. ترتيب الأعمدة النهائية
# =====================================================

final_columns = [

    "Inventory_Date",

    "Store",

    "Department",

    "Product_Name",

    "Size",

    "Color",

    "Product_Group",

    "Product_Code",

    "Barcode",

    "Stock_Quantity"

]

df = df[final_columns]


# =====================================================
# 12. مراجعة أنواع البيانات
# =====================================================

df["Inventory_Date"] = pd.to_datetime(
    df["Inventory_Date"],
    errors="coerce"
)

df["Stock_Quantity"] = pd.to_numeric(
    df["Stock_Quantity"],
    errors="coerce"
)


# =====================================================
# 13. ترتيب البيانات
# =====================================================

df = df.sort_values(

    by=[
        "Department",
        "Product_Name",
        "Size"
    ],

    na_position="last"

).reset_index(drop=True)


# =====================================================
# 14. فحص جودة البيانات
# =====================================================

print("\n" + "=" * 50)
print("FINAL DATA QUALITY CHECK")
print("=" * 50)

print("Total Rows:", len(df))

print("Total Columns:", len(df.columns))

print("\nMissing Values:")

print(df.isnull().sum())

print("\nDepartment Distribution:")

print(df["Department"].value_counts())

print("\nSize Distribution:")

print(df["Size"].value_counts())

print("\nColor Distribution:")

print(df["Color"].value_counts())

print("\nProduct Group Distribution:")

print(df["Product_Group"].value_counts())

print("\nTotal Stock Quantity:")

print(df["Stock_Quantity"].sum())

print("\nProducts with Zero Stock:")

print(
    (df["Stock_Quantity"] == 0).sum()
)


# =====================================================
# 15. تصدير Excel
# =====================================================

with pd.ExcelWriter(

    OUTPUT_FILE,

    engine="openpyxl",

    datetime_format="yyyy-mm-dd"

) as writer:

    df.to_excel(

        writer,

        sheet_name="Inventory",

        index=False

    )


# =====================================================
# 16. تنسيق ملف Excel
# =====================================================

workbook = load_workbook(OUTPUT_FILE)

worksheet = workbook["Inventory"]

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

    "A": 18,

    "B": 22,

    "C": 20,

    "D": 35,

    "E": 18,

    "F": 18,

    "G": 20,

    "H": 18,

    "I": 20,

    "J": 18

}

for column, width in column_widths.items():

    worksheet.column_dimensions[column].width = width


# تنسيق التاريخ
for cell in worksheet["A"][1:]:

    cell.number_format = "yyyy-mm-dd"


# تنسيق الرصيد
for cell in worksheet["J"][1:]:

    cell.number_format = '#,##0.00'


workbook.save(OUTPUT_FILE)


# =====================================================
# 17. النتيجة النهائية
# =====================================================

print("\n" + "=" * 50)

print("CLEANING COMPLETED SUCCESSFULLY")

print("=" * 50)

print("Output File:")

print(OUTPUT_FILE)

print("Rows:", len(df))

print("Columns:", len(df.columns))

print("\nInventory data is ready for Power BI.")
