
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

ZIP_FILE = BASE_DIR / "Supplier Account Statement.zip"

OUTPUT_FILE = BASE_DIR / "Cottonil_Supplier_Account_PowerBI_Ready.xlsx"

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

    # إزالة الرموز والمسافات المخفية
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


def clean_supplier(value):

    value = clean_text(value)

    if value is None:
        return None

    # حذف عبارة الفترة وما بعدها من اسم المورد
    value = re.sub(
        r"\s*للفترة\b.*$",
        "",
        value
    )

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


def clean_date(value):

    value = clean_text(value)

    if value is None:
        return pd.NaT

    return pd.to_datetime(
        value,
        format="%d/%m/%Y",
        errors="coerce"
    )


def get_balance_type(value):

    value = clean_text(value)

    if value is None:
        return UNKNOWN_VALUE

    # بعض الكلمات تظهر مقسمة بين خليتين أو جدولين
    if value in ["دا", "دائن", "ئن"]:
        return "دائن"

    if value in ["مد", "مدين"]:
        return "مدين"

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

        # قراءة الملفات بترميز Windows-1256
        html_text = html_content.decode(
            "cp1256",
            errors="replace"
        )

        soup = BeautifulSoup(
            html_text,
            "lxml"
        )

        tables = soup.find_all("table")

        # ---------------------------------------------
        # 4.1 استخراج اسم المورد
        # ---------------------------------------------

        supplier = None

        if len(tables) > 1:

            header_values = [
                cell.get_text(" ", strip=True)
                for cell in tables[1].find_all(["td", "th"])
            ]

            if len(header_values) > 3:

                supplier = clean_supplier(
                    header_values[3]
                )

        if supplier is None:

            supplier = UNKNOWN_VALUE

        # ---------------------------------------------
        # 4.2 استخراج الرصيد الافتتاحي
        # ---------------------------------------------

        opening_balance = 0

        if len(tables) > 2:

            opening_values = [
                cell.get_text(" ", strip=True)
                for cell in tables[2].find_all(["td", "th"])
            ]

            if len(opening_values) > 5:

                opening_balance = clean_number(
                    opening_values[5]
                )

        if opening_balance is None:

            opening_balance = 0

        # ---------------------------------------------
        # 4.3 استخراج حركات المورد
        # ---------------------------------------------

        for table_index, table in enumerate(tables):

            rows = table.find_all("tr")

            if not rows:
                continue

            for row in rows:

                cells = row.find_all(["td", "th"])

                values = [
                    cell.get_text(" ", strip=True)
                    for cell in cells
                ]

                # صف الحركة يحتوي على 10 خلايا
                if len(values) != 10:
                    continue

                description = clean_text(values[1])

                transaction_date = clean_date(values[2])

                # تجاهل صفوف العناوين والإجماليات
                if description is None:
                    continue

                if "الإجمالى" in description or "الإجمالي" in description:
                    continue

                if pd.isna(transaction_date):
                    continue

                # -------------------------------------
                # 4.4 استخراج نوع الحركة
                # -------------------------------------

                if "ارتجاع" in description:

                    transaction_type = "مرتجع"

                elif "استلام" in description:

                    transaction_type = "استلام"

                else:

                    transaction_type = "أخرى"

                # -------------------------------------
                # 4.5 استخراج نوع الرصيد
                # -------------------------------------

                balance_type = clean_text(values[9])

                # في بعض الصفحات تكملة الكلمة موجودة
                # في الجدول التالي
                if table_index + 1 < len(tables):

                    next_table_text = tables[
                        table_index + 1
                    ].get_text(" ", strip=True)

                    if balance_type == "دا" and "ئن" in next_table_text:

                        balance_type = "دائن"

                    elif balance_type == "مد" and "ين" in next_table_text:

                        balance_type = "مدين"

                balance_type = get_balance_type(
                    balance_type
                )

                # -------------------------------------
                # 4.6 إنشاء السجل
                # -------------------------------------

                record = {

                    "Store": STORE_NAME,

                    "Supplier": supplier,

                    "Transaction_Type": transaction_type,

                    "Description": description,

                    "Transaction_Date": transaction_date,

                    "Invoices": clean_number(values[3]),

                    "Returns": clean_number(values[4]),

                    "Payments": clean_number(values[5]),

                    "Discount": clean_number(values[6]),

                    "Carried_Balance": clean_number(values[7]),

                    "Balance": clean_number(values[8]),

                    "Balance_Type": balance_type,

                    "Opening_Balance": opening_balance

                }

                all_records.append(record)


# =====================================================
# 5. تحويل البيانات إلى DataFrame
# =====================================================

df = pd.DataFrame(all_records)

if df.empty:

    raise ValueError(
        "لم يتم استخراج أي حركات. راجع ملفات HTML وطريقة عرض الجداول."
    )

print("\nData extraction completed.")

print("Extracted Rows:", len(df))


# =====================================================
# 6. تنظيف الأعمدة النصية
# =====================================================

text_columns = [

    "Store",

    "Supplier",

    "Transaction_Type",

    "Description",

    "Balance_Type"

]

for column in text_columns:

    df[column] = df[column].apply(clean_text)


# =====================================================
# 7. تنظيف اسم المورد
# =====================================================

df["Supplier"] = df["Supplier"].apply(clean_supplier)


# =====================================================
# 8. معالجة القيم الفارغة
# =====================================================

unknown_columns = [

    "Store",

    "Supplier",

    "Transaction_Type",

    "Description",

    "Balance_Type"

]

for column in unknown_columns:

    df[column] = df[column].fillna(
        UNKNOWN_VALUE
    )


# تثبيت اسم الفرع في جميع الصفوف
df["Store"] = STORE_NAME


# =====================================================
# 9. تنظيف التاريخ
# =====================================================

df["Transaction_Date"] = pd.to_datetime(
    df["Transaction_Date"],
    errors="coerce"
)


# =====================================================
# 10. تنظيف الأعمدة الرقمية
# =====================================================

numeric_columns = [

    "Invoices",

    "Returns",

    "Payments",

    "Discount",

    "Carried_Balance",

    "Balance",

    "Opening_Balance"

]

for column in numeric_columns:

    df[column] = pd.to_numeric(
        df[column],
        errors="coerce"
    )


# =====================================================
# 11. ترتيب الأعمدة النهائية
# =====================================================

final_columns = [

    "Store",

    "Supplier",

    "Transaction_Type",

    "Description",

    "Transaction_Date",

    "Invoices",

    "Returns",

    "Payments",

    "Discount",

    "Carried_Balance",

    "Balance",

    "Balance_Type",

    "Opening_Balance"

]

df = df[final_columns]


# =====================================================
# 12. ترتيب البيانات
# =====================================================

df = df.sort_values(

    by=[
        "Supplier",
        "Transaction_Date"
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

print("\nTransaction Types:")

print(df["Transaction_Type"].value_counts())

print("\nBalance Types:")

print(df["Balance_Type"].value_counts())

print("\nTotal Invoices:")

print(df["Invoices"].sum())

print("\nTotal Returns:")

print(df["Returns"].sum())

print("\nTotal Payments:")

print(df["Payments"].sum())

print("\nTotal Discounts:")

print(df["Discount"].sum())

print("\nTotal Balance:")

print(df["Balance"].sum())


# =====================================================
# 14. تصدير ملف Excel
# =====================================================

with pd.ExcelWriter(

    OUTPUT_FILE,

    engine="openpyxl",

    datetime_format="yyyy-mm-dd"

) as writer:

    df.to_excel(

        writer,

        sheet_name="Supplier_Account",

        index=False

    )


# =====================================================
# 15. تنسيق ملف Excel
# =====================================================

workbook = load_workbook(OUTPUT_FILE)

worksheet = workbook["Supplier_Account"]

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

    "A": 22,

    "B": 22,

    "C": 20,

    "D": 30,

    "E": 20,

    "F": 18,

    "G": 18,

    "H": 18,

    "I": 18,

    "J": 20,

    "K": 18,

    "L": 18,

    "M": 20

}

for column, width in column_widths.items():

    worksheet.column_dimensions[column].width = width


# تنسيق التاريخ
for cell in worksheet["E"][1:]:

    cell.number_format = "yyyy-mm-dd"


# تنسيق القيم المالية
for column in ["F", "G", "H", "I", "J", "K", "M"]:

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

print("Store:", STORE_NAME)

print("\nSupplier Account data is ready for Power BI.")
