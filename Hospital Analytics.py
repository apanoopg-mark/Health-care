import pandas as pd

# 1. Load the original file
file_path = "Hospital Analytics.xlsx"
df = pd.read_excel(file_path)

# 2. Fix Dates (Swap dates if checkout is before entry to keep all rows)
df["تاريخ الدخول"] = pd.to_datetime(df["تاريخ الدخول"])
df["تاريخ الخروج"] = pd.to_datetime(df["تاريخ الخروج"])
mask_dates = df["تاريخ الخروج"] < df["تاريخ الدخول"]
temp = df.loc[mask_dates, "تاريخ الدخول"].copy()
df.loc[mask_dates, "تاريخ الدخول"] = df.loc[mask_dates, "تاريخ الخروج"]
df.loc[mask_dates, "تاريخ الخروج"] = temp

# 3. Fix Invoices (Recalculate placeholder values 999 and 3852)
df["مدة_الإقامة"] = (df["تاريخ الخروج"] - df["تاريخ الدخول"]).dt.days
df["مدة_الإقامة"] = df["مدة_الإقامة"].apply(lambda x: x if x > 1 else 1)
mask_placeholders = df["الفاتورة"].isin([999, 3852])
df.loc[mask_placeholders, "الفاتورة"] = 500 + (df["مدة_الإقامة"] * 250)

# 4. Fix Gender Mismatches based on Arabic First Names
female_names = [
    "ندى",
    "ياسمين",
    "مريم",
    "فاطمة",
    "عائشة",
    "زينب",
    "سارة",
    "ريم",
    "نور",
    "سعاد",
    "هدى",
    "منى",
    "ليلى",
    "أمل",
]
male_names = [
    "يوسف",
    "أحمد",
    "محمد",
    "علي",
    "إبراهيم",
    "محمود",
    "خالد",
    "حسن",
    "عبد",
    "عمر",
    "حسين",
    "مصطفى",
]


def correct_gender(row):
  first_name = str(row["اسم المريض"]).split()[0]
  if first_name in female_names:
    return "أنثى"
  elif first_name in male_names:
    return "ذكر"
  return row["الجنس"]


df["الجنس"] = df.apply(correct_gender, axis=1)

# 5. Export to a new Excel file
output_path = "Hospital_Analytics_Fully_Cleaned.xlsx"
df.to_excel(output_path, index=False)
print(f"Fully cleaned file exported successfully with {len(df)} rows!")