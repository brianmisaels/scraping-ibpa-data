import pandas as pd
from pathlib import Path
import re

BASE_DIR = Path("Scrape PHEI")
OUTPUT_DIR = Path("output_eom")
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = OUTPUT_DIR / "Selected_EoM_Yield_Curve.xlsx"

month_map = {
    "Januari": 1,
    "Februari": 2,
    "Maret": 3,
    "April": 4,
    "Mei": 5,
    "Juni": 6,
    "Juli": 7,
    "Agustus": 8,
    "September": 9,
    "Oktober": 10,
    "November": 11,
    "Desember": 12,
}

def parse_indonesian_date(date_text):
    parts = date_text.split("-")

    if len(parts) != 3:
        raise ValueError(f"Unexpected date format: {date_text}")

    day = int(parts[0])
    month = month_map[parts[1]]
    year = int(parts[2])

    return pd.Timestamp(year=year, month=month, day=day)

records = []

for file in BASE_DIR.rglob("Yield-Curve-*.xlsx"):
    match = re.search(r"Yield-Curve-\s*(.+)\.xlsx$", file.name)

    if not match:
        continue

    date_text = match.group(1).strip()

    try:
        file_date = parse_indonesian_date(date_text)
    except Exception as e:
        print(f"Skip file because date cannot be parsed: {file.name}. Error: {e}")
        continue

    keep_2026_monthly = file_date.year == 2026
    keep_dec_2024_2025 = file_date.year in [2024, 2025] and file_date.month == 12

    if keep_2026_monthly or keep_dec_2024_2025:
        records.append({
            "file_path": file,
            "date": file_date,
            "year": file_date.year,
            "month": file_date.month,
            "file_name": file.name
        })

files_df = pd.DataFrame(records)

if files_df.empty:
    raise ValueError("No matching Yield-Curve files found for 2026 monthly or Dec 2024/Dec 2025.")

selected_files = (
    files_df
    .sort_values("date")
    .groupby(["year", "month"], as_index=False)
    .tail(1)
    .sort_values("date")
)

compiled_data = []

for _, row in selected_files.iterrows():
    source_file = row["file_path"]
    source_date = row["date"]

    df = pd.read_excel(source_file, index_col=0)
    df = df.reset_index()

    first_col = df.columns[0]
    df = df.rename(columns={first_col: "Tenor Year"})

    df.insert(0, "Date", source_date)
    df.insert(1, "Year", row["year"])
    df.insert(2, "Month", row["month"])

    if row["year"] in [2024, 2025] and row["month"] == 12:
        df.insert(3, "Period Type", "Year End")
    else:
        df.insert(3, "Period Type", "Month End")

    df["Source File"] = row["file_name"]

    compiled_data.append(df)

final_df = pd.concat(compiled_data, ignore_index=True)

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    final_df.to_excel(writer, sheet_name="Selected_EoM_Long", index=False)
    selected_files.to_excel(writer, sheet_name="Selected_Source_Files", index=False)

print(f"Done. Output saved to: {OUTPUT_FILE}")