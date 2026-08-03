import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# Configuration
# ============================================================

BASE_DIR = Path("Scrape PHEI")
OUTPUT_DIR = Path("output_eom")
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = OUTPUT_DIR / "Selected_EoM_Yield_Curve.xlsx"

MONTH_MAP = {
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

MAX_TENOR = 30


# ============================================================
# Helper functions
# ============================================================

def parse_indonesian_date(date_text):
    """
    Expected file name pattern:
    Yield-Curve-30-Desember-2024.xlsx
    or
    Yield-Curve- 30-Desember-2024.xlsx
    """
    date_text = str(date_text).strip()
    parts = date_text.split("-")

    if len(parts) != 3:
        raise ValueError(f"Unexpected date format: {date_text}")

    day = int(parts[0].strip())
    month_name = parts[1].strip()
    year = int(parts[2].strip())

    if month_name not in MONTH_MAP:
        raise ValueError(f"Unknown Indonesian month name: {month_name}")

    month = MONTH_MAP[month_name]

    return pd.Timestamp(year=year, month=month, day=day)


def find_available_files():
    """
    Find all available Yield-Curve files in the repository.

    This version does not restrict to Dec 2024, Dec 2025, or 2026.
    It takes all available spot rate files under Scrape PHEI.
    """
    records = []

    for file in BASE_DIR.rglob("Yield-Curve-*.xlsx"):
        match = re.search(r"Yield-Curve-\s*(.+)\.xlsx$", file.name)

        if not match:
            continue

        date_text = match.group(1).strip()

        try:
            file_date = parse_indonesian_date(date_text)
        except Exception as error:
            print(f"Skip file because date cannot be parsed: {file.name}. Error: {error}")
            continue

        records.append(
            {
                "file_path": file,
                "date": file_date,
                "year": file_date.year,
                "month": file_date.month,
                "file_name": file.name,
            }
        )

    files_df = pd.DataFrame(records)

    if files_df.empty:
        raise ValueError("No Yield-Curve files found under Scrape PHEI.")

    files_df = files_df.sort_values("date").reset_index(drop=True)

    return files_df


def read_and_compile_long(files_df):
    """
    Read all available Yield-Curve files and combine them into one long-format table.
    Tenor is restricted to maximum 30.
    """
    compiled_data = []

    for _, row in files_df.iterrows():
        source_file = row["file_path"]
        source_date = row["date"]

        df = pd.read_excel(source_file, index_col=0)
        df = df.reset_index()

        first_col = df.columns[0]
        df = df.rename(columns={first_col: "Tenor Year"})

        required_columns = ["Tenor Year", "Spot-Rate"]
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            raise ValueError(
                f"Missing required columns in {source_file}: {missing_columns}"
            )

        df["Tenor Year"] = pd.to_numeric(df["Tenor Year"], errors="coerce")
        df = df[df["Tenor Year"].notna()]
        df = df[df["Tenor Year"] <= MAX_TENOR]

        df.insert(0, "Date", source_date)
        df.insert(1, "Year", int(row["year"]))
        df.insert(2, "Month", int(row["month"]))
        df["Source File"] = row["file_name"]

        compiled_data.append(df)

    if not compiled_data:
        raise ValueError("No usable data after tenor filtering.")

    final_df = pd.concat(compiled_data, ignore_index=True)

    return final_df


def make_spot_wide(final_df):
    """
    Convert long Spot-Rate data into wide format:
    rows = tenor
    columns = available dates
    values = spot rates
    """
    spot_wide = (
        final_df
        .pivot_table(
            index="Tenor Year",
            columns="Date",
            values="Spot-Rate",
            aggfunc="first",
        )
        .sort_index()
    )

    spot_wide = spot_wide.reindex(sorted(spot_wide.columns), axis=1)

    return spot_wide


def get_integer_tenors(spot_wide):
    """
    Use only integer tenors for forward rate section:
    1, 2, 3, ..., 30.
    """
    integer_tenors = []

    for tenor in spot_wide.index:
        try:
            tenor_float = float(tenor)

            if (
                tenor_float >= 1
                and tenor_float <= MAX_TENOR
                and abs(tenor_float - round(tenor_float)) < 1e-9
            ):
                integer_tenors.append(int(round(tenor_float)))
        except Exception:
            continue

    return sorted(set(integer_tenors))


def build_spot_row_map(spot_wide, spot_start_row):
    """
    Map integer tenor to its Excel row in the spot rate section.
    """
    spot_row_map = {}

    for row_idx, tenor in enumerate(spot_wide.index, start=spot_start_row):
        try:
            tenor_float = float(tenor)

            if abs(tenor_float - round(tenor_float)) < 1e-9:
                integer_tenor = int(round(tenor_float))

                if integer_tenor <= MAX_TENOR:
                    spot_row_map[integer_tenor] = row_idx
        except Exception:
            continue

    return spot_row_map


# ============================================================
# Excel writing
# ============================================================

def create_idr_workbook(spot_wide):
    """
    Create one-sheet Excel workbook with only IDR sheet:
    - Spot Rate (IDR)
    - Forward Rate (IDR)

    Forward rate formulas reference the Tenor Year column rather than hardcoded powers.

    Example output formula:
    =((1+D11)^$A11/((1+D10)^$A10))-1
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "IDR"

    date_cols = list(spot_wide.columns)
    integer_tenors = get_integer_tenors(spot_wide)

    # --------------------------------------------------------
    # Styles
    # --------------------------------------------------------
    dark_fill = PatternFill("solid", fgColor="1F4E78")
    light_fill = PatternFill("solid", fgColor="D9EAF7")
    white_font = Font(color="FFFFFF", bold=True)
    bold_font = Font(bold=True)
    thin_bottom = Border(bottom=Side(style="thin", color="808080"))

    # --------------------------------------------------------
    # Sheet title
    # --------------------------------------------------------
    ws["A1"] = "IDR"
    ws["A1"].font = Font(bold=True, size=14)

    # ========================================================
    # Spot Rate section
    # ========================================================
    spot_title_row = 3
    spot_header_row = 5
    spot_start_row = 8

    ws.cell(row=spot_title_row, column=1).value = (
        "Spot Rate (IDR) ---- Spot rate of t=3 means the interest rate is effective from t=0 to t=3"
    )

    ws.cell(row=spot_header_row, column=1).value = "Time Period"
    ws.cell(row=spot_header_row + 1, column=1).value = "Rate Type"
    ws.cell(row=spot_header_row + 2, column=1).value = "Tenor Year"

    for col_idx, date_value in enumerate(date_cols, start=2):
        date_cell = ws.cell(row=spot_header_row, column=col_idx)
        date_cell.value = date_value
        date_cell.number_format = "mm/dd/yyyy"

        ws.cell(row=spot_header_row + 1, column=col_idx).value = "Spot"

    for row_idx, tenor in enumerate(spot_wide.index, start=spot_start_row):
        ws.cell(row=row_idx, column=1).value = tenor

        for col_idx, date_value in enumerate(date_cols, start=2):
            value_cell = ws.cell(row=row_idx, column=col_idx)
            value_cell.value = spot_wide.loc[tenor, date_value]
            value_cell.number_format = "0.0000%"

    spot_end_row = spot_start_row + len(spot_wide.index) - 1
    spot_row_map = build_spot_row_map(spot_wide, spot_start_row)

    # ========================================================
    # Forward Rate section
    # ========================================================
    forward_title_row = spot_end_row + 4
    forward_header_row = forward_title_row + 2
    forward_start_row = forward_header_row + 3

    ws.cell(row=forward_title_row, column=1).value = (
        "Forward Rate (IDR) ---- Forward rate of t=3 means the interest rate is effective from t=2 to t=3"
    )

    ws.cell(row=forward_header_row, column=1).value = "Time Period"
    ws.cell(row=forward_header_row + 1, column=1).value = "Rate Type"
    ws.cell(row=forward_header_row + 2, column=1).value = "Tenor Year"

    for col_idx, date_value in enumerate(date_cols, start=2):
        date_cell = ws.cell(row=forward_header_row, column=col_idx)
        date_cell.value = date_value
        date_cell.number_format = "mm/dd/yyyy"

        ws.cell(row=forward_header_row + 1, column=col_idx).value = "Forward"

    for row_idx, tenor in enumerate(integer_tenors, start=forward_start_row):
        ws.cell(row=row_idx, column=1).value = tenor

        for col_idx, _ in enumerate(date_cols, start=2):
            col_letter = get_column_letter(col_idx)

            spot_current_row = spot_row_map.get(tenor)
            spot_previous_row = spot_row_map.get(tenor - 1)

            formula_cell = ws.cell(row=row_idx, column=col_idx)

            if spot_current_row is None:
                formula_cell.value = None

            elif tenor == 1:
                current_ref = f"{col_letter}{spot_current_row}"
                formula_cell.value = f"={current_ref}"

            elif spot_previous_row is None:
                formula_cell.value = None

            else:
                current_ref = f"{col_letter}{spot_current_row}"
                previous_ref = f"{col_letter}{spot_previous_row}"
                current_tenor_ref = f"$A{spot_current_row}"
                previous_tenor_ref = f"$A{spot_previous_row}"

                formula_cell.value = (
                    f"=((1+{current_ref})^{current_tenor_ref}/"
                    f"((1+{previous_ref})^{previous_tenor_ref}))-1"
                )

            formula_cell.number_format = "0.0000%"

    # ========================================================
    # Styling
    # ========================================================
    max_col = len(date_cols) + 1

    section_rows = [
        spot_title_row,
        forward_title_row,
    ]

    header_rows = [
        spot_header_row,
        spot_header_row + 1,
        spot_header_row + 2,
        forward_header_row,
        forward_header_row + 1,
        forward_header_row + 2,
    ]

    for row in section_rows:
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = dark_fill
            cell.font = white_font
            cell.alignment = Alignment(vertical="center")

    for row in header_rows:
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = light_fill
            cell.font = bold_font
            cell.border = thin_bottom
            cell.alignment = Alignment(vertical="center")

    for row in range(1, ws.max_row + 1):
        ws.cell(row=row, column=1).alignment = Alignment(vertical="center")

    for row in range(1, ws.max_row + 1):
        for col in range(2, max_col + 1):
            ws.cell(row=row, column=col).alignment = Alignment(
                vertical="center",
                horizontal="right",
            )

    ws.freeze_panes = "B8"

    ws.column_dimensions["A"].width = 18

    for col in range(2, max_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 13

    return wb


# ============================================================
# Main process
# ============================================================

def main():
    files_df = find_available_files()
    final_df = read_and_compile_long(files_df)
    spot_wide = make_spot_wide(final_df)

    wb = create_idr_workbook(spot_wide)
    wb.save(OUTPUT_FILE)

    print(f"Done. Output saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()