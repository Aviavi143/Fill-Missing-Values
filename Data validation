"""
app.py
------
Streamlit web app: upload your two Excel files in the browser, pick the
sheets + key column, click a button, and download the completed workbook.
Every value that got filled in from the reference sheet is highlighted
in red in the output file, so it's easy to spot what changed.

Handles the messy real-world stuff that silently breaks naive matching:
  - extra/leading/trailing spaces and non-breaking spaces in keys
  - different capitalization between the two files ('a2' vs 'A2')
  - numeric codes that Excel stores as 12345.0 instead of 12345
  - column headers with stray trailing spaces ('Size ' vs 'Size')

RUN LOCALLY:
    pip install streamlit pandas openpyxl
    streamlit run app.py

This opens automatically in your browser (usually http://localhost:8501).
No files ever leave your machine — everything runs locally in this script.
"""

import io
import re
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill

st.set_page_config(page_title="Fill Missing Values", layout="centered")
st.title("📋 Fill Missing Values from a Reference Sheet")
st.write(
    "Upload the file that has gaps, upload the file that might contain the "
    "missing values, pick the matching key column, and get a completed "
    "workbook back — newly filled cells are highlighted in red — plus a "
    "list of anything that couldn't be resolved."
)


def normalize_key(value):
    """Make key matching resilient to spacing, case, and Excel's numeric quirks."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip()
    s = s.replace("\xa0", " ")          # non-breaking spaces
    s = re.sub(r"\s+", " ", s)          # collapse internal whitespace
    if re.fullmatch(r"-?\d+\.0", s):    # "12345.0" -> "12345"
        s = s[:-2]
    return s.upper()


def is_blank(v):
    return pd.isna(v) or (isinstance(v, str) and v.strip() == "")


# ---------------------------------------------------------------------
# Step 1: Upload files
# ---------------------------------------------------------------------
st.header("1. Upload your files")

col1, col2 = st.columns(2)
with col1:
    main_upload = st.file_uploader("Main file (has missing values)", type=["xlsx", "xls"], key="main")
with col2:
    ref_upload = st.file_uploader("Reference file (may hold the values)", type=["xlsx", "xls"], key="ref")

if not main_upload or not ref_upload:
    st.info("Upload both files to continue.")
    st.stop()

# ---------------------------------------------------------------------
# Step 2: Pick sheets
# ---------------------------------------------------------------------
st.header("2. Pick the sheets")

main_bytes = main_upload.getvalue()
ref_bytes = ref_upload.getvalue()

main_sheets = pd.ExcelFile(io.BytesIO(main_bytes)).sheet_names
ref_sheets = pd.ExcelFile(io.BytesIO(ref_bytes)).sheet_names

col1, col2 = st.columns(2)
with col1:
    main_sheet_name = st.selectbox("Sheet in main file", main_sheets, key="main_sheet")
with col2:
    ref_sheet_name = st.selectbox("Sheet in reference file", ref_sheets, key="ref_sheet")

main_df = pd.read_excel(io.BytesIO(main_bytes), sheet_name=main_sheet_name, dtype=str)
ref_df = pd.read_excel(io.BytesIO(ref_bytes), sheet_name=ref_sheet_name, dtype=str)

# Strip stray whitespace from headers so "Size " and "Size" are recognized as the same column
main_df.columns = [str(c).strip() for c in main_df.columns]
ref_df.columns = [str(c).strip() for c in ref_df.columns]

# ---------------------------------------------------------------------
# Step 3: Pick the key column
# ---------------------------------------------------------------------
st.header("3. Pick the key column")

common_columns = [c for c in main_df.columns if c in ref_df.columns]
if not common_columns:
    st.error("These two sheets share no column names, so there's nothing to match on. "
              "Make sure both files use the exact same header for the key column (e.g. 'Material Code').")
    st.stop()

default_idx = 0
for guess in ["Material Code", "Material code", "material code", "Material_Code"]:
    if guess in common_columns:
        default_idx = common_columns.index(guess)
        break

key_column = st.selectbox("Key column (used to match rows between the two sheets)",
                           common_columns, index=default_idx)

# ---------------------------------------------------------------------
# Step 4: Run
# ---------------------------------------------------------------------
st.header("4. Run it")

if st.button("Fill missing values", type="primary"):
    # Build a normalized-key lookup from the reference sheet.
    # If duplicate normalized keys exist, the last row wins.
    ref_df["_norm_key"] = ref_df[key_column].map(normalize_key)
    ref_lookup = ref_df.set_index("_norm_key").to_dict(orient="index")

    fillable_columns = [c for c in main_df.columns if c != key_column and c in ref_df.columns]

    not_found_records = []   # (Material Code, Column Name, Reason)
    filled_cells = []        # (row_index_in_df, column_name)
    filled_count = 0

    for idx, row in main_df.iterrows():
        norm_code = normalize_key(row[key_column])

        for col in fillable_columns:
            if not is_blank(row[col]):
                continue

            ref_row = ref_lookup.get(norm_code)
            if ref_row is None:
                not_found_records.append((row[key_column], col, "Key not found in reference sheet"))
                continue

            ref_value = ref_row.get(col)
            if is_blank(ref_value):
                not_found_records.append((row[key_column], col, "Key found, but that column is blank in reference sheet too"))
            else:
                main_df.at[idx, col] = ref_value
                filled_cells.append((idx, col))
                filled_count += 1

    not_found_df = pd.DataFrame(not_found_records, columns=[key_column, "Column Name", "Reason"])

    st.success(f"Filled {filled_count} value(s). {len(not_found_df)} value(s) still missing.")

    st.subheader("Preview: Filled Data")
    st.dataframe(main_df, use_container_width=True)

    if len(not_found_df) > 0:
        st.subheader("Preview: Not Found (with reason)")
        st.dataframe(not_found_df, use_container_width=True)

    # Build downloadable workbook in memory
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        main_df.to_excel(writer, sheet_name="Filled Data", index=False)
        not_found_df.to_excel(writer, sheet_name="Not Found", index=False)

        # --- Highlight every filled cell in red ---
        worksheet = writer.sheets["Filled Data"]
        red_font = Font(color="FF0000", bold=True)
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

        col_to_excel_idx = {col: i + 1 for i, col in enumerate(main_df.columns)}  # 1-based
        for df_row_idx, col_name in filled_cells:
            excel_row = df_row_idx + 2  # +1 for header row, +1 for 1-based indexing
            excel_col = col_to_excel_idx[col_name]
            cell = worksheet.cell(row=excel_row, column=excel_col)
            cell.font = red_font
            cell.fill = red_fill

    output_buffer.seek(0)

    st.download_button(
        label="⬇️ Download output_filled.xlsx",
        data=output_buffer,
        file_name="output_filled.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
