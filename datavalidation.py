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
No files ever leave your machine -- everything runs locally in this script.
"""

import io
import re
import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill

# ---------------------------------------------------------------------
# Pentland Brands logo
# ---------------------------------------------------------------------
LOGO_URL = "https://pbs.twimg.com/media/GyJLJ06W4AAQRrN?format=png&name=large"


def logo_data_uri():
    return LOGO_URL



st.set_page_config(
    page_title="Fill Missing Values | Pentland Brands",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------
# Styling -- Pentland Brands black / white palette
# ---------------------------------------------------------------------
PRIMARY_BLACK = "#0A0A0A"
ACCENT_RED = "#E4002B"
LIGHT_GREY = "#F5F5F7"
MID_GREY = "#8A8D93"
BORDER_GREY = "#E2E4E8"

st.markdown(
    f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

        html, body, [class*="css"] {{
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
        }}

        .stApp {{
            background-color: {LIGHT_GREY};
        }}

        /* Hide default Streamlit chrome for a cleaner look */
        #MainMenu {{visibility: hidden;}}
        footer {{visibility: hidden;}}
        header {{visibility: hidden;}}

        .block-container {{
            padding-top: 1.5rem;
            padding-bottom: 3rem;
            max-width: 820px;
        }}

        /* ---- Header / brand bar ---- */
        .brand-header {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            background-color: {PRIMARY_BLACK};
            border-radius: 14px;
            padding: 22px 28px;
            margin-bottom: 28px;
            box-shadow: 0 6px 20px rgba(0,0,0,0.12);
        }}

        .brand-header img {{
            height: 42px;
        }}

        .brand-header .brand-title {{
            color: white;
            text-align: right;
        }}

        .brand-header .brand-title h1 {{
            font-size: 1.05rem;
            font-weight: 700;
            margin: 0;
            letter-spacing: 0.2px;
        }}

        .brand-header .brand-title p {{
            font-size: 0.8rem;
            color: #B9BBC0;
            margin: 2px 0 0 0;
        }}

        /* ---- Section / step cards ---- */
        .step-card {{
            background-color: white;
            border: 1px solid {BORDER_GREY};
            border-radius: 12px;
            padding: 22px 26px 8px 26px;
            margin-bottom: 20px;
            box-shadow: 0 2px 8px rgba(16,17,20,0.04);
        }}

        .step-badge {{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 26px;
            height: 26px;
            border-radius: 50%;
            background-color: {PRIMARY_BLACK};
            color: white;
            font-size: 0.8rem;
            font-weight: 700;
            margin-right: 10px;
        }}

        .step-title {{
            display: flex;
            align-items: center;
            font-size: 1.05rem;
            font-weight: 700;
            color: {PRIMARY_BLACK};
            margin-bottom: 4px;
        }}

        .step-subtitle {{
            color: {MID_GREY};
            font-size: 0.88rem;
            margin: 0 0 18px 36px;
        }}

        /* ---- Buttons ---- */
        .stButton > button {{
            background-color: {PRIMARY_BLACK};
            color: white;
            border-radius: 8px;
            border: none;
            padding: 0.65rem 1.4rem;
            font-weight: 600;
            font-size: 0.95rem;
            transition: all 0.15s ease-in-out;
            width: 100%;
        }}

        .stButton > button:hover {{
            background-color: {ACCENT_RED};
            color: white;
            transform: translateY(-1px);
            box-shadow: 0 4px 12px rgba(228,0,43,0.28);
        }}

        .stDownloadButton > button {{
            background-color: {ACCENT_RED};
            color: white;
            border-radius: 8px;
            border: none;
            padding: 0.7rem 1.4rem;
            font-weight: 700;
            font-size: 0.98rem;
            width: 100%;
        }}

        .stDownloadButton > button:hover {{
            background-color: #C10024;
        }}

        /* ---- File uploader ---- */
        [data-testid="stFileUploaderDropzone"] {{
            background-color: {LIGHT_GREY};
            border: 1.5px dashed {MID_GREY};
            border-radius: 10px;
        }}

        /* ---- Metrics ---- */
        [data-testid="stMetric"] {{
            background-color: white;
            border: 1px solid {BORDER_GREY};
            border-radius: 10px;
            padding: 14px 16px;
        }}

        [data-testid="stMetricLabel"] {{
            color: {MID_GREY};
        }}

        /* ---- Dataframes ---- */
        [data-testid="stDataFrame"] {{
            border: 1px solid {BORDER_GREY};
            border-radius: 10px;
            overflow: hidden;
        }}

        /* ---- Footer note ---- */
        .privacy-note {{
            display: flex;
            align-items: center;
            gap: 8px;
            color: {MID_GREY};
            font-size: 0.82rem;
            margin-top: 6px;
        }}

        .footer-note {{
            text-align: center;
            color: {MID_GREY};
            font-size: 0.78rem;
            margin-top: 36px;
        }}

        hr {{
            border-color: {BORDER_GREY};
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------
st.markdown(
    f"""
    <div class="brand-header">
        <img src="{logo_data_uri()}" alt="Pentland Brands" />
        <div class="brand-title">
            <h1>Fill Missing Values</h1>
            <p>Reconcile two spreadsheets from a shared key column</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <p style="color:#4B4E54; font-size:0.98rem; margin-top:-6px; margin-bottom:26px;">
    Upload the file that has gaps, upload the file that might contain the
    missing values, pick the matching key column, and get a completed
    workbook back &mdash; newly filled cells are highlighted in red &mdash;
    plus a list of anything that couldn't be resolved.
    </p>
    """,
    unsafe_allow_html=True,
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
st.markdown(
    """
    <div class="step-card">
        <div class="step-title"><span class="step-badge">1</span> Upload your files</div>
        <div class="step-subtitle">Both files should be .xlsx or .xls workbooks</div>
    """,
    unsafe_allow_html=True,
)

col1, col2 = st.columns(2)
with col1:
    main_upload = st.file_uploader("Main file (has missing values)", type=["xlsx", "xls"], key="main")
with col2:
    ref_upload = st.file_uploader("Reference file (may hold the values)", type=["xlsx", "xls"], key="ref")

st.markdown(
    """
    <div class="privacy-note">🔒 No files ever leave your machine &mdash; everything runs locally.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

if not main_upload or not ref_upload:
    st.info("⬆️ Upload both files above to continue.")
    st.stop()

# ---------------------------------------------------------------------
# Step 2: Pick sheets
# ---------------------------------------------------------------------
main_bytes = main_upload.getvalue()
ref_bytes = ref_upload.getvalue()

main_sheets = pd.ExcelFile(io.BytesIO(main_bytes)).sheet_names
ref_sheets = pd.ExcelFile(io.BytesIO(ref_bytes)).sheet_names

st.markdown(
    """
    <div class="step-card">
        <div class="step-title"><span class="step-badge">2</span> Pick the sheets</div>
        <div class="step-subtitle">Choose which sheet in each workbook to use</div>
    """,
    unsafe_allow_html=True,
)

col1, col2 = st.columns(2)
with col1:
    main_sheet_name = st.selectbox("Sheet in main file", main_sheets, key="main_sheet")
with col2:
    ref_sheet_name = st.selectbox("Sheet in reference file", ref_sheets, key="ref_sheet")

st.markdown("</div>", unsafe_allow_html=True)

main_df = pd.read_excel(io.BytesIO(main_bytes), sheet_name=main_sheet_name, dtype=str)
ref_df = pd.read_excel(io.BytesIO(ref_bytes), sheet_name=ref_sheet_name, dtype=str)

# Strip stray whitespace from headers so "Size " and "Size" are recognized as the same column
main_df.columns = [str(c).strip() for c in main_df.columns]
ref_df.columns = [str(c).strip() for c in ref_df.columns]

# ---------------------------------------------------------------------
# Step 3: Pick the key column
# ---------------------------------------------------------------------
common_columns = [c for c in main_df.columns if c in ref_df.columns]

st.markdown(
    """
    <div class="step-card">
        <div class="step-title"><span class="step-badge">3</span> Pick the key column</div>
        <div class="step-subtitle">Used to match rows between the two sheets</div>
    """,
    unsafe_allow_html=True,
)

if not common_columns:
    st.error(
        "These two sheets share no column names, so there's nothing to match on. "
        "Make sure both files use the exact same header for the key column (e.g. 'Material Code')."
    )
    st.markdown("</div>", unsafe_allow_html=True)
    st.stop()

default_idx = 0
for guess in ["Material Code", "Material code", "material code", "Material_Code"]:
    if guess in common_columns:
        default_idx = common_columns.index(guess)
        break

key_column = st.selectbox(
    "Key column",
    common_columns,
    index=default_idx,
    label_visibility="collapsed",
)

st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------------------
# Step 4: Run
# ---------------------------------------------------------------------
st.markdown(
    """
    <div class="step-card">
        <div class="step-title"><span class="step-badge">4</span> Run it</div>
        <div class="step-subtitle">Fill the gaps and download the completed workbook</div>
    """,
    unsafe_allow_html=True,
)

run_clicked = st.button("✨ Fill missing values", type="primary")
st.markdown("</div>", unsafe_allow_html=True)

if run_clicked:
    with st.spinner("Matching rows and filling gaps..."):
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
                    not_found_records.append(
                        (row[key_column], col, "Key found, but that column is blank in reference sheet too")
                    )
                else:
                    main_df.at[idx, col] = ref_value
                    filled_cells.append((idx, col))
                    filled_count += 1

        not_found_df = pd.DataFrame(not_found_records, columns=[key_column, "Column Name", "Reason"])

    st.markdown("### Results")
    m1, m2, m3 = st.columns(3)
    m1.metric("Rows processed", len(main_df))
    m2.metric("Values filled", filled_count)
    m3.metric("Still missing", len(not_found_df))

    if filled_count > 0:
        st.success(f"✅ Filled {filled_count} value(s) from the reference sheet.")
    else:
        st.warning("No values were filled -- nothing matched, or there were no gaps to fill.")

    tab1, tab2 = st.tabs(["📄 Filled Data", "⚠️ Not Found"])
    with tab1:
        st.dataframe(main_df, use_container_width=True)
    with tab2:
        if len(not_found_df) > 0:
            st.dataframe(not_found_df, use_container_width=True)
        else:
            st.info("Nothing to show -- every gap was filled.")

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

    st.markdown("<br>", unsafe_allow_html=True)
    st.download_button(
        label="⬇️  Download output_filled.xlsx",
        data=output_buffer,
        file_name="output_filled.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown(
    """
    <div class="footer-note">Pentland Brands &middot; Internal Tool &middot; Runs entirely on your machine</div>
    """,
    unsafe_allow_html=True,
)
