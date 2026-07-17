import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
import os

# ------------------------------
# Page Configuration
# ------------------------------
st.set_page_config(
    page_title="Excel Sheet Combiner",
    page_icon="📊",
    layout="wide"
)

# ------------------------------
# Helper Function
# ------------------------------
def combine_excel_files(folder_path, target_sheet):
    folder = Path(folder_path)

    if not folder.exists():
        raise FileNotFoundError("Folder does not exist.")

    excel_files = list(folder.glob("*.xlsx"))

    if not excel_files:
        raise FileNotFoundError("No .xlsx files found in the selected folder.")

    combined_data = []

    progress_bar = st.progress(0)
    status_text = st.empty()

    total_files = len(excel_files)

    for index, file in enumerate(excel_files):

        try:
            status_text.info(
                f"Processing {index + 1}/{total_files}: {file.name}"
            )

            # Read ONLY required sheet
            df = pd.read_excel(
                file,
                sheet_name=target_sheet,
                engine="openpyxl"
            )

            df["Source_File"] = file.name
            df["Source_Sheet"] = target_sheet

            combined_data.append(df)

        except ValueError:
            st.warning(
                f"Sheet '{target_sheet}' not found in {file.name}"
            )

        except Exception as e:
            st.error(
                f"Error processing {file.name}: {str(e)}"
            )

        progress_bar.progress((index + 1) / total_files)

    status_text.empty()

    if combined_data:
        return pd.concat(
            combined_data,
            ignore_index=True,
            sort=False
        )

    return pd.DataFrame()


def save_output(df, output_file):
    df.to_excel(
        output_file,
        index=False,
        engine="openpyxl"
    )


# ------------------------------
# UI
# ------------------------------
st.title("📊 Excel Directory Combiner")

st.markdown("""
Combine a specific worksheet from multiple Excel files.

**Recommended for large files (>200 MB)**

Features:
- Process files directly from folder
- Reads only required worksheet
- Adds Source File tracking
- Progress monitoring
- Export combined output
""")

folder_path = st.text_input(
    "Input Folder Path",
    value=r"D:\DCT_Files"
)

sheet_name = st.text_input(
    "Sheet Name",
    value="Asset Attribute Mappings"
)

output_file = st.text_input(
    "Output File Path",
    value=r"D:\DCT_Files\Combined_DCT.xlsx"
)

remove_duplicates = st.checkbox(
    "Remove Duplicate Rows",
    value=False
)

# ------------------------------
# Process Button
# ------------------------------
if st.button("Combine Files", type="primary"):

    try:

        start_time = datetime.now()

        result_df = combine_excel_files(
            folder_path=folder_path,
            target_sheet=sheet_name
        )

        if result_df.empty:
            st.warning("No data found.")
            st.stop()

        if remove_duplicates:

            before = len(result_df)

            result_df = result_df.drop_duplicates()

            after = len(result_df)

            st.success(
                f"Removed {before - after:,} duplicate rows."
            )

        st.success(
            f"Successfully combined {len(result_df):,} rows."
        )

        save_output(
            result_df,
            output_file
        )

        elapsed = datetime.now() - start_time

        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric(
                "Rows",
                f"{len(result_df):,}"
            )

        with col2:
            st.metric(
                "Columns",
                result_df.shape[1]
            )

        with col3:
            st.metric(
                "Time",
                str(elapsed).split(".")[0]
            )

        st.subheader("Preview")

        st.dataframe(
            result_df.head(100),
            use_container_width=True
        )

        st.success(
            f"Output saved successfully:\n\n{output_file}"
        )

    except Exception as e:
        st.error(str(e))
