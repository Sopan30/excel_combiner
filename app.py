import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path

st.set_page_config(
    page_title="Excel Directory Combiner",
    page_icon="📊",
    layout="wide"
)


class ExcelDirectoryCombiner:
    def __init__(self, uploaded_files, target_sheets=None):
        self.uploaded_files = uploaded_files
        self.target_sheets = target_sheets or []

    def combine_files(self):
        combined_data = []

        for uploaded_file in self.uploaded_files:
            try:
                sheets = pd.read_excel(uploaded_file, sheet_name=None)

                for sheet_name, df in sheets.items():

                    # Process only selected sheets
                    if not self.target_sheets or sheet_name in self.target_sheets:

                        df["Source_File"] = uploaded_file.name
                        df["Source_Sheet"] = sheet_name

                        combined_data.append(df)

            except Exception as e:
                st.error(f"Error processing {uploaded_file.name}: {e}")

        if combined_data:
            return pd.concat(combined_data, ignore_index=True, sort=False)

        return pd.DataFrame()

    def save_to_excel(self, dataframe):
        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            dataframe.to_excel(
                writer,
                sheet_name="Combined_Data",
                index=False
            )

        output.seek(0)
        return output


st.title("📊 Excel Directory Combiner")

st.markdown(
    """
    Upload multiple Excel files and combine data from a specific worksheet
    across all files.
    """
)

uploaded_files = st.file_uploader(
    "Upload Excel Files",
    type=["xlsx"],
    accept_multiple_files=True
)

sheet_name = st.text_input(
    "Worksheet Name to Combine",
    value="Asset Attribute Mappings"
)

if uploaded_files:

    st.success(f"{len(uploaded_files)} file(s) uploaded")

    if st.button("Combine Files"):

        with st.spinner("Processing files..."):

            combiner = ExcelDirectoryCombiner(
                uploaded_files=uploaded_files,
                target_sheets=[sheet_name]
            )

            combined_df = combiner.combine_files()

        if not combined_df.empty:

            st.success(
                f"Successfully combined {len(combined_df)} rows."
            )

            st.dataframe(
                combined_df,
                use_container_width=True,
                height=500
            )

            excel_file = combiner.save_to_excel(combined_df)

            st.download_button(
                label="📥 Download Combined Excel",
                data=excel_file,
                file_name="Combined_DCT.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        else:
            st.warning(
                "No matching worksheet found in uploaded files."
            )