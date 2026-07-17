import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="Excel Sheet Combiner",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Excel Sheet Combiner")

st.markdown("""
Combine a specific sheet from multiple Excel files.

**Example Sheet:** Asset Attribute Mappings
""")

sheet_name = st.text_input(
    "Sheet Name to Combine",
    value="Asset Attribute Mappings"
)

uploaded_files = st.file_uploader(
    "Upload Excel Files",
    type=["xlsx"],
    accept_multiple_files=True
)

remove_duplicates = st.checkbox(
    "Remove Duplicate Rows",
    value=False
)

def create_excel_file(df):
    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:
        df.to_excel(
            writer,
            index=False,
            sheet_name="Combined_Data"
        )

    output.seek(0)
    return output

if st.button("Combine Files", type="primary"):

    if not uploaded_files:
        st.warning("Please upload Excel files.")
        st.stop()

    start_time = datetime.now()

    combined_data = []

    progress = st.progress(0)
    status = st.empty()

    total_files = len(uploaded_files)

    for i, file in enumerate(uploaded_files):

        status.info(
            f"Processing {i+1}/{total_files}: {file.name}"
        )

        try:
            df = pd.read_excel(
                file,
                sheet_name=sheet_name,
                engine="openpyxl"
            )

            df["Source_File"] = file.name
            df["Source_Sheet"] = sheet_name

            combined_data.append(df)

        except ValueError:
            st.warning(
                f"Sheet '{sheet_name}' not found in {file.name}"
            )

        except Exception as e:
            st.error(
                f"{file.name}: {str(e)}"
            )

        progress.progress((i + 1) / total_files)

    if not combined_data:
        st.error("No matching sheet found.")
        st.stop()

    result_df = pd.concat(
        combined_data,
        ignore_index=True,
        sort=False
    )

    if remove_duplicates:
        before = len(result_df)
        result_df = result_df.drop_duplicates()
        removed = before - len(result_df)

        st.success(
            f"Removed {removed:,} duplicate rows"
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

    excel_output = create_excel_file(result_df)

    st.download_button(
        label="📥 Download Combined Excel",
        data=excel_output,
        file_name="Combined_DCT.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    csv_output = result_df.to_csv(index=False).encode("utf-8")

    st.download_button(
        label="📄 Download CSV (Recommended for Large Data)",
        data=csv_output,
        file_name="Combined_DCT.csv",
        mime="text/csv"
    )

    st.success("Processing completed successfully ✅")
