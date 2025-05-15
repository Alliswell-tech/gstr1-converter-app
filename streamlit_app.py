import streamlit as st
import pandas as pd
import json
from io import BytesIO

# Import conversion function from gstr1_converter.py
from gstr1_converter import get_all_sections_from_json

# Set page config
st.set_page_config(page_title="GSTR-1 JSON to Excel", layout="wide")

# Title and description
st.title("📊 GSTR-1 JSON to Excel Converter")
st.markdown("""
Upload your **GSTR-1 JSON file** (downloaded from the GST Portal) and get an **Excel file** with each section in separate sheets.
""")

# File upload
uploaded_file = st.file_uploader("Choose a GSTR-1 JSON file", type=["json"])

if uploaded_file is not None:
    try:
        # Load JSON data from uploaded file
        json_data = json.load(uploaded_file)

        # Extract all sections into DataFrames
        dfs = get_all_sections_from_json(json_data)

        # Show summary of extracted data
        st.subheader("📄 Conversion Summary")
        col1, col2 = st.columns(2)
        total_sheets = len(dfs)
        non_empty_sheets = sum(1 for df in dfs.values() if not df.empty)
        col1.metric("Total Sheets", total_sheets)
        col1.metric("Non-empty Sheets", non_empty_sheets)
        col2.info(f"Sections: {', '.join(dfs.keys())}")

        # Generate Excel file in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in dfs.items():
                if not df.empty:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)

        # Add download button
        st.download_button(
            label="📥 Download Excel File",
            data=output.getvalue(),
            file_name="converted_gstr1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except json.JSONDecodeError:
        st.error("❌ Invalid JSON file. Please ensure it's a valid GSTR-1 JSON file.")
    except Exception as e:
        st.error(f"❌ Error during conversion: {e}")
        st.exception(e)