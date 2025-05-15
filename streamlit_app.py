import streamlit as st
import pandas as pd
import json
from io import BytesIO
from gstr1_converter import convert_gstr1_json_to_excel, get_all_sections_from_json

st.set_page_config(page_title="GSTR-1 JSON to Excel", layout="wide")
st.title("📊 GSTR-1 JSON to Excel Converter")

st.markdown("""
Upload your GSTR-1 JSON file and download an Excel file with each section in separate sheets.
""")

uploaded_file = st.file_uploader("Choose a GSTR-1 JSON file", type=["json"])

if uploaded_file is not None:
    try:
        json_data = json.load(uploaded_file)

        # Extract all sections
        dfs = get_all_sections_from_json(json_data)

        # Show summary
        st.subheader("📄 Conversion Summary")
        col1, col2 = st.columns(2)
        total_sheets = len(dfs)
        non_empty_sheets = sum(1 for df in dfs.values() if not df.empty)
        col1.metric("Total Sheets", total_sheets)
        col1.metric("Non-empty Sheets", non_empty_sheets)
        col2.info(f"Detected sections: {', '.join(dfs.keys())}")

        # Generate Excel in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in dfs.items():
                if not df.empty:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)

        # Download button
        st.download_button(
            label="📥 Download Excel File",
            data=output.getvalue(),
            file_name="converted_gstr1.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error: {e}")
        st.exception(e)