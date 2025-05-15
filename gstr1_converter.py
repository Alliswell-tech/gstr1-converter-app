import json
import pandas as pd
from io import BytesIO

# --- Extraction Functions (from your original file) ---

def extract_b2b_data(data):
    """Extracts B2B invoice data with calculated Invoice Value"""
    b2b_data = []
    try:
        b2b_section = data.get("b2b", [])
        for supplier in b2b_section:
            recipient_gstin = supplier.get("ctin")
            invoices = supplier.get("inv", [])
            for invoice in invoices:
                items = invoice.get("itms", [])
                total_taxable_value = 0
                total_igst = 0
                total_cgst = 0
                total_sgst = 0
                total_cess = 0
                for item in items:
                    item_details = item.get("itm_det", {})
                    taxable_value = item_details.get("txval", 0)
                    total_taxable_value += taxable_value
                    total_igst += item_details.get("iamt", 0)
                    total_cgst += item_details.get("camt", 0)
                    total_sgst += item_details.get("samt", 0)
                    total_cess += item_details.get("csamt", 0)
                # Calculate Invoice Value if not provided
                invoice_value = invoice.get("val")
                if invoice_value is None:
                    invoice_value = total_taxable_value + total_igst + total_cgst + total_sgst + total_cess
                b2b_data.append({
                    "Recipient GSTIN": recipient_gstin,
                    "Invoice Number": invoice.get("inum"),
                    "Invoice Date": invoice.get("idt"),
                    "Place of Supply": invoice.get("pos"),
                    "Reverse Charge": invoice.get("rchg"),
                    "Taxable Value": total_taxable_value,
                    "Total IGST": total_igst,
                    "Total CGST": total_cgst,
                    "Total SGST": total_sgst,
                    "Total CESS": total_cess,
                    "Invoice Value": invoice_value
                })
    except Exception as e:
        print(f"Error extracting B2B data: {e}")
        return []
    return b2b_data


# Include all other extraction functions exactly as they are in your code:
# extract_b2cl_data(), extract_b2cs_data(), extract_export_data(),
# extract_cdnr_data(), extract_cdunr_data(), extract_hsn_data(),
# extract_doc_issued_data()
# (You can copy them directly from your knowledge base or original file)


# --- Main Conversion Function (for local use) ---
def convert_gstr1_json_to_excel(json_file_path, excel_file_path):
    """Converts GSTR-1 JSON to Excel with multiple sheets"""
    try:
        # Load JSON data
        with open(json_file_path, "r", encoding="utf-8") as f:
            gstr1_data = json.load(f)

        sections = {
            "B2B Invoices": extract_b2b_data(gstr1_data),
            "B2C Large Invoices": extract_b2cl_data(gstr1_data),
            "B2C Small Summary": extract_b2cs_data(gstr1_data),
            "Exports": extract_export_data(gstr1_data),
            "Credit Debit Notes (Reg)": extract_cdnr_data(gstr1_data),
            "Credit Debit Notes (Unreg)": extract_cdunr_data(gstr1_data),
            "HSN Summary": extract_hsn_data(gstr1_data),
            "Document Issued Summary": extract_doc_issued_data(gstr1_data),
        }

        dfs = {sheet_name: pd.DataFrame(data) for sheet_name, data in sections.items()}

        with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
            for sheet_name, df in dfs.items():
                if not df.empty:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        print("Excel file created successfully.")
    except Exception as e:
        print(f"Error during conversion: {e}")


# --- New Function for Streamlit App ---
def get_all_sections_from_json(json_data):
    """
    Returns a dictionary of DataFrames for each section.
    Suitable for use in Streamlit apps where we work with in-memory data.
    """
    try:
        if isinstance(json_data, str):
            gstr1_data = json.loads(json_data)
        else:
            gstr1_data = json_data  # Assume it's already a dict

        sections = {
            "B2B Invoices": extract_b2b_data(gstr1_data),
            "B2C Large Invoices": extract_b2cl_data(gstr1_data),
            "B2C Small Summary": extract_b2cs_data(gstr1_data),
            "Exports": extract_export_data(gstr1_data),
            "Credit Debit Notes (Reg)": extract_cdnr_data(gstr1_data),
            "Credit Debit Notes (Unreg)": extract_cdunr_data(gstr1_data),
            "HSN Summary": extract_hsn_data(gstr1_data),
            "Document Issued Summary": extract_doc_issued_data(gstr1_data),
        }

        dfs = {sheet_name: pd.DataFrame(data) for sheet_name, data in sections.items()}
        return dfs

    except Exception as e:
        raise RuntimeError(f"Error extracting sections: {e}")


# --- For Local Use Only ---
if __name__ == "__main__":
    print("This script is designed to be imported. Run the CLI version instead.")