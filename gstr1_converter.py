import json
import pandas as pd
from io import BytesIO

# --- Extraction Functions ---

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
                    "Reverse Charge": invoice.get("rchrg"),
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


def extract_b2cl_data(data):
    """Extracts B2C Large invoices with calculated Invoice Value"""
    b2cl_data = []
    try:
        b2cl_section = data.get("b2cl", [])
        for invoice in b2cl_section:
            items = invoice.get("itms", [])
            total_taxable_value = 0
            total_igst = 0
            for item in items:
                item_details = item.get("itm_det", {})
                taxable_value = item_details.get("txval", 0)
                total_taxable_value += taxable_value
                total_igst += item_details.get("iamt", 0)
            # Calculate Invoice Value if not provided
            invoice_value = invoice.get("val")
            if invoice_value is None:
                invoice_value = total_taxable_value + total_igst
            b2cl_data.append({
                "Invoice Number": invoice.get("inum"),
                "Invoice Date": invoice.get("idt"),
                "Place of Supply": invoice.get("pos"),
                "Taxable Value": total_taxable_value,
                "Total IGST": total_igst,
                "Invoice Value": invoice_value
            })
    except Exception as e:
        print(f"Error extracting B2CL data: {e}")
        return []
    return b2cl_data


def extract_b2cs_data(data):
    """Extracts B2C Small summary data"""
    b2cs_data = []
    try:
        b2cs_section = data.get("b2cs", [])
        for item in b2cs_section:
            b2cs_data.append({
                "Place of Supply": item.get("pos"),
                "Taxable Value": item.get("txval"),
                "Rate": item.get("rt"),
                "IGST": item.get("iamt", 0),
                "CGST": item.get("camt", 0),
                "SGST": item.get("samt", 0),
                "CESS": item.get("csamt", 0)
            })
    except Exception as e:
        print(f"Error extracting B2CS data: {e}")
        return []
    return b2cs_data


def extract_export_data(data):
    """Extracts Export invoices with calculated Invoice Value"""
    export_data = []
    try:
        export_section = data.get("exp", [])
        for invoice in export_section:
            items = invoice.get("itms", [])
            total_taxable_value = 0
            total_igst = 0
            for item in items:
                item_details = item.get("itm_det", {})
                taxable_value = item_details.get("txval", 0)
                total_taxable_value += taxable_value
                total_igst += item_details.get("iamt", 0)
            # Calculate Invoice Value if not provided
            invoice_value = invoice.get("val")
            if invoice_value is None:
                invoice_value = total_taxable_value + total_igst
            export_data.append({
                "Invoice Number": invoice.get("inum"),
                "Invoice Date": invoice.get("idt"),
                "Port Code": invoice.get("pcode"),
                "Shipping Bill Number": invoice.get("sbnum"),
                "Shipping Bill Date": invoice.get("sbdt"),
                "Taxable Value": total_taxable_value,
                "Total IGST": total_igst,
                "Invoice Value": invoice_value
            })
    except Exception as e:
        print(f"Error extracting Export data: {e}")
        return []
    return export_data


def extract_cdnr_data(data):
    """Extracts Credit/Debit Notes (Registered) with calculated Note Value"""
    cdnr_data = []
    try:
        cdnr_section = data.get("cdnr", [])
        for supplier_note in cdnr_section:
            recipient_gstin = supplier_note.get("ctin")
            notes = supplier_note.get("nt", [])
            for note in notes:
                items = note.get("itms", [])
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
                # Calculate Note Value if not provided
                note_value = note.get("val")
                if note_value is None:
                    note_value = total_taxable_value + total_igst + total_cgst + total_sgst + total_cess
                cdnr_data.append({
                    "Recipient GSTIN": recipient_gstin,
                    "Note Number": note.get("nt_num"),
                    "Note Date": note.get("nt_dt"),
                    "Note Type": note.get("ntty"),
                    "Original Invoice Number": note.get("oinum"),
                    "Original Invoice Date": note.get("oidt"),
                    "Place of Supply": note.get("pos"),
                    "Reason": note.get("rsn"),
                    "Taxable Value": total_taxable_value,
                    "Total IGST": total_igst,
                    "Total CGST": total_cgst,
                    "Total SGST": total_sgst,
                    "Total CESS": total_cess,
                    "Note Value": note_value
                })
    except Exception as e:
        print(f"Error extracting CDNR data: {e}")
        return []
    return cdnr_data


def extract_cdunr_data(data):
    """Extracts Credit/Debit Notes (Unregistered) with calculated Note Value"""
    cdunr_data = []
    try:
        cdunr_section = data.get("cdunr", [])
        for note in cdunr_section:
            items = note.get("itms", [])
            total_taxable_value = 0
            total_igst = 0
            total_cess = 0
            for item in items:
                item_details = item.get("itm_det", {})
                taxable_value = item_details.get("txval", 0)
                total_taxable_value += taxable_value
                total_igst += item_details.get("iamt", 0)
                total_cess += item_details.get("csamt", 0)
            # Calculate Note Value if not provided
            note_value = note.get("val")
            if note_value is None:
                note_value = total_taxable_value + total_igst + total_cess
            cdunr_data.append({
                "Note Number": note.get("nt_num"),
                "Note Date": note.get("nt_dt"),
                "Note Type": note.get("ntty"),
                "Original Invoice Number": note.get("oinum"),
                "Original Invoice Date": note.get("oidt"),
                "Place of Supply": note.get("pos"),
                "Reason": note.get("rsn"),
                "Taxable Value": total_taxable_value,
                "Total IGST": total_igst,
                "Total CESS": total_cess,
                "Note Value": note_value
            })
    except Exception as e:
        print(f"Error extracting CDUNR data: {e}")
        return []
    return cdunr_data


def extract_hsn_data(data):
    """Extracts HSN summary data"""
    hsn_data = []
    try:
        hsn_section = data.get("hsn", {})
        hsn_data_list = hsn_section.get("data", [])
        for hsn_item in hsn_data_list:
            hsn_data.append({
                "HSN Code": hsn_item.get("hsn_sc"),
                "Description": hsn_item.get("desc"),
                "UQC": hsn_item.get("uqc"),
                "Quantity": hsn_item.get("qty", 0),
                "Total Value": hsn_item.get("val", 0),
                "Taxable Value": hsn_item.get("txval", 0),
                "IGST": hsn_item.get("iamt", 0),
                "CGST": hsn_item.get("camt", 0),
                "SGST": hsn_item.get("samt", 0),
                "CESS": hsn_item.get("csamt", 0),
            })
    except Exception as e:
        print(f"Error extracting HSN data: {e}")
        return []
    return hsn_data


def extract_doc_issued_data(data):
    """Extracts document issued summary data"""
    doc_data = []
    try:
        doc_section = data.get("doc_issue", {})
        doc_list = doc_section.get("doc_det", [])
        for doc_series in doc_list:
            doc_num = doc_series.get("doc_num", 0)
            docs = doc_series.get("docs", [])
            for doc_detail in docs:
                doc_data.append({
                    "Document Type Index": doc_num,
                    "Serial Number From": doc_detail.get("from"),
                    "Serial Number To": doc_detail.get("to"),
                    "Total Issued": doc_detail.get("totnum"),
                    "Cancelled": doc_detail.get("cancel"),
                    "Net Issued": doc_detail.get("net_issue")
                })
    except Exception as e:
        print(f"Error extracting Document Issued data: {e}")
        return []
    return doc_data


def extract_nil_rated_data(data):
    """Extracts Nil Rated Supplies data"""
    nil_rated_data = []
    try:
        nil_rated_section = data.get("nil", [])
        for item in nil_rated_section:
            nil_rated_data.append({
                "Place of Supply": item.get("pos"),
                "Taxable Value": item.get("txval"),
                "Rate": item.get("rt"),
                "IGST": item.get("iamt", 0),
                "CGST": item.get("camt", 0),
                "SGST": item.get("samt", 0),
                "CESS": item.get("csamt", 0)
            })
    except Exception as e:
        print(f"Error extracting Nil Rated data: {e}")
        return []
    return nil_rated_data


def extract_amended_b2b_data(data):
    """Extracts Amended B2B Invoices data"""
    amended_b2b_data = []
    try:
        amended_b2b_section = data.get("amend_b2b", [])
        for supplier in amended_b2b_section:
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
                amended_b2b_data.append({
                    "Recipient GSTIN": recipient_gstin,
                    "Invoice Number": invoice.get("inum"),
                    "Invoice Date": invoice.get("idt"),
                    "Place of Supply": invoice.get("pos"),
                    "Reverse Charge": invoice.get("rchrg"),
                    "Taxable Value": total_taxable_value,
                    "Total IGST": total_igst,
                    "Total CGST": total_cgst,
                    "Total SGST": total_sgst,
                    "Total CESS": total_cess,
                    "Invoice Value": invoice_value
                })
    except Exception as e:
        print(f"Error extracting Amended B2B data: {e}")
        return []
    return amended_b2b_data


def extract_amended_b2cl_data(data):
    """Extracts Amended B2C (Large) Invoices data"""
    amended_b2cl_data = []
    try:
        amended_b2cl_section = data.get("amend_b2cl", [])
        for invoice in amended_b2cl_section:
            items = invoice.get("itms", [])
            total_taxable_value = 0
            total_igst = 0
            for item in items:
                item_details = item.get("itm_det", {})
                taxable_value = item_details.get("txval", 0)
                total_taxable_value += taxable_value
                total_igst += item_details.get("iamt", 0)
            # Calculate Invoice Value if not provided
            invoice_value = invoice.get("val")
            if invoice_value is None:
                invoice_value = total_taxable_value + total_igst
            amended_b2cl_data.append({
                "Invoice Number": invoice.get("inum"),
                "Invoice Date": invoice.get("idt"),
                "Place of Supply": invoice.get("pos"),
                "Taxable Value": total_taxable_value,
                "Total IGST": total_igst,
                "Invoice Value": invoice_value
            })
    except Exception as e:
        print(f"Error extracting Amended B2CL data: {e}")
        return []
    return amended_b2cl_data


def extract_amended_export_data(data):
    """Extracts Amended Export Invoices data"""
    amended_export_data = []
    try:
        amended_export_section = data.get("amend_exp", [])
        for invoice in amended_export_section:
            items = invoice.get("itms", [])
            total_taxable_value = 0
            total_igst = 0
            for item in items:
                item_details = item.get("itm_det", {})
                taxable_value = item_details.get("txval", 0)
                total_taxable_value += taxable_value
                total_igst += item_details.get("iamt", 0)
            # Calculate Invoice Value if not provided
            invoice_value = invoice.get("val")
            if invoice_value is None:
                invoice_value = total_taxable_value + total_igst
            amended_export_data.append({
                "Invoice Number": invoice.get("inum"),
                "Invoice Date": invoice.get("idt"),
                "Port Code": invoice.get("pcode"),
                "Shipping Bill Number": invoice.get("sbnum"),
                "Shipping Bill Date": invoice.get("sbdt"),
                "Taxable Value": total_taxable_value,
                "Total IGST": total_igst,
                "Invoice Value": invoice_value
            })
    except Exception as e:
        print(f"Error extracting Amended Export data: {e}")
        return []
    return amended_export_data


def extract_amended_cdnr_data(data):
    """Extracts Amended Credit/Debit Notes (Registered) data"""
    amended_cdnr_data = []
    try:
        amended_cdnr_section = data.get("amend_cdnr", [])
        for supplier_note in amended_cdnr_section:
            recipient_gstin = supplier_note.get("ctin")
            notes = supplier_note.get("nt", [])
            for note in notes:
                items = note.get("itms", [])
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
                # Calculate Note Value if not provided
                note_value = note.get("val")
                if note_value is None:
                    note_value = total_taxable_value + total_igst + total_cgst + total_sgst + total_cess
                amended_cdnr_data.append({
                    "Recipient GSTIN": recipient_gstin,
                    "Note Number": note.get("nt_num"),
                    "Note Date": note.get("nt_dt"),
                    "Note Type": note.get("ntty"),
                    "Original Invoice Number": note.get("oinum"),
                    "Original Invoice Date": note.get("oidt"),
                    "Place of Supply": note.get("pos"),
                    "Reason": note.get("rsn"),
                    "Taxable Value": total_taxable_value,
                    "Total IGST": total_igst,
                    "Total CGST": total_cgst,
                    "Total SGST": total_sgst,
                    "Total CESS": total_cess,
                    "Note Value": note_value
                })
    except Exception as e:
        print(f"Error extracting Amended CDNR data: {e}")
        return []
    return amended_cdnr_data


def extract_amended_cdunr_data(data):
    """Extracts Amended Credit/Debit Notes (Unregistered) data"""
    amended_cdunr_data = []
    try:
        amended_cdunr_section = data.get("amend_cdunr", [])
        for note in amended_cdunr_section:
            items = note.get("itms", [])
            total_taxable_value = 0
            total_igst = 0
            total_cess = 0
            for item in items:
                item_details = item.get("itm_det", {})
                taxable_value = item_details.get("txval", 0)
                total_taxable_value += taxable_value
                total_igst += item_details.get("iamt", 0)
                total_cess += item_details.get("csamt", 0)
            # Calculate Note Value if not provided
            note_value = note.get("val")
            if note_value is None:
                note_value = total_taxable_value + total_igst + total_cess
            amended_cdunr_data.append({
                "Note Number": note.get("nt_num"),
                "Note Date": note.get("nt_dt"),
                "Note Type": note.get("ntty"),
                "Original Invoice Number": note.get("oinum"),
                "Original Invoice Date": note.get("oidt"),
                "Place of Supply": note.get("pos"),
                "Reason": note.get("rsn"),
                "Taxable Value": total_taxable_value,
                "Total IGST": total_igst,
                "Total CESS": total_cess,
                "Note Value": note_value
            })
    except Exception as e:
        print(f"Error extracting Amended CDUNR data: {e}")
        return []
    return amended_cdunr_data


def extract_amended_hsn_data(data):
    """Extracts Amended HSN summary data"""
    amended_hsn_data = []
    try:
        amended_hsn_section = data.get("amend_hsn", {})
        amended_hsn_list = amended_hsn_section.get("data", [])
        for hsn_item in amended_hsn_list:
            amended_hsn_data.append({
                "HSN Code": hsn_item.get("hsn_sc"),
                "Description": hsn_item.get("desc"),
                "UQC": hsn_item.get("uqc"),
                "Quantity": hsn_item.get("qty", 0),
                "Total Value": hsn_item.get("val", 0),
                "Taxable Value": hsn_item.get("txval", 0),
                "IGST": hsn_item.get("iamt", 0),
                "CGST": hsn_item.get("camt", 0),
                "SGST": hsn_item.get("samt", 0),
                "CESS": hsn_item.get("csamt", 0),
            })
    except Exception as e:
        print(f"Error extracting Amended HSN data: {e}")
        return []
    return amended_hsn_data


def extract_amended_doc_issued_data(data):
    """Extracts Amended Document Issued summary data"""
    amended_doc_data = []
    try:
        amended_doc_section = data.get("amend_doc_issue", {})
        amended_doc_list = amended_doc_section.get("doc_det", [])
        for doc_series in amended_doc_list:
            doc_num = doc_series.get("doc_num", 0)
            docs = doc_series.get("docs", [])
            for doc_detail in docs:
                amended_doc_data.append({
                    "Document Type Index": doc_num,
                    "Serial Number From": doc_detail.get("from"),
                    "Serial Number To": doc_detail.get("to"),
                    "Total Issued": doc_detail.get("totnum"),
                    "Cancelled": doc_detail.get("cancel"),
                    "Net Issued": doc_detail.get("net_issue")
                })
    except Exception as e:
        print(f"Error extracting Amended Document Issued data: {e}")
        return []
    return amended_doc_data


def extract_amended_nil_rated_data(data):
    """Extracts Amended Nil Rated Supplies data"""
    amended_nil_rated_data = []
    try:
        amended_nil_rated_section = data.get("amend_nil", [])
        for item in amended_nil_rated_section:
            amended_nil_rated_data.append({
                "Place of Supply": item.get("pos"),
                "Taxable Value": item.get("txval"),
                "Rate": item.get("rt"),
                "IGST": item.get("iamt", 0),
                "CGST": item.get("camt", 0),
                "SGST": item.get("samt", 0),
                "CESS": item.get("csamt", 0)
            })
    except Exception as e:
        print(f"Error extracting Amended Nil Rated data: {e}")
        return []
    return amended_nil_rated_data


# --- Main Conversion Logic ---

def convert_gstr1_json_to_excel_bytes(json_data):
    """
    Converts GSTR-1 JSON data (as string or dict) to an Excel file in memory.
    Returns: BytesIO object containing the Excel file
    """
    try:
        # Parse JSON if it's a string
        if isinstance(json_data, str):
            gstr1_data = json.loads(json_data)
        else:
            gstr1_data = json_data  # Assume it's already a dict

        # Extract sections
        sections = {
            "B2B Invoices": extract_b2b_data(gstr1_data),
            "B2C Large Invoices": extract_b2cl_data(gstr1_data),
            "B2C Small Summary": extract_b2cs_data(gstr1_data),
            "Exports": extract_export_data(gstr1_data),
            "Credit Debit Notes (Reg)": extract_cdnr_data(gstr1_data),
            "Credit Debit Notes (Unreg)": extract_cdunr_data(gstr1_data),
            "HSN Summary": extract_hsn_data(gstr1_data),
            "Document Issued Summary": extract_doc_issued_data(gstr1_data),
            "Nil Rated Supplies": extract_nil_rated_data(gstr1_data),
            "Amended B2B Invoices": extract_amended_b2b_data(gstr1_data),
            "Amended B2C Large Invoices": extract_amended_b2cl_data(gstr1_data),
            "Amended Exports": extract_amended_export_data(gstr1_data),
            "Amended Credit Debit Notes (Reg)": extract_amended_cdnr_data(gstr1_data),
            "Amended Credit Debit Notes (Unreg)": extract_amended_cdunr_data(gstr1_data),
            "Amended HSN Summary": extract_amended_hsn_data(gstr1_data),
            "Amended Document Issued Summary": extract_amended_doc_issued_data(gstr1_data),
            "Amended Nil Rated Supplies": extract_amended_nil_rated_data(gstr1_data),
        }

        # Create DataFrames
        dfs = {sheet_name: pd.DataFrame(data) for sheet_name, data in sections.items() if data}

        # Write to BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in dfs.items():
                if not df.empty:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        output.seek(0)
        return output

    except Exception as e:
        raise RuntimeError(f"Error during conversion: {e}")


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
            "Nil Rated Supplies": extract_nil_rated_data(gstr1_data),
            "Amended B2B Invoices": extract_amended_b2b_data(gstr1_data),
            "Amended B2C Large Invoices": extract_amended_b2cl_data(gstr1_data),
            "Amended Exports": extract_amended_export_data(gstr1_data),
            "Amended Credit Debit Notes (Reg)": extract_amended_cdnr_data(gstr1_data),
            "Amended Credit Debit Notes (Unreg)": extract_amended_cdunr_data(gstr1_data),
            "Amended HSN Summary": extract_amended_hsn_data(gstr1_data),
            "Amended Document Issued Summary": extract_amended_doc_issued_data(gstr1_data),
            "Amended Nil Rated Supplies": extract_amended_nil_rated_data(gstr1_data),
        }

        dfs = {sheet_name: pd.DataFrame(data) for sheet_name, data in sections.items() if data}
        return dfs

    except Exception as e:
        raise RuntimeError(f"Error extracting sections: {e}")