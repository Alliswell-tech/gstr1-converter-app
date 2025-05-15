import json
import pandas as pd
from io import BytesIO

# --- Enhanced Extraction Functions ---

def extract_b2b_data(data):
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
    except Exception:
        return []
    return b2b_data

def extract_b2cl_data(data):
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
    except Exception:
        return []
    return b2cl_data

def extract_b2cs_data(data):
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
    except Exception:
        return []
    return b2cs_data

def extract_export_data(data):
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
    except Exception:
        return []
    return export_data

def extract_cdnr_data(data):
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
    except Exception:
        return []
    return cdnr_data

def extract_cdunr_data(data):
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
    except Exception:
        return []
    return cdunr_data

def extract_hsn_data(data):
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
    except Exception:
        return []
    return hsn_data

def extract_doc_issued_data(data):
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
    except Exception:
        return []
    return doc_data

# --- Main Conversion Function ---

def convert_gstr1_json_to_excel_bytes(json_data):
    try:
        if isinstance(json_data, str):
            gstr1_data = json.loads(json_data)
        else:
            gstr1_data = json_data

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

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in dfs.items():
                if not df.empty:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
        output.seek(0)
        return output

    except Exception as e:
        raise RuntimeError(f"Error during conversion: {e}")