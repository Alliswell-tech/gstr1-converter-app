# GSTR-1 JSON to Excel Converter

A tool to convert **GSTR-1 JSON files** (downloaded from the GST Portal) into **Excel spreadsheets with multiple sheets**, making it easy to analyze and share data.

This project includes:
- A **Python script** to convert GSTR-1 JSON to Excel.
- A **Streamlit web app** so users can upload JSON files via a browser and download Excel files without installing Python.

---

## 📄 Features

✅ Converts the following GSTR-1 sections into separate Excel sheets:
- B2B Invoices
- B2C Large Invoices
- B2C Small Summary
- Export Invoices
- Credit/Debit Notes (Registered)
- Credit/Debit Notes (Unregistered)
- HSN/SAC Summary
- Document Issued Summary

✅ Automatically calculates missing values like:
- Taxable Value
- Invoice Value
- Total IGST / CGST / SGST / CESS

✅ User-friendly **web interface** using **Streamlit**

---

## 🔧 How It Works

The script parses the GSTR-1 JSON file, extracts invoice-level data from each section, organizes them into structured tables, and writes all of them into a single Excel file with one sheet per section.

---

## 🚀 How to Use (Local Version)

### Prerequisites
- Python 3.8+
- `pandas`, `openpyxl`

### Steps

1. Clone the repo:
   ```bash
   git clone https://github.com/Alliswell-tech/gstr1-converter-app.git 
   cd gstr1-converter-app