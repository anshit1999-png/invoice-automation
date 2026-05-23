# =========================================
# IMPORT LIBRARIES
# =========================================

import streamlit as st
import tempfile
import zipfile
import os
import pdfplumber
import pandas as pd
import re

# =========================================
# PAGE CONFIG
# =========================================

st.set_page_config(
    page_title="Invoice Reconciliation Automation",
    layout="wide"
)

st.title("Invoice Reconciliation Automation System")

# =========================================
# FILE UPLOADS
# =========================================

zip_file = st.file_uploader(
    "Upload ZIP File",
    type=["zip"]
)

excel_file = st.file_uploader(
    "Upload Excel File",
    type=["xlsx"]
)

run_button = st.button("Run Automation")

# =========================================
# MAIN PROCESS
# =========================================

if run_button:

    if zip_file is None or excel_file is None:

        st.error(
            "Please upload both ZIP and Excel files"
        )

    else:

        st.success(
            "Files uploaded successfully"
        )

        with tempfile.TemporaryDirectory() as temp_dir:

            # =========================================
            # SAVE ZIP
            # =========================================

            zip_path = os.path.join(
                temp_dir,
                zip_file.name
            )

            with open(zip_path, "wb") as f:

                f.write(zip_file.read())

            # =========================================
            # EXTRACT ZIP
            # =========================================

            extract_path = os.path.join(
                temp_dir,
                "Extracted_Files"
            )

            with zipfile.ZipFile(
                zip_path,
                'r'
            ) as zip_ref:

                zip_ref.extractall(
                    extract_path
                )

            st.success(
                "ZIP Extracted Successfully"
            )

            # =========================================
            # READ EXCEL
            # =========================================

            excel_sheets = pd.ExcelFile(
                excel_file
            )

            selected_sheet = None

            for sheet in excel_sheets.sheet_names:

                try:

                    temp_df = pd.read_excel(
                        excel_file,
                        sheet_name=sheet,
                        nrows=5
                    )

                    cols = [

                        str(c).lower()

                        for c in temp_df.columns
                    ]

                    invoice_found = any(

                        "invoice" in c

                        for c in cols
                    )

                    amount_found = any(

                        "amt" in c
                        or
                        "amount" in c

                        for c in cols
                    )

                    if (
                        invoice_found
                        and
                        amount_found
                    ):

                        selected_sheet = sheet
                        break

                except:

                    pass

            # =========================================
            # NO SHEET FOUND
            # =========================================

            if selected_sheet is None:

                st.error(
                    "No valid billing sheet found"
                )

                st.stop()

            # =========================================
            # LOAD SHEET
            # =========================================

            detail_df = pd.read_excel(
                excel_file,
                sheet_name=selected_sheet
            )

            st.success(
                f"Detected Sheet : {selected_sheet}"
            )

            # =========================================
            # CLEAN COLUMN NAMES
            # =========================================

            detail_df.columns = (

                detail_df.columns
                .astype(str)
                .str.strip()

            )

            # =========================================
            # DYNAMIC COLUMN DETECTION
            # =========================================

            invoice_col = None
            vendor_col = None
            amount_col = None

            for col in detail_df.columns:

                col_lower = col.lower().strip()

                # =========================================
                # INVOICE COLUMN
                # =========================================

                if (
                    "invoice" in col_lower
                    and
                    (
                        "no" in col_lower
                        or
                        "number" in col_lower
                    )
                ):

                    invoice_col = col

                # =========================================
                # VENDOR COLUMN
                # =========================================

                elif (

                    "vendor" in col_lower
                    or
                    "tpt" in col_lower
                    or
                    "code" in col_lower

                ):

                    vendor_col = col

                # =========================================
                # AMOUNT COLUMN
                # =========================================

                elif (

                    col_lower.startswith("t")
                    or
                    col_lower.startswith("to")
                    or
                    col_lower.startswith("tot")

                ):

                    words = col_lower.split()

                    if len(words) >= 2:

                        second_word = words[-1]

                        if (

                            second_word.startswith("am")
                            or
                            second_word.startswith("amt")
                            or
                            second_word.startswith("a")

                        ):

                            amount_col = col

            # =========================================
            # COLUMN CHECK
            # =========================================

            if (
                invoice_col is None
                or
                vendor_col is None
                or
                amount_col is None
            ):

                st.error(
                    "Required columns not found in Excel"
                )

                st.write(
                    detail_df.columns.tolist()
                )

                st.stop()

            st.success(f"""

            Invoice Column : {invoice_col}

            Vendor Column : {vendor_col}

            Amount Column : {amount_col}

            """)

            # =========================================
            # CLEAN EXCEL DATA
            # =========================================

            detail_df[invoice_col] = (

                detail_df[invoice_col]
                .astype(str)
                .str.strip()
                .str.replace(
                    ".0",
                    "",
                    regex=False
                )

            )

            detail_df[vendor_col] = (

                detail_df[vendor_col]
                .astype(str)
                .str.strip()
                .str.replace(
                    ".0",
                    "",
                    regex=False
                )

            )

            detail_df[amount_col] = (

                detail_df[amount_col]
                .astype(str)
                .str.replace(",", "")

            )

            detail_df[amount_col] = pd.to_numeric(

                detail_df[amount_col],
                errors="coerce"

            )

            # =========================================
            # PROCESS PDFS
            # =========================================

            results = []

            folders = os.listdir(
                extract_path
            )

            for folder in folders:

                folder_path = os.path.join(
                    extract_path,
                    folder
                )

                if os.path.isdir(folder_path):

                    for file in os.listdir(folder_path):

                        if file.lower().endswith(".pdf"):

                            # Ignore cheque files

                            if re.search(

                                r'Cheque Request|Chq\. Req|Cheq\. Req',

                                file,

                                re.IGNORECASE
                            ):

                                continue

                            pdf_path = os.path.join(
                                folder_path,
                                file
                            )

                            full_text = ""

                            try:

                                with pdfplumber.open(
                                    pdf_path
                                ) as pdf:

                                    for page in pdf.pages:

                                        text = page.extract_text()

                                        if text:

                                            full_text += text

                            except Exception as e:

                                st.write(
                                    f"PDF Error : {file}"
                                )

                                st.write(e)

                            # =========================================
                            # EXTRACT INVOICE NUMBER
                            # =========================================

                            invoice_match = re.search(

                                r'Invoice\s*No\.?\s*[:\-]?\s*([A-Z0-9\/\-]+)',

                                full_text,

                                re.IGNORECASE
                            )

                            if invoice_match:

                                invoice_number = (
                                    invoice_match.group(1)
                                )

                            else:

                                numbers = re.findall(
                                    r'\d+',
                                    file
                                )

                                if numbers:

                                    invoice_number = numbers[-1]

                                else:

                                    invoice_number = "NOT FOUND"

                            # =========================================
                            # SAVE RESULT
                            # =========================================

                            results.append({

                                "Folder": folder,
                                "File": file,
                                "Invoice_Number": invoice_number,
                                "PDF_Text": full_text

                            })

            # =========================================
            # CREATE REPORT DF
            # =========================================

            report_df = pd.DataFrame(
                results
            )

            # =========================================
            # VALIDATION FUNCTION
            # =========================================

            def validate_invoice(row):

                invoice = str(
                    row["Invoice_Number"]
                ).strip()

                pdf_text = str(
                    row["PDF_Text"]
                )

                pdf_text_clean = re.sub(
                    r'[^A-Za-z0-9.]',
                    '',
                    pdf_text
                )

                # =========================================
                # FIND INVOICE ROWS
                # =========================================

                invoice_rows = detail_df[

                    detail_df[invoice_col]
                    .astype(str)
                    .str.strip() == invoice

                ]

                # =========================================
                # INVOICE NOT FOUND
                # =========================================

                if len(invoice_rows) == 0:

                    return pd.Series([

                        "NOT FOUND",
                        "NOT FOUND",
                        "NOT FOUND",
                        "INVOICE NOT FOUND"

                    ])

                # =========================================
                # MATCH VENDOR + INVOICE
                # =========================================

                matched_row = None

                for idx, excel_row in invoice_rows.iterrows():

                    vendor_code = str(
                        excel_row[vendor_col]
                    ).strip()

                    vendor_code_clean = re.sub(
                        r'\.0$',
                        '',
                        vendor_code
                    )

                    vendor_found = (

                        vendor_code_clean
                        in
                        pdf_text_clean

                    )

                    # BOTH CONDITIONS REQUIRED

                    if (
                        vendor_found
                        and
                        invoice == str(
                            excel_row[invoice_col]
                        ).strip()
                    ):

                        matched_row = excel_row
                        break

                # =========================================
                # VENDOR MISMATCH
                # =========================================

                if matched_row is None:

                    return pd.Series([

                        "NOT FOUND",
                        "NOT FOUND",
                        "NOT FOUND",
                        "VENDOR CODE MISMATCH"

                    ])

                # =========================================
                # EXCEL VALUES
                # =========================================

                excel_amount = matched_row[
                    amount_col
                ]

                vendor_code = str(
                    matched_row[vendor_col]
                ).strip()

                # =========================================
                # AMOUNT CHECK
                # =========================================

                amount_found = False

                matched_amount = "NOT FOUND"

                pdf_amount_text = (
                    pdf_text.replace(",", "")
                )

                for diff in range(-2, 3):

                    try:

                        check_amount = round(
                            excel_amount + diff,
                            2
                        )

                        amount_formats = [

                            f"{check_amount:,.2f}",
                            f"{check_amount:.2f}",
                            str(int(check_amount)),
                            str(check_amount)

                        ]

                        for amt in amount_formats:

                            amt_clean = amt.replace(
                                ",",
                                ""
                            )

                            if amt_clean in pdf_amount_text:

                                amount_found = True
                                matched_amount = amt
                                break

                        if amount_found:

                            break

                    except:

                        pass

                # =========================================
                # FINAL STATUS
                # =========================================

                if amount_found:

                    status = "FULL MATCHED"

                else:

                    status = "AMOUNT MISMATCH"

                # =========================================
                # RETURN EXACTLY 4 VALUES
                # =========================================

                return pd.Series([

                    vendor_code,
                    excel_amount,
                    matched_amount,
                    status

                ])

            # =========================================
            # APPLY VALIDATION
            # =========================================

            report_df[[
                "Vendor_Code",
                "Excel_Amount",
                "Matched_Amount_In_PDF",
                "Final_Status"
            ]] = report_df.apply(
                validate_invoice,
                axis=1
            )

            # =========================================
            # DUPLICATE CHECK
            # =========================================

            duplicate_df = report_df[

                report_df.duplicated(

                    subset=[
                        "Vendor_Code",
                        "Invoice_Number"
                    ],

                    keep=False

                )

            ]

            duplicate_df = duplicate_df[

                (
                    duplicate_df["Vendor_Code"]
                    != "NOT FOUND"
                )

                &

                (
                    duplicate_df["Invoice_Number"]
                    != "NOT FOUND"
                )

            ]

            duplicate_df = duplicate_df.sort_values(

                by=[
                    "Vendor_Code",
                    "Invoice_Number"
                ]

            )

            # =========================================
            # MISMATCH REPORT
            # =========================================

            mismatch_df = report_df[

                report_df[
                    "Final_Status"
                ] == "AMOUNT MISMATCH"

            ]

            # =========================================
            # SUMMARY
            # =========================================

            total_files = len(report_df)

            matched_count = len(

                report_df[
                    report_df[
                        "Final_Status"
                    ] == "FULL MATCHED"
                ]

            )

            mismatch_count = len(
                mismatch_df
            )

            duplicate_count = len(
                duplicate_df
            )

            # =========================================
            # DASHBOARD
            # =========================================

            st.subheader(
                "Validation Summary"
            )

            col1, col2, col3, col4 = st.columns(4)

            col1.metric(
                "Total Files",
                total_files
            )

            col2.metric(
                "Matched",
                matched_count
            )

            col3.metric(
                "Mismatched",
                mismatch_count
            )

            col4.metric(
                "Duplicates",
                duplicate_count
            )

            # =========================================
            # SHOW REPORT
            # =========================================

            st.subheader(
                "Final Validation Report"
            )

            st.dataframe(report_df)

            # =========================================
            # SAVE EXCEL REPORTS
            # =========================================

            final_report_path = os.path.join(
                temp_dir,
                "Final_Validated_Report.xlsx"
            )

            duplicate_report_path = os.path.join(
                temp_dir,
                "Duplicate_Report.xlsx"
            )

            mismatch_report_path = os.path.join(
                temp_dir,
                "Mismatch_Report.xlsx"
            )

            report_df.to_excel(
                final_report_path,
                index=False
            )

            duplicate_df.to_excel(
                duplicate_report_path,
                index=False
            )

            mismatch_df.to_excel(
                mismatch_report_path,
                index=False
            )

            # =========================================
            # DOWNLOAD BUTTONS
            # =========================================

            final_file = open(
                final_report_path,
                "rb"
            ).read()

            duplicate_file = open(
                duplicate_report_path,
                "rb"
            ).read()

            mismatch_file = open(
                mismatch_report_path,
                "rb"
            ).read()

            st.download_button(
                label="Download Final Report",
                data=final_file,
                file_name="Final_Validated_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.download_button(
                label="Download Duplicate Report",
                data=duplicate_file,
                file_name="Duplicate_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.download_button(
                label="Download Mismatch Report",
                data=mismatch_file,
                file_name="Mismatch_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
