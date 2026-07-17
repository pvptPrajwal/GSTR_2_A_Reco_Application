## GSTR-2A Reconciliation Tool 
# User Guide & Reference Document 

/# What is this tool? 
This tool automates the GSTR-2A vs Purchase Book reconciliation process. 
It compares invoices filed by your suppliers on the GST Portal (GSTR-2A) 
with your internal Purchase Book — and tells you exactly what matches

what is missing, and what has value differences. 
You select your Excel files, click Run — and get a ready report in seconds. 

1.	How to Run the Tool
   Follow these steps every time you want to run a reconciliation:
  	Step 1 — Open the Application
  	•	gstr2a_app.exe on your computer.
  	•	The main window will open. No installation needed.
  	Step 2 — Select GST Portal File(s)
  	•	Browse Files next to 'GST Portal Files'.
  	•	You can select one file or multiple files at the same time.
  	•	Example: Select April_GST.xlsx, May_GST.xlsx, June_GST.xlsx together.
  	•	The tool will automatically read all sheets inside each file (B2B, CDNR, etc.).
  	Step 3 — Select Purchase Book File(s)
  	•	Browse Files next to 'Purchase Book Files'.
  	•	Again, you can select one or multiple files.
  	•	Example: Select April_PB.xlsx, May_PB.xlsx, June_PB.xlsx together.
  	Step 4 — Select Output Folder
  	•	Browse next to 'Output Folder'.
  	•	Choose the folder where you want the report to be saved.
  	•	Default is your Desktop.
  	Step 5 — Click Run
  	•	▶  Run Reconciliation button.
  	•	Watch the Activity Log on the right side — it shows live progress.
  	•	The progress bar will fill up as each step completes.
  	•	When done, a popup will appear showing the summary counts.
  	•	Final_Reco_Report.xlsx in the selected folder.
  	 
<img width="1056" height="1" alt="image" src="https://github.com/user-attachments/assets/284e6a55-7e30-4ccd-aac6-669cf4bd9efd" />
 
1.	About the Input Files
   1.1	GST Portal File (GSTR-2A)
  	  This is the Excel file you download from the GST Portal. It contains invoices that your suppliers have filed in their GSTR-1. 
        •	GST Portal → Return Dashboard → GSTR-2A → Download Excel
  	    •	The file has multiple sheets: B2B, B2BUR, CDNR, CDNUR, IMPG, IMPS, etc.
  	    •	The tool automatically reads ALL sheets — you do not need to manually open or edit the file. 
        •	Required columns the tool looks for: GSTIN, Invoice No, Taxable Value, IGST, CGST, SGST. 
        •	Optional but useful: Trade/Legal Name (for vendor name in report), ITC Eligibility. 
    Supported Sheet Types
  	B2B      — Regular invoices from registered suppliers 
    B2BUR    — Invoices from unregistered suppliers 
    CDNR     — Credit/Debit notes from registered suppliers 
    CDNUR    — Credit/Debit notes from unregistered suppliers 
    IMPG     — Import of goods (Bill of Entry) 
    IMPS     — Import of services 
    Any other sheet — auto-detected if it has invoice data
  	 
  1.2	Purchase Book File 
   This is your internal record of purchase invoices — usually exported from Tally, SAP, or any accounting software. 
     •	Only the first sheet is read. 
     •	Required columns: Invoice Number, GSTIN, Taxable Value, IGST/CGST/SGST.
     •	Optional: Party Name (shown as Vendor Name in report), Invoice Date. 
     •	Column names can vary — the tool automatically recognises common variants. 
   Accepted Column Name Variants 
     GSTIN          : 'gstin', 'Supplier GSTIN', 'Party GSTIN', 'Vendor GSTIN' 
     Invoice Number : 'Invoice No', 'Invoice No.', 'Bill No', 'Document No' 
     Taxable Value  : 'Taxable Value', 'Taxable Amount', 'Assessable Value' 
     IGST           : 'IGST', 'IGST Paid', 'IGST Amount', 'Integrated Tax' 
     CGST           : 'CGST', 'CGST Paid', 'CGST Amount', 'Central Tax' 
     SGST           : 'SGST', 'SGST Paid', 'SGST/UT Paid', 'State Tax' 
     Vendor Name    : 'Party Name', 'Supplier Name', 'Trade Name' 

2.	Output Report — Sheet by Sheet
   The output file is named Final_Reco_Report.xlsx and contains these sheets:
  	<img width="1056" height="1" alt="image" src="https://github.com/user-attachments/assets/b4605879-5d42-435c-ae82-c4758ad91dd9" />

    <img width="748" height="178" alt="image" src="https://github.com/user-attachments/assets/63537e7e-766f-4d78-bf3b-9e4c07b52131" />
    <img width="744" height="456" alt="image" src="https://github.com/user-attachments/assets/59f2d509-1f92-4518-ae7e-abbc070a7d37" />

<img width="1056" height="1" alt="image" src="https://github.com/user-attachments/assets/6f3b41a8-abab-4ce4-a422-3149f49a86a7" />

1.	How the Tool Works (Simple Explanation)
   1.3	Reading Files
  	•	All GST Portal files are stacked together into one list.
  	•	All Purchase Book files are stacked together into one list.
  	•	Column names are automatically detected and mapped to standard names.
  	•	Blank rows, NaN values, and formatting issues are cleaned automatically. 
   1.1	Invoice Aggregation
  	One invoice can have multiple rows in the file — for example, same bill has 5% GST on one item and 18% GST on another item.   
    •	The tool groups by GSTIN + Invoice Number and sums all values.
  	•	Result: one row per invoice with total Taxable Value, IGST, CGST, SGST.
  	•	This is done for BOTH the GST file and the Purchase Book.
  	<img width="1056" height="1" alt="image" src="https://github.com/user-attachments/assets/68e7595b-bc9d-4189-98d0-81777b1652cf" />
   Example: Aggregation
  	INV001  |  27AAFFK5555F1ZH  |  ₹5,000  |  CGST ₹250  |  SGST ₹250   (row 1)
  	INV001  |  27AAFFK5555F1ZH  |  ₹3,000  |  CGST ₹540  |  SGST ₹540   (row 2)
  	                                    ↓  combined into one row:
  	INV001  |  27AAFFK5555F1ZH  |  ₹8,000  |  CGST ₹790  |  SGST ₹790
   1.1	Matching Logic
  	•	Match is done on GSTIN + Invoice Number together.
  	•	Invoice numbers are cleaned before matching — all spaces and special characters (/, -, etc.) are removed and converted to uppercase.
  	•	So INV/001 and INV-001 and INV 001 all become INV001 and match correctly. 
   1.1	Mismatch Rule
  	•	If GSTIN + Invoice Number matches in both files BUT total tax difference is more than ₹1 — it goes to Mismatched sheet.
  	•	The ₹1 tolerance handles minor rounding differences between systems.
  	•	In the Mismatched sheet, the specific cell that differs is highlighted in red. 
   1.1	GSTIN Validation
  	•	All GSTINs are checked against the standard 15-character format.
  	•	Invalid GSTINs are flagged as warnings in the Activity Log and Summary sheet.
  	•	Processing still continues — invalid GSTINs are not removed. 
   <img width="1056" height="1" alt="image" src="https://github.com/user-attachments/assets/562579b9-292c-4073-8516-c09856927cbf" />
<img width="1056" height="1" alt="image" src="https://github.com/user-attachments/assets/e9ac241b-2631-4998-9e88-4bb4436281f2" />

1.	Output Column Reference
   Standard Columns (all sheets)
  	<img width="737" height="329" alt="image" src="https://github.com/user-attachments/assets/5ad372ca-01b1-4398-8206-5f9cdeff8254" />
   Mismatched Sheet Extra Columns
  	•	— Value from Purchase Book and GST Portal respectively
  	•	— IGST from each side
  	•	— CGST from each side
  	•	— SGST from each side
  	•	— Difference (PB minus GST)
  	•	— Total absolute tax difference. If this is more than ₹1, invoice is Mismatched. 
<img width="1056" height="1" alt="image" src="https://github.com/user-attachments/assets/46cb1583-5321-4316-ba3a-5a390f6ec56a" />

   Vendor Summary Sheet Columns
  	•	— Total Taxable Value from PB and GST Portal
  	•	— Total IGST from each side
  	•	— Total CGST from each side
  	•	— Total SGST from each side
  	•	— Difference (PB minus GST). Red = gap exists. 
<img width="1056" height="1" alt="image" src="https://github.com/user-attachments/assets/dd4d97d4-98ba-4464-b144-26c4791508ac" />
      
1.	Common Questions 
  Q: Why does the same invoice appear in both Missing sheets? 
  This means the supplier has filed SOME invoices but not ALL. For example — they filed INV001 and INV002 on the portal, but INV003 and INV004 are missing. Meanwhile your Purchase Book has INV003 and INV004 but not INV001. This is shown in the Cross_Match sheet. 
 
  Q: What does ITC At Risk mean? 
  ITC (Input Tax Credit) is the tax you can claim back. If an invoice is Missing in GST Portal or Mismatched — you may not be able to claim that tax back. The Summary sheet shows the total tax amount at risk so you know the financial impact. 
 
  Q: The tool says 'no usable data' for a sheet — why? 
  That sheet does not have the required columns (Invoice Number + at least one tax column). The tool skips it automatically. This is normal for sheets like 'Home', 'Summary', 'Cover Page', etc. 
 
  Q: Can I select files from different months at once? 
  Yes. You can select January, February, March files together. All data is combined and reconciled as one report. If you want month-wise separation, run the tool separately for each month. 
 
  Q: What if my column names are different from the list? 
  The tool has a built-in dictionary of 30+ column name variants. If your column name is not recognised, the tool will fill it with zero and show a warning. In that case, rename your column in the Excel file to match one of the standard names listed in Section 2. 
 
 <img width="1056" height="1" alt="image" src="https://github.com/user-attachments/assets/7e194404-cc42-48d0-990c-7a4940eb6c35" />

1.	Tips for Best Results

  	•	Always download a fresh GSTR-2A file from the portal before reconciliation.
  	•	Make sure your Purchase Book includes the GSTIN column — without it, matching will fail.
  	•	If a vendor has no GSTIN (unregistered), leave the GSTIN column blank — the tool handles it.
  	•	Run the reconciliation month by month for better tracking rather than the full year at once.
  	•	After reconciliation, follow up on Missing_in_GST_Portal invoices before filing GSTR-3B to avoid losing ITC.
  	•	The Cross_Match sheet is the most important for follow-up — these vendors need immediate attention.
  	 
  Quick Reference — What to Do with Each Sheet 
    Summary                  →  Review totals, check ITC At Risk 
    Matched                  →  No action needed 
    Mismatched               →  Call supplier to correct filing or fix your PB 
    Missing_in_GST_Portal    →  Follow up with supplier to file GSTR-1 
    Missing_in_Purchase_Book →  Check if invoice was missed in your books 
    Cross_Match              →  Priority follow-up — supplier has partial filing 
    Vendor_Summary           →  Sort by Diff_TV to find biggest gaps 

