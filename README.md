Retailer Analysis App

A simple, professional Streamlit app for the Franchisee_CSC report workflow.

What it does

- Generate Report tab: upload your own POS Details CSV, Monthly Activations (BCD)
  CSV, and Retailer_DBN_Division_Wise.xlsx — or leave any of them blank to use the
  bundled default sample files in `defaults/`. Press Execute to build the full
  multi-sheet workbook (the same structure as pos_details.xlsx: POS Details,
  Active_POS, Franchisee_CSC_Wise, Monthly_Activations, Daywise_Pivot,
  Retailer_DBN_Division_Wise, Username_Lookup, and one sheet per Franchisee_CSC),
  then download it.
- Username Lookup tab: pick any active POS username from a searchable dropdown
  and instantly see their POS profile, franchisee/division details, total Sims
  Sold this month, and a day-by-day activation chart — computed directly in the
  app, without needing to open the generated Excel file.

How to run

1. Install Python 3.9+ if you don't already have it.
2. In this folder, install the dependencies:
   pip install -r requirements.txt
3. Start the app:
   streamlit run app.py
4. Your browser will open automatically at http://localhost:8501. If it
   doesn't, open that address manually.

Updating the default files

To change what "default" means (e.g. a new month's BCD export), replace the
matching file in the `defaults/` folder, keeping the same filename, or update
the `DEFAULT_POS` / `DEFAULT_BCD` / `DEFAULT_DBN` paths at the top of `app.py`.

Files in this folder

- app.py — the Streamlit UI (two tabs: Generate Report, Username Lookup)
- report_builder.py — all the data-loading and workbook-building logic, reused
  by both tabs (kept separate from the UI so it can also be tested or reused
  on its own)
- defaults/ — the three bundled sample input files
- requirements.txt — Python dependencies
