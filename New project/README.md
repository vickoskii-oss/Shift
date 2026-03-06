# Self-Employed Shift Manager App

This app was generated from your Excel workbook structure:
`Self_Employed_Shift_Manager.xlsx`.

## What it does

- Loads shift records from the workbook (`Shift Records` sheet)
- Lets you edit/add rows in an interactive table
- Calculates:
  - Gross pay
  - Total expenses
  - Net pay
- Shows:
  - Monthly summary
  - Client summary
  - Tax estimate (20% tax, 9% NI, 5% pension)
- Exports an updated `.xlsx` file using the same workbook template

## Run locally

1. Install dependencies:

   ```powershell
   pip install -r requirements.txt
   ```

2. Start the app:

   ```powershell
   streamlit run app.py
   ```

3. In the sidebar:
- Leave the default path if your workbook is at:
  `C:/Users/olugb/Downloads/Self_Employed_Shift_Manager.xlsx`
- Or upload the workbook manually.

## Notes

- The app writes up to 500 shift rows back to the workbook template.
- Client names are also populated into `Client Summary` automatically.
- Summary/tax formulas in the workbook remain intact.
