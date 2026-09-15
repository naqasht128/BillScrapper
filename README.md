# PITC Bill Scrapper & MDI Analyzer

Desktop UI application to scrape PITC (PESCO) electricity bills from account/reference numbers in Excel files and perform automated MDI (Maximum Demand Indicator) difference and fixed charges analysis.

---

## ✨ Features & Column Details

1. **`prv_cum_mdi_1` and `prv_cum_mdi_2`**:
   - Fetches the **PRV CUMM MDI** values from the bill (located directly before `PRS CUMM MDI`).
   - Positioned immediately before `cum_mdi_1` and `cum_mdi_2`.

2. **`cum_mdi_1` and `cum_mdi_2`**:
   - Present cumulative MDI readings (`PRS CUMM MDI`).

3. **`Diff_cum_mdi1` and `diff_cum_mdi2`**:
   - The actual numeric difference between cumulative registers:
     - **Normal status**:
       - `Diff_cum_mdi1 = |cum_mdi_1 - prv_cum_mdi_1|`
       - `diff_cum_mdi2 = |cum_mdi_2 - prv_cum_mdi_2|`
     - **`Replaced` status**:
       - `Diff_cum_mdi1 = |cum_mdi_2 - prv_cum_mdi_2|`
       - `diff_cum_mdi2 = |cum_mdi_4 - prv_cum_mdi_4|`
     - **`P.Disc` or `T.Disc` status**:
       - Subtraction is **skipped** (left blank), as cumulative counter differences do not reflect billing demand for disconnected connections.

4. **`Actual MDI`**:
   - `max(Diff_cum_mdi1, diff_cum_mdi2) * mdi_mf_1`
   - Evaluates the maximum recorded MDI difference and multiplies it by the MDI Multiplying Factor (`mdi_mf_1`).
   - Skipped for `P.Disc` and `T.Disc`.

5. **`Actual_Fix_Charges`**:
   - `Actual MDI * 1250`
   - Calculates the actual fixed charges from `Actual MDI` at the rate of Rs. 1,250 per kW.
   - Skipped for `P.Disc` and `T.Disc`.

6. **`In_Bill_Fix_Charges (Estimated)`**:
   - `mdi_units_1 * 1250`
   - Calculates the estimated fixed charges billed on the invoice using `mdi_units_1` at Rs. 1,250 per kW for instant audit comparison.

7. **`payment` and `payment_date`**:
   - Captures the bill's **Amount Paid** and **Payment Date** when payment information is available.
   - Positioned immediately after `grand_total`; unpaid bills remain blank in these columns.

8. **Automatic Unit Separation**:
   - Automatically separates Off-Peak `(O)` and Peak `(P)` or space-separated units in `kwh_units`, `kvarh_units`, and `mdi_units` into discrete columns (`_units_1`, `_units_2`).

9. **Status Preservation & Clean Alignment**:
   - The consumer's connection/meter status is captured in the `status` column.
   - Excel rows are mapped strictly by column names to prevent any header shifting or data misalignment.

---

## 🚀 How to Run

### Method 1: Double Click `run.bat`
Simply double-click the `run.bat` file in `D:\Projects\MDIScrapper` to launch the application.

### Method 2: Command Line
Open PowerShell or Command Prompt in the project folder and execute:
```bash
cd D:\Projects\MDIScrapper
python main.py
```

---

## 📋 Input File Format

- Prepare an Excel workbook (`.xlsx`).
- Put the 14-digit Account / Reference numbers in **Column A** (Row 1 can be `AC No.` or account numbers directly).
- Optional: If you have existing statuses in a `status` column, they will be recognized.
- Open the application, click **Browse...**, select your Excel file, and click **Start Scraping**.
- The file is updated in place with all columns added and formatted.

---

## 📦 Dependencies

- Python 3.8+
- `requests`
- `beautifulsoup4`
- `lxml`
- `openpyxl`

To install:
```bash
pip install -r requirements.txt
pip install pyinstaller
```

---

## 🛠️ How to Build `.exe` (PyInstaller)

Run the following command in your VS Code terminal / PowerShell in `D:\Projects\MDIScrapper`:

```powershell
pyinstaller --noconsole --onefile --name "MDIScrapper" --icon "assets\mdiscrapper.ico" --add-data "assets\mdiscrapper.ico;assets" --clean main.py
```

The standalone executable `MDIScrapper.exe` will be generated in the `D:\Projects\MDIScrapper\dist\` folder.
