# 📈 Excel VBA Stock Analysis Tool

This project contains a powerful VBA macro for analyzing stock data across multiple worksheets, grouped by ticker.

---

## 💡 What the Macro Does

For each worksheet in the workbook (excluding the 'Overall Summary'):

1. Groups stock data by ticker symbol.
2. Calculates the **quarterly change**: closing price - opening price.
3. Calculates the **percentage change**: (change / opening price) × 100.
4. Sums the **total volume** for each ticker.
5. Writes the results **to the right of the original data** (starting from column I).
6. Identifies and writes the **per-sheet stats**:
   - Greatest % Increase
   - Greatest % Decrease
   - Greatest Total Volume

After looping through all sheets, it creates a new sheet named **Overall Summary**, which highlights:

- 📈 Stock with the **Greatest % Increase** overall
- 📉 Stock with the **Greatest % Decrease** overall
- 📊 Stock with the **Greatest Total Volume** overall

---

## 🛠️ How to Use

1. Open the Excel file with stock data (each quarter or group in a separate sheet).
2. Open the **VBA Editor** (`Alt + F11`).
3. Paste the macro into a **new Module**.
4. Run `AnalyzeStocksWithOverallSummary`.

---

## 📷 Screenshots (To Add)

Include screenshots of results for each sheet showing:

- Calculated columns: `Quarterly Change`, `Percentage Change`, and `Total Volume`
- Per-sheet summaries beneath those columns
- Final `Overall Summary` sheet

You can take screenshots by:
- Pressing `Windows + Shift + S` (Snipping Tool) on Windows
- Using `Command + Shift + 4` on macOS
- Then paste or attach the images in your documentation

---

## 📁 File Structure (Example)

```
📊 Excel_Stock_Workbook.xlsm
├── Sheet A
├── Sheet B
├── Sheet C
├── Overall Summary
└── VBA Module (AnalyzeStocksWithOverallSummary)
```

---

## ✅ Requirements

- Excel with macros enabled (`.xlsm`)
- Data should have the following columns (example structure):
  - Column A: `<Ticker>`
  - Column C: `<Open>`
  - Column F: `<Close>`
  - Column G: `<Volume>`

You can modify the macro easily if your layout differs.
