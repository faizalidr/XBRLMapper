# MBRS Excel Mapper (XBRLMapper)

A client-side web application that transfers data from a filled **MBRS** (Malaysian Business Reporting System) source `.xlsm` file into an empty MBRS destination template.

## 🚀 Live App (GitHub Pages)

Visit the hosted app:  
👉 **`https://faizalidr.github.io/XBRLMapper/`**

---

## Features

- **100% In-Browser Execution**: Files never leave your browser; no server or cloud upload required, ensuring complete confidentiality for financial data.
- **Full Structure Preservation**: Supports `.xlsm` macro-enabled workbooks, keeping VBA code and sheets intact.
- **Smart Cell Mapping**: Accurately maps values by sheet name and coordinate across all 32 MBRS financial statements and notes.
- **Shared Strings Management**: Automatically extracts, resolves, and appends unique text items into the destination workbook's shared string table.
- **Live Progress & Detailed Logging**: Real-time feedback with a visual progress bar and log stream during file generation.

---

## How to Use

1. Open **[https://faizalidr.github.io/XBRLMapper/](https://faizalidr.github.io/XBRLMapper/)**.
2. **Upload** your filled source MBRS `.xlsm` file.
3. **Upload** your empty destination MBRS `.xlsm` template.
4. Click **Run Mapper**.
5. Once complete, click **Download Filled File** to get your ready-to-file `.xlsm` document.

---

## Project Structure

```
├── index.html   # Main interface and upload cards
├── style.css    # Responsive dark-theme styling
├── app.js       # Core ZIP extraction, XML parsing, and mapping engine
└── README.md    # Documentation and usage guide
```
