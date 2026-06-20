<div align="center">
  <h1>IDX Ownership Data Pipeline</h1>
  <p>
    Scrape Indonesian Stock Exchange disclosure PDFs, extract complex 5%+ ownership tables, and export clean Excel workbooks.
  </p>

  <img src="https://img.shields.io/badge/Python-Data%20Pipeline-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python Data Pipeline" />
  <img src="https://img.shields.io/badge/IDX-Disclosure%20Scraper-0A66C2?style=for-the-badge" alt="IDX Disclosure Scraper" />
  <img src="https://img.shields.io/badge/PDF-Table%20Extraction-EA4335?style=for-the-badge&logo=adobeacrobatreader&logoColor=white" alt="PDF Table Extraction" />
  <img src="https://img.shields.io/badge/Excel-Export-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white" alt="Excel Export" />
</div>

---

## About the Project

**IDX Ownership Data Pipeline** is a Python project for collecting and processing Indonesian Stock Exchange disclosure documents.

The pipeline searches the IDX **Keterbukaan Informasi** page, downloads relevant PDF attachments, extracts **KSEI/BEI “Kepemilikan Efek di Atas 5%”** ownership tables, and saves the results as structured Excel files.

It is designed to handle difficult PDF table layouts where normal text extraction often fails, including wrapped address fields, inconsistent spacing, duplicated glyphs, and multi-column ownership records.

## What It Does

```text
Search IDX disclosures    Download PDF attachments    Extract 5%+ ownership tables
Clean parsed records      Validate numeric fields      Export structured Excel files
```

## Main Features

- Scrapes IDX disclosure results using a keyword such as `5%`
- Downloads valid PDF attachments from IDX
- Extracts KSEI/BEI ownership tables into `.xlsx`
- Handles multi-column and irregular PDF table layouts
- Uses positional PDF parsing with `pdfplumber`
- Adds parse warnings and confidence scores
- Supports command-line usage
- Includes a Streamlit GUI for easier operation
- Can merge multiple extracted Excel files into one master workbook

## Tech Stack

### Core Language

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)

### Web Scraping and Automation

![Playwright](https://img.shields.io/badge/Playwright-2EAD33?style=for-the-badge&logo=playwright&logoColor=white)
![Requests](https://img.shields.io/badge/Requests-20232A?style=for-the-badge)

### PDF and Data Processing

![pdfplumber](https://img.shields.io/badge/pdfplumber-000000?style=for-the-badge)
![Pandas](https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white)
![OpenPyXL](https://img.shields.io/badge/OpenPyXL-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)

### Interface

![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)

### Utilities

![tqdm](https://img.shields.io/badge/tqdm-FFC107?style=for-the-badge)
![CLI](https://img.shields.io/badge/CLI-000000?style=for-the-badge&logo=gnubash&logoColor=white)

## Project Structure

```text
idx-ownership-data-pipeline/
│
├── scrape_and_download.py
├── extract_ownership_table.py
├── run_gui.py
├── requirements.txt
├── requirements-dev.txt
├── examples/
│   └── sample_ownership_output.xlsx
├── tests/
│   ├── fixtures/
│   │   ├── ownership_row_spaced.txt
│   │   └── ownership_row_standard.txt
│   └── test_text_parsing.py
├── README.md
│
└── outputs/
    ├── pdfs/
    └── extracted/
```

## File Overview

| File | Purpose |
|---|---|
| `scrape_and_download.py` | Searches IDX disclosures, collects PDF links, downloads PDFs, and can optionally run extraction. |
| `extract_ownership_table.py` | Extracts KSEI/BEI ownership records from a PDF into a clean Excel file. |
| `run_gui.py` | Streamlit interface for running scrape, extraction, merge, cleanup, and download steps. |
| `requirements.txt` | Python package dependencies. |
| `requirements-dev.txt` | Runtime dependencies plus pytest. |
| `tests/` | Focused tests and sanitized parser fixtures. |
| `examples/sample_ownership_output.xlsx` | Fictional sample of the generated workbook schema. |
| `outputs/pdfs/` | Default folder for downloaded PDFs. |
| `outputs/extracted/` | Default folder for extracted Excel outputs. |

## Installation

Clone the repository:

```bash
git clone https://github.com/INo-xious/idx-ownership-data-pipeline.git
cd idx-ownership-data-pipeline
```

Create and activate a virtual environment:

```bash
python -m venv .venv
```

On macOS/Linux:

```bash
source .venv/bin/activate
```

On Windows:

```bash
.venv\Scripts\activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Install Playwright Chromium:

```bash
python -m playwright install chromium
```

If you want to use the GUI, install Streamlit:

```bash
pip install streamlit
```

## Command-Line Usage

### Scrape and Download PDFs

```bash
python scrape_and_download.py
```

By default, the scraper searches IDX using:

```text
5%
```

PDF files are saved to:

```text
outputs/pdfs/
```

### Limit the Number of PDFs

```bash
python scrape_and_download.py --max-pdfs 5
```

### Use a Custom Keyword

```bash
python scrape_and_download.py --keyword "5%"
```

### Scrape and Extract Automatically

```bash
python scrape_and_download.py --extract
```

This downloads PDFs into:

```text
outputs/pdfs/
```

And saves extracted Excel files into:

```text
outputs/extracted/
```

### Overwrite Existing Files

```bash
python scrape_and_download.py --extract --overwrite
```

## Extract a Single PDF

To extract one PDF manually:

```bash
python extract_ownership_table.py \
  --pdf "outputs/pdfs/example.pdf" \
  --out "outputs/extracted/example.ownership_table.xlsx"
```

## Process Specific Pages Only

```bash
python extract_ownership_table.py \
  --pdf "outputs/pdfs/example.pdf" \
  --out "outputs/extracted/example.ownership_table.xlsx" \
  --page-from 2 \
  --page-to 10
```

## Include Raw Debug Rows

```bash
python extract_ownership_table.py \
  --pdf "outputs/pdfs/example.pdf" \
  --out "outputs/extracted/example.ownership_table.xlsx" \
  --include-raw-debug
```

## Write Debug Artifacts

```bash
python extract_ownership_table.py \
  --pdf "outputs/pdfs/example.pdf" \
  --out "outputs/extracted/example.ownership_table.xlsx" \
  --debug-dir "outputs/debug/example"
```

This can create debug files such as:

```text
outputs/debug/example/words/words_p001.csv
```

These files help inspect PDF word positions when table extraction needs checking.

## GUI Usage

Start the Streamlit app:

```bash
streamlit run run_gui.py
```

The GUI allows you to:

- set the IDX search keyword,
- limit the number of PDFs,
- run the full scraping pipeline,
- extract all downloaded PDFs,
- merge extracted files into one master Excel workbook,
- clean up intermediate files,
- open the output folder,
- download the final Excel file.

The default merged output is:

```text
outputs/extracted/ownership_table.xlsx
```

## Output Columns

The extracted Excel file contains structured ownership records with columns such as:

| Column | Description |
|---|---|
| `source_file` | Source PDF filename. |
| `page` | Page number where the row was extracted. |
| `row_no` | Original table row number. |
| `ticker` | Stock ticker / kode efek. |
| `emiten` | Listed company name. |
| `broker` | Broker or participant field. |
| `shareholder` | Shareholder name. |
| `account_name` | Account name. |
| `address` | Shareholder address. |
| `nationality` | Shareholder nationality. |
| `domicile` | Shareholder domicile. |
| `status` | Ownership status, usually `L` or `A`. |
| `shares_after` | Number of shares after the reporting change. |
| `combined_after` | Combined shares after the reporting change. |
| `pct_after` | Ownership percentage after the reporting change. |
| `shares_before` | Number of shares before the reporting change. |
| `combined_before` | Combined shares before the reporting change. |
| `pct_before` | Ownership percentage before the reporting change. |
| `delta_shares_or_delta_field` | Change or delta field extracted from the table. |
| `raw_text` | Raw row text used during parsing. |
| `parse_warnings` | Parser warnings for missing or unusual fields. |
| `confidence` | Basic extraction confidence score. |

## Example Workflow

Run the full pipeline from the command line:

```bash
python scrape_and_download.py --keyword "5%" --extract
```

Or use the GUI:

```bash
streamlit run run_gui.py
```

Then check the output folder:

```text
outputs/extracted/
```

## Tests and Sanitized Samples

Install the development dependencies and run the parser tests:

```bash
pip install -r requirements-dev.txt
pytest -q
```

The fixtures in `tests/fixtures/` contain fictional, sanitized ownership rows.
They cover both normal text extraction and glyph-by-glyph PDF text without
publishing real shareholder information. A representative workbook is available
at `examples/sample_ownership_output.xlsx`.

## Suggested `requirements.txt`

The GUI uses Streamlit, so the dependency list should include it.

```text
playwright==1.41.2
requests>=2.31.0
tqdm>=4.66.0
pdfplumber>=0.10.3
pandas>=2.1.0
openpyxl>=3.1.2
streamlit>=1.30.0
```

## Notes and Limitations

- The scraper depends on the current IDX website structure.
- If IDX changes its layout, the scraper selectors may need updates.
- The extractor is designed for KSEI/BEI 5%+ ownership table PDFs.
- Scanned PDFs or heavily distorted files may not extract correctly.
- Some PDF links may be skipped if they are not valid PDF files.
- Extracted results should be reviewed before being used for important analysis.

## Future Improvements

- Expand tests with additional sanitized table layouts
- Add configuration file support
- Add better error summaries after batch extraction
- Add duplicate detection for downloaded PDFs
- Add a cleaner merged output report
- Add packaged executable or Docker support

---

<div align="center">
  <b>Automating IDX disclosure collection, PDF reconstruction, and ownership data extraction.</b>
</div>
