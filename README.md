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

⸻

About the Project

IDX Ownership Data Pipeline is a Python project for collecting and processing Indonesian Stock Exchange disclosure documents.

The pipeline searches the IDX Keterbukaan Informasi page, downloads relevant PDF attachments, extracts KSEI/BEI “Kepemilikan Efek di Atas 5%” ownership tables, and saves the results as structured Excel files.

It is designed to handle difficult PDF table layouts where normal text extraction often fails, including wrapped address fields, inconsistent spacing, duplicated glyphs, and multi-column ownership records.

What It Does

Search IDX disclosures    Download PDF attachments    Extract 5%+ ownership tables
Clean parsed records      Validate numeric fields      Export structured Excel files

Main Features

* Scrapes IDX disclosure results using a keyword such as 5%
* Downloads valid PDF attachments from IDX
* Extracts KSEI/BEI ownership tables into .xlsx
* Handles multi-column and irregular PDF table layouts
* Uses positional PDF parsing with pdfplumber
* Adds parse warnings and confidence scores
* Supports command-line usage
* Includes a Streamlit GUI for easier operation
* Can merge multiple extracted Excel files into one master workbook

Tech Stack

Core Language

Web Scraping and Automation

PDF and Data Processing

Interface

Utilities

Project Structure

idx-ownership-data-pipeline/
│
├── scrape_and_download.py
├── extract_ownership_table.py
├── run_gui.py
├── requirements.txt
├── README.md
│
└── outputs/
    ├── pdfs/
    └── extracted/

File Overview

File	Purpose
scrape_and_download.py	Searches IDX disclosures, collects PDF links, downloads PDFs, and can optionally run extraction.
extract_ownership_table.py	Extracts KSEI/BEI ownership records from a PDF into a clean Excel file.
run_gui.py	Streamlit interface for running scrape, extraction, merge, cleanup, and download steps.
requirements.txt	Python package dependencies.
outputs/pdfs/	Default folder for downloaded PDFs.
outputs/extracted/	Default folder for extracted Excel outputs.

Installation

Clone the repository:

git clone https://github.com/INo-xious/idx-ownership-data-pipeline.git
cd idx-ownership-data-pipeline

Create and activate a virtual environment:

python -m venv .venv

On macOS/Linux:

source .venv/bin/activate

On Windows:

.venv\Scripts\activate

Install dependencies:

pip install -r requirements.txt

Install Playwright Chromium:

python -m playwright install chromium

If you want to use the GUI, install Streamlit:

pip install streamlit

Command-Line Usage

1. Scrape and Download PDFs

python scrape_and_download.py

By default, the scraper searches IDX using:

5%

PDF files are saved to:

outputs/pdfs/

2. Limit the Number of PDFs

Useful for testing:

python scrape_and_download.py --max-pdfs 5

3. Use a Custom Keyword

python scrape_and_download.py --keyword "5%"

4. Scrape and Extract Automatically

python scrape_and_download.py --extract

This downloads PDFs into:

outputs/pdfs/

And saves extracted Excel files into:

outputs/extracted/

5. Overwrite Existing Files

python scrape_and_download.py --extract --overwrite

Extract a Single PDF

To extract one PDF manually:

python extract_ownership_table.py \
  --pdf "outputs/pdfs/example.pdf" \
  --out "outputs/extracted/example.ownership_table.xlsx"

Process Specific Pages Only

python extract_ownership_table.py \
  --pdf "outputs/pdfs/example.pdf" \
  --out "outputs/extracted/example.ownership_table.xlsx" \
  --page-from 2 \
  --page-to 10

Include Raw Debug Rows

python extract_ownership_table.py \
  --pdf "outputs/pdfs/example.pdf" \
  --out "outputs/extracted/example.ownership_table.xlsx" \
  --include-raw-debug

Write Debug Artifacts

python extract_ownership_table.py \
  --pdf "outputs/pdfs/example.pdf" \
  --out "outputs/extracted/example.ownership_table.xlsx" \
  --debug-dir "outputs/debug/example"

This can create debug files such as:

outputs/debug/example/words/words_p001.csv

These files help inspect PDF word positions when table extraction needs checking.

GUI Usage

Start the Streamlit app:

streamlit run run_gui.py

The GUI allows you to:

* set the IDX search keyword,
* limit the number of PDFs,
* run the full scraping pipeline,
* extract all downloaded PDFs,
* merge extracted files into one master Excel workbook,
* clean up intermediate files,
* open the output folder,
* download the final Excel file.

The default merged output is:

outputs/extracted/ownership_table.xlsx

Output Columns

The extracted Excel file contains structured ownership records with columns such as:

Column	Description
source_file	Source PDF filename.
page	Page number where the row was extracted.
row_no	Original table row number.
ticker	Stock ticker / kode efek.
emiten	Listed company name.
broker	Broker or participant field.
shareholder	Shareholder name.
account_name	Account name.
address	Shareholder address.
nationality	Shareholder nationality.
domicile	Shareholder domicile.
status	Ownership status, usually L or A.
shares_after	Number of shares after the reporting change.
combined_after	Combined shares after the reporting change.
pct_after	Ownership percentage after the reporting change.
shares_before	Number of shares before the reporting change.
combined_before	Combined shares before the reporting change.
pct_before	Ownership percentage before the reporting change.
delta_shares_or_delta_field	Change or delta field extracted from the table.
raw_text	Raw row text used during parsing.
parse_warnings	Parser warnings for missing or unusual fields.
confidence	Basic extraction confidence score.

Example Workflow

Run the full pipeline from the command line:

python scrape_and_download.py --keyword "5%" --extract

Or use the GUI:

streamlit run run_gui.py

Then check the output folder:

outputs/extracted/

Suggested requirements.txt

The GUI uses Streamlit, so the dependency list should include it.

playwright==1.41.2
requests>=2.31.0
tqdm>=4.66.0
pdfplumber>=0.10.3
pandas>=2.1.0
openpyxl>=3.1.2
streamlit>=1.30.0

Notes and Limitations

* The scraper depends on the current IDX website structure.
* If IDX changes its layout, the scraper selectors may need updates.
* The extractor is designed for KSEI/BEI 5%+ ownership table PDFs.
* Scanned PDFs or heavily distorted files may not extract correctly.
* Some PDF links may be skipped if they are not valid PDF files.
* Extracted results should be reviewed before being used for important analysis.

Future Improvements

* Add automated tests with sample PDFs
* Add configuration file support
* Add better error summaries after batch extraction
* Add duplicate detection for downloaded PDFs
* Add a cleaner merged output report
* Add packaged executable or Docker support

⸻

<div align="center">
  <b>Automating IDX disclosure collection, PDF reconstruction, and ownership data extraction.</b>
</div>
