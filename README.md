# Payslip Payment Details Extractor

A Python tool to automatically extract payment details and summary information from PDF payslips of UniSA and Adelaide University.

## Features

- Extracts detailed payment records (work date, hours, rate, amount)
- Extracts summary information (gross pay, tax, net pay, YTD totals)
- Processes multiple PDF files in batch
- Exports data to Excel with two separate sheets
- Provides statistical summaries

## Requirements

- Python 3.8+
- Dependencies (managed via Poetry):
  - pdfplumber
  - pandas
  - openpyxl

## Installation

1. Install Poetry if you haven't already:
```bash
git clone https://github.com/jet-w/unisa-uniadelaide-payslip-extractor.git
```

2. Install dependencies:
```bash
poetry install
```

## Usage

1. Place your PDF payslip files in the `data/` directory

2. Activate the virtual environment and run the script:
```bash
poetry run python extract_payslip.py
```

Or activate the virtual environment first:
```bash
poetry shell
python extract_payslip.py
```

3. The script will:
   - Scan all PDF files in the `data/` directory
   - Extract payment details and summary information
   - Save results to `payslip_details.xlsx`
   - Display statistics and sample data in the console

## Output

The script generates an Excel file (`payslip_details.xlsx`) with two sheets:

### Sheet 1: Payment Details
Contains individual payment records with columns:
- PDF File
- Page
- Pay Period
- Paid Date
- Work Date
- Hours
- Rate
- Amount

### Sheet 2: Summary
Contains pay period summaries with columns:
- PDF File
- Page
- Pay Period
- Paid Date
- Gross Pay
- Tax
- Nett Pay
- YTD Gross Pay
- YTD Tax
- YTD Nett Pay
- Disbursement Amount

## Example Output

```
Found 2 PDF files

Processing: 6-2 Payslip-02 Jun 2023 to 30 Jun 2024.pdf
  - Extracted 123 payment records
  - Extracted 18 pay period summaries

Processing: 6-2 Payslip-01 Jul 2024 to 16 Mar 2025.pdf
  - Extracted 181 payment records
  - Extracted 19 pay period summaries

✓ Data saved to: payslip_details.xlsx
  - Payment Details: 304 records
  - Summary: 37 records

=== Statistics ===
Total Hours Worked: 1605.10
Total Gross Income: $82,289.19
Total Tax: $0.00
Total Net Income: $67,315.19
```

## Project Structure

```
payslip/
├── data/                          # Place PDF files here
├── extract_payslip.py            # Main extraction script
├── payslip_details.xlsx          # Generated output file
├── pyproject.toml                # Poetry configuration
└── README.md                     # This file
```

## How It Works

The extractor uses a modular design with the following components:

1. **Data Models**: `PayPeriodInfo`, `PaymentRecord`, `SummaryRecord`
2. **Parsers**:
   - `PayPeriodParser`: Extracts pay period information
   - `PaymentParser`: Extracts individual payment lines
   - `SummaryParser`: Extracts summary totals
3. **PDF Processor**: Handles PDF file reading and text extraction
4. **Data Exporter**: Converts data to DataFrames and exports to Excel

## Supported Payslip Format

The script is designed to extract data from payslips with the following format:
- Payment section with columns: Hours, Rate, Reference (date), Amount
- Payment type: "CAS OrdPay (incCASloading)"
- Summary section with Gross Pay, Tax, Nett Pay, and YTD totals
- Pay period format: "Pay Period DD MMM YYYY to DD MMM YYYY Paid DD MMM YYYY"

## Troubleshooting

**No data extracted?**
- Ensure PDF files are in the `data/` directory
- Check that PDF files are not password-protected
- Verify the payslip format matches the expected format

**Incorrect data?**
- Check the console output for parsing errors
- Verify the regex patterns match your payslip format
- Review the extracted text format by enabling debug output

## License

This tool is for personal use. Ensure you have permission to process payslip data in your jurisdiction.

## Author

Created for extracting payment details from university casual employment payslips.
