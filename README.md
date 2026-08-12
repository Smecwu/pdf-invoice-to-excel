# PDF Invoice To Excel

Free Python tool for office staff to convert Chinese VAT e-invoice PDFs into a clean Excel ledger.

Many PDF-to-Excel tools are paid or too complicated for daily clerical work. This project keeps the workflow simple: put invoice PDFs in one folder, provide an Excel template, and run one PowerShell command.

## What It Does

- Reads Chinese electronic VAT invoice PDFs.
- Extracts invoice number, issue date, material code, quantity, amount, tax, total amount, and remark.
- Writes the result into a single-sheet Excel workbook.
- Keeps only `Sheet1` in the output file.
- Skips merged PDF files by default to avoid duplicate imports.
- Skips duplicate invoice numbers by default.

## Install

Install Python 3.10 or later, then install dependencies:

```powershell
pip install -r requirements.txt
```

## Usage

Example:

```powershell
python .\pdf_invoice_to_sheet1_excel.py --month-dir ".\data\2026.05" --template ".\templates\invoice_template.xlsx" --output "2026.05_invoice_ledger.xlsx"
```

Arguments:

- `--month-dir`: folder containing invoice PDFs.
- `--template`: Excel template path. Only the first sheet is used.
- `--output`: output Excel path.
- `--dry-run`: parse PDFs and print summary without writing Excel.
- `--keep-duplicates`: keep duplicate invoice numbers.
- `--include-merged`: include PDFs whose filename contains `merge`.

## Important Privacy Note

Do not upload real invoice PDFs, real Excel ledgers, tax IDs, customer names, or company financial data to a public GitHub repository.

This repository should contain code, documentation, and anonymized examples only.

## Suggested Folder Layout

```text
pdf-invoice-to-excel/
  pdf_invoice_to_sheet1_excel.py
  requirements.txt
  README.md
  templates/
    invoice_template.xlsx
  data/
    2026.05/
      put-pdfs-here.txt
```

## License

MIT License.
