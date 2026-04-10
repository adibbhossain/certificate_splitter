# Certificate Splitter

A Python CLI tool to split a multi-page certificate PDF into individual PDF files named after recipients.

This project is especially useful for Canva-generated certificates where multiple certificate pages are exported as a single PDF.

## Features

- Split one multi-page PDF into one PDF per page
- Detect recipient names automatically
- Preview detected names before writing files
- Choose a custom output directory
- Manually fix failed detections
- Export a CSV report
- Handle duplicate names safely

## Use Case

You design certificates in Canva, export all pages as a single PDF, and want each certificate saved as a separate PDF named after the recipient.

Example:

- Input: one 90-page PDF
- Output:
  - `Name01.pdf`
  - `Name02.pdf`
  - `Name03.pdf`

## Installation

Clone the repository:

```bash
git clone https://github.com/adibbhossain/certificate-splitter.git
cd certificate-splitter
```

Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Preview detected names

```bash
python certsplit.py certificates.pdf --preview
```

### Write individual PDFs

```bash
python certsplit.py certificates.pdf --write
```

### Write files to a custom folder

```bash
python certsplit.py certificates.pdf --write --output "./output"
```

### Prompt for manual fix when detection fails

```bash
python certsplit.py certificates.pdf --write --prompt-on-fail
```

### Export CSV report

```bash
python certsplit.py certificates.pdf --write --export-csv report.csv
```

### Full example

```bash
python certsplit.py certificates.pdf --write --output "./output" --prompt-on-fail --export-csv report.csv
```

## Arguments

- `input_pdf` : path to the input PDF
- `--preview` : preview detected names without writing files
- `--write` : write individual PDFs
- `--output` : output directory for saved PDFs
- `--prompt-on-fail` : manually enter a name if detection fails
- `--export-csv` : export processing results as CSV
- `--fallback-prefix` : prefix for fallback filenames (default: `page`)

## How it works

This tool uses layout-aware PDF text extraction to detect likely recipient names from certificate pages. It is designed for structured certificate templates, especially Canva-exported PDFs where normal text order may be inconsistent.

## Current limitations

- Works best with text-based PDFs, not scanned image PDFs
- Detection may fail on highly decorative or unusual layouts
- Some templates may require future tuning or OCR support

## Roadmap

- [ ] Add OCR support for scanned/image-only PDFs
- [ ] Add multiple extraction modes
- [ ] Add installable CLI entry point
- [ ] Add tests
- [ ] Add support for image export

## License

MIT
