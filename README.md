# PDF Comment & Annotation Extractor

Extract PDF comments/annotations and export to CSV, XLSX, or JSON.

## Install

```bash
pip install -r requirements.txt
```

## Usage

```bash
python extractor.py <file1.pdf> [file2.pdf ...] [options]
python extractor.py --folder ./pdfs/ [options]
```

Options:

- --output PATH Output file path (default: annotations_output.xlsx)
- --format csv Force CSV output regardless of file extension
- --format xlsx Force XLSX output regardless of file extension
- --format json Force JSON output regardless of file extension
- --no-color Disable color extraction (faster)
- --no-text Skip highlighted text extraction (faster)
- --sheet-per-file (XLSX only) Put each PDF's annotations on a separate sheet
- --verbose Print progress and per-file annotation counts
- --skip-empty Skip PDFs with no annotations
- --overwrite Overwrite output file instead of appending

Examples:

```bash
python extractor.py report.pdf --output comments.xlsx
python extractor.py --folder ./docs/ --output all_comments.csv --verbose
python extractor.py doc1.pdf doc2.pdf --output review.xlsx --sheet-per-file
```

## Web UI

```bash
python web_ui.py
```

Then open the local URL printed in the terminal to upload PDFs and download the results.

## Output Columns

- source_file: Name of the source PDF file
- page_number: 1-indexed page number
- annotation_type: Human-readable type
- author: Author/creator of the annotation
- date_created: Creation date in ISO 8601 (if present)
- date_modified: Modified date in ISO 8601 (if present)
- subject: Subject field (if present)
- contents: Annotation text body
- highlighted_text: Extracted text under highlights/underline/strikeout
- color: Annotation color as hex
- rect: Bounding box [x0, y0, x1, y1]
- flags: Raw annotation flags integer
- popup_open: Boolean; popup open when saved

## Notes

- Encrypted PDFs are skipped with a warning.
- By default, existing output files are appended to. Use --overwrite to replace.
- XLSX output includes formatted headers, alternating row fills, and color-coded annotation_type cells.
