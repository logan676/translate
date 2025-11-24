# Document Translation Tool

A Python-based document translation pipeline that splits, translates, and merges DOCX files using AI translation services.

## Features

- Split large DOCX files into manageable page-by-page chunks
- Parallel translation processing with configurable concurrency (up to 30 workers)
- Support for multiple translation backends:
  - DeepSeek AI translation
  - Google Translate
  - Plain text translation
- Automatic merging of translated documents
- Batch processing of multiple documents

## Prerequisites

- Python 3.8 or higher
- `libreoffice` (for PDF conversion in splitting process)

## Installation

### 1. Clone or Download the Repository

```bash
cd /path/to/translate
```

### 2. Create and Activate Virtual Environment

```bash
# Create virtual environment
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate  # On macOS/Linux
# OR
venv\Scripts\activate     # On Windows
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

## Usage

### Complete Translation Workflow

#### Step 1: Split Original Document

Split your DOCX file into individual pages for processing:

```bash
./docx_pdf_docx.sh my_document.docx
```

This creates a `pages_output` directory containing individual DOCX files (one per page).

#### Step 2: Translate Split Files

Run batch translation on all split files:

```bash
./batch_translate.py
```

This processes all files in `pages_output` using parallel workers (default: 30 concurrent translations).

#### Step 3: Merge Translated Files

Combine all translated pages into a single document:

```bash
python3 merge_docx.py pages_output
```

Output: `merged_output.docx`

## Project Structure

```
.
├── batch_translate.py      # Parallel batch translation processor
├── docx_pdf_docx.sh        # DOCX splitting script (via PDF conversion)
├── ds_translate.py         # DeepSeek translation implementation
├── ds_plain_text_ts.py     # Plain text translation with DeepSeek
├── google_translate.py     # Google Translate implementation
├── merge_docx.py           # Document merging utility
├── requirements.txt        # Python dependencies
└── README.md              # This file
```

## Configuration

### Adjusting Concurrency

Edit `batch_translate.py` line 28 to change the number of parallel workers:

```python
max_workers = 30  # Adjust based on your API rate limits
```

### Changing Translation Backend

Modify `batch_translate.py` line 11 to use different translation scripts:

```python
# Current: DeepSeek translation
subprocess.run(["python3", "ds_translate.py", file_path], ...)

# Alternative: Google Translate
subprocess.run(["python3", "google_translate.py", file_path], ...)
```

## Notes

- Ensure you have sufficient API credits/quota for your chosen translation service
- Processing time depends on document size, concurrency settings, and API rate limits
- The `pages_output` directory is automatically created during the splitting process
- Intermediate files are preserved for debugging; delete manually if not needed

## Troubleshooting

### "No DOCX files found in pages_output"
Ensure you've run the splitting script first: `./docx_pdf_docx.sh your_file.docx`

### PDF Conversion Fails
Verify LibreOffice is installed: `libreoffice --version`

### Translation Errors
Check your API credentials and rate limits for the translation service being used