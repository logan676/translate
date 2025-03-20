#!/bin/bash
# split_docx.sh
# This script converts a DOCX file to PDF, splits the PDF into single-page PDFs,
# and converts each page to a DOCX file using the pdf2docx command-line tool.
#
# Requirements:
#   - LibreOffice (for DOCX -> PDF conversion)
#   - poppler-utils (for pdfseparate)
#   - pdf2docx (for PDF -> DOCX conversion; install via pip3)
#
# Usage: ./split_docx.sh input.docx

# Check if input file is provided
if [ "$#" -ne 1 ]; then
    echo "Usage: $0 <input_docx>"
    exit 1
fi

INPUT_DOCX="$1"

# Verify that the file exists
if [ ! -f "$INPUT_DOCX" ]; then
    echo "Error: File '$INPUT_DOCX' not found."
    exit 1
fi

echo "Converting DOCX to PDF..."
# Convert DOCX to PDF using LibreOffice in headless mode
libreoffice --headless --convert-to pdf "$INPUT_DOCX"
if [ $? -ne 0 ]; then
    echo "Error converting DOCX to PDF."
    exit 1
fi

# The output PDF will have the same name as the DOCX but with a .pdf extension.
PDF_FILE="${INPUT_DOCX%.*}.pdf"

# Check if the PDF was created successfully
if [ ! -f "$PDF_FILE" ]; then
    echo "Error: PDF file '$PDF_FILE' was not created."
    exit 1
fi

# Create an output directory for the split pages
OUTPUT_DIR="pages_output"
mkdir -p "$OUTPUT_DIR"

echo "Splitting PDF into individual pages..."
# Split the PDF into single-page PDFs named page-1.pdf, page-2.pdf, etc.
pdfseparate "$PDF_FILE" "$OUTPUT_DIR/page-%d.pdf"
if [ $? -ne 0 ]; then
    echo "Error splitting PDF."
    exit 1
fi

echo "Converting each PDF page to DOCX..."
# Loop through each single-page PDF and convert it to DOCX using pdf2docx
for page_pdf in "$OUTPUT_DIR"/page-*.pdf; do
    base=$(basename "$page_pdf" .pdf)
    echo "Converting $page_pdf to ${base}.docx..."
    pdf2docx convert "$page_pdf" "$OUTPUT_DIR/${base}.docx"
    if [ $? -ne 0 ]; then
        echo "Error converting $page_pdf to DOCX."
    fi
done

echo "Conversion complete. Check the '$OUTPUT_DIR' directory for individual DOCX files."
