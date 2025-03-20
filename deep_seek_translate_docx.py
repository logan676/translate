#!/usr/bin/env python3

import sys
import os
import datetime
import time  # Added for timing
from docx import Document
from docx.oxml.ns import qn

# If you normally use the official OpenAI library, you can do:
#   import openai
#   openai.api_base = DEESEEK_BASE_URL
#   openai.api_key = DEESEEK_API_KEY
# and so on.
# For demonstration, we'll assume you have 'OpenAI' class from 'openai' matching your original usage:
from openai import OpenAI

# ---------------------------------------------------------
# DeepSeek / OpenAI Configuration
# ---------------------------------------------------------
DEESEEK_API_KEY = "sk-8df2d0cbcd594a349762d33de5b9df3f"
DEESEEK_BASE_URL = "https://api.deepseek.com"
MODEL_NAME = "deepseek-reasoner"

# How many pages to include in each output DOCX before moving on
PAGES_PER_SEGMENT = 5
# ---------------------------------------------------------

def init_client():
    """
    Initialize and return the OpenAI (DeepSeek) client.
    """
    return OpenAI(
        api_key=DEESEEK_API_KEY,
        base_url=DEESEEK_BASE_URL
    )

def run_has_page_break(run):
    """
    Checks if the run contains a page break by examining its XML.
    """
    nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    for br in run._element.findall('.//w:br', namespaces=nsmap):
        if br.get(qn('w:type')) == 'page':
            return True
    return False

def assign_page_numbers(doc):
    """
    Assign a page number to each paragraph by detecting page breaks in runs.
    Returns a list of tuples: [(page_number, paragraph), ...] plus the total page count.
    """
    page_num = 1
    paragraphs_with_page = []

    for paragraph in doc.paragraphs:
        paragraphs_with_page.append((page_num, paragraph))
        # If any run has a page break, increment page_num after this paragraph.
        for run in paragraph.runs:
            if run_has_page_break(run):
                page_num += 1
                break

    return paragraphs_with_page, page_num

def group_into_segments(paragraphs_with_page):
    """
    Group paragraphs into segments, each containing up to PAGES_PER_SEGMENT.
    Returns a dict: {segment_index: [(page_num, paragraph), ...]}.
    """
    segments = {}
    for page_num, paragraph in paragraphs_with_page:
        segment_index = ((page_num - 1) // PAGES_PER_SEGMENT) + 1
        if segment_index not in segments:
            segments[segment_index] = []
        segments[segment_index].append((page_num, paragraph))

    return segments

def translate_text(client, text, progress_info=""):
    """
    Translates the given Chinese text into English using DeepSeek / OpenAI.
    Includes a system message for context in mechanical & electrical systems.
    Added debug logs that group request details.
    """
    system_message = (
        "行業背景：我們所從事的機電系統工程，涵蓋的範圍包括：\n"
        "（1）機：空調、給水、排水（含雨排水、污廢排水）、消防、防火填塞；\n"
        "（2）電：電力（含接地與避雷系統等）、弱電、通訊、安全、能源（柴油發電機組的燃料來源）。\n"
        "翻譯時請採用國家工程主管機構的中英文對照版本作為專業名詞的依據。"
    )

    start_time = time.time()
    try:
        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "system", "content": system_message},
                {"role": "user", "content": f"Translate the following Chinese text into English:\n\n{text}"}
            ],
            stream=False
        )
        translated_text = response.choices[0].message.content.strip()
    except Exception as e:
        translated_text = f"[Translation Error: {str(e)}]"
    end_time = time.time()
    time_cost = end_time - start_time

    # Grouped debug log for this translation request
    debug_log = (
        "\n--- Debug Log for Translation Request ---\n"
        f"Request Text: {text}\n"
        f"Response Text: {translated_text}\n"
        f"Time Cost: {time_cost:.2f} seconds\n"
        f"Progress: {progress_info}\n"
        "--- End Debug Log ---\n"
    )
    print(debug_log)
    return translated_text

def process_segment(client, segment_paragraphs, segment_range, input_path):
    """
    Translates paragraphs (line by line) in one segment,
    then saves a DOCX file for that segment.
    """
    seg_doc = Document()
    total_paragraphs = len(segment_paragraphs)

    # For progress
    print(f"  >> Processing Pages {segment_range[0]}–{segment_range[1]} "
          f"({total_paragraphs} paragraphs in this segment)")

    for idx, (page_num, paragraph) in enumerate(segment_paragraphs, start=1):
        raw_text = paragraph.text.strip()

        # Skip empty paragraphs
        if not raw_text:
            continue

        print(f"    - Paragraph {idx}/{total_paragraphs}, Page {page_num}")

        # Split paragraph into lines, then translate each line.
        lines = raw_text.split('\n')
        for line_idx, line in enumerate(lines, start=1):
            line_text = line.strip()
            if not line_text:
                continue

            # Show line-level progress (optional)
            progress_info = f"Paragraph {idx}/{total_paragraphs}, Page {page_num}, Line {line_idx}/{len(lines)}"
            print(f"       * Translating line {line_idx} of {len(lines)} in paragraph {idx}")

            # 1) Original text
            seg_doc.add_paragraph(line_text)

            # 2) Translated text with debug logging
            translated_line = translate_text(client, line_text, progress_info)
            seg_doc.add_paragraph(translated_line)

    # Create an output path with a timestamp so each segment is saved separately
    base, ext = os.path.splitext(input_path)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = f"{base}_translated_{segment_range[0]}-{segment_range[1]}_{timestamp}{ext}"

    try:
        seg_doc.save(output_path)
        print(f"  >> Segment saved to: {output_path}\n")
    except Exception as e:
        print(f"  !! Error saving segment {segment_range[0]}–{segment_range[1]}: {e}")

def main():
    """
    Script workflow:
      1. Read input DOCX file from argv.
      2. Assign page numbers to paragraphs using page breaks.
      3. Group paragraphs into segments of PAGES_PER_SEGMENT pages.
      4. For each segment, translate line-by-line and save a partial DOCX.
      5. This approach yields multiple smaller DOCX files, ensuring partial results
         are always saved if the script is interrupted or fails later.
    """
    if len(sys.argv) != 2:
        print("Usage: python3 translate_docx.py <path_to_docx_file>")
        sys.exit(1)

    input_path = sys.argv[1]
    if not os.path.isfile(input_path):
        print(f"Error: File not found => {input_path}")
        sys.exit(1)

    # Load the DOCX
    try:
        doc = Document(input_path)
    except Exception as e:
        print(f"Error: Could not open the DOCX file. Details: {e}")
        sys.exit(1)

    # Initialize DeepSeek client
    client = init_client()

    # 1) Assign pages
    paragraphs_with_page, final_page_count = assign_page_numbers(doc)
    # 2) Group into segments of PAGES_PER_SEGMENT
    segments = group_into_segments(paragraphs_with_page)

    print(f"\nDocument has ~{final_page_count} pages.")
    print(f"Dividing into segments of {PAGES_PER_SEGMENT} pages each => total segments: {len(segments)}\n")

    # 3) Process each segment
    for seg_index in sorted(segments.keys()):
        start_page = (seg_index - 1) * PAGES_PER_SEGMENT + 1
        end_page = seg_index * PAGES_PER_SEGMENT
        process_segment(client, segments[seg_index], (start_page, end_page), input_path)

    print("All segments processed successfully.")

if __name__ == "__main__":
    main()
