#!/usr/bin/env python3

import sys
import os
import datetime
import time
import copy
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx import Document
from docx.oxml.ns import qn

TRANSLATE_FONT_SIZE = Pt(9)  # 9号字
TRANSLATE_ITALIC = True      # 斜体

# If you normally use the official OpenAI library, you can do:
#   import openai
#   openai.api_base = DEESEEK_BASE_URL
#   openai.api_key = DEESEEK_API_KEY
import openai

# ---------------------------------------------------------
# DeepSeek / OpenAI Configuration
# ---------------------------------------------------------
DEESEEK_API_KEY = "sk-8df2d0cbcd594a349762d33de5b9df3f"
DEESEEK_BASE_URL = "https://api.deepseek.com"
MODEL_NAME = "deepseek-reasoner"

# How many pages to include in each output before moving on
PAGES_PER_SEGMENT = 5
# ---------------------------------------------------------

def init_client():
    """
    Initialize and return the OpenAI (DeepSeek) client.
    """
    openai.api_key = DEESEEK_API_KEY
    openai.api_base = DEESEEK_BASE_URL
    return openai

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
        for run in paragraph.runs:
            if run_has_page_break(run):
                page_num += 1
                break

    return paragraphs_with_page, page_num

def group_into_segments(paragraphs_with_page):
    """
    Group paragraphs into segments, each containing up to PAGES_PER_SEGMENT pages.
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
    The system message enforces domain context and instructs the model to return
    only the translated text with no extra explanations or summaries.
    """
    system_message = (
        "你是一位高級譯者，專門負責機電系統領域的中英翻譯。"
        "以下是機電系統工程的行業背景與專業名詞依據：\n"
        "（1）機：空調、給水、排水（含雨排水、污廢排水）、消防、防火填塞；\n"
        "（2）電：電力（含接地與避雷系統等）、弱電、通訊、安全、能源（柴油發電機組的燃料來源）。\n\n"
        "請僅提供譯文，不要提供任何解釋、提示或總結。"
    )

    start_time = time.time()
    try:
        # The user message is simply the text to be translated,
        # with no extra instructions or disclaimers.
        response = client.ChatCompletion.create(
            model=MODEL_NAME,
            messages=[
                {"role": "system", "content": system_message},
                {"role": "user", "content": text}
            ],
            stream=False
        )
        translated_text = response.choices[0].message.content.strip()
    except Exception as e:
        translated_text = f"[Translation Error: {str(e)}]"
    end_time = time.time()
    time_cost = end_time - start_time

    # Print debug log to console (not included in the final output text).
    print(
        "\n--- Debug Log for Translation Request ---\n"
        f"Original Text: {text}\n"
        f"Translated  : {translated_text}\n"
        f"Time Cost   : {time_cost:.2f} seconds\n"
        f"Progress    : {progress_info}\n"
        "--- End Debug Log ---\n"
    )

    return translated_text

def clone_paragraph_style(src_para, dest_para):
    """增强的样式克隆函数，支持多级列表等复杂格式"""
    # 克隆段落格式属性
    dest_para.paragraph_format.left_indent = src_para.paragraph_format.left_indent
    dest_para.paragraph_format.right_indent = src_para.paragraph_format.right_indent
    dest_para.paragraph_format.space_before = src_para.paragraph_format.space_before
    dest_para.paragraph_format.space_after = src_para.paragraph_format.space_after
    
    # 强制创建目标段落的run
    if not dest_para.runs:
        dest_para.add_run("")  # 创建空run占位
    
    # 克隆run级格式（当源段落有内容时）
    if src_para.runs:
        try:
            src_rPr = src_para.runs[0]._element.rPr
            dest_rPr = dest_para.runs[0]._element.get_or_add_rPr()
            dest_rPr.append(copy.deepcopy(src_rPr))
        except Exception as e:
            print(f"Run样式克隆警告: {str(e)}")
    
    # 克隆段落样式（含多级列表）
    try:
        # 标准样式克隆
        dest_para.style = src_para.style
        
        # 处理多级列表编号
        if src_para._element.pPr.numPr is not None:
            dest_pPr = dest_para._element.get_or_add_pPr()
            dest_numPr = copy.deepcopy(src_para._element.pPr.numPr)
            dest_pPr.append(dest_numPr)
            
        # 克隆对齐方式
        if src_para.paragraph_format.alignment:
            dest_para.paragraph_format.alignment = src_para.paragraph_format.alignment
            
    except Exception as e:
        print(f"段落样式克隆异常: {str(e)}，使用安全模式")
        dest_para.style = 'Normal'
    
    # 强制中文字体兼容
    dest_para.runs[0].font.name = '微软雅黑'
    dest_para.runs[0]._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')

def process_tables(doc, client):
    for table in doc.tables:
        # 在表格底部添加翻译行
        new_row = table.add_row()
        
        for idx, orig_cell in enumerate(table.rows[-2].cells):  # 假设原文在倒数第二行
            new_cell = new_row.cells[idx]
            
            # 克隆单元格样式
            new_cell._element.get_or_add_tcPr().append(
                copy.deepcopy(orig_cell._element.get_or_add_tcPr())
            )
            
            # 添加翻译文本
            translated = translate_text(client, orig_cell.text)
            new_run = new_cell.paragraphs[0].add_run(translated)
            new_run.font.size = TRANSLATE_FONT_SIZE
            new_run.italic = TRANSLATE_ITALIC

def process_paragraphs(doc, client):
    for para in list(doc.paragraphs):  # 转为list防止遍历错乱
        if not para.text.strip():
            continue
        
        # 创建新段落并克隆样式
        new_para = doc.add_paragraph()
        clone_paragraph_style(para, new_para)
        
        # 设置翻译格式
        translated = translate_text(client, para.text)
        new_run = new_para.add_run(translated)
        new_run.font.size = TRANSLATE_FONT_SIZE
        new_run.italic = TRANSLATE_ITALIC
        
        # 解决中文乱码问题
        new_run.font.name = 'Times New Roman'
        new_run._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')


def process_segment(client, segment_paragraphs, segment_range, input_path):
    """
    Translates paragraphs in a segment line by line, and outputs results to a TXT file.
    """
    print(f"  >> Processing Pages {segment_range[0]}–{segment_range[1]} "
          f"({len(segment_paragraphs)} paragraphs in this segment)")

    base, ext = os.path.splitext(input_path)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = f"{base}_translated_{segment_range[0]}-{segment_range[1]}_{timestamp}.txt"

    try:
        with open(output_path, "w", encoding="utf-8") as out_file:
            for idx, (page_num, paragraph) in enumerate(segment_paragraphs, start=1):
                raw_text = paragraph.text.strip()
                if not raw_text:
                    continue

                print(f"    - Paragraph {idx}/{len(segment_paragraphs)}, Page {page_num}")

                lines = raw_text.split('\n')
                for line_idx, line in enumerate(lines, start=1):
                    line_text = line.strip()
                    if not line_text:
                        continue

                    # Indicate progress in console
                    progress_info = (f"Paragraph {idx}/{len(segment_paragraphs)}, "
                                     f"Page {page_num}, Line {line_idx}/{len(lines)}")
                    print(f"       * Translating line {line_idx} of {len(lines)} in paragraph {idx}")

                    # Perform translation
                    translated_line = translate_text(client, line_text, progress_info)

                    # Write the translated line to the TXT file
                    out_file.write(line_text + "\n")
                    out_file.write(translated_line + "\n")
                    out_file.write("\n")  # extra newline for separation
                    out_file.flush()

        print(f"  >> Segment saved to: {output_path}\n")

    except Exception as e:
        print(f"  !! Error saving segment {segment_range[0]}–{segment_range[1]}: {e}")

def main():
    """
    Script workflow:
      1. Read input DOCX file from argv.
      2. Assign page numbers to paragraphs using page breaks.
      3. Group paragraphs into segments of PAGES_PER_SEGMENT pages.
      4. For each segment, translate line-by-line and save a partial TXT file.
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

    # Initialize DeepSeek / OpenAI client
    client = init_client()

    # # 1) Assign pages to paragraphs
    # paragraphs_with_page, final_page_count = assign_page_numbers(doc)

    # # 2) Group into segments
    # segments = group_into_segments(paragraphs_with_page)

    # print(f"\nDocument has ~{final_page_count} pages.")
    # print(f"Dividing into segments of {PAGES_PER_SEGMENT} pages each => total segments: {len(segments)}\n")

    # # 3) Process each segment -> produces one TXT file per segment
    # for seg_index in sorted(segments.keys()):
    #     start_page = (seg_index - 1) * PAGES_PER_SEGMENT + 1
    #     end_page = seg_index * PAGES_PER_SEGMENT
    #     process_segment(client, segments[seg_index], (start_page, end_page), input_path)

    # print("All segments processed successfully.")


     # 处理顺序：先表格后段落
    process_tables(doc, client)
    process_paragraphs(doc, client)
    
    # 保存时保留原文档格式[8](@ref)
    output_path = os.path.splitext(sys.argv[1])[0] + "_translated.docx"
    doc.save(output_path)
    print(f"成功生成翻译文档：{output_path}")

if __name__ == "__main__":
    main()
