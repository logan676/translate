# README

This repository contains Python scripts to split a large DOCX file into smaller segments, translate each segment, and (optionally) merge the translated segments into one final DOCX file. Follow the steps below to set up your environment and run the scripts on Ubuntu.

---

## 1. Enable Virtual Environment

Using a virtual environment is recommended to isolate your project dependencies. Open your terminal and run the following commands:

```bash
# Create a virtual environment named 'venv'
python3 -m venv venv

# Activate the virtual environment
source venv/bin/activate
```

When activated, your prompt should show `(venv)`. To deactivate, simply run:

```bash
deactivate
```

---

## 2. Install Dependencies

With your virtual environment activated, install the required libraries using pip. You can install them individually or via a `requirements.txt` file if provided.

### Option A: Using `requirements.txt`
```bash
pip install -r requirements.txt
```

---

## 3. Run the Scripts

### 3.1 Pre-Process: Split Original DOCX

Run the `split_docx.py` script to split your large DOCX into smaller segments. You need to specify:
- **Input DOCX file**: Your original DOCX document.
- **Output folder**: The folder where the split segments will be saved.
- **Segment page count**: The approximate number of pages per segment (based on page breaks).

Example command:
```bash
python3 split_docx.py my_document.docx splitted 5
```
- `my_document.docx`: Your original document.
- `splitted`: Folder to store the split files.
- `5`: Number of pages per segment.

### 3.2 Process: Translate the Splitted Files

Run the `translate_docx.py` script to translate each split file into English. You need to provide:
- **Splitted folder**: The folder containing the split DOCX files.
- **Translated folder**: The folder where the translated files will be saved.

Example command:
```bash
python3 google_translate.py splitted/xx 
```
- `splitted`: Folder containing the split files from step 3.1.
- `translated`: Folder where the translated files will be placed.

### 3.3 (Optional) Merge: Combine Translated Files

If you have a merge script (e.g., `merge_docx.py`), you can merge the translated segments into one final DOCX file.

Example command:
```bash
python3 merge_docx.py translated final_merged.docx
```
- `translated`: Folder with the translated DOCX files.
- `final_merged.docx`: Name for the final merged document.

---

## Additional Notes

- **Error Logs**: Each script generates an error log (e.g., `split_error.log`, `translate_error.log`, or `merge_error.log`) if issues occur during processing.
- **Usage**: Always ensure your virtual environment is activated before running any script.
- **Customization**: You can adjust the segment page count and other parameters within the scripts as needed.

Happy processing!