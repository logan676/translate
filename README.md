## 1. Enable Virtual Environment

Using a virtual environment is recommended to isolate your project dependencies. Open your terminal and run the following commands:

```bash
# Create a virtual environment named 'venv'
# Activate the virtual environment
python3 -m venv venv
source venv/bin/activate
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

```bash
./docx_pdf_docx.sh my_document.docx
```
### 3.2 Process: Translate the Splitted Files

Run the `./batch_translate.py` script to translate each split file into English. 

Example command:
```bash
./batch_translate.py
```

### 3.3 Merge: Combine Translated Files

```bash
python3 merge_docx.py pages_output
```
---

Happy processing!