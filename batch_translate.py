#!/usr/bin/env python3
import glob
import concurrent.futures
import subprocess
import os

def process_file(file_path):
    print(f"Processing {file_path} ...")
    # Call the translation script using the command:
    #   python3 ds_translate.py <file_path>
    result = subprocess.run(["python3", "ds_translate.py", file_path], capture_output=True, text=True)
    
    # Print standard output and error messages for debugging.
    print(result.stdout)
    if result.stderr:
        print(f"Error processing {file_path}:")
        print(result.stderr)
    print(f"Finished processing {file_path}\n")

def main():
    # Find all DOCX files in the 'pages_output' folder.
    docx_files = glob.glob(os.path.join("pages_output", "*.docx"))
    if not docx_files:
        print("No DOCX files found in pages_output.")
        return

    # Set number of workers; adjust max_workers according to your API rate limits.
    max_workers = 30
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        executor.map(process_file, docx_files)

if __name__ == "__main__":
    main()
