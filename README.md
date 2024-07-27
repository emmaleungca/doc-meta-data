# File Metadata Extraction Script

This script recursively scans a specified folder for files of various types (PDF, JPG/JPEG, DOCX, DOC, XLSX, and MSG), extracts relevant metadata, and saves the information to an Excel file.

## Features

- Extracts creation and modification dates from:
  - PDF files
  - JPG/JPEG images
  - DOCX and DOC Word documents
  - XLSX Excel files
- Extracts "From", "To", and "CC" email addresses from MSG files
- Saves the extracted metadata to an Excel file

## Prerequisites

Ensure you have the following libraries installed:

```bash
pip install extract-msg pdfminer.six pillow python-docx openpyxl pandas
```
## Usage
Update the folder_path variable in the script to point to the folder you want to scan.
Run the script.
Example
Replace 'path_to_your_folder' with the path to the folder containing your files:

```python
folder_path = 'path_to_your_folder'
```
Run the script:
```bash
python file_metadata_extraction.py
```
The script will create an Excel file named file_metadata.xlsx in the current directory, containing the following columns:
- Filename
- Creation Date
- Modification Date
- File Type
- Path (relative to the base folder)
- Folder Name
- From Email (for MSG files)
- To Email (for MSG files)
- CC Email (for MSG files)
