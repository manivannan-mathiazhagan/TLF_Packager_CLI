# 📦 TLF_PACKAGER

**Automate the packaging, bookmarking, and Table of Contents (TOC) creation for clinical TLF (Tables, Listings, Figures) outputs.**  
Seamlessly combines SAS macros and Python scripts for an efficient, consistent workflow.

---

## 🧩 Key Features

- **Unified Packaging:** Combines RTF, PDF, and DOCX outputs into a single, bookmarked PDF.

- **Automated Table of Contents:** Generates a fully clickable TOC for easy navigation within the merged PDF.

- **Smart Title Extraction:** Automatically extracts and formats section titles (Listing, Table, Figure, Appendix) to use as bookmarks and TOC entries.

- **User-Driven Workflow:** Lets users review, edit, and approve bookmarks, order, and conversion method (Word/LibreOffice) in Excel before final PDF creation.

- **Flexible & Folder-Agnostic:** Adapts to various study output folder structures and naming conventions.

- **Automatic Cleanup:** Optionally deletes intermediate files (e.g., converted PDFs) after processing to keep folders tidy.

- **Cross-Platform & Easy Integration:** Runs with SAS 9.4 and Python 3+; supports both Windows and Mac environments.
---

## 📚 Included Macro

###  `TLF_Packager.sas`

Packages all RTF, PDF, and DOCX files from a given folder into a single, bookmarked PDF with an optional Table of Contents (TOC), leveraging Python for content extraction and PDF merging.

**Features:**
- Detects input type (**RTF**, **PDF**, or **DOCX**) and processes each accordingly.
- Dynamically extracts the first three lines (including **Listing/Table/Figure**) for use as bookmark text.
- Can delete temporary PDFs after merge (for **RTF** and **DOCX** input).
- **WORD/LIBREOFFICE** Conversion for **RTF** and **DOCX** to **PDF** based on user input.
- Output **PDF** is auto-named by folder and timestamp.

**Input Requirements:**
- Folder containing TLFs as **RTF**, **PDF**, or **DOCX** files.
---

### Macro Parameters

| **Parameter**     | **Required** | **Default** | **Description**                                                                    |
|-------------------|--------------|-------------|------------------------------------------------------------------------------------|
| `input_path`      | ✅ Yes       | —           | Path to the folder containing RTF, PDF, and/or DOCX files.                         |
| `delete_pdfs`     | ❌ No        | `Y`         | Y/N flag to delete intermediate PDFs (for RTF and DOCX input only).                         |
| `default_conv`             | ❌ No        | `WORD`         |LIBREOFFICE/WORD: Default converter for all RTFs. User can override in Excel.|

---

### Example Usage

%TLF_Packager(input_path=E:\Study\Output\Listings);                                 
%TLF_Packager(input_path=E:\Study\Output\Tables, delete_pdfs=N, default_conv=LIBREOFFICE); 

---

# Python Script

This toolkit includes a Python script that drive the core automation for merging, bookmarking, and TOC generation.

###  `TLF_Packager.py`

- **Purpose:**  
TLF_Packager.py automates the end-to-end packaging of clinical Tables, Listings, and Figures (TLFs) by merging **RTF**, **PDF**, and **DOCX** outputs into a single, bookmarked PDF with a clickable Table of Contents (TOC). The script enables users to review, reorder, and customize section bookmarks via Excel, and select the desired conversion method (Microsoft Word or LibreOffice) for each **RTF** or **DOCX** output, streamlining and standardizing the TLF delivery process.

- **Requirements:**
  - Python 3+
  - [PyMuPDF](https://pymupdf.readthedocs.io/en/latest/) (`pip install pymupdf`)
  - [Openpyxl](https://openpyxl.readthedocs.io/en/stable/) (`pip install openpyxl`)
  - [Pywin32](https://pypi.org/project/pywin32/) (`pip install pywin32`)
  - [python-docx](https://python-docx.readthedocs.io/en/latest/) (`pip install python-docx`)
  - [LibreOffice](https://www.libreoffice.org/download/download/) (for RTF to PDF conversion)
  - Microsoft Word 

- **Usage:**

 python TLF_Packager.py "input_folder_path" "output_file_name.pdf" Y WORD

---

# 🗂️ How the Script Works

## File Detection
  - The script scans the specified folder and lists all RTF and/or PDF files for processing.

## Title Extraction
  - For each file, the script extracts key section titles—identifying the first line containing Listing, Table, Figure, or Appendix, plus up to two additional lines.

  - Extracted titles are validated and used for PDF bookmarks and Table of Contents (TOC) entries.

## User Review & Bookmark Approval
  - All extracted titles and proposed bookmarks are written to an Excel worksheet.

  - Users can review, edit, reorder, and specify the preferred RTF-to-PDF converter (Word or LibreOffice) directly in Excel before the PDF is created.

## PDF Generation with Bookmarks & TOC
  - Upon user confirmation, the script converts RTF files to PDF (using the selected converter), then merges all outputs in the approved order.

  - The final PDF includes:

    -   Bookmarks at each section for navigation

    - An automatically generated, clickable Table of Contents (TOC)

## Temporary File Handling
- PDFs generated from RTFs during conversion are automatically deleted after merging (if selected), keeping the working folder clean.

---
