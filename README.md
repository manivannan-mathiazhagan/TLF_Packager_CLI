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

# TLF Packager

The **TLF Packager** is a Python and SAS-integrated automation tool designed to streamline the process of packaging Tables, Listings, Figures, and Appendices for clinical study reports. It supports RTF, PDF, and DOCX files and generates a single bookmarked PDF with a clickable Table of Contents (TOC), allowing user-defined ordering and customization via Excel.

---

## Evolution Summary

### 1. Initial Phase: Separate Scripts
- Separate scripts for PDF and RTF.
- No unified handling or ordering.
- Bookmarks generated directly from extracted titles.

### 2. Unified Packager Script Introduced
- Combines RTF and PDF logic into a single script.
- Folder-based scanning.
- Title extraction and bookmark merging into one PDF.

### 3. Intermediate Excel Output Introduced
- Excel shows file details, titles, bookmark names, and order.
- Allows review and customization.

### 4. Confirmation and Popup Control
- Tkinter popup prompts for user confirmation after Excel edit.

### 5. DOCX Format Support Added
- DOCX titles extracted using `python-docx`.
- DOCX converted using MS Word automation.

### 6. Converter Method Customization
- Excel column `Converter` allows:
  - `WORD` – Microsoft Word
  - `LIBREOFFICE` – LibreOffice

### 7. Title Validation Logic Improved
- Validates based on keywords: Table, Listing, Figure, Appendix.
- Skips invalid/blank titles with logging.

### 8. Appendix Sorting Logic
- Appendices placed last in final PDF.
- Maintains separate hierarchy in bookmarks.

### 9. Alignment-Based RTF Title Extraction
- Extracts up to 3 center-aligned lines for title.
- Non-centered lines skipped.

### 10. Enhanced TOC and Bookmark Features
- TOC auto-generated and inserted as the first page.
- Reflects titles and Excel order.

### 11. Cleanup and Logging Enhancements
- Deletes temp files optionally.
- Logs steps, conversions, and errors.

### 12. SAS Integration via Macro
- SAS macro wraps the script.
- Parameters: `input_path`, `output_pdf`, `delete_flag`.
- Uses X command/PIPE to execute Python.

### 13. Packaging Improvements
- Final output includes:
  - Bookmarked PDF
  - Excel metadata file
  - Optional ZIP bundle

---

## Planned Enhancements

| Feature                        | Description |
|-------------------------------|-------------|
| Clickable TOC Links           | Direct PDF links from TOC to content |
| Section Headers/Grouping      | Group entries (e.g., Tables, Listings) |
| Preview Mode                  | Generate Excel metadata without merging |
| Hyperlinked Annotations       | Add internal page reference hyperlinks |
| Dynamic Bookmark Styling      | Customize fonts and levels |
| Error Recovery and Skipping   | Continue on failure, log errors |
| Command-Line Utility          | CLI with help, argument parsing |

---

## License

Distributed internally at Veristat. External use subject to licensing policy.

---

## Author

**Manivannan Mathialagan**  
Associate Manager – Statistical Programming  
---
