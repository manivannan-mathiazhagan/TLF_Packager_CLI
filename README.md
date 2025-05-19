# 📦 TLF_PACKAGER

**Automate the packaging, bookmarking, and Table of Contents (TOC) creation for clinical TLF (Tables, Listings, Figures) outputs.**  
Seamlessly combines SAS macros and Python scripts for an efficient, consistent workflow.

---

## 🧩 Key Features

- Packages both **RTF** and **PDF** outputs into a single bookmarked PDF.
- Automatically generates a clickable **Table of Contents (TOC)** for easy navigation.
- Extracts and formats titles from outputs for use as bookmarks and TOC entries.
- Cleans up intermediate files after processing (optional).
- Flexible for various study output folder structures.

---

## 📚 Included Macro

### 1. `TLF_Packager`

Packages all RTF or PDF files from a given folder into a single, bookmarked PDF with an optional TOC, leveraging Python for content extraction and PDF merging.

**Features:**
- Detects input type (RTF/PDF) and processes accordingly.
- Dynamically extracts the first three lines (including Listing/Table/Figure) for use as bookmark text.
- Can delete temporary PDFs after merge (for RTF input).
- Output PDF is auto-named by folder and timestamp.

**Input Requirements:**
- Folder containing TLFs as **RTF** or **PDF** files.

---

### Macro Parameters

| **Parameter**     | **Required** | **Default** | **Description**                                                                    |
|-------------------|--------------|-------------|------------------------------------------------------------------------------------|
| `input_path`      | ✅ Yes       | —           | Path to the folder containing RTF or PDF files.                                    |
| `input_type`      | ✅ Yes       | —           | `RTF` or `PDF` – determines processing workflow.                                   |
| `toc`             | ❌ No        | `Y`         | Y/N flag for TOC title line. `Y` adds "Table of Contents" at top of TOC; `N` omits.|
| `delete_pdfs`     | ❌ No        | `Y`         | Y/N flag to delete intermediate PDFs (for RTF input only).                         |

---

### Example Usage

%TLF_Packager(input_path=E:\Study\Output\Listings, input_type=RTF, toc=Y, delete_pdfs=Y);
%TLF_Packager(input_path=E:\Study\Output\Tables, input_type=PDF, toc=N);

---

# Python Scripts

This toolkit includes Python scripts that drive the core automation for merging, bookmarking, and TOC generation.

### 1. `pack_rtfs_with_toc.py`

- **Purpose:**  
  Converts RTF files to PDFs (using LibreOffice), extracts key titles, merges PDFs, adds bookmarks, and inserts a clickable TOC.

- **Requirements:**
  - Python 3+
  - [PyMuPDF](https://pymupdf.readthedocs.io/en/latest/) (`pip install pymupdf`)
  - [LibreOffice](https://www.libreoffice.org/download/download/) (for RTF to PDF conversion)

- **Usage:**

  python pack_rtfs_with_toc.py "<input_folder_path>" "<output_file_name.pdf>" [TOC_TITLE:Y/N] [DELETE_PDFS:Y/N]
  DELETE_PDFS: Y deletes temporary PDFs after merging (default is Y, only applies to RTF input).

---

### 2. `pack_pdfs_with_toc.py`

- **Purpose:** 
Merges multiple PDF files in a folder, extracts section titles for bookmarks, and generates a clickable TOC.

- **Requirements:**
  - Python 3+
  - [PyMuPDF](https://pymupdf.readthedocs.io/en/latest/) (`pip install pymupdf`)

- **Usage:**
python pack_pdfs_with_toc.py "<input_folder_path>" "<output_file_name.pdf>" [TOC_TITLE:Y/N]
TOC_TITLE: Y adds "Table of Contents" line at the top of the TOC, N omits it (default is Y).


---

# 🗂️ How the Scripts Work

## Title Extraction
- The scripts extract the first line containing Listing/Table/Figure (Title line 1), plus up to two subsequent lines.
- This extracted text is used for both PDF bookmarks and Table of Contents (TOC) entries.

## TOC Generation
- The scripts automatically paginate the TOC based on content length.
- Clickable links are inserted in the TOC, allowing direct navigation to each output section in the PDF.

## Bookmarks
- Bookmarks are created at the beginning of each merged output, enabling quick navigation throughout the final PDF.

## File Handling
- For RTF input, temporary PDFs generated during conversion are optionally deleted after merging to save space and maintain folder cleanliness.

---

# 💻 Requirements

- **SAS** (tested with SAS 9.4)
- **Python 3+** (with [PyMuPDF](https://pymupdf.readthedocs.io/en/latest/) installed: `pip install pymupdf`)
- **LibreOffice** (for RTF conversion; required by `pack_rtfs_with_toc.py`)
- **OS:** Windows or Mac (scripts use system calls for automation)

---
