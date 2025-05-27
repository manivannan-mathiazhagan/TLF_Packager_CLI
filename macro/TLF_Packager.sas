/***-------------------------------------------------------------------------------------------------***/
/*** Macro Name:    TLF_Packager.sas                                                                 ***/
/***                                                                                                 ***/
/*** Purpose:       Packages all RTF and PDF files from a specified folder into a single,            ***/
/***                bookmarked PDF with a clickable Table of Contents (TOC).                         ***/
/***                Utilizes the Python script TLF_Packager.py for all core operations.              ***/
/***                RTF-to-PDF conversion is controlled per file in Excel ("Converter" column):      ***/
/***                  - Use Microsoft Word for files marked "WORD"                                   ***/
/***                  - Use LibreOffice for files marked "LIBREOFFICE"                               ***/
/***                The default converter (for all RTFs) can be specified via macro argument.        ***/
/***                Supports user-driven review and approval of bookmarks/order via Excel.           ***/
/***-------------------------------------------------------------------------------------------------***/
/*** Programmed By: Manivannan Mathialagan                                                           ***/
/***                                                                                                 ***/
/*** Created On:    22-May-2025                                                                      ***/
/***-------------------------------------------------------------------------------------------------***/
/*** Parameters:                                                                                     ***/
/***                                                                                                 ***/
/***-------------------------------------------------------------------------------------------------***/
/*** Name           | Description                                 | Default value   | Required       ***/
/***----------------|---------------------------------------------|-----------------|----------------***/
/*** input_path     | Path to folder containing RTF and/or PDF    |   None          |    Yes         ***/
/***                | files to be packaged.                       |                 |                ***/
/***----------------|---------------------------------------------|-----------------|----------------***/
/*** delete_pdfs    | Y/N flag: delete PDFs converted from RTFs   |   Y             |    No          ***/
/***                | after merging.                              |                 |                ***/
/***----------------|---------------------------------------------|-----------------|----------------***/
/*** default_conv   | LIBREOFFICE/WORD: Default converter         |   WORD          |    No          ***/
/***                | for all RTFs. User can override in Excel.   |                 |                ***/
/***-------------------------------------------------------------------------------------------------***/
/*** Output(s):                                                                                      ***/
/***                                                                                                 ***/
/***   - PDF:      Single merged PDF with bookmarks and clickable TOC, placed in the input folder.   ***/
/***   - Excel:    Worksheet for user review/approval of bookmark titles, order, and converter.      ***/
/***   - Log:      TXT log file streaming all Python output for audit/troubleshooting.               ***/
/***-------------------------------------------------------------------------------------------------***/
/*** Macro Variables:    None                                                                        ***/
/*** Data sets:          None                                                                        ***/
/*** Variables:          None                                                                        ***/
/*** Other Files:        Dynamically named PDF, Excel, and TXT log files in input folder.            ***/
/***-------------------------------------------------------------------------------------------------***/
/*** Dependencies:                                                                                   ***/
/***                                                                                                 ***/
/*** - Python 3 (with openpyxl, PyMuPDF, pywin32 installed)                                          ***/
/*** - LibreOffice (for RTF-to-PDF conversion)                                                       ***/
/*** - Microsoft Word (for RTF-to-PDF conversion)                                                    ***/
/*** - TLF_Packager.py script in &glibroot\Unvalidated\                                              ***/
/***-------------------------------------------------------------------------------------------------***/
/*** Example Usage:                                                                                  ***/
/***                                                                                                 ***/
/*** %TLF_Packager(input_path=E:\Study\Output\Listings);                                             ***/
/*** %TLF_Packager(input_path=E:\Study\Output\Tables, delete_pdfs=N, default_conv=WORD);             ***/
/***-------------------------------------------------------------------------------------------------***/
/*** Notes:                                                                                          ***/
/***                                                                                                 ***/
/***   1. The Python script creates an Excel worksheet listing all files, their titles/bookmarks,    ***/
/***      and a "Converter" column.                                                                  ***/
/***      User may change converter for any RTF (dropdown: WORD/LIBREOFFICE)                         ***/
/***   2. Macro passes default_conv value to Python to pre-fill Excel.                               ***/
/***   3. After Excel review, merging and conversion proceed as indicated in worksheet.              ***/
/***   4. Both LibreOffice and Word must be installed and accessible for full function.              ***/
/***-------------------------------------------------------------------------------------------------***/

%macro TLF_Packager(input_path=, delete_pdfs=Y, default_conv=WORD);

    %local scriptpath pyexe ts base folder_name output_pdf logfile batfile;

    /* Path to Python script and executable Python file */
    %let scriptpath=&glibroot\Unvalidated\TLF_Packager\TLF_Packager.py;
    %let pyexe=D:\Python\Python37\python.exe;

    /* Generate timestamp as a macro variable for uniqueness */
    %let __ST = %sysfunc(datetime());
    data _null_;
        st = &__ST;
        datestr = put(datepart(st), yymmddn8.);
        timestr = put(timepart(st), tod8.);
        ts_clean = cats(datestr, '_T', compress(timestr, ':'));
        call symputx('ts', ts_clean);
    run;

    /* Set base name for output files (customize as needed) */
    %let folder_name = %scan(%sysfunc(reverse(%sysfunc(scan(%sysfunc(reverse(&input_path)),1,"\")))),1," ");
    %if %index(%upcase(&folder_name), TABLE) %then %let base = Tables;
    %else %if %index(%upcase(&folder_name), LISTING) %then %let base = Listings;
    %else %if %index(%upcase(&folder_name), FIGURE) %then %let base = Figures;
    %else %let base = TLFs;

    %let batfile = &input_path.\run_tlf_packager_&ts..bat;
    %let logfile = &input_path.\log_&base._&ts..txt;
    %let output_pdf = &base._&ts..pdf;

    %put NOTE: BATFILE: &batfile LOGFILE: &logfile SCRIPTPATH: &scriptpath PYEXE: &pyexe OUTPUT_PDF: &output_pdf;

     /* Generate BAT file with resolved macro values */
    data _null_;
        file "&batfile.";
        put '@echo off';
        put 
          '"' "&pyexe." '" ' 
          '"' "&scriptpath." '" '
          '"' "&input_path." '" '
          '"' "&output_pdf." '" '
          '"' "&delete_pdfs." '" '
          '"' "&default_conv." '" '
          '> "' "&logfile." '" 2>&1';
    run;

    /* Execute BAT file */
    options noxwait;
    X """&batfile.""";

    /* Optional: Clean up BAT file afterwards */
    X "del ""&batfile."" ";

    %put NOTE: Log saved to &logfile. Output PDF: &output_pdf.;

%mend;

/* Example usage:
%TLF_Packager(input_path=E:\Projects\Bristol Myers Squibb\KarXT Kar-012\Output\Draft 1 TLFs\Check\tables_v20241014_d20250516);
%TLF_Packager(input_path=E:\Projects\Bristol Myers Squibb\KarXT Kar-012\Output\Draft 1 TLFs\Check\both files, delete_pdfs=N);
*/

