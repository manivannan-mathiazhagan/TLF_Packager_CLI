***-------------------------------------------------------------------------------------------------***;
*** Macro Name:    TLF_Packager.sas                                                                 ***;
***                                                                                                 ***;
*** Purpose:       Packages all RTF, DOCX, and PDF files from a specified folder into a single,     ***;
***                bookmarked PDF with a clickable Table of Contents (TOC).                         ***;
***                Utilizes the Python script TLF_Packager.py for all core operations.              ***;
***                Supports user-driven review and approval of bookmarks/order via Excel.           ***;
***                                                                                                 ***;
***                New Parameter: wait_approval = Y/N                                               ***;
***                - Y: Pause after Excel export for user review (default)                          ***;
***                - N: Skip approval step and proceed directly                                     ***;
***-------------------------------------------------------------------------------------------------***;
*** Programmed By: Manivannan Mathialagan                                                           ***;
*** Created On:    22-May-2025                                                                      ***;
*** Updated On:    08-Jul-2025                                                                      ***;
***-------------------------------------------------------------------------------------------------***;
*** Parameters:                                                                                     ***;
***                                                                                                 ***;
*** Name           | Description                                 | Default value   | Required       ***;
***----------------|---------------------------------------------|-----------------|----------------***;
*** input_path     | Path to folder containing RTF/DOCX/PDF files|     —           |     Yes        ***;
*** delete_pdfs    | Y/N: Delete intermediate PDFs post-merge    |     Y           |     No         ***;
*** wait_approval  | Y/N: Pause after Excel for manual review    |     Y           |     No         ***;
***-------------------------------------------------------------------------------------------------***;
*** Output(s):                                                                                      ***;
***   - PDF:   Single merged PDF with bookmarks and TOC                                             ***;
***   - Excel: Bookmark worksheet for user edits (titles/order)                                     ***;
***   - Log:   TXT log file with full Python output                                                 ***;
***-------------------------------------------------------------------------------------------------***;
*** Dependencies:                                                                                   ***;
*** - Python 3 with required packages (openpyxl, PyMuPDF, pywin32, docx)                            ***;
*** - Microsoft Word (RTF/DOCX conversion)                                                          ***;
*** - TLF_Packager.py located in P:\Biostatistics\PDF Tool\                                         ***;
***-------------------------------------------------------------------------------------------------***;
*** Example Usage:                                                                                  ***;
*** %TLF_Packager(input_path=E:\Study\Output\Listings)                                              ***;
*** %TLF_Packager(input_path=E:\Study\Output\Tables, delete_pdfs=N, wait_approval=N)                ***;

***-------------------------------------------------------------------------------------------------***;
%macro TLF_Packager(input_path=, delete_pdfs=N, wait_approval=N);
    %local scriptpath pyexe ts base folder_name output_pdf logfile batfile;

    /* Path to Python script and executable */
    %let scriptpath=P:\BSP_LocalDev\Manivannan.Mathialag\zzzz My SAS
        Files\Presentation\PHUSE 2026\Tool\Version 1.0\TLF_Packager.py;
    %let pyexe=C:\Program Files\Python313\python.exe;

    /* Timestamp */
    %let ts=%sysfunc(putn(%sysfunc(today()),
        yymmddn8.))_%sysfunc(compress(%sysfunc(putn(%sysfunc(time()), time8.)),
        :));

    /* Extract folder name */
    %let folder_name=%scan(&input_path, -1, \);

    %if %sysfunc(index(%upcase(&folder_name), TABLE)) %then %let base=Tables;
    %else %if %sysfunc(index(%upcase(&folder_name), LISTING)) %then %let base=
        Listings;
    %else %if %sysfunc(index(%upcase(&folder_name), FIGURE)) %then %let base=
        Figures;
    %else %let base=TLFs;

    %let batfile=&input_path.\run_tlf_packager_&ts..bat;
    %let logfile=&input_path.\log_&base._&ts..txt;
    %let output_pdf=&base._&ts..pdf;

    %put NOTE: Generating output using:;
    %put NOTE- Python Executable: &pyexe;
    %put NOTE- Python Script : &scriptpath;
    %put NOTE- Input Path : &input_path;
    %put NOTE- Output PDF Name : &output_pdf;
    %put NOTE- Delete PDFs : &delete_pdfs;
    %put NOTE- Wait for Review : &wait_approval;
    %put NOTE- Log File : &logfile;
    %put NOTE- BAT File : &batfile;

    /* Create BAT file */
    data _null_;
        file "&batfile.";
        put '@echo off';
        put '"' "&pyexe." '" "' "&scriptpath." '" "' "&input_path." '" "'
            "&output_pdf." '" "' "&delete_pdfs." '" "' "&wait_approval." '" > "'
            "&logfile." '" 2>&1';
    run;

    /* Execute BAT */
    options noxwait;
    x """&batfile.""";

    /* Delete temporary BAT file */
    x "del ""&batfile.""";

%mend;

/* Example Call */
/*
%TLF_Packager(input_path=P:\Biostatistics\PDF Tool\Testing Veristat PDF\RRP-002\1. Tables);


%TLF_Packager(input_path=P:\Biostatistics\PDF Tool\Testing Veristat PDF\RRP-002\2. Listings,delete_pdfs=Y);

%TLF_Packager(input_path=P:\Biostatistics\PDF Tool\Testing Veristat PDF\KAR-012\TLF\Tables);

 */
