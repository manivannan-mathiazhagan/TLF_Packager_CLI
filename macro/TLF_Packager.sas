/*****************************************************************************************************/
/* Macro Name     : TLF_Packager                                                                     */
/*                                                                                                   */
/* Purpose        : Packages RTF or PDF files in a given folder into a single PDF with bookmarks and */
/*                  an optional clickable Table of Contents (TOC), using Python scripts.             */
/*                                                                                                   */
/* Author         : Manivannan Mathialagan                                                           */
/* Created On     : 19-May-2025                                                                      */
/*                                                                                                   */
/* Parameters                                                                                        */
/*   - input_path   (Required) : Path to the folder containing RTF or PDF files.                     */
/*   - input_type   (Required) : RTF or PDF.                                                         */
/*   - toc          (Optional) : Y/N flag to indicate if TOC should be added. Default is Y.          */
/*   - delete_pdfs  (Optional) : Y/N flag to delete intermediate PDFs after merge (RTF only).        */
/*                                                                                                   */
/* Example Usage                                                                                     */
/*   %TLF_Packager(input_path=E:\Study\Output\Listings, input_type=RTF, toc=Y, delete_pdfs=Y);       */
/*   %TLF_Packager(input_path=E:\Study\Output\Tables, input_type=PDF, toc=Y);                        */
/*                                                                                                   */
/* Notes                                                                                             */
/*   - Python must be installed and accessible via system PATH or specified directly.                */
/*   - Python scripts `pack_rtfs_with_toc.py` and `pack_pdfs_with_toc.py` must be in &glibroot path. */
/*   - Uses `filename ... pipe` to stream Python output directly into the SAS log.                   */
/*****************************************************************************************************/

%macro TLF_Packager(input_path=, input_type=, toc=Y, delete_pdfs=Y);

  %local script output base folder_name ts logfile pycmd shellcmd;

  /* Generate timestamp */
  %let __ST = %sysfunc(datetime());
  data _null_;
    st = &__ST;
    datestr = put(datepart(st), yymmddn8.);
    timestr = put(timepart(st), tod8.);
    ts_clean = cats(datestr, '_T', compress(timestr, ':'));
    call symputx('ts', ts_clean);
  run;

  %put [INFO] Timestamp generated: &ts;

  /* Detect folder type for output name prefix */
  %let folder_name = %scan(%sysfunc(reverse(%sysfunc(scan(%sysfunc(reverse(&input_path)),1,"\")))),1," ");
  %if %index(%upcase(&folder_name), TABLE) %then %let base = Tables;
  %else %if %index(%upcase(&folder_name), LISTING) %then %let base = Listings;
  %else %if %index(%upcase(&folder_name), FIGURE) %then %let base = Figures;
  %else %let base = TLFs;

  %let output = &base._&ts..pdf;
  %let logfile = &input_path.\log_&base._&ts..txt;

  /* Determine script to use */
  %let input_type = %upcase(&input_type);
  %if &input_type = RTF %then %let script = pack_rtfs_with_toc.py;
  %else %if &input_type = PDF %then %let script = pack_pdfs_with_toc.py;
  %else %do;
    %put ERROR: input_type must be RTF or PDF.;
    %return;
  %end;

  /* Build Python execution command */
  %if &input_type = RTF %then %do;
    %let pycmd = D:\Python\Python37\python &glibroot\Unvalidated\&script.
                   "&input_path." "&input_path.\&output." &toc &delete_pdfs;
  %end;
  %else %do; /* PDF */
    %let pycmd = D:\Python\Python37\python &glibroot\Unvalidated\&script.
                   "&input_path." "&input_path.\&output." &toc;
  %end;

  /* Construct shell command with output redirection */
  %let shellcmd = %sysfunc(quote(&pycmd. > "&logfile." 2>&1));

  /* Pipe the execution and capture the log in SAS */
  filename pycall pipe &shellcmd;

  data _null_;
    infile pycall;
    input;
    putlog "[Python] " _infile_;
  run;

  %put NOTE: Log saved to &logfile.;
  %put NOTE: Output PDF: &output.;

%mend;

/* Example call:
%TLF_Packager(input_path=E:\Projects\Bristol Myers Squibb\KarXT Kar-012\Output\Draft 1 TLFs\tables_v20250508_d20250516, input_type=RTF);
%TLF_Packager(input_path=E:\Projects\Bristol Myers Squibb\KarXT Kar-012\Output\Draft 1 TLFs\Check\tables_v20241014_d20250516, input_type=PDF);
*/
