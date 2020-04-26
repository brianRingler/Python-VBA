***
# About
***
Using VBA to import and run Python scripts to enhance Excel. First release is 

***
# Python Script Merge PDFs
***
The Python script `merge_pdf.py` is called from an Excel file which will locate all PDF files where the Excel file resides.  The script uses PyPDF2 version `(1.26.0)` to **append** the PDFs into one PDF and saves the file as `merged_files_yyyy_mm_dd`. 

The script takes one argument from the Excel file which is the path of the PDF files that the used would like merged into one PDF.  

A custom error function `error_msg()` with a simple message box is used to notify user if an error occurred. It provides user option to email the developer by selecting `Okay` or do nothing by selecting `Cancel`. 

If there is an error running the Python file the current configuration is set to email from users Outlook.  The file `SendMail.py` contains the function to send the `pdf_merger.log`. 

The script saves a logging file to `C:\python_logging` which is the required location of the merge_pdf.py file. 
1. `merge_pdf.py`
2. `pdf_merger.log`

Both `merge_pdf.py` and `SendMail.py` are coded to be stored in `C:\\python_logging` directory.  

***
# Excel Merge PDF
***
This Excel file has a one sheet `RingVision` with a button to run the macro, `run_python_pdf_merger`. There is an additional macro `pdf_merger_shell32` that uses  code from the `pdf_merge` module.  The sheet has two spots for entering file location of python script and `python.exe` if using the "Wscript.Shell"). 

The sheet `RingVision` has two (2) spots for entering the path of two files:
1. The first path `C:\\python_logging` is to the required Python scripts. This can be modified but needs to be addressed in the the code. `merge_pdf.py` creates and saves the log file to that location if it does not exist.
   1. `merge_pdf.py`
   2. `SendMail.py`
2. The second path is the Python installation which contains the `python.exe` file.  This will be different depending on Python install. **FIX** would be to create VBA function that locates path ONE TIME and saves to the file.  
