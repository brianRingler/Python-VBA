***
# About
***
Using VBA to import and run Python scripts to enhance Excel. 

***
# Python Script PDF Converter 
***
Using PyPDF2 the Python script `merge_pdf.py` will locate all PDF files in a folder indicated by the user and merge into one PDF and save in the same folder. 
1. `merge_pdf.py`

Added logging to see if the Python script is running when called from Excel. Does not appear to be. 

***
# Excel Merge PDF
***

This file has a one sheet with a button to run `pdf_merger_shell32`  or `run_python_pdf_merger` code from the `pdf_merge` module.  The sheet has two spots for entering file location of python script and `python.exe` if using the "Wscript.Shell"). 


