#! python
import os
import sys
import logging
import ctypes
import datetime as dt
from SendMail import send_email_outlook

def error_msg(error_type):
    '''If an error occurs notify user of error and provide option to email developer 
    of the error of the error. Okay = 1 and send notice and cancel = 2 do not send.
    Function takes a string and returns int 1 or 2.'''

    msg1 = ": Select Okay to send email or Cancel to not send notification"
    msg2 = "Python Error notify RingVision"

    # returns 1 if okay and 2 cancel
    value = ctypes.windll.user32.MessageBoxW(0,error_type + msg1, msg2, 1)
    return value

# check if path is exists already 
logging_path = 'C:\\python_logging'
if os.path.exists(logging_path):
    pass
else:
    os.mkdir(logging_path)
#  path to logging file attach to email if errrors
file_name = logging_path + "\\" + "pdf_merger.log"

logging.basicConfig(level=logging.INFO,
                    filename=file_name, 
                    filemode='w',
                    format='%(name)s - %(levelname)s - %(asctime)s - %(message)s - %(pathname)s')

try:
    from PyPDF2 import PdfFileMerger
except ImportError:
    ret_val = error_msg('PyPDF2NotFound')
    if ret_val == 1:
        # if okay need to send email notice of error
        logging.info("PyPDF2 - Okay selected")
        send_email_outlook("Merge PDf","PyPDF2 import Error: merge_pdf.py",file_name)
        #sys.exit()
    else:
        logging.info("PyPDF2 - Cancel selected")
        sys.exit(1)

# sys.argv[0] returns full path plus script name
# sys.argv[1] returns just the path of ActiveWorkbook
logging.info(f"sys.argv[1]: {sys.argv[1]}")
path = sys.argv[1]

# cd to Excel file 
os.chdir(path)
logging.info(f"Directory os.chdir(path): {path}")
logging.info(f"All files: {os.listdir(path)}")

# current date used for naming merged pdf 
y_m_d = dt.datetime.today().strftime("%Y%m%d")
output_filename = 'merged_files_' + y_m_d + '.pdf'


pdfs_merge = []
for filename in os.listdir(path):
    # find all files that end with pdf
    if filename.endswith(".PDF") or filename.endswith(".pdf"):
        pdfs_merge.append(filename)


# create object from PdfFileMerger class
merger = PdfFileMerger()
for pdf in pdfs_merge:
    merger.append(pdf)

merger.write(output_filename)
merger.close()

logging.info(f'List of all files to merge "pdfs_merge": {pdfs_merge}')
