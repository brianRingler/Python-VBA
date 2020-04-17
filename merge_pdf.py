import os
import logging
import datetime as dt

try:
    from PyPDF2 import PdfFileMerger
except ImportError:
    print("PyPDF2 is required. Install and try again")

logging.basicConfig(filename='app.log', filemode='w',format='%(name)s - %(levelname)s - %(message)s')

path = os.getcwd()
logging.warning(f"This is the path: {path}")


y_m_d = dt.datetime.today().strftime("%Y%m%d")
output_filename = 'merged_files_' + y_m_d + '.pdf'
pdfs_merge = []


for filename in os.listdir(path):
    if filename.endswith(".PDF") or filename.endswith(".pdf"):
        pdfs_merge.append(filename)


# create object from PdfFileMerger class
merger = PdfFileMerger()

for pdf in pdfs_merge:
    merger.append(pdf)

merger.write(output_filename)
merger.close()

logging.warning(f"This is output_filename: {output_filename}")
logging.warning(f"This is pdfs_merge: {pdfs_merge}")