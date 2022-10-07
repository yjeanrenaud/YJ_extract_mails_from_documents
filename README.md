# YJ_extract_mails_from_documents
A small python project for collecting (camouflaged) emailadresses from local documents in PDF, doc and docx format


##requirements
- Python 3 with standard libraries
- [pyMuPDF] (https://github.com/pymupdf/PyMuPDF)
- [python-docx] (https://pypi.org/project/python-docx/)
##usage
- Just download (or clone) and run `yj_extract_mails_from_documents.py`.
- it will crawl all document for emails in the folder `PUT_YOUR_DOCUMENTS_HERE` and corresponding subfolders.
- Results are spilled out to a CSV file called `emails_collected.csv`.
