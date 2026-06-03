"""
word_convert.py
---------------
Convert a DOCX to PDF using a PRIVATE, invisible Microsoft Word instance via
win32com — so the user's already-open Word session is never touched or closed.

Why not docx2pdf:
  • docx2pdf uses win32com.client.Dispatch("Word.Application"), which ATTACHES
    to the user's running Word and then calls Quit() — closing their open
    documents.
  • docx2pdf also imports tqdm, whose progress bar writes to sys.stderr. In a
    no-console / windowed build sys.stderr is None, so it crashes with
    "'NoneType' object has no attribute 'write'".

This module avoids both: DispatchEx() spawns a brand-new private Word process,
we keep it invisible, and only that instance is quit. The conversion engine is
identical to docx2pdf — Word's own "Save As PDF" (WdSaveFormat = 17).
"""

import os

WD_FORMAT_PDF = 17  # WdSaveFormat.wdFormatPDF


def convert_docx_to_pdf(docx_path, pdf_path):
    """
    Convert docx_path -> pdf_path with a private, invisible Word instance.

    Safe to call from the GUI's daemon thread: it initializes the thread's COM
    apartment itself. Raises on failure (caught by the pipeline's try/except).
    """
    import pythoncom
    import win32com.client

    docx_path = os.path.abspath(docx_path)
    pdf_path  = os.path.abspath(pdf_path)

    # The pipeline runs in a daemon thread; COM needs the apartment initialized.
    try:
        pythoncom.CoInitialize()
        com_inited = True
    except Exception:
        com_inited = False

    word = None
    doc  = None
    try:
        # DispatchEx => a brand-new, PRIVATE Word process (not the user's).
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        try:
            word.DisplayAlerts = 0  # wdAlertsNone — never block on a dialog
        except Exception:
            pass

        doc = word.Documents.Open(docx_path, ReadOnly=True)
        doc.SaveAs(pdf_path, FileFormat=WD_FORMAT_PDF)
        doc.Close(False)
        doc = None
    finally:
        if doc is not None:
            try:
                doc.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()            # quits ONLY this private instance
            except Exception:
                pass
        if com_inited:
            pythoncom.CoUninitialize()

    return pdf_path
