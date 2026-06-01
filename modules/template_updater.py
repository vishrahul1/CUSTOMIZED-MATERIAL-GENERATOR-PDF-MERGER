import os

from docx2pdf import convert

from spectropy_index import generate_index


class TemplateUpdater:
    def __init__(self, template_path=None):
        pass  # template_path kept for API compatibility but no longer used

    def generate_index_pdf(self, index_data, output_pdf_path):
        """
        Build the SPECTROPY "TABLE OF CONTENTS" index using spectropy_index.py
        and convert it to PDF.

        index_data is the list of dicts produced by
        PDFMerger.merge_pdfs_with_index_tracking:
            {'chapter_name', 'topic_name', 'page_number'}

        spectropy_index.generate_index expects a list of tuples:
            (chapter_name, topic_text, page_number)
        so we just reshape the input here — spectropy_index.py is untouched.
        """
        # ---- reshape input: list of dicts -> list of tuples ----
        data = [
            (
                str(entry['chapter_name']),
                str(entry['topic_name']),
                entry['page_number'],
            )
            for entry in index_data
        ]

        # ---- generate the .docx next to the target .pdf ----
        docx_path = os.path.splitext(output_pdf_path)[0] + ".docx"
        generate_index(data, docx_path)
        print(f"[TemplateUpdater] Generated index DOCX at {docx_path}")

        # ---- convert .docx -> .pdf (Word COM via docx2pdf) ----
        # The GUI runs the pipeline in a daemon thread; Word COM needs the
        # thread's COM apartment initialized before use.
        try:
            import pythoncom
            pythoncom.CoInitialize()
            _com_inited = True
        except Exception:
            _com_inited = False

        try:
            convert(docx_path, output_pdf_path)
        finally:
            if _com_inited:
                pythoncom.CoUninitialize()

        print(f"[TemplateUpdater] Generated index PDF at {output_pdf_path}")
