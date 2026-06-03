import os

from modules.word_convert import convert_docx_to_pdf

from spectropy_index import generate_index


class TemplateUpdater:
    def __init__(self, template_path=None):
        pass  # template_path kept for API compatibility but no longer used

    def generate_old_index_pdf(self, index_data, output_pdf_path):
        """
        Build the original ReportLab "TABLE OF CONTENTS" index PDF.

        Kept as an optional reference output (not merged into the final PDF).
        index_data is the list of dicts produced by
        PDFMerger.merge_pdfs_with_index_tracking.
        """
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.lib import colors
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
        )

        doc = SimpleDocTemplate(
            output_pdf_path,
            pagesize=A4,
            rightMargin=20 * mm,
            leftMargin=20 * mm,
            topMargin=20 * mm,
            bottomMargin=20 * mm,
        )

        styles = getSampleStyleSheet()
        story = []

        title_style = ParagraphStyle(
            'IndexTitle',
            parent=styles['Title'],
            fontSize=20,
            spaceAfter=8 * mm,
            textColor=colors.HexColor('#2c3e50'),
        )
        story.append(Paragraph("TABLE OF CONTENTS", title_style))
        story.append(Spacer(1, 4 * mm))

        col_widths = [55 * mm, 95 * mm, 20 * mm]
        data = [['Chapter', 'Topic', 'Page']]
        for entry in index_data:
            data.append([
                str(entry['chapter_name']),
                str(entry['topic_name']),
                str(entry['page_number']),
            ])

        tbl = Table(data, colWidths=col_widths, repeatRows=1)
        tbl.setStyle(TableStyle([
            ('BACKGROUND',    (0, 0), (-1, 0),  colors.HexColor('#2c3e50')),
            ('TEXTCOLOR',     (0, 0), (-1, 0),  colors.white),
            ('FONTNAME',      (0, 0), (-1, 0),  'Helvetica-Bold'),
            ('FONTSIZE',      (0, 0), (-1, 0),  11),
            ('TOPPADDING',    (0, 0), (-1, 0),  8),
            ('BOTTOMPADDING', (0, 0), (-1, 0),  8),
            ('FONTNAME',      (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE',      (0, 1), (-1, -1), 9),
            ('TOPPADDING',    (0, 1), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 5),
            ('ALIGN',         (2, 0), (2, -1),  'CENTER'),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
            ('ROWBACKGROUNDS',(0, 1), (-1, -1),
             [colors.white, colors.HexColor('#f0f2f5')]),
            ('GRID',          (0, 0), (-1, -1), 0.5, colors.HexColor('#bdc3c7')),
            ('LINEBELOW',     (0, 0), (-1, 0),  1.5, colors.HexColor('#2c3e50')),
        ]))
        story.append(tbl)

        doc.build(story)
        print(f"[TemplateUpdater] Generated OLD index PDF at {output_pdf_path}")

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

        # ---- convert .docx -> .pdf (private, invisible Word via win32com) ----
        # Uses a dedicated Word instance so the user's open Word is never closed,
        # and avoids docx2pdf's tqdm (which crashes in no-console builds).
        convert_docx_to_pdf(docx_path, output_pdf_path)

        print(f"[TemplateUpdater] Generated index PDF at {output_pdf_path}")
