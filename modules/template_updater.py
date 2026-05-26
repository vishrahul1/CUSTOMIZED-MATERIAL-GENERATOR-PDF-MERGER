from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle


class TemplateUpdater:
    def __init__(self, template_path=None):
        pass  # template_path kept for API compatibility but no longer used

    def generate_index_pdf(self, index_data, output_pdf_path):
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
        print(f"[TemplateUpdater] Generated index PDF at {output_pdf_path}")
