from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor
from io import BytesIO


class PDFMerger:

    def merge_pdfs_with_index_tracking(self, file_entries, collected_paths, output_path):
        merger = PdfWriter()
        index_data = []
        current_page = 1

        for entry, pdf_path in zip(file_entries, collected_paths):
            reader = PdfReader(pdf_path)
            num_pages = len(reader.pages)

            index_data.append({
                'chapter_name': entry['chapter_name'],
                'topic_name':   entry['topic_name'],
                'page_number':  current_page,
            })

            merger.append(reader)
            print(f"[PDFMerger] Added {pdf_path} ({num_pages} pages)")
            current_page += num_pages

        with open(output_path, 'wb') as f_out:
            merger.write(f_out)

        print(f"[PDFMerger] Saved merged PDF → {output_path}")
        return index_data

    def add_header_footer(self, input_pdf_path, output_pdf_path, header_text, style=None):
        """
        Overlays header and footer on every page.

        style dict keys (all optional — defaults match the original hardcoded values):
            header_x, header_y, header_font, header_size, header_color
            footer_x, footer_y, footer_font, footer_size, footer_color
        """
        if style is None:
            style = {}

        h_x     = style.get('header_x',     502)
        h_y     = style.get('header_y',     796)
        h_font  = style.get('header_font',  'Times-Bold')
        h_size  = style.get('header_size',  10)
        h_color = style.get('header_color', '#000000')

        f_x     = style.get('footer_x',     528)
        f_y     = style.get('footer_y',     37)
        f_font  = style.get('footer_font',  'Times-Bold')
        f_size  = style.get('footer_size',  12)
        f_color = style.get('footer_color', '#ffffff')

        original = PdfReader(input_pdf_path)
        writer   = PdfWriter()

        for page_num, page in enumerate(original.pages):
            packet = BytesIO()
            can = canvas.Canvas(packet, pagesize=A4)

            can.setFont(h_font, h_size)
            can.setFillColor(HexColor(h_color))
            can.drawString(h_x, h_y, header_text)

            can.setFont(f_font, f_size)
            can.setFillColor(HexColor(f_color))
            can.drawString(f_x, f_y, str(page_num + 1))

            can.save()
            packet.seek(0)

            overlay = PdfReader(packet)
            page.merge_page(overlay.pages[0])
            writer.add_page(page)

        with open(output_pdf_path, 'wb') as f_out:
            writer.write(f_out)

        print(f"[PDFMerger] Header/footer added → {output_pdf_path}")
