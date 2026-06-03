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

    def insert_chapter_title_pages(self, content_pdf_path, titles_pdf_path,
                                   chapter_start_pages, output_path):
        """
        Insert chapter divider pages into an already-merged (and header/footer'd)
        content PDF, placing each divider immediately before its chapter's first
        content page.

        Parameters
        ----------
        content_pdf_path    : the merged content PDF (post header/footer).
        titles_pdf_path     : PDF with one divider page per chapter, in chapter
                              order — as produced by
                              chapter_titles.generate_chapter_titles_pdf.
        chapter_start_pages : list of 1-based content page numbers where each
                              chapter starts, in the SAME order as the divider
                              pages in titles_pdf_path.
        output_path         : where to write the combined PDF.

        Divider pages are inserted as-is (no header/footer) and are unnumbered,
        so the content's footer numbers and the index page numbers are unaffected.
        """
        content = PdfReader(content_pdf_path)
        titles  = PdfReader(titles_pdf_path)
        writer  = PdfWriter()

        # content start page (1-based) -> index of its divider page in titles
        start_to_title = {sp: i for i, sp in enumerate(chapter_start_pages)}

        inserted = 0
        for i, page in enumerate(content.pages):
            page_no = i + 1
            if page_no in start_to_title:
                writer.add_page(titles.pages[start_to_title[page_no]])
                inserted += 1
            writer.add_page(page)

        with open(output_path, 'wb') as f_out:
            writer.write(f_out)

        print(f"[PDFMerger] Inserted {inserted} chapter title page(s) → {output_path}")
