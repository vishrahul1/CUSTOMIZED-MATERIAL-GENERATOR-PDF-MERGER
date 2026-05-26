from pypdf import PdfReader, PdfWriter

class IndexManager:
    def combine_index_and_content(self, index_pdf_path, content_pdf_path, output_pdf_path):
        writer = PdfWriter()
        writer.append(index_pdf_path)
        writer.append(content_pdf_path)

        with open(output_pdf_path, 'wb') as f_out:
            writer.write(f_out)

        print(f"[IndexManager] Combined index + content and saved to {output_pdf_path}")
