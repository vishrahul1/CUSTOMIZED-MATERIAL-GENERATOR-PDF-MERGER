from modules.input_handler import InputHandler
from modules.file_collector import FileCollector
from modules.pdf_merger import PDFMerger
from modules.template_updater import TemplateUpdater
from modules.index_manager import IndexManager
from pypdf import PdfReader
import os


def run_main_pipeline(excel_path, source_folder, output_pdf_path,
                      header_footer_style=None, update_status=print):
    try:
        base_output             = os.path.dirname(output_pdf_path)
        merged_pdf_path         = os.path.join(base_output, "merged_content.pdf")
        merged_with_header_path = os.path.join(base_output, "merged_with_header.pdf")
        final_index_pdf_path    = os.path.join(base_output, "final_index.pdf")

        update_status("🔄 Starting PDF merge process...")

        # Step 1: Read Excel
        update_status("📥 Reading Excel...")
        handler = InputHandler(excel_path)
        handler.read_excel()
        handler.validate_columns()
        file_entries = handler.extract_file_list()
        update_status(f"    ✓ Found {len(file_entries)} entries.")

        # Step 2: Collect files (recursive search through all subfolders)
        update_status("📂 Searching for PDF files (including subfolders)...")
        collector = FileCollector(source_folder)
        collected_paths, duplicates = collector.collect_files(file_entries)
        update_status(f"    ✓ {len(collected_paths)} file(s) located.")

        if duplicates:
            update_status(
                f"⚠️  Duplicate filenames found — using the first match for each:\n" +
                "\n".join(duplicates)
            )

        # Step 3: Page size check (21 × 27.8 cm) — warning only, merge continues
        update_status("📐 Checking page sizes (21 × 27.8 cm)...")
        violations = collector.check_page_sizes(collected_paths)
        if violations:
            bad_files = sorted(set(v['file'] for v in violations))
            lines = [
                f"  - {v['file']}  page {v['page']}: {v['size']}"
                for v in violations
            ]
            update_status(
                f"⚠️  Page size warning — {len(bad_files)} file(s) have pages "
                f"not matching 21 × 27.8 cm:\n" +
                "\n".join(lines) +
                "\n    Continuing merge..."
            )
        else:
            update_status(f"    ✓ All {len(collected_paths)} PDFs are 21 × 27.8 cm.")

        # Step 4: Merge PDFs and track index
        update_status("📑 Merging PDFs...")
        merger = PDFMerger()
        index_data = merger.merge_pdfs_with_index_tracking(
            file_entries, collected_paths, merged_pdf_path
        )
        update_status(f"    ✓ Merged {len(file_entries)} PDFs.")

        # Step 5: Add headers / footers
        class_name = file_entries[0]['class_name']
        update_status(f"🔤 Adding headers and footers (class: {class_name})...")
        merger.add_header_footer(
            merged_pdf_path,
            merged_with_header_path,
            header_text=class_name,
            style=header_footer_style,
        )

        # Step 6: Generate index PDF with ReportLab
        update_status("📄 Generating table of contents PDF...")
        TemplateUpdater().generate_index_pdf(index_data, final_index_pdf_path)

        # Step 7: Combine index + content
        update_status("📎 Combining index and content into final PDF...")
        IndexManager().combine_index_and_content(
            final_index_pdf_path,
            merged_with_header_path,
            output_pdf_path,
        )

        final_pages = len(PdfReader(output_pdf_path).pages)
        update_status(
            f"\n✅ Done!  Final PDF → {output_pdf_path}\n"
            f"   Total pages: {final_pages}"
        )

    except Exception as e:
        update_status(f"\n❌ ERROR: {e}")


if __name__ == "__main__":
    run_main_pipeline(
        excel_path='./resources/excel.xlsx',
        source_folder='./resources/source_pdfs/',
        output_pdf_path='./output/final_combined.pdf',
    )
