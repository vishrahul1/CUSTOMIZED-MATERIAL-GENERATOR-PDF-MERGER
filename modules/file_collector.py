from pathlib import Path
from pypdf import PdfReader

# 21 × 27.8 cm in PDF points (1 pt = 1/72 inch = 0.03528 cm)
PAGE_W     = 595.28   # 21.0 cm
PAGE_H     = 787.97   # 27.8 cm
PAGE_LABEL = "21 × 27.8 cm"
TOLERANCE  = 2.0


class FileCollector:
    def __init__(self, source_folder):
        self.source_folder = source_folder

    def collect_files(self, file_entries):
        """
        Recursively scans source_folder for .pdf files at all depths.
        Builds a filename → [all matching paths] map in one pass.

        Duplicates are only reported for filenames that appear in file_entries
        (Excel-listed files), not for every PDF found in the folder.

        Returns (collected_paths, duplicate_warnings).
        """
        # Single pass: map every PDF name → list of all paths where it exists
        all_paths: dict[str, list[Path]] = {}
        for path in Path(self.source_folder).rglob("*.pdf"):
            all_paths.setdefault(path.name, []).append(path)

        collected:  list[str] = []
        missing:    list[str] = []
        duplicates: list[str] = []

        for entry in file_entries:
            fname = entry['file_name']
            if fname in all_paths:
                paths = all_paths[fname]
                collected.append(str(paths[0]))   # first match wins
                # Only warn if this Excel-listed file has multiple copies
                if len(paths) > 1:
                    ignored = "\n".join(f"    Ignored: {p}" for p in paths[1:])
                    duplicates.append(
                        f"  '{fname}'\n"
                        f"    Using  : {paths[0]}\n{ignored}"
                    )
            else:
                missing.append(fname)

        if missing:
            formatted = "\n".join(f"  - {f}" for f in missing)
            raise FileNotFoundError(
                f"[FileCollector] {len(missing)} file(s) not found "
                f"anywhere under '{self.source_folder}':\n{formatted}"
            )

        total = sum(len(v) for v in all_paths.values())
        print(f"[FileCollector] Collected {len(collected)} files "
              f"(scanned {total} PDFs across all subfolders).")
        return collected, duplicates

    def check_page_sizes(self, file_paths):
        """
        Returns a list of dicts for every page whose dimensions do not match
        PAGE_LABEL. Does NOT raise — caller decides severity.
        Each dict: {'file': str, 'page': int, 'size': str}
        """
        violations = []
        for file_path in file_paths:
            try:
                reader = PdfReader(file_path)
                for i, page in enumerate(reader.pages):
                    w = float(page.mediabox.width)
                    h = float(page.mediabox.height)
                    if abs(w - PAGE_W) > TOLERANCE or abs(h - PAGE_H) > TOLERANCE:
                        violations.append({
                            'file': file_path,
                            'page': i + 1,
                            'size': f"{w:.0f} × {h:.0f} pt",
                        })
            except Exception as e:
                violations.append({
                    'file': file_path,
                    'page': '?',
                    'size': f'Read error: {e}',
                })
        return violations
