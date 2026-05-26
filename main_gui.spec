# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all, collect_data_files

# ── PyMuPDF — has native DLLs that must travel with the exe ──────────────────
fitz_d,   fitz_b,   fitz_h   = collect_all('fitz')
pymupdf_d, pymupdf_b, pymupdf_h = collect_all('pymupdf')

# ── ReportLab — ships a fonts/ directory it loads at runtime ─────────────────
rl_datas = collect_data_files('reportlab')

# ── Pillow — needs its imaging plugins ───────────────────────────────────────
pil_d, pil_b, pil_h = collect_all('PIL')

a = Analysis(
    ['main_gui.py'],
    pathex=[],
    binaries=fitz_b + pymupdf_b + pil_b,
    datas=[
        # Sample / template resources bundled inside the exe
        ('resources/excel.xlsx', 'resources'),

        # Library data files
        *fitz_d,
        *pymupdf_d,
        *rl_datas,
        *pil_d,
    ],
    hiddenimports=[
        # PyMuPDF
        'fitz', 'fitz.utils', 'fitz.table',
        'pymupdf',
        *fitz_h,
        *pymupdf_h,

        # Pillow
        'PIL', 'PIL.Image', 'PIL.ImageTk', 'PIL._imagingtk',
        *pil_h,

        # pandas + Excel stack
        'pandas',
        'pandas.io.excel._openpyxl',
        'openpyxl',
        'openpyxl.cell._writer',
        'et_xmlfile',

        # ReportLab
        'reportlab',
        'reportlab.graphics',
        'reportlab.platypus',
        'reportlab.platypus.tables',
        'reportlab.lib.styles',
        'reportlab.lib.pagesizes',
        'reportlab.lib.units',
        'reportlab.lib.colors',
        'reportlab.pdfgen.canvas',
        'reportlab.pdfbase.ttfonts',
        'reportlab.pdfbase.pdfmetrics',

        # Utilities
        'chardet',
        'lxml',
        'lxml.etree',
        'six',
        'pytz',
        'dateutil',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Packages no longer used — keep them out of the bundle
    excludes=['docx2pdf', 'docx', 'python-docx', 'win32com'],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='PDF Merge Dashboard v3',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    # Don't UPX-compress native MuPDF DLLs — it breaks them
    upx_exclude=['fitz*.dll', 'mupdf*.dll', '_fitz*.pyd'],
    runtime_tmpdir=None,
    console=False,                  # no black console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
