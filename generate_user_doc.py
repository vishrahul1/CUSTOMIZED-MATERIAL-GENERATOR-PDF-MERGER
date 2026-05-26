"""
Run once to generate the user guide PDF.
Usage:  python generate_user_doc.py
Output: User Guide - PDF Merge Dashboard.pdf
"""

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak
)
from reportlab.platypus.flowables import KeepTogether

W, H = A4
OUT  = "User Guide - PDF Merge Dashboard.pdf"

# ── Colour palette ────────────────────────────────────────────────────────────
NAVY    = colors.HexColor("#2c3e50")
BLUE    = colors.HexColor("#3498db")
GREEN   = colors.HexColor("#27ae60")
LIGHT   = colors.HexColor("#f0f2f5")
BORDER  = colors.HexColor("#dde1e7")
WHITE   = colors.white
RED     = colors.HexColor("#e74c3c")

# ── Styles ────────────────────────────────────────────────────────────────────
base = getSampleStyleSheet()

def style(name, parent="Normal", **kw):
    s = ParagraphStyle(name, parent=base[parent], **kw)
    return s

H1 = style("H1", "Heading1",
           fontSize=18, textColor=NAVY, spaceAfter=4*mm,
           spaceBefore=6*mm, fontName="Helvetica-Bold")

H2 = style("H2", "Heading2",
           fontSize=13, textColor=BLUE, spaceAfter=3*mm,
           spaceBefore=5*mm, fontName="Helvetica-Bold",
           borderPad=2, leftIndent=0)

BODY = style("BODY", fontSize=10, leading=16, spaceAfter=2*mm,
             fontName="Helvetica", textColor=colors.HexColor("#2d2d2d"))

NOTE = style("NOTE", fontSize=9, leading=14, fontName="Helvetica-Oblique",
             textColor=colors.HexColor("#555555"), leftIndent=6*mm,
             spaceAfter=2*mm)

STEP = style("STEP", fontSize=10, leading=16, fontName="Helvetica-Bold",
             textColor=NAVY, leftIndent=4*mm, spaceAfter=1*mm)

STEP_BODY = style("STEP_BODY", fontSize=10, leading=15,
                  fontName="Helvetica", leftIndent=8*mm, spaceAfter=1*mm,
                  textColor=colors.HexColor("#2d2d2d"))

CODE = style("CODE", fontSize=9, leading=13, fontName="Courier",
             textColor=colors.HexColor("#c0392b"),
             backColor=colors.HexColor("#fdf2f2"),
             leftIndent=6*mm, spaceAfter=2*mm)

TIP  = style("TIP", fontSize=9, leading=13, fontName="Helvetica",
             textColor=colors.HexColor("#1a5c2e"),
             backColor=colors.HexColor("#eafaf1"),
             leftIndent=6*mm, spaceAfter=2*mm, borderPad=4)

# ── Table helper ──────────────────────────────────────────────────────────────
def make_table(header_row, data_rows, col_widths):
    rows = [header_row] + data_rows
    tbl  = Table(rows, colWidths=col_widths, repeatRows=1)
    style_cmds = [
        ("BACKGROUND",    (0, 0), (-1, 0),  NAVY),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  WHITE),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0),  10),
        ("TOPPADDING",    (0, 0), (-1, 0),  6),
        ("BOTTOMPADDING", (0, 0), (-1, 0),  6),
        ("FONTNAME",      (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE",      (0, 1), (-1, -1), 9),
        ("TOPPADDING",    (0, 1), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 5),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [WHITE, LIGHT]),
        ("GRID",          (0, 0), (-1, -1), 0.4, BORDER),
        ("LINEBELOW",     (0, 0), (-1, 0),  1.2, NAVY),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
    ]
    tbl.setStyle(TableStyle(style_cmds))
    return tbl

# ── Cover page ────────────────────────────────────────────────────────────────
def cover_page(canvas, doc):
    canvas.saveState()
    # Navy header band
    canvas.setFillColor(NAVY)
    canvas.rect(0, H - 80*mm, W, 80*mm, fill=True, stroke=False)
    # Title
    canvas.setFillColor(WHITE)
    canvas.setFont("Helvetica-Bold", 28)
    canvas.drawCentredString(W/2, H - 38*mm, "PDF Merge Dashboard")
    canvas.setFont("Helvetica", 14)
    canvas.drawCentredString(W/2, H - 52*mm, "User Guide")
    # Subtitle band
    canvas.setFillColor(BLUE)
    canvas.rect(0, H - 92*mm, W, 12*mm, fill=True, stroke=False)
    canvas.setFillColor(WHITE)
    canvas.setFont("Helvetica", 10)
    canvas.drawCentredString(W/2, H - 87*mm, "Educational PDF Generator  ·  Simple Step-by-Step Guide")
    # Footer
    canvas.setFillColor(LIGHT)
    canvas.rect(0, 0, W, 18*mm, fill=True, stroke=False)
    canvas.setFillColor(NAVY)
    canvas.setFont("Helvetica", 8)
    canvas.drawCentredString(W/2, 7*mm, "PDF Merge Dashboard  |  Educational PDF Generator")
    canvas.restoreState()

# ── Running header / footer on body pages ─────────────────────────────────────
def body_page(canvas, doc):
    canvas.saveState()
    # Top line
    canvas.setStrokeColor(BORDER)
    canvas.setLineWidth(0.5)
    canvas.line(15*mm, H - 12*mm, W - 15*mm, H - 12*mm)
    canvas.setFillColor(NAVY)
    canvas.setFont("Helvetica-Bold", 8)
    canvas.drawString(15*mm, H - 10*mm, "PDF Merge Dashboard  —  User Guide")
    # Bottom
    canvas.line(15*mm, 12*mm, W - 15*mm, 12*mm)
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#888888"))
    canvas.drawRightString(W - 15*mm, 8*mm, f"Page {doc.page}")
    canvas.restoreState()

# ── Document content ──────────────────────────────────────────────────────────
def build():
    doc = SimpleDocTemplate(
        OUT, pagesize=A4,
        leftMargin=18*mm, rightMargin=18*mm,
        topMargin=20*mm, bottomMargin=20*mm,
    )

    story = []
    P = Paragraph
    SP = lambda n=4: Spacer(1, n*mm)
    HR = lambda: HRFlowable(width="100%", thickness=0.5, color=BORDER,
                             spaceAfter=3*mm, spaceBefore=3*mm)

    # ── COVER (blank page driven by onFirstPage) ──────────────────────────────
    story.append(PageBreak())       # triggers cover_page callback

    # ── SECTION 1: What does this app do? ─────────────────────────────────────
    story.append(P("What Does This App Do?", H1))
    story.append(HR())
    story.append(P(
        "PDF Merge Dashboard combines multiple PDF files into one complete document. "
        "It automatically adds a <b>class name header</b> on every page, "
        "<b>page numbers</b> in the footer, and a <b>table of contents</b> at the start — "
        "all driven by a simple Excel file you prepare.", BODY))
    story.append(SP())

    # What you get
    story.append(P("What the final PDF contains:", H2))
    rows = [
        ["Part", "Description"],
        ["Table of Contents", "Lists every chapter and topic with page numbers"],
        ["Merged Content",    "All your PDFs combined in Excel order"],
        ["Header",            "Class / grade name printed on every page"],
        ["Footer",            "Page number printed on every page"],
    ]
    story.append(make_table(
        rows[0], rows[1:],
        [60*mm, 105*mm]))
    story.append(SP(6))

    # ── SECTION 2: How to launch ───────────────────────────────────────────────
    story.append(P("How to Launch the App", H1))
    story.append(HR())
    story.append(P(
        "Double-click <b>PDF Merge Dashboard.exe</b>. "
        "The app opens maximised and is ready to use — "
        "no installation or internet connection required.", BODY))
    story.append(SP(6))

    # ── SECTION 3: Step-by-step ────────────────────────────────────────────────
    story.append(P("Step-by-Step Usage", H1))
    story.append(HR())

    # Step 1
    story.append(KeepTogether([
        P("Step 1 — Select Files", H2),
        P("Fill in the three file paths at the top of the window:", BODY),
        P("① Excel File (.xlsx)", STEP),
        P("Your configuration file listing all PDFs in order. See the Excel Format section below.", STEP_BODY),
        P("② Source PDFs Folder", STEP),
        P("The folder on your computer that contains all the PDF files listed in the Excel.", STEP_BODY),
        P("③ Output PDF Path", STEP),
        P("Where the final combined PDF will be saved. Click Save As and type a filename.", STEP_BODY),
        P("💡  Use the Browse buttons — you do not need to type the paths manually.", TIP),
    ]))
    story.append(SP(4))

    # Step 2
    story.append(KeepTogether([
        P("Step 2 — Preview &amp; Adjust Header / Footer  (Optional)", H2),
        P("After selecting the source folder, click the purple <b>Preview</b> button.", BODY),
        P("A full-page view of your first PDF opens. You will see:", BODY),
        P("● Blue crosshair  =  Header position  (class name)", STEP_BODY),
        P("● Red crosshair   =  Footer position  (page number)", STEP_BODY),
        SP(2),
        P("To move the header or footer, change the <b>X</b> and <b>Y</b> values "
          "in the left panel. The crosshairs update instantly on the page.", BODY),
        P("You can also change the <b>font</b>, <b>size</b> and <b>colour</b> for each element.", BODY),
        P("💡  Click  ✓ Close &amp; Apply  when you are happy. Your settings are saved.", TIP),
    ]))
    story.append(SP(4))

    # Step 3
    story.append(KeepTogether([
        P("Step 3 — Start the Merge", H2),
        P('Click the green <b>▶ START MERGE</b> button.', BODY),
        P("The Process Log at the bottom shows each step in real time:", BODY),
        P("● Green lines = steps completed successfully", STEP_BODY),
        P("● Red lines   = errors that stopped the process", STEP_BODY),
        P("When finished you will see:", BODY),
        P("✅  Done!  Final PDF → [your chosen path]   Total pages: XX", CODE),
        P("💡  The app also saves intermediate files in the same output folder for troubleshooting.", TIP),
    ]))
    story.append(SP(6))

    # ── SECTION 4: Excel Format ────────────────────────────────────────────────
    story.append(P("Excel File Format", H1))
    story.append(HR())
    story.append(P(
        "Your Excel file (<b>.xlsx</b>) must have exactly these four column headings "
        "in the first row. Each row is one PDF file:", BODY))
    story.append(SP(2))

    rows = [
        ["Column Name", "What to Enter", "Example"],
        ["File_Name",    "Exact filename of the PDF (with .pdf extension)",   "chapter1.pdf"],
        ["Chapter_Name", "Chapter heading shown in the table of contents",    "Chapter 1"],
        ["Topic_Name",   "Topic heading shown in the table of contents",      "Motion and Time"],
        ["Class",        "Class or grade name — printed as header on pages",  "Class 8A"],
    ]
    story.append(make_table(rows[0], rows[1:], [38*mm, 80*mm, 40*mm]))
    story.append(SP(3))
    story.append(P(
        "⚠  The column names are case-sensitive. "
        "File_Name must match the actual PDF filename exactly, including uppercase and lowercase letters.", NOTE))
    story.append(SP(4))

    # ── SECTION 5: Coordinate system ──────────────────────────────────────────
    story.append(P("Understanding X / Y Coordinates", H1))
    story.append(HR())
    story.append(P(
        "The position of the header and footer is controlled by X and Y values. "
        "These follow the same system used inside PDF files:", BODY))
    story.append(SP(2))

    rows = [
        ["Setting", "Meaning"],
        ["Origin (0, 0)",   "Bottom-LEFT corner of the page"],
        ["X increases →",   "Moving right across the page  (max 595 for A4)"],
        ["Y increases ↑",   "Moving up the page  (max 842 for A4)"],
        ["Header default",  "X = 502,  Y = 796  (top-right area)"],
        ["Footer default",  "X = 528,  Y = 37   (bottom-right area)"],
    ]
    story.append(make_table(rows[0], rows[1:], [55*mm, 110*mm]))
    story.append(SP(2))
    story.append(P(
        "💡  Use the Preview window to drag the crosshairs visually — "
        "you do not need to calculate coordinates manually.", TIP))
    story.append(SP(6))

    # ── SECTION 6: Requirements ────────────────────────────────────────────────
    story.append(P("Requirements &amp; Limitations", H1))
    story.append(HR())

    rows = [
        ["Requirement",         "Detail"],
        ["PDF page size",       "All PDFs must be A4 (210 × 297 mm). The app checks this automatically."],
        ["PDF filenames",       "Must match the File_Name column in Excel exactly (case-sensitive)."],
        ["Excel format",        "Must be .xlsx format with the four required columns."],
        ["Output folder",       "Must exist on your computer and be writable."],
        ["Microsoft Word",      "Not required — the table of contents is generated internally."],
    ]
    story.append(make_table(rows[0], rows[1:], [45*mm, 120*mm]))
    story.append(SP(6))

    # ── SECTION 7: Troubleshooting ────────────────────────────────────────────
    story.append(P("Troubleshooting", H1))
    story.append(HR())

    rows = [
        ["Error / Problem",             "What to Do"],
        ["Non-A4 pages found",
         "Open the listed PDFs in any PDF editor and resize/export them as A4."],
        ["File missing error",
         "Check that the filename in the Excel File_Name column matches the actual PDF filename exactly."],
        ["Header or footer not visible",
         "If you set the colour to white, text is invisible on a white page. Change the colour."],
        ["Preview shows no PDF",
         "Select a Source PDFs Folder first — the preview uses the first PDF it finds there."],
        ["Process log is empty",
         "Click Start Merge — the log only fills when a merge is running."],
        ["App does not open",
         "Run the .exe directly from your local drive, not from a network or USB drive."],
    ]
    story.append(make_table(rows[0], rows[1:], [52*mm, 113*mm]))
    story.append(SP(6))

    # ── SECTION 8: Output files ────────────────────────────────────────────────
    story.append(KeepTogether([
        P("Output Files Explained", H1),
        HR(),
        P("After a successful merge, the output folder contains:", BODY),
        SP(2),
    ]))

    rows = [
        ["File",                      "What it is"],
        ["final_combined.pdf",        "Your final document — this is the one to use"],
        ["final_index.pdf",           "The table of contents page only"],
        ["merged_with_header.pdf",    "All PDFs merged with header and footer added"],
        ["merged_content.pdf",        "All PDFs merged (no header/footer yet)"],
    ]
    story.append(make_table(rows[0], rows[1:], [62*mm, 103*mm]))
    story.append(SP(3))
    story.append(P(
        "💡  The intermediate files are kept so you can inspect each stage if something looks wrong.", TIP))

    # ── Build ─────────────────────────────────────────────────────────────────
    doc.build(story,
              onFirstPage=cover_page,
              onLaterPages=body_page)
    print(f"Done! User guide saved: {OUT}")


if __name__ == "__main__":
    build()
