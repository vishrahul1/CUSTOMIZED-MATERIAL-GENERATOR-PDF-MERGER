import tkinter as tk
import tkinter.font as tkFont
from tkinter import ttk, filedialog, messagebox, colorchooser
from tkinter.scrolledtext import ScrolledText
from main import run_main_pipeline
import threading
import os

# ── Palette ───────────────────────────────────────────────────────────────────
BG     = "#f4f6f9"
CARD   = "#ffffff"
ACCENT = "#2c3e50"
BLUE   = "#3498db"
GREEN  = "#27ae60"
MUTED  = "#95a5a6"
BORDER = "#dde1e7"
DARK   = "#1e272e"

# A4 in ReportLab points
A4_W, A4_H = 595, 842

PDF_FONTS = [
    "Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique",
    "Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic",
    "Courier", "Courier-Bold", "Courier-Oblique",
]

# Maps ReportLab PDF font names → (Windows system family, weight, slant)
# These are the closest screen equivalents; metrics won't be pixel-perfect but
# character proportions and ascender/descender ratios will match closely.
_FONT_MAP: dict[str, tuple[str, str, str]] = {
    "Helvetica":             ("Arial",           "normal", "roman"),
    "Helvetica-Bold":        ("Arial",           "bold",   "roman"),
    "Helvetica-Oblique":     ("Arial",           "normal", "italic"),
    "Helvetica-BoldOblique": ("Arial",           "bold",   "italic"),
    "Times-Roman":           ("Times New Roman", "normal", "roman"),
    "Times-Bold":            ("Times New Roman", "bold",   "roman"),
    "Times-Italic":          ("Times New Roman", "normal", "italic"),
    "Times-BoldItalic":      ("Times New Roman", "bold",   "italic"),
    "Courier":               ("Courier New",     "normal", "roman"),
    "Courier-Bold":          ("Courier New",     "bold",   "roman"),
    "Courier-Oblique":       ("Courier New",     "normal", "italic"),
}


# ─────────────────────────────────────────────────────────────────────────────
#  Preview window — shows actual PDF page with live crosshair overlay
# ─────────────────────────────────────────────────────────────────────────────
class PreviewWindow:
    """
    Opens a Toplevel with the first PDF from source_folder rendered at ~96 DPI.
    Header (blue) and footer (red) crosshairs update in real time as the user
    edits any coordinate / font / size / color control.

    All IntVar / StringVar objects are SHARED with MergeApp, so changes here
    automatically update the main window controls and vice-versa.
    """

    # Render at 96/72 ≈ 1.333× so 1 PDF point ≈ 1 screen pixel at 96 DPI
    SCALE = 96 / 72

    def __init__(self, app: "MergeApp", pdf_path: str):
        self.app    = app
        self.s      = self.SCALE
        # Actual page dimensions — overwritten by _render() from the real PDF.
        # Defaults to A4 so _build() has valid values even if rendering fails.
        self.page_w = float(A4_W)
        self.page_h = float(A4_H)
        self.dw     = int(A4_W * self.s)
        self.dh     = int(A4_H * self.s)
        self._traces: list[tuple] = []

        self.win = tk.Toplevel(app.root)
        self.win.title(f"Preview  ·  {os.path.basename(pdf_path)}")
        self.win.minsize(900, 600)
        self.win.configure(bg=BG)
        self.win.protocol("WM_DELETE_WINDOW", self._close)
        self.win.state("zoomed")   # open maximised

        self._bg_photo = self._render(pdf_path)
        self._build()
        self._draw()
        self._attach_traces()

    # ── PDF rendering ─────────────────────────────────────────────────────────
    def _render(self, path: str):
        try:
            import fitz
            from PIL import Image, ImageTk
            doc  = fitz.open(path)
            page = doc[0]

            # Read the REAL page dimensions in PDF points and update
            # dw/dh so the canvas, scrollregion, and coordinate mapping
            # all match the actual paper size (not just A4).
            self.page_w = page.rect.width
            self.page_h = page.rect.height
            self.dw     = int(self.page_w * self.s)
            self.dh     = int(self.page_h * self.s)

            pix = page.get_pixmap(matrix=fitz.Matrix(self.s, self.s))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()
            return ImageTk.PhotoImage(img)
        except Exception as e:
            print(f"[Preview] render error: {e}")
            return None

    # ── UI ────────────────────────────────────────────────────────────────────
    def _build(self):
        # ── Left panel: controls ─────────────────────────────────────────────
        left = tk.Frame(self.win, bg=CARD, width=270,
                        highlightbackground=BORDER, highlightthickness=1)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)

        tk.Label(left, text="Header & Footer", bg=CARD, fg=ACCENT,
                 font=("Segoe UI", 12, "bold")).pack(
            anchor="w", padx=14, pady=(16, 2))
        tk.Label(left, text="Changes apply instantly to the preview →",
                 bg=CARD, fg=MUTED, font=("Segoe UI", 8)).pack(
            anchor="w", padx=14, pady=(0, 8))

        ttk.Separator(left, orient="horizontal").pack(fill="x", padx=10)

        self._controls(left, "HEADER ——",
                       self.app.hdr_x, self.app.hdr_y,
                       self.app.hdr_font, self.app.hdr_size,
                       is_header=True)

        ttk.Separator(left, orient="horizontal").pack(fill="x", padx=10, pady=6)

        self._controls(left, "FOOTER ——",
                       self.app.ftr_x, self.app.ftr_y,
                       self.app.ftr_font, self.app.ftr_size,
                       is_header=False)

        # Legend
        leg = tk.Frame(left, bg=CARD)
        leg.pack(side="bottom", fill="x", padx=14, pady=(8, 0))
        for col, txt in [("#3498db", "● Header"), ("#e74c3c", "● Footer")]:
            tk.Label(leg, text=txt, bg=CARD, fg=col,
                     font=("Segoe UI", 8, "bold")).pack(side="left", padx=(0, 10))

        tk.Button(left, text="✓  Close & Apply",
                  command=self._close,
                  bg=GREEN, fg="white", relief="flat",
                  font=("Segoe UI", 10, "bold"),
                  padx=12, pady=9, cursor="hand2").pack(
            side="bottom", fill="x", padx=14, pady=14)

        # ── Right panel: scrollable canvas ───────────────────────────────────
        right = tk.Frame(self.win, bg="#4a4a4a")
        right.pack(side="left", fill="both", expand=True)

        vbar = tk.Scrollbar(right, orient="vertical")
        hbar = tk.Scrollbar(right, orient="horizontal")
        vbar.pack(side="right",  fill="y")
        hbar.pack(side="bottom", fill="x")

        self.canvas = tk.Canvas(
            right, bg="#4a4a4a",
            xscrollcommand=hbar.set,
            yscrollcommand=vbar.set,
            scrollregion=(0, 0, self.dw, self.dh),
            highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        vbar.config(command=self.canvas.yview)
        hbar.config(command=self.canvas.xview)

        # Page shadow
        self.canvas.create_rectangle(
            6, 6, self.dw + 4, self.dh + 4,
            fill="#222222", outline="", tags="bg")

        # PDF page (or blank white if rendering failed)
        if self._bg_photo:
            self.canvas.create_image(2, 2, anchor="nw",
                                     image=self._bg_photo, tags="bg")
        else:
            self.canvas.create_rectangle(
                2, 2, self.dw + 2, self.dh + 2,
                fill="white", outline="#cccccc", tags="bg")
            self.canvas.create_text(
                self.dw // 2, self.dh // 2,
                text="PDF preview unavailable\n(install pymupdf)",
                fill="#aaaaaa", font=("Segoe UI", 13), tags="bg")

        # Mouse-wheel scrolling
        self.canvas.bind(
            "<MouseWheel>",
            lambda e: self.canvas.yview_scroll(-(e.delta // 120), "units"))

    def _controls(self, parent, label, xv, yv, fontv, sizev, is_header):
        L = dict(bg=CARD, font=("Segoe UI", 9))

        tk.Label(parent, text=label, bg=CARD, fg=ACCENT,
                 font=("Segoe UI", 9, "bold")).pack(
            anchor="w", padx=14, pady=(10, 4))

        # Header text input (header only)
        if is_header:
            rt = tk.Frame(parent, bg=CARD)
            rt.pack(anchor="w", fill="x", padx=14, pady=(0, 4))
            tk.Label(rt, text="Text:", **L).pack(side="left")
            tk.Entry(rt, textvariable=self.app.preview_text,
                     relief="solid", bd=1,
                     font=("Segoe UI", 9), bg=CARD).pack(
                side="left", fill="x", expand=True, padx=(5, 0), ipady=3)

        # X / Y
        r1 = tk.Frame(parent, bg=CARD)
        r1.pack(anchor="w", padx=14, pady=2)
        tk.Label(r1, text="X:", **L).pack(side="left")
        self._spin(r1, xv, 0, 595).pack(side="left", padx=(3, 14))
        tk.Label(r1, text="Y:", **L).pack(side="left")
        self._spin(r1, yv, 0, 842).pack(side="left", padx=(3, 0))

        # Font / Size
        r2 = tk.Frame(parent, bg=CARD)
        r2.pack(anchor="w", padx=14, pady=2)
        tk.Label(r2, text="Font:", **L).pack(side="left")
        ttk.Combobox(r2, textvariable=fontv, values=PDF_FONTS,
                     state="readonly", width=14,
                     font=("Segoe UI", 9)).pack(side="left", padx=(3, 10))
        tk.Label(r2, text="Sz:", **L).pack(side="left")
        self._spin(r2, sizev, 6, 72, w=4).pack(side="left", padx=(3, 0))

        # Color
        attr        = "hdr_color"  if is_header else "ftr_color"
        swatch_attr = "_pw_hsw"    if is_header else "_pw_fsw"
        pick        = (lambda: self._pick(True)) if is_header \
               else  (lambda: self._pick(False))

        r3 = tk.Frame(parent, bg=CARD)
        r3.pack(anchor="w", padx=14, pady=(2, 4))
        tk.Label(r3, text="Color:", **L).pack(side="left")
        sw = tk.Label(r3, bg=getattr(self.app, attr),
                      width=3, height=1, relief="solid", bd=1)
        sw.pack(side="left", padx=(5, 5))
        setattr(self, swatch_attr, sw)
        tk.Button(r3, text="Pick…", command=pick,
                  relief="flat", bg="#ecf0f1", fg=ACCENT,
                  font=("Segoe UI", 8), padx=8, pady=2,
                  cursor="hand2").pack(side="left")

    def _spin(self, parent, var, lo, hi, w=5):
        return tk.Spinbox(
            parent, textvariable=var, from_=lo, to=hi, width=w,
            font=("Segoe UI", 9), relief="solid", bd=1,
            command=self._draw)

    # ── Color picker ──────────────────────────────────────────────────────────
    def _pick(self, is_header: bool):
        attr        = "hdr_color"  if is_header else "ftr_color"
        swatch_attr = "_pw_hsw"    if is_header else "_pw_fsw"
        main_sw     = "hdr_swatch" if is_header else "ftr_swatch"

        # parent=self.win keeps the dialog owned by the preview window,
        # so focus returns here (not to the main window) after it closes.
        result = colorchooser.askcolor(
            color=getattr(self.app, attr),
            title="Header Text Color" if is_header else "Footer Text Color",
            parent=self.win)

        if result and result[1]:
            setattr(self.app, attr, result[1])
            getattr(self, swatch_attr).config(bg=result[1])
            main = getattr(self.app, main_sw, None)
            if main:
                main.config(bg=result[1])
            self._draw()

        # Always bring preview back to front after the dialog closes
        self.win.lift()
        self.win.focus_force()

    # ── Font helpers ──────────────────────────────────────────────────────────
    @staticmethod
    def _resolve_font(pdf_name: str, size: int) -> tkFont.Font:
        """
        Convert a ReportLab PDF font name + size (pts) into a tkFont.Font
        using the closest Windows system font equivalent.
        Font size is in points — same unit as ReportLab — no DPI scaling needed.
        """
        family, weight, slant = _FONT_MAP.get(
            pdf_name, ("Arial", "normal", "roman"))
        return tkFont.Font(family=family, size=size,
                           weight=weight, slant=slant)

    # ── Overlay drawing ───────────────────────────────────────────────────────
    def _draw(self, *_):
        """
        Redraws header + footer crosshairs accurately.

        Coordinate accuracy:
          canvas_x = pdf_x * scale + PAGE_OFFSET
          canvas_y = (page_h - pdf_y) * scale + PAGE_OFFSET
          where PAGE_OFFSET = 2 (page image placed at canvas (2,2))

        Font accuracy:
          - PDF font → mapped Windows system font (same proportions)
          - Font size stays in POINTS (Tkinter and ReportLab share the unit)
          - ReportLab drawString anchors at the TEXT BASELINE; Tkinter anchor='sw'
            anchors at the BOTTOM of the bounding box (descent below baseline).
            Fix: shift anchor_y down by font.metrics('descent') so baseline aligns.
        """
        try:
            if not self.canvas.winfo_exists():
                return
            c = self.canvas
            c.delete("ov")

            s   = self.s
            OFF = 2   # page image offset on canvas (shadow)

            def px(rx, ry):
                # Uses self.page_h (actual page height from PDF, not hardcoded A4)
                return rx * s + OFF, (self.page_h - ry) * s + OFF

            # ── Header ───────────────────────────────────────────────────────
            hx, hy = px(self.app.hdr_x.get(), self.app.hdr_y.get())
            H_COL  = "#3498db"

            c.create_line(OFF, hy, self.dw + OFF, hy,
                          fill=H_COL, width=1, dash=(10, 6), tags="ov")
            c.create_line(hx, OFF, hx, self.dh + OFF,
                          fill=H_COL, width=1, dash=(10, 6), tags="ov")
            c.create_oval(hx - 5, hy - 5, hx + 5, hy + 5,
                          fill=H_COL, outline="white", width=2, tags="ov")

            hfsz    = max(6, self.app.hdr_size.get())
            h_font  = self._resolve_font(self.app.hdr_font.get(), hfsz)
            h_desc  = h_font.metrics("descent")   # pixels below baseline
            h_col   = self.app.hdr_color
            if h_col.lower() in ("#ffffff", "#fff"):
                h_col = H_COL
            # anchor='sw' bottom aligns at hy+h_desc → baseline lands exactly at hy
            c.create_text(hx, hy + h_desc,
                          text=self.app.preview_text.get() or "Header",
                          anchor="sw", font=h_font,
                          fill=h_col, tags="ov")

            # ── Footer ───────────────────────────────────────────────────────
            fx, fy = px(self.app.ftr_x.get(), self.app.ftr_y.get())
            F_COL  = "#e74c3c"

            c.create_line(OFF, fy, self.dw + OFF, fy,
                          fill=F_COL, width=1, dash=(10, 6), tags="ov")
            c.create_line(fx, OFF, fx, self.dh + OFF,
                          fill=F_COL, width=1, dash=(10, 6), tags="ov")
            c.create_oval(fx - 5, fy - 5, fx + 5, fy + 5,
                          fill=F_COL, outline="white", width=2, tags="ov")

            ffsz    = max(6, self.app.ftr_size.get())
            f_font  = self._resolve_font(self.app.ftr_font.get(), ffsz)
            f_desc  = f_font.metrics("descent")
            f_col   = self.app.ftr_color
            if f_col.lower() in ("#ffffff", "#fff"):
                f_col = F_COL
            c.create_text(fx, fy + f_desc,
                          text="1",
                          anchor="sw", font=f_font,
                          fill=f_col, tags="ov")

            # ── Info bar (pinned at top of visible canvas, not page) ──────────
            c.create_rectangle(OFF, OFF, self.dw + OFF, OFF + 22,
                               fill="#2c3e50", outline="", tags="ov")
            info = (f"  ● Header  X={self.app.hdr_x.get()}"
                    f"  Y={self.app.hdr_y.get()}      "
                    f"● Footer  X={self.app.ftr_x.get()}"
                    f"  Y={self.app.ftr_y.get()}")
            c.create_text(OFF + 6, OFF + 11, text=info,
                          anchor="w", font=("Consolas", 8),
                          fill="white", tags="ov")

        except Exception:
            pass

    # ── Traces ────────────────────────────────────────────────────────────────
    def _attach_traces(self):
        for v in (self.app.hdr_x, self.app.hdr_y,
                  self.app.hdr_font, self.app.hdr_size,
                  self.app.ftr_x, self.app.ftr_y,
                  self.app.ftr_font, self.app.ftr_size,
                  self.app.preview_text):
            tid = v.trace_add("write", self._draw)
            self._traces.append((v, tid))

    def _close(self):
        for v, tid in self._traces:
            try:
                v.trace_remove("write", tid)
            except Exception:
                pass
        self.win.destroy()


# ─────────────────────────────────────────────────────────────────────────────
#  Main application window
# ─────────────────────────────────────────────────────────────────────────────
class MergeApp:
    def __init__(self, root):
        self.root = root
        root.title("PDF Merge Dashboard")
        root.minsize(780, 660)
        root.configure(bg=BG)
        root.state("zoomed")   # open maximised

        # File paths
        self.excel_path    = tk.StringVar()
        self.source_folder = tk.StringVar()
        self.output_pdf    = tk.StringVar()

        # Header style (defaults match original hardcoded values)
        self.hdr_x     = tk.IntVar(value=502)
        self.hdr_y     = tk.IntVar(value=796)
        self.hdr_font  = tk.StringVar(value="Times-Bold")
        self.hdr_size  = tk.IntVar(value=10)
        self.hdr_color = "#000000"

        # Live header text shown in preview (populated from Excel when available)
        self.preview_text = tk.StringVar(value="Class Name")

        # Footer style
        self.ftr_x     = tk.IntVar(value=528)
        self.ftr_y     = tk.IntVar(value=37)
        self.ftr_font  = tk.StringVar(value="Times-Bold")
        self.ftr_size  = tk.IntVar(value=12)
        self.ftr_color = "#ffffff"

        # Log zoom state
        self._log_font_size = 9

        # Compute a scale factor from screen DPI so spacing/fonts feel right
        # on both 96 DPI (1×) and 144 DPI (1.5×) screens
        raw_dpi = root.winfo_fpixels('1i')
        self._sc  = max(1.0, raw_dpi / 96)   # scale multiplier

        self._configure_ttk()
        self._build_ui()

    # ── TTK theme ─────────────────────────────────────────────────────────────
    def _configure_ttk(self):
        s = ttk.Style(self.root)
        s.theme_use("clam")
        s.configure("TCombobox",
                    fieldbackground=CARD, background=CARD,
                    font=("Segoe UI", 9))
        s.configure("TSeparator", background=BORDER)

    # ── Layout ────────────────────────────────────────────────────────────────
    def _build_ui(self):
        sc = self._sc
        pad = int(16 * sc)

        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True, padx=pad, pady=int(12 * sc))

        # Title bar — height scales with DPI
        bar_h = int(54 * sc)
        hdr = tk.Frame(outer, bg=ACCENT, height=bar_h)
        hdr.pack(fill="x", pady=(0, int(14 * sc)))
        hdr.pack_propagate(False)

        tk.Label(hdr, text="  PDF Merge Dashboard",
                 bg=ACCENT, fg="white",
                 font=("Segoe UI", int(15 * sc), "bold")).pack(
            side="left", padx=int(10 * sc), pady=int(10 * sc))
        tk.Label(hdr, text="Educational PDF Generator",
                 bg=ACCENT, fg="#7f8c8d",
                 font=("Segoe UI", int(9 * sc))).pack(
            side="right", padx=int(14 * sc))

        self._build_file_card(outer)
        self._build_log_card(outer)

    # ── Card factory ──────────────────────────────────────────────────────────
    def _card(self, parent, heading, expand=False, zoom_controls=False):
        sc   = self._sc
        wrap = tk.Frame(parent, bg=BG)
        wrap.pack(fill="both" if expand else "x",
                  expand=expand, pady=(0, int(10 * sc)))

        hdr_row = tk.Frame(wrap, bg=BG)
        hdr_row.pack(fill="x", padx=4, pady=(0, int(3 * sc)))

        tk.Label(hdr_row, text=heading, bg=BG, fg=ACCENT,
                 font=("Segoe UI", int(8 * sc), "bold")).pack(side="left")

        if zoom_controls:
            tk.Button(hdr_row, text=" − ",
                      command=lambda: self._zoom_log(-1),
                      bg=BORDER, fg=ACCENT, relief="flat",
                      font=("Segoe UI", int(8 * sc), "bold"),
                      cursor="hand2").pack(side="right", padx=(2, 0))
            tk.Button(hdr_row, text=" + ",
                      command=lambda: self._zoom_log(1),
                      bg=BORDER, fg=ACCENT, relief="flat",
                      font=("Segoe UI", int(8 * sc), "bold"),
                      cursor="hand2").pack(side="right", padx=(0, 2))
            tk.Label(hdr_row, text="zoom:", bg=BG, fg=MUTED,
                     font=("Segoe UI", int(7 * sc))).pack(side="right", padx=4)

        body = tk.Frame(wrap, bg=CARD,
                        highlightbackground=BORDER, highlightthickness=1)
        body.pack(fill="both", expand=expand)
        return body

    def _btn(self, parent, text, cmd, bg=BLUE, fs=9):
        sc  = self._sc
        btn = tk.Button(parent, text=text, command=cmd,
                        bg=bg, fg="white", relief="flat",
                        font=("Segoe UI", int(fs * sc), "bold"),
                        padx=int(12 * sc), pady=int(4 * sc),
                        activebackground=self._dk(bg),
                        cursor="hand2")
        btn.bind("<Enter>", lambda e: btn.config(bg=self._dk(bg)))
        btn.bind("<Leave>", lambda e: btn.config(bg=bg))
        return btn

    @staticmethod
    def _dk(h):
        c = h.lstrip('#')
        return "#{:02x}{:02x}{:02x}".format(
            *[max(0, int(c[i:i+2], 16) - 28) for i in (0, 2, 4)])

    # ── FILE card ─────────────────────────────────────────────────────────────
    def _build_file_card(self, parent):
        sc   = self._sc
        pad  = int(16 * sc)
        card = self._card(parent, "SELECT FILES")

        rows = [
            ("Excel File (.xlsx):",  self.excel_path,    self.browse_excel,  "Browse"),
            ("Source PDFs Folder:",  self.source_folder, self.browse_folder,  "Browse"),
            ("Output PDF Path:",     self.output_pdf,    self.save_as_pdf,   "Save As"),
        ]
        for i, (lbl, var, cmd, btn) in enumerate(rows):
            tk.Label(card, text=lbl, bg=CARD, fg=MUTED,
                     font=("Segoe UI", int(8 * sc))).pack(
                anchor="w", padx=pad,
                pady=(int(10 * sc) if i == 0 else int(4 * sc), 0))
            row = tk.Frame(card, bg=CARD)
            row.pack(fill="x", padx=pad, pady=(int(2 * sc), 0))
            tk.Entry(row, textvariable=var, relief="solid", bd=1,
                     font=("Segoe UI", int(9 * sc)), bg=CARD).pack(
                side="left", fill="x", expand=True, ipady=int(4 * sc))
            self._btn(row, btn, cmd).pack(side="left", padx=(int(8 * sc), 0))

        # Preview button
        prev_row = tk.Frame(card, bg=CARD)
        prev_row.pack(fill="x", padx=pad, pady=(int(10 * sc), int(12 * sc)))
        self._btn(prev_row, "👁  Preview Header & Footer on PDF",
                  self._open_preview, bg="#8e44ad").pack(side="left")
        tk.Label(prev_row,
                 text="  Opens the first PDF listed in your Excel file with live crosshairs",
                 bg=CARD, fg=MUTED,
                 font=("Segoe UI", int(8 * sc))).pack(side="left")

    # ── LOG card ──────────────────────────────────────────────────────────────
    def _build_log_card(self, parent):
        sc  = self._sc
        pad = int(10 * sc)

        # Start button row
        btn_row = tk.Frame(parent, bg=BG)
        btn_row.pack(fill="x", pady=(0, int(10 * sc)))

        self.run_btn = tk.Button(
            btn_row, text="▶   START MERGE",
            command=self.start_process,
            bg=GREEN, fg="white", relief="flat",
            font=("Segoe UI", int(12 * sc), "bold"),
            padx=int(28 * sc), pady=int(10 * sc),
            activebackground=self._dk(GREEN),
            cursor="hand2")
        self.run_btn.pack(side="left")
        self.run_btn.bind("<Enter>",
            lambda e: self.run_btn.config(bg=self._dk(GREEN)))
        self.run_btn.bind("<Leave>",
            lambda e: self.run_btn.config(bg=GREEN))

        self.status_lbl = tk.Label(btn_row, text="", bg=BG,
                                    font=("Segoe UI", int(9 * sc)), fg=MUTED)
        self.status_lbl.pack(side="left", padx=int(14 * sc))

        # Log card with zoom controls in heading
        card = self._card(parent, "PROCESS LOG",
                          expand=True, zoom_controls=True)

        self.log_box = ScrolledText(
            card, font=("Consolas", self._log_font_size),
            bg=DARK, fg="#dfe6e9",
            insertbackground="white",
            relief="flat", height=10, wrap="word")
        self.log_box.pack(fill="both", expand=True, padx=pad, pady=pad)
        self.log_box.tag_config("err", foreground="#e74c3c")
        self.log_box.tag_config("ok",  foreground="#2ecc71")
        self.log_box.tag_config("dim", foreground="#7f8c8d")

        # Keyboard zoom: Ctrl+= / Ctrl+- and Ctrl+scroll
        self.root.bind("<Control-equal>",
                       lambda e: self._zoom_log(1))
        self.root.bind("<Control-minus>",
                       lambda e: self._zoom_log(-1))
        self.root.bind("<Control-plus>",
                       lambda e: self._zoom_log(1))
        self.log_box.bind("<Control-MouseWheel>",
                          lambda e: self._zoom_log(1 if e.delta > 0 else -1))

    def _zoom_log(self, delta: int):
        self._log_font_size = max(7, min(22, self._log_font_size + delta))
        self.log_box.config(font=("Consolas", self._log_font_size))

    # ── File pickers ──────────────────────────────────────────────────────────
    def browse_excel(self):
        p = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])
        if p:
            self.excel_path.set(p)

    def browse_folder(self):
        f = filedialog.askdirectory(title="Select Source PDFs Folder")
        if f:
            self.source_folder.set(f)

    def save_as_pdf(self):
        f = filedialog.asksaveasfilename(
            title="Save Output PDF As",
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")])
        if f:
            self.output_pdf.set(f)

    # ── Preview ───────────────────────────────────────────────────────────────
    def _open_preview(self):
        excel  = self.excel_path.get()
        folder = self.source_folder.get()

        if not excel or not os.path.isfile(excel):
            messagebox.showwarning("No Excel File",
                                   "Please select an Excel file first.")
            return
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning("No Folder",
                                   "Please select a Source PDFs Folder first.")
            return

        # Read the first File_Name entry from the Excel
        try:
            import pandas as pd
            df       = pd.read_excel(excel, usecols=["File_Name"])
            filename = str(df["File_Name"].dropna().iloc[0]).strip()
        except Exception as e:
            messagebox.showerror("Excel Read Error",
                                 f"Could not read Excel file:\n{e}")
            return

        # Find that file recursively under the source folder
        from pathlib import Path
        matches = list(Path(folder).rglob(filename))
        if not matches:
            messagebox.showwarning(
                "File Not Found",
                f"'{filename}' (first entry in Excel) was not found\n"
                f"anywhere under:\n{folder}")
            return

        PreviewWindow(self, str(matches[0]))

    # ── Pipeline ──────────────────────────────────────────────────────────────
    def _get_style(self):
        return {
            'header_x':     self.hdr_x.get(),
            'header_y':     self.hdr_y.get(),
            'header_font':  self.hdr_font.get(),
            'header_size':  self.hdr_size.get(),
            'header_color': self.hdr_color,
            'footer_x':     self.ftr_x.get(),
            'footer_y':     self.ftr_y.get(),
            'footer_font':  self.ftr_font.get(),
            'footer_size':  self.ftr_size.get(),
            'footer_color': self.ftr_color,
        }

    def update_status(self, message):
        """Worker-thread safe — schedules the insert on the main thread."""
        def _insert():
            if "❌" in message or "ERROR" in message:
                tag = "err"
            elif "✅" in message:
                tag = "ok"
            elif message.strip().startswith("    ✓"):
                tag = "dim"
            else:
                tag = ""
            self.log_box.insert(tk.END, message + "\n", tag)
            self.log_box.see(tk.END)
        self.root.after(0, _insert)

    def start_process(self):
        if not self.excel_path.get() or \
           not self.source_folder.get() or \
           not self.output_pdf.get():
            messagebox.showwarning("Missing Input",
                                   "Please fill in all three file paths first.")
            return

        self.log_box.delete(1.0, tk.END)
        self.run_btn.config(state="disabled", text="⏳  Running…")
        self.status_lbl.config(text="Processing…", fg=BLUE)

        def _worker():
            try:
                run_main_pipeline(
                    excel_path=self.excel_path.get(),
                    source_folder=self.source_folder.get(),
                    output_pdf_path=self.output_pdf.get(),
                    header_footer_style=self._get_style(),
                    update_status=self.update_status,
                )
            finally:
                self.root.after(0, self._done)

        threading.Thread(target=_worker, daemon=True).start()

    def _done(self):
        self.run_btn.config(state="normal", text="▶   START MERGE")
        self.status_lbl.config(text="", fg=MUTED)


if __name__ == "__main__":
    root = tk.Tk()
    MergeApp(root)
    root.mainloop()
