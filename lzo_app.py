#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Aplikacija za upravljanje i evidenciju Lične Zaštitne Opreme (LZO)
Aplikacija omogućava praćenje zaduženja opreme, automatizovani proračun isteka rokova,
napredno filtriranje, generisanje statistike, izvoz u Excel i PDF formatu, kao i uvoz podataka.
"""

import os
import sys
import json
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import pandas as pd
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as ReportLabImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


def get_pdf_unicode_font():
    """Registruje i vraća sistemski font sa podrškom za balkanske karaktere (Č, Ć, Š, Đ, Ž) za ReportLab."""
    if not REPORTLAB_AVAILABLE:
        return "Helvetica"
    
    font_candidates = [
        ("Arial", "C:\\Windows\\Fonts\\arial.ttf"),
        ("Arial", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
        ("Arial", "/Library/Fonts/Arial.ttf"),
        ("DejaVuSans", "DejaVuSans.ttf")
    ]
    
    for font_name, font_path in font_candidates:
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont(font_name, font_path))
                return font_name
            except Exception:
                continue
    return "Helvetica"


class RecordDialog(tk.Toplevel):
    """Dijalog prozor za unos novog ili izmjenu postojećeg zaduženja LZO."""

    def __init__(self, parent, title="Zaduženje LZO", record=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("520x680")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.result = None
        self.record = record or {}

        self.setup_ui()
        self.center_window()

    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f'{width}x{height}+{x}+{y}')

    def setup_ui(self):
        padding = {'padx': 12, 'pady': 6}

        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        lbl_title = ttk.Label(main_frame, text=self.title(), font=("Segoe UI", 12, "bold"))
        lbl_title.grid(row=0, column=0, columnspan=2, pady=(0, 15), sticky="w")

        fields = [
            ("Zaposleni (Ime i Prezime):", "zaposleni"),
            ("Radno mjesto:", "rm"),
            ("Organizaciona jedinica:", "org_jedinica"),
            ("Grad / Lokacija:", "grad"),
            ("Veličina odjeće:", "vel_odjeca"),
            ("Veličina obuće:", "vel_obuca"),
            ("Naziv LZO opreme:", "oprema"),
            ("Jedinica mjere:", "jm"),
            ("Normativ (kol.):", "normativ"),
            ("Rok trajanja (mjeseci):", "rok_mjeseci"),
            ("Izdata količina:", "izdata_kol"),
            ("Datum zaduženja (YYYY-MM-DD):", "datum_zaduzenja"),
            ("Napomena:", "napomena")
        ]

        self.entries = {}

        for idx, (label_text, key) in enumerate(fields, start=1):
            lbl = ttk.Label(main_frame, text=label_text, font=("Segoe UI", 9))
            lbl.grid(row=idx, column=0, sticky="w", **padding)

            entry = ttk.Entry(main_frame, width=32, font=("Segoe UI", 9))
            entry.grid(row=idx, column=1, sticky="ew", **padding)

            # Postavljanje podrazumijevanih ili postojećih vrijednosti
            default_val = self.record.get(key, "")
            if not default_val:
                if key == "jm": default_val = "KOM"
                elif key in ["normativ", "izdata_kol"]: default_val = "1"
                elif key == "rok_mjeseci": default_val = "12"
                elif key == "datum_zaduzenja": default_val = datetime.now().strftime("%Y-%m-%d")

            entry.insert(0, str(default_val))
            self.entries[key] = entry

        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=len(fields) + 1, column=0, columnspan=2, pady=(20, 0), sticky="e")

        btn_cancel = ttk.Button(btn_frame, text="Otkaži", command=self.destroy)
        btn_cancel.pack(side=tk.RIGHT, padx=5)

        btn_save = ttk.Button(btn_frame, text="Sačuvaj", command=self.on_save)
        btn_save.pack(side=tk.RIGHT, padx=5)

    def on_save(self):
        zaposleni = self.entries["zaposleni"].get().strip()
        oprema = self.entries["oprema"].get().strip()
        datum_zad = self.entries["datum_zaduzenja"].get().strip()

        if not zaposleni or not oprema or not datum_zad:
            messagebox.showwarning("Upozorenje", "Polja 'Zaposleni', 'Naziv opreme' i 'Datum zaduženja' su obavezna!", parent=self)
            return

        try:
            datetime.strptime(datum_zad, "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Greška", "Datum zaduženja mora biti u formatu YYYY-MM-DD (npr. 2026-09-17)!", parent=self)
            return

        try:
            normativ = int(self.entries["normativ"].get().strip() or 1)
            rok_mjeseci = int(self.entries["rok_mjeseci"].get().strip() or 12)
            izdata_kol = int(self.entries["izdata_kol"].get().strip() or 1)
        except ValueError:
            messagebox.showerror("Greška", "Normativ, rok trajanja i izdata količina moraju biti cijeli brojevi!", parent=self)
            return

        self.result = {
            "id": self.record.get("id"),
            "zaposleni": zaposleni,
            "rm": self.entries["rm"].get().strip(),
            "org_jedinica": self.entries["org_jedinica"].get().strip(),
            "grad": self.entries["grad"].get().strip(),
            "vel_odjeca": self.entries["vel_odjeca"].get().strip(),
            "vel_obuca": self.entries["vel_obuca"].get().strip(),
            "oprema": oprema,
            "jm": self.entries["jm"].get().strip() or "KOM",
            "normativ": normativ,
            "rok_mjeseci": rok_mjeseci,
            "izdata_kol": izdata_kol,
            "datum_zaduzenja": datum_zad,
            "napomena": self.entries["napomena"].get().strip()
        }
        self.destroy()


class LZOApp:
    """Glavna klasa desktop aplikacije za evidenciju LZO."""

    def __init__(self, root):
        self.root = root
        self.root.title("Sistem za Evidenciju i Praćenje Lične Zaštitne Opreme (LZO)")
        self.root.geometry("1380x820")
        self.root.minsize(1050, 650)

        self.data_file = "lzo_data.json"
        self.data = []
        self.current_filtered_data = []

        # Undo / Redo istorija
        self.undo_stack = []
        self.redo_stack = []

        self.setup_styles()
        self.load_data()

        self.create_menu()
        self.create_toolbar()
        self.create_main_layout()

        self.recalculate_and_refresh()
        self.bind_shortcuts()

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # Podešavanje boja
        self.bg_color = "#F8FAFC"
        self.card_bg = "#FFFFFF"
        self.primary_color = "#1E293B"

        self.root.configure(bg=self.bg_color)

        self.style.configure("Treeview", font=("Segoe UI", 9), rowheight=26, background="#FFFFFF", fieldbackground="#FFFFFF")
        self.style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#E2E8F0", foreground="#0F172A")
        self.style.map("Treeview", background=[('selected', '#3B82F6')], foreground=[('selected', '#FFFFFF')])

        self.style.configure("Card.TFrame", background=self.card_bg, relief="solid", borderwidth=1)
        self.style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"), background=self.bg_color, foreground="#0F172A")
        self.style.configure("SubHeader.TLabel", font=("Segoe UI", 10), background=self.bg_color, foreground="#64748B")

    def create_menu(self):
        menubar = tk.Menu(self.root)

        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Uvezi iz Excel-a...", command=self.import_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Izvezi u Excel...", command=self.export_excel_report)
        file_menu.add_command(label="Izvezi u PDF...", command=self.export_pdf_report)
        file_menu.add_separator()
        file_menu.add_command(label="Izlaz", command=self.root.quit)
        menubar.add_cascade(label="Datoteka", menu=file_menu)

        edit_menu = tk.Menu(menubar, tearoff=0)
        edit_menu.add_command(label="Novo zaduženje", accelerator="Ctrl+N", command=self.open_add_dialog)
        edit_menu.add_command(label="Izmijeni selektovano", command=self.open_edit_dialog)
        edit_menu.add_command(label="Obriši selektovano", accelerator="Delete", command=self.delete_selected)
        edit_menu.add_separator()
        edit_menu.add_command(label="Poništi (Undo)", accelerator="Ctrl+Z", command=self.undo)
        edit_menu.add_command(label="Ponovi (Redo)", accelerator="Ctrl+Y", command=self.redo)
        menubar.add_cascade(label="Uređivanje", menu=edit_menu)

        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="O aplikaciji", command=self.show_about)
        menubar.add_cascade(label="Pomoć", menu=help_menu)

        self.root.config(menu=menubar)

    def create_toolbar(self):
        tb = ttk.Frame(self.root, padding=(10, 8))
        tb.pack(fill=tk.X)

        btn_add = ttk.Button(tb, text="+ Novo zaduženje", command=self.open_add_dialog)
        btn_add.pack(side=tk.LEFT, padx=3)

        btn_edit = ttk.Button(tb, text="Izmijeni", command=self.open_edit_dialog)
        btn_edit.pack(side=tk.LEFT, padx=3)

        btn_del = ttk.Button(tb, text="Obriši", command=self.delete_selected)
        btn_del.pack(side=tk.LEFT, padx=3)

        ttk.Separator(tb, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8)

        btn_excel = ttk.Button(tb, text="Izvoz u Excel", command=self.export_excel_report)
        btn_excel.pack(side=tk.LEFT, padx=3)

        btn_pdf = ttk.Button(tb, text="Izvoz u PDF", command=self.export_pdf_report)
        btn_pdf.pack(side=tk.LEFT, padx=3)

        btn_import = ttk.Button(tb, text="Uvoz iz Excel-a", command=self.import_excel)
        btn_import.pack(side=tk.LEFT, padx=3)

        ttk.Separator(tb, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8)

        btn_undo = ttk.Button(tb, text="↺ Undo", command=self.undo)
        btn_undo.pack(side=tk.LEFT, padx=3)

        btn_redo = ttk.Button(tb, text="↻ Redo", command=self.redo)
        btn_redo.pack(side=tk.LEFT, padx=3)

    def create_main_layout(self):
        main_container = ttk.Frame(self.root, padding=10)
        main_container.pack(fill=tk.BOTH, expand=True)

        # Kartice sa statistikom na vrhu
        self.create_stats_cards(main_container)

        # Filter panel
        self.create_filter_panel(main_container)

        # Središnji dio: Tabela i Grafikon
        content_paned = ttk.PanedWindow(main_container, orient=tk.HORIZONTAL)
        content_paned.pack(fill=tk.BOTH, expand=True, pady=10)

        # Lijevo: Tabela
        table_frame = ttk.Frame(content_paned)
        content_paned.add(table_frame, weight=3)
        self.create_main_table(table_frame)

        # Desno: Grafikon
        chart_frame = ttk.LabelFrame(content_paned, text=" Statistički pregled statusa ", padding=10)
        content_paned.add(chart_frame, weight=1)
        self.create_chart_panel(chart_frame)

        # Statusna traka
        self.status_bar = ttk.Label(self.root, text="Spremno", relief=tk.SUNKEN, anchor=tk.W, font=("Segoe UI", 8), padding=4)
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)

    def create_stats_cards(self, parent):
        cards_frame = ttk.Frame(parent)
        cards_frame.pack(fill=tk.X, pady=(0, 10))

        self.card_total_var = tk.StringVar(value="0")
        self.card_valid_var = tk.StringVar(value="0")
        self.card_warning_var = tk.StringVar(value="0")
        self.card_expired_var = tk.StringVar(value="0")

        cards_data = [
            ("Ukupno zaduženja", self.card_total_var, "#1E293B", "#F1F5F9"),
            ("Važeće LZO", self.card_valid_var, "#15803D", "#DCFCE7"),
            ("Uskoro ističe (30 dana)", self.card_warning_var, "#A16207", "#FEF08A"),
            ("Isteklo LZO", self.card_expired_var, "#B91C1C", "#FECDD3")
        ]

        for idx, (title, var, fg_color, bg_color) in enumerate(cards_data):
            f = tk.Frame(cards_frame, bg=bg_color, highlightbackground="#CBD5E1", highlightthickness=1, padx=15, pady=10)
            f.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=4)

            lbl_title = tk.Label(f, text=title, font=("Segoe UI", 9, "bold"), fg=fg_color, bg=bg_color)
            lbl_title.pack(anchor="w")

            lbl_val = tk.Label(f, textvariable=var, font=("Segoe UI", 18, "bold"), fg=fg_color, bg=bg_color)
            lbl_val.pack(anchor="w")

    def create_filter_panel(self, parent):
        filter_frame = ttk.LabelFrame(parent, text=" Pretraga i filtriranje ", padding=8)
        filter_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(filter_frame, text="Pretraga (Radnik / Oprema / RM):").pack(side=tk.LEFT, padx=(5, 2))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.apply_filters())
        entry_search = ttk.Entry(filter_frame, textvariable=self.search_var, width=25)
        entry_search.pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(filter_frame, text="Status:").pack(side=tk.LEFT, padx=(5, 2))
        self.status_filter_var = tk.StringVar(value="SVE")
        cb_status = ttk.Combobox(filter_frame, textvariable=self.status_filter_var, values=["SVE", "VAŽEĆE", "USKORO", "ISTEKLO"], state="readonly", width=12)
        cb_status.pack(side=tk.LEFT, padx=(0, 15))
        cb_status.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        ttk.Label(filter_frame, text="Org. jedinica:").pack(side=tk.LEFT, padx=(5, 2))
        self.org_filter_var = tk.StringVar(value="SVE")
        self.cb_org = ttk.Combobox(filter_frame, textvariable=self.org_filter_var, state="readonly", width=18)
        self.cb_org.pack(side=tk.LEFT, padx=(0, 15))
        self.cb_org.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        btn_reset = ttk.Button(filter_frame, text="Poništi filtere", command=self.reset_filters)
        btn_reset.pack(side=tk.RIGHT, padx=5)

    def create_main_table(self, parent):
        columns = ("id", "zaposleni", "rm", "org_jedinica", "grad", "oprema", "jm", "normativ", "rok", "izdato", "datum_zad", "datum_ist", "preostalo", "status")
        
        self.tree = ttk.Treeview(parent, columns=columns, show="headings", selectmode="extended")

        headers = [
            ("id", "ID", 40),
            ("zaposleni", "Zaposleni", 140),
            ("rm", "Radno mjesto", 130),
            ("org_jedinica", "Org. jedinica", 120),
            ("grad", "Grad", 90),
            ("oprema", "Naziv LZO opreme", 150),
            ("jm", "J.M.", 45),
            ("normativ", "Nor.", 40),
            ("rok", "Rok(m)", 50),
            ("izdato", "Kol.", 40),
            ("datum_zad", "Zaduženo", 80),
            ("datum_ist", "Ističe", 80),
            ("preostalo", "Preostalo", 65),
            ("status", "Status", 75)
        ]

        for col, title, width in headers:
            self.tree.heading(col, text=title, command=lambda _col=col: self.sort_table_column(_col, False))
            self.tree.column(col, width=width, anchor=tk.CENTER if col in ["id", "jm", "normativ", "rok", "izdato", "datum_zad", "datum_ist", "preostalo", "status"] else tk.W)

        vsb = ttk.Scrollbar(parent, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(parent, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        parent.grid_rowconfigure(0, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        # Oznake boja za status redova
        self.tree.tag_configure("ISTEKLO", background="#FECDD3", foreground="#881337")
        self.tree.tag_configure("USKORO", background="#FEF08A", foreground="#713F12")
        self.tree.tag_configure("VAŽEĆE", background="#FFFFFF", foreground="#0F172A")

        self.tree.bind("<Double-1>", self.open_edit_dialog)

    def create_chart_panel(self, parent):
        self.fig, self.ax = plt.subplots(figsize=(4, 4), dpi=90)
        self.fig.patch.set_facecolor('#FFFFFF')
        
        self.canvas = FigureCanvasTkAgg(self.fig, master=parent)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    def bind_shortcuts(self):
        self.root.bind("<Control-n>", lambda e: self.open_add_dialog())
        self.root.bind("<Control-z>", lambda e: self.undo())
        self.root.bind("<Control-y>", lambda e: self.redo())
        self.root.bind("<Delete>", lambda e: self.delete_selected())

    def load_data(self):
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, "r", encoding="utf-8") as f:
                    self.data = json.load(f)
            except Exception as e:
                messagebox.showerror("Greška", f"Nije moguće učitati podatke iz datoteke: {e}")
                self.data = self.get_sample_data()
        else:
            self.data = self.get_sample_data()

    def save_data(self):
        try:
            with open(self.data_file, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Greška", f"Nije moguće sačuvati podatke: {e}")

    def get_sample_data(self):
        """Generiše ogledne podatke ukoliko lokalna baza ne postoji."""
        return [
            {
                "id": 1, "zaposleni": "Marko Petrović", "rm": "Električar", "org_jedinica": "Održavanje",
                "grad": "Podgorica", "vel_odjeca": "XL", "vel_obuca": "44", "oprema": "Zaštitne šlem štitnik",
                "jm": "KOM", "normativ": 1, "rok_mjeseci": 24, "izdata_kol": 1,
                "datum_zaduzenja": "2024-10-15", "napomena": "Ispravno"
            },
            {
                "id": 2, "zaposleni": "Nikola Đurović", "rm": "Monterski radnik", "org_jedinica": "Mreža",
                "grad": "Herceg Novi", "vel_odjeca": "L", "vel_obuca": "43", "oprema": "Radne cipele S3",
                "jm": "PAR", "normativ": 1, "rok_mjeseci": 12, "izdata_kol": 1,
                "datum_zaduzenja": "2025-09-20", "napomena": "Zamjena planirana"
            },
            {
                "id": 3, "zaposleni": "Petar Jovanović", "rm": "Laborant", "org_jedinica": "Laboratorija",
                "grad": "Podgorica", "vel_odjeca": "M", "vel_obuca": "42", "oprema": "Zaštitne rukavice hem.",
                "jm": "PAR", "normativ": 2, "rok_mjeseci": 6, "izdata_kol": 2,
                "datum_zaduzenja": "2026-04-10", "napomena": ""
            }
        ]

    def push_undo_state(self):
        """Snima trenutno stanje radi implementacije Undo opcije."""
        import copy
        self.undo_stack.append(copy.deepcopy(self.data))
        self.redo_stack.clear()

    def undo(self):
        if self.undo_stack:
            import copy
            self.redo_stack.append(copy.deepcopy(self.data))
            self.data = self.undo_stack.pop()
            self.recalculate_and_refresh()

    def redo(self):
        if self.redo_stack:
            import copy
            self.undo_stack.append(copy.deepcopy(self.data))
            self.data = self.redo_stack.pop()
            self.recalculate_and_refresh()

    def calculate_status_and_days(self, record):
        """Kalkuliše datum isticanja, preostale dane i status zaduženja LZO."""
        try:
            d_zad = datetime.strptime(record["datum_zaduzenja"], "%Y-%m-%d")
            rok = int(record.get("rok_mjeseci", 12))
            
            # Dodavanje mjeseci na datum
            year = d_zad.year + (d_zad.month + rok - 1) // 12
            month = (d_zad.month + rok - 1) % 12 + 1
            day = min(d_zad.day, [31, 29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1])
            d_ist = datetime(year, month, day)

            danas = datetime.now()
            preostalo = (d_ist - danas).days

            record["datum_isticanja"] = d_ist.strftime("%Y-%m-%d")
            record["preostalo_dana"] = preostalo

            if preostalo < 0:
                record["status"] = "ISTEKLO"
            elif preostalo <= 30:
                record["status"] = "USKORO"
            else:
                record["status"] = "VAŽEĆE"
        except Exception:
            record["datum_isticanja"] = "-"
            record["preostalo_dana"] = 0
            record["status"] = "NEPOZNATO"

    def recalculate_and_refresh(self):
        for rec in self.data:
            self.calculate_status_and_days(rec)

        # Ažuriranje padajuće liste organizacionih jedinica
        orgs = sorted(list(set(r.get("org_jedinica", "") for r in self.data if r.get("org_jedinica"))))
        self.cb_org['values'] = ["SVE"] + orgs

        self.apply_filters()
        self.save_data()

    def apply_filters(self):
        q = self.search_var.get().strip().lower()
        st_filter = self.status_filter_var.get()
        org_filter = self.org_filter_var.get()

        filtered = []
        for r in self.data:
            # Tekstualna pretraga
            matches_text = (
                q in r.get("zaposleni", "").lower() or
                q in r.get("oprema", "").lower() or
                q in r.get("rm", "").lower() or
                q in r.get("grad", "").lower()
            )

            # Status filter
            matches_status = (st_filter == "SVE" or r.get("status") == st_filter)

            # Org filter
            matches_org = (org_filter == "SVE" or r.get("org_jedinica") == org_filter)

            if matches_text and matches_status and matches_org:
                filtered.append(r)

        self.current_filtered_data = filtered
        self.refresh_table_view()
        self.update_stats_and_chart()

    def reset_filters(self):
        self.search_var.set("")
        self.status_filter_var.set("SVE")
        self.org_filter_var.set("SVE")
        self.apply_filters()

    def refresh_table_view(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for r in self.current_filtered_data:
            d_zad_str = datetime.strptime(r["datum_zaduzenja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if r.get("datum_zaduzenja") else ""
            d_ist_str = datetime.strptime(r["datum_isticanja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if r.get("datum_isticanja") else ""

            self.tree.insert("", tk.END, values=(
                r.get("id"),
                r.get("zaposleni"),
                r.get("rm"),
                r.get("org_jedinica"),
                r.get("grad"),
                r.get("oprema"),
                r.get("jm"),
                r.get("normativ"),
                r.get("rok_mjeseci"),
                r.get("izdata_kol"),
                d_zad_str,
                d_ist_str,
                r.get("preostalo_dana"),
                r.get("status")
            ), tags=(r.get("status"),))

        self.status_bar.config(text=f"Prikazano zapisa: {len(self.current_filtered_data)} od ukupno {len(self.data)}")

    def update_stats_and_chart(self):
        total = len(self.current_filtered_data)
        valid = sum(1 for r in self.current_filtered_data if r.get("status") == "VAŽEĆE")
        warning = sum(1 for r in self.current_filtered_data if r.get("status") == "USKORO")
        expired = sum(1 for r in self.current_filtered_data if r.get("status") == "ISTEKLO")

        self.card_total_var.set(str(total))
        self.card_valid_var.set(str(valid))
        self.card_warning_var.set(str(warning))
        self.card_expired_var.set(str(expired))

        # Crtanje pita grafikona
        self.ax.clear()
        if total > 0:
            labels = []
            sizes = []
            colors_list = []

            if valid > 0:
                labels.append(f"Važeće\n({valid})")
                sizes.append(valid)
                colors_list.append("#22C55E")
            if warning > 0:
                labels.append(f"Uskoro\n({warning})")
                sizes.append(warning)
                colors_list.append("#EAB308")
            if expired > 0:
                labels.append(f"Isteklo\n({expired})")
                sizes.append(expired)
                colors_list.append("#EF4444")

            self.ax.pie(sizes, labels=labels, colors=colors_list, autopct='%1.1f%%', startangle=90, textprops={'fontsize': 8})
            self.ax.axis('equal')
        else:
            self.ax.text(0.5, 0.5, "Nema podataka", horizontalalignment='center', verticalalignment='center', transform=self.ax.transAxes)

        self.fig.tight_layout()
        self.canvas.draw()

    def get_summary_percentages(self):
        total = len(self.current_filtered_data)
        if total == 0:
            return "Nema selektovanih zapisa.", 0, 0, 0

        valid = sum(1 for r in self.current_filtered_data if r.get("status") == "VAŽEĆE")
        warning = sum(1 for r in self.current_filtered_data if r.get("status") == "USKORO")
        expired = sum(1 for r in self.current_filtered_data if r.get("status") == "ISTEKLO")

        text = f"Ukupno: {total} | Važeće: {valid} ({valid/total*100:.1f}%) | Uskoro ističe: {warning} ({warning/total*100:.1f}%) | Isteklo: {expired} ({expired/total*100:.1f}%)"
        return text, valid, warning, expired

    def generate_status_chart(self, filename="temp_chart.png"):
        """Generiše i čuva privremenu sliku grafikona za Excel/PDF izvještaj."""
        fig, ax = plt.subplots(figsize=(5, 2.5), dpi=150)
        total = len(self.current_filtered_data)
        
        if total > 0:
            valid = sum(1 for r in self.current_filtered_data if r.get("status") == "VAŽEĆE")
            warning = sum(1 for r in self.current_filtered_data if r.get("status") == "USKORO")
            expired = sum(1 for r in self.current_filtered_data if r.get("status") == "ISTEKLO")

            labels, sizes, colors_list = [], [], []
            if valid > 0:
                labels.append("Važeće")
                sizes.append(valid)
                colors_list.append("#22C55E")
            if warning > 0:
                labels.append("Uskoro")
                sizes.append(warning)
                colors_list.append("#EAB308")
            if expired > 0:
                labels.append("Isteklo")
                sizes.append(expired)
                colors_list.append("#EF4444")

            ax.pie(sizes, labels=labels, colors=colors_list, autopct='%1.1f%%', startangle=90, textprops={'fontsize': 9})
            ax.axis('equal')
        else:
            ax.text(0.5, 0.5, "Nema podataka", horizontalalignment='center', verticalalignment='center')

        plt.tight_layout()
        chart_path = os.path.join(os.getcwd(), filename)
        fig.savefig(chart_path, dpi=150, bbox_inches='tight')
        plt.close(fig)
        return chart_path

    def sort_table_column(self, col, reverse):
        l = [(self.tree.set(k, col), k) for k in self.tree.get_children('')]
        try:
            l.sort(key=lambda t: int(t[0]), reverse=reverse)
        except ValueError:
            l.sort(reverse=reverse)

        for index, (val, k) in enumerate(l):
            self.tree.move(k, '', index)

        self.tree.heading(col, command=lambda: self.sort_table_column(col, not reverse))

    def export_excel_report(self):
        """Generiše strukturirani i stilizovani Excel izvještaj sa slikom grafikona."""
        if not self.current_filtered_data:
            messagebox.showwarning("Upozorenje", "Nema podataka za izvoz.", parent=self.root)
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile="Izvjestaj_LZO_Zaduzenja.xlsx"
        )
        if not save_path:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Zaduženja LZO"

            # Naslov
            ws.merge_cells("A1:P1")
            title_cell = ws["A1"]
            title_cell.value = "IZVJEŠTAJ O ZADUŽENJU LIČNE ZAŠTITNE OPREME (LZO)"
            title_cell.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
            title_cell.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")

            # Podnaslov sa statistikom
            summary_text, _, _, _ = self.get_summary_percentages()
            ws.merge_cells("A2:P2")
            sub_cell = ws["A2"]
            sub_cell.value = f"Datum generisanja: {datetime.now().strftime('%d.%m.%Y. u %H:%M')} | {summary_text}"
            sub_cell.font = Font(name="Calibri", size=10, italic=True)
            sub_cell.alignment = Alignment(horizontal="center", vertical="center")

            headers = [
                "Zaposleni", "Radno mjesto", "Org. jedinica", "Grad", "Vel. odjeća",
                "Vel. obuća", "Oprema / Naziv", "J.M.", "Normativ", "Rok (mj)",
                "Izdata kol.", "Zaduženo", "Ističe", "Preostalo dana", "Status", "Napomena"
            ]

            header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            header_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
            thin_border = Border(
                left=Side(style='thin', color='CBD5E1'),
                right=Side(style='thin', color='CBD5E1'),
                top=Side(style='thin', color='CBD5E1'),
                bottom=Side(style='thin', color='CBD5E1')
            )

            for col_num, header_title in enumerate(headers, 1):
                cell = ws.cell(row=4, column=col_num)
                cell.value = header_title
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            fill_expired = PatternFill(start_color="FECDD3", end_color="FECDD3", fill_type="solid")
            fill_warning = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
            fill_valid = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")

            for row_idx, item in enumerate(self.current_filtered_data, 5):
                d_zad = datetime.strptime(item["datum_zaduzenja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if item.get("datum_zaduzenja") else ""
                d_ist = datetime.strptime(item["datum_isticanja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if item.get("datum_isticanja") else ""
                
                row_values = [
                    item.get("zaposleni", ""),
                    item.get("rm", ""),
                    item.get("org_jedinica", ""),
                    item.get("grad", ""),
                    item.get("vel_odjeca", ""),
                    item.get("vel_obuca", ""),
                    item.get("oprema", ""),
                    item.get("jm", "KOM"),
                    item.get("normativ", 1),
                    item.get("rok_mjeseci", 12),
                    item.get("izdata_kol", 1),
                    d_zad,
                    d_ist,
                    item.get("preostalo_dana", "-"),
                    item.get("status", ""),
                    item.get("napomena", "")
                ]

                for col_idx, val in enumerate(row_values, 1):
                    c = ws.cell(row=row_idx, column=col_idx, value=val)
                    c.border = thin_border
                    c.font = Font(name="Calibri", size=10)
                    if col_idx in [5, 6, 8, 9, 10, 11, 12, 13, 14, 15]:
                        c.alignment = Alignment(horizontal="center")
                    
                    if col_idx == 15:
                        st = item.get("status")
                        if st == "ISTEKLO":
                            c.fill = fill_expired
                        elif st == "USKORO":
                            c.fill = fill_warning
                        elif st == "VAŽEĆE":
                            c.fill = fill_valid

            # Ugradnja grafikona u Excel izvještaj
            chart_img_path = self.generate_status_chart()
            if os.path.exists(chart_img_path):
                img = OpenpyxlImage(chart_img_path)
                img.width = 450
                img.height = 225
                ws.add_image(img, f"A{len(self.current_filtered_data) + 7}")

            # Automatsko prilagođavanje širine kolona
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

            wb.save(save_path)
            messagebox.showinfo("Uspjeh", "Izvještaj je uspješno izvezen u Excel!", parent=self.root)
        except Exception as e:
            messagebox.showerror("Greška", f"Greška pri izvozu u Excel: {e}", parent=self.root)

    def export_pdf_report(self):
        """Generiše landscape PDF izvještaj sa tabelom i grafikonom pomoću ReportLab biblioteke."""
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Greška", "ReportLab biblioteka nije instalirana. Instalirajte je naredbom: pip install reportlab", parent=self.root)
            return

        if not self.current_filtered_data:
            messagebox.showwarning("Upozorenje", "Nema podataka za izvoz.", parent=self.root)
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile="Izvjestaj_LZO.pdf"
        )
        if not save_path:
            return

        try:
            pdf_font = get_pdf_unicode_font()
            doc = SimpleDocTemplate(
                save_path,
                pagesize=landscape(A4),
                rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20
            )

            elements = []
            styles = getSampleStyleSheet()

            title_style = ParagraphStyle(
                'PdfTitle', parent=styles['Heading1'], fontName=pdf_font,
                fontSize=14, leading=16, textColor=colors.HexColor('#1E293B'), alignment=1
            )
            sub_style = ParagraphStyle(
                'PdfSubTitle', parent=styles['Normal'], fontName=pdf_font,
                fontSize=9, leading=12, textColor=colors.HexColor('#475569'), alignment=1
            )
            cell_style = ParagraphStyle('CellText', parent=styles['Normal'], fontName=pdf_font, fontSize=7, leading=9)
            cell_header_style = ParagraphStyle('CellHeader', parent=styles['Normal'], fontName=pdf_font, fontSize=7.5, leading=9, textColor=colors.white)

            elements.append(Paragraph("IZVJEŠTAJ ZADUŽENJA LIČNE ZAŠTITNE OPREME (LZO)", title_style))
            elements.append(Spacer(1, 6))

            summary_text, _, _, _ = self.get_summary_percentages()
            elements.append(Paragraph(f"Datum generisanja: {datetime.now().strftime('%d.%m.%Y. u %H:%M')} | {summary_text}", sub_style))
            elements.append(Spacer(1, 10))

            headers = ["Zaposleni", "Radno mjesto", "Org. jed.", "Grad", "Oprema", "J.M.", "Nor.", "Rok", "Kol.", "Zaduženo", "Ističe", "Ostalo", "Status"]
            table_data = [[Paragraph(h, cell_header_style) for h in headers]]

            for row in self.current_filtered_data:
                d_zad = datetime.strptime(row["datum_zaduzenja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if row.get("datum_zaduzenja") else ""
                d_ist = datetime.strptime(row["datum_isticanja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if row.get("datum_isticanja") else ""
                
                r = [
                    Paragraph(str(row.get("zaposleni", "")), cell_style),
                    Paragraph(str(row.get("rm", "")), cell_style),
                    Paragraph(str(row.get("org_jedinica", "")), cell_style),
                    Paragraph(str(row.get("grad", "")), cell_style),
                    Paragraph(str(row.get("oprema", "")), cell_style),
                    Paragraph(str(row.get("jm", "KOM")), cell_style),
                    Paragraph(str(row.get("normativ", 1)), cell_style),
                    Paragraph(f"{row.get('rok_mjeseci', 12)}m", cell_style),
                    Paragraph(str(row.get("izdata_kol", 1)), cell_style),
                    Paragraph(d_zad, cell_style),
                    Paragraph(d_ist, cell_style),
                    Paragraph(f"{row.get('preostalo_dana', '-')}", cell_style),
                    Paragraph(str(row.get("status", "")), cell_style)
                ]
                table_data.append(r)

            col_widths = [100, 90, 85, 60, 110, 30, 30, 30, 30, 60, 60, 40, 50]
            t = Table(table_data, colWidths=col_widths, repeatRows=1)
            t_style = [
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E293B')),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ]

            for idx, row in enumerate(self.current_filtered_data, 1):
                st = row.get("status")
                if st == "ISTEKLO":
                    t_style.append(('BACKGROUND', (0, idx), (-1, idx), colors.HexColor('#FECDD3')))
                elif st == "USKORO":
                    t_style.append(('BACKGROUND', (0, idx), (-1, idx), colors.HexColor('#FEF08A')))

            t.setStyle(TableStyle(t_style))
            elements.append(t)

            chart_path = self.generate_status_chart()
            if os.path.exists(chart_path):
                elements.append(Spacer(1, 15))
                elements.append(ReportLabImage(chart_path, width=400, height=200))

            doc.build(elements)
            messagebox.showinfo("Uspjeh", "PDF izvještaj je uspješno generisan!", parent=self.root)
        except Exception as e:
            messagebox.showerror("Greška", f"Greška pri generisanju PDF-a: {e}", parent=self.root)

    def import_excel(self):
        """Uvozi zaduženja iz spoljnog Excel (.xlsx) fajla sa automatskim mapiranjem kolona."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
            self.push_undo_state()

            col_map = {
                "zaposleni": ["zaposleni", "ime i prezime", "radnik", "zaposleni / ime i prezime"],
                "rm": ["rm", "radno mjesto", "radno_mjesto", "posao"],
                "org_jedinica": ["org_jedinica", "org. jedinica", "organizaciona jedinica", "sektor"],
                "grad": ["grad", "mjesto", "mjesto / grad", "lokacija"],
                "vel_odjeca": ["vel_odjeca", "vel. odjeća", "velicina odjece", "odjeća"],
                "vel_obuca": ["vel_obuca", "vel. obuća", "velicina obuce", "obuća"],
                "oprema": ["oprema", "naziv opreme", "lzo oprema", "oprema / naziv"],
                "jm": ["jm", "j.m.", "jedinica mjere"],
                "normativ": ["normativ", "norma"],
                "rok_mjeseci": ["rok_mjeseci", "rok (mj)", "rok", "trajanje (mjeseci)"],
                "izdata_kol": ["izdata_kol", "izdata kol.", "količina", "kol"],
                "datum_zaduzenja": ["datum_zaduzenja", "zaduženo", "datum zaduženja", "datum zaduzenja"],
                "napomena": ["napomena", "napomene", "komentar"]
            }

            def find_col(df_cols, possible):
                for col in df_cols:
                    if str(col).strip().lower() in possible:
                        return col
                return None

            df.columns = [str(c).strip() for c in df.columns]
            imported = []
            max_id = max([r.get("id", 0) for r in self.data], default=0)

            for _, row in df.iterrows():
                max_id += 1
                rec = {"id": max_id}
                for key, possible_names in col_map.items():
                    matched_col = find_col(df.columns, possible_names)
                    val = row[matched_col] if matched_col else ""
                    if pd.isna(val):
                        val = ""
                    
                    if key in ["normativ", "rok_mjeseci", "izdata_kol"]:
                        try: val = int(val)
                        except: val = 1 if key != "rok_mjeseci" else 12
                    elif key == "datum_zaduzenja" and val:
                        if isinstance(val, datetime):
                            val = val.strftime("%Y-%m-%d")
                        else:
                            try:
                                val = pd.to_datetime(val).strftime("%Y-%m-%d")
                            except:
                                val = datetime.now().strftime("%Y-%m-%d")

                    rec[key] = str(val).strip() if isinstance(val, str) else val

                if not rec.get("jm"): rec["jm"] = "KOM"
                if not rec.get("normativ"): rec["normativ"] = 1
                if not rec.get("rok_mjeseci"): rec["rok_mjeseci"] = 12
                if not rec.get("izdata_kol"): rec["izdata_kol"] = 1

                imported.append(rec)

            self.data.extend(imported)
            self.recalculate_and_refresh()
            messagebox.showinfo("Uvoz završen", f"Uspješno uvezeno {len(imported)} zapisa iz Excel fajla.", parent=self.root)
        except Exception as e:
            messagebox.showerror("Greška pri uvozu", f"Nije moguće uvesti podatke: {e}", parent=self.root)

    def open_add_dialog(self):
        """Otvara prozor za unos novog zaduženja."""
        max_id = max([r.get("id", 0) for r in self.data], default=0) + 1
        dlg = RecordDialog(self.root, title="Novo zaduženje LZO", record={"id": max_id})
        self.root.wait_window(dlg)
        if dlg.result:
            self.push_undo_state()
            self.data.append(dlg.result)
            self.recalculate_and_refresh()

    def open_edit_dialog(self, event=None):
        """Otvara prozor za izmjenu odabranog zaduženja."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Upozorenje", "Selektujte zapis za izmjenu.", parent=self.root)
            return

        item_id = self.tree.item(selected[0])['values'][0]
        record = next((r for r in self.data if r.get("id") == item_id), None)
        if not record:
            return

        dlg = RecordDialog(self.root, title="Izmjena zaduženja LZO", record=record)
        self.root.wait_window(dlg)
        if dlg.result:
            self.push_undo_state()
            idx = next(i for i, r in enumerate(self.data) if r.get("id") == item_id)
            self.data[idx] = dlg.result
            self.recalculate_and_refresh()

    def delete_selected(self):
        """Briše selektovane zapise iz baze."""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Upozorenje", "Selektujte zaduženje(a) koje želite obrisati.", parent=self.root)
            return

        if messagebox.askyesno("Potvrda brisanja", f"Da li ste sigurni da želite obrisati {len(selected)} selektovanih zapisa?", parent=self.root):
            self.push_undo_state()
            ids_to_del = [self.tree.item(s)['values'][0] for s in selected]
            self.data = [r for r in self.data if r.get("id") not in ids_to_del]
            self.recalculate_and_refresh()

    def show_about(self):
        messagebox.showinfo("O aplikaciji", "Sistem za evidenciju LZO v2.0\n\nFunkcionalnosti:\n- Praćenje rokova trajanja LZO\n- Automatska kalkulacija statusa\n- Izvoz u PDF i Excel sa grafikonima\n- Uvoz podataka iz Excel-a\n- Undo/Redo podrška", parent=self.root)


if __name__ == "__main__":
    root = tk.Tk()
    app = LZOApp(root)
    root.mainloop()
