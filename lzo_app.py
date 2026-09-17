import os
import json
import copy
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenpyxlImage

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

# ReportLab biblioteka za PDF generisanje
try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as ReportLabImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

DB_FILE = "lzo_baza.json"

def get_pdf_unicode_font():
    """Registruje i vraća sistemski font koji podržava naša slova (Č, Š, Ć, Đ, Ž) za ReportLab PDF."""
    if not REPORTLAB_AVAILABLE:
        return "Helvetica"
    
    font_paths = [
        "C:\\Windows\\Fonts\\arial.ttf",
        "C:\\Windows\\Fonts\\calibri.ttf",
        "C:\\Windows\\Fonts\\tahoma.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/TTF/DejaVuSans.ttf",
        "/Library/Fonts/Arial.ttf"
    ]
    
    for path in font_paths:
        if os.path.exists(path):
            try:
                font_name = "CustomUnicodeFont"
                pdfmetrics.registerFont(TTFont(font_name, path))
                return font_name
            except Exception:
                continue
    return "Helvetica"


class LoginWindow(tk.Toplevel):
    def __init__(self, parent, on_success):
        super().__init__(parent)
        self.parent = parent
        self.on_success = on_success
        self.title("Prijava na sistem - LZO v4.0")
        self.geometry("400x320")
        self.resizable(False, False)
        self.configure(bg="#f8fafc")
        
        # Centriranje prozora
        self.tk.eval('tk::PlaceWindow . center')
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.create_widgets()

    def create_widgets(self):
        # Header Baner
        header_frame = tk.Frame(self, bg="#1e293b", height=70)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)

        lbl_header = tk.Label(
            header_frame, text="🔒 PRIJAVA NA SISTEM", 
            font=("Segoe UI", 12, "bold"), fg="white", bg="#1e293b"
        )
        lbl_header.pack(expand=True)

        # Forma za unos
        form_frame = tk.Frame(self, bg="#f8fafc", padx=30, pady=20)
        form_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(form_frame, text="Korisničko ime:", font=("Segoe UI", 9, "bold"), bg="#f8fafc", fg="#334155").pack(anchor="w", pady=(0, 2))
        self.entry_user = ttk.Entry(form_frame, font=("Segoe UI", 10))
        self.entry_user.pack(fill=tk.X, pady=(0, 10))
        self.entry_user.focus()

        tk.Label(form_frame, text="Lozinka:", font=("Segoe UI", 9, "bold"), bg="#f8fafc", fg="#334155").pack(anchor="w", pady=(0, 2))
        self.entry_pass = ttk.Entry(form_frame, font=("Segoe UI", 10), show="•")
        self.entry_pass.pack(fill=tk.X, pady=(0, 15))
        self.entry_pass.bind("<Return>", lambda e: self.check_login())

        btn_login = tk.Button(
            form_frame, text="Prijavi se", command=self.check_login,
            bg="#2563eb", fg="white", font=("Segoe UI", 10, "bold"),
            activebackground="#1d4ed8", activeforeground="white", relief="flat", cursor="hand2", pady=6
        )
        btn_login.pack(fill=tk.X)

    def check_login(self):
        user = self.entry_user.get().strip()
        password = self.entry_pass.get().strip()

        if user.upper() == "REGION 4" and password == "12345678":
            self.destroy()
            self.on_success()
        else:
            messagebox.showerror("Greška pri prijavi", "Neispravno korisničko ime ili lozinka!", parent=self)

    def on_close(self):
        self.parent.destroy()


class RecordDialog(tk.Toplevel):
    """Dijalog za unos i izmjenu zaduženja LZO."""
    def __init__(self, parent, title="Zaduženje LZO", record=None):
        super().__init__(parent)
        self.title(title)
        self.geometry("520x620")
        self.resizable(False, False)
        self.configure(bg="#f8fafc")
        self.result = None
        self.record = record or {}

        self.transient(parent)
        self.grab_set()

        self.create_widgets()
        self.tk.eval('tk::PlaceWindow . center')

    def create_widgets(self):
        header = tk.Frame(self, bg="#1e293b", pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="UNOS / IZMJENA ZADUŽENJA", font=("Segoe UI", 11, "bold"), fg="white", bg="#1e293b").pack()

        body = tk.Frame(self, bg="#f8fafc", padx=20, pady=15)
        body.pack(fill=tk.BOTH, expand=True)

        fields = [
            ("Zaposleni (Ime i prezime):", "zaposleni"),
            ("Radno mjesto:", "rm"),
            ("Organizaciona jedinica:", "org_jedinica"),
            ("Mjesto / Grad:", "grad"),
            ("Veličina odjeće:", "vel_odjeca"),
            ("Veličina obuće:", "vel_obuca"),
            ("Naziv opreme:", "oprema"),
            ("Jedinica mjere (J.M.):", "jm"),
            ("Normativ:", "normativ"),
            ("Rok trajanja (mjeseci):", "rok_mjeseci"),
            ("Izdata količina:", "izdata_kol"),
            ("Datum zaduženja (YYYY-MM-DD):", "datum_zaduzenja"),
            ("Napomena:", "napomena")
        ]

        self.entries = {}
        for idx, (label_text, key) in enumerate(fields):
            tk.Label(body, text=label_text, font=("Segoe UI", 8, "bold"), bg="#f8fafc", fg="#334155").grid(row=idx, column=0, sticky="w", pady=2)
            ent = ttk.Entry(body, font=("Segoe UI", 9))
            ent.grid(row=idx, column=1, sticky="ew", pady=2, padx=(10, 0))
            
            val = self.record.get(key, "")
            if key == "jm" and not val: val = "KOM"
            if key == "normativ" and not val: val = "1"
            if key == "rok_mjeseci" and not val: val = "12"
            if key == "izdata_kol" and not val: val = "1"
            if key == "datum_zaduzenja" and not val: val = datetime.now().strftime("%Y-%m-%d")

            ent.insert(0, str(val))
            self.entries[key] = ent

        body.columnconfigure(1, weight=1)

        btn_frame = tk.Frame(self, bg="#f8fafc", pady=10)
        btn_frame.pack(fill=tk.X)

        tk.Button(btn_frame, text="Sačuvaj", command=self.on_save, bg="#16a34a", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=15, pady=4, cursor="hand2").pack(side=tk.RIGHT, padx=15)
        tk.Button(btn_frame, text="Otkaži", command=self.destroy, bg="#64748b", fg="white", font=("Segoe UI", 9), relief="flat", padx=15, pady=4, cursor="hand2").pack(side=tk.RIGHT, padx=5)

    def on_save(self):
        try:
            d_zad = self.entries["datum_zaduzenja"].get().strip()
            if d_zad:
                datetime.strptime(d_zad, "%Y-%m-%d")
        except ValueError:
            messagebox.showerror("Greška", "Datum zaduženja mora biti u formatu YYYY-MM-DD!", parent=self)
            return

        try:
            normativ = int(self.entries["normativ"].get().strip() or 1)
            rok = int(self.entries["rok_mjeseci"].get().strip() or 12)
            kol = int(self.entries["izdata_kol"].get().strip() or 1)
        except ValueError:
            messagebox.showerror("Greška", "Normativ, rok i količina moraju biti cijeli brojevi!", parent=self)
            return

        self.result = {
            "id": self.record.get("id"),
            "zaposleni": self.entries["zaposleni"].get().strip(),
            "rm": self.entries["rm"].get().strip(),
            "org_jedinica": self.entries["org_jedinica"].get().strip(),
            "grad": self.entries["grad"].get().strip(),
            "vel_odjeca": self.entries["vel_odjeca"].get().strip(),
            "vel_obuca": self.entries["vel_obuca"].get().strip(),
            "oprema": self.entries["oprema"].get().strip(),
            "jm": self.entries["jm"].get().strip() or "KOM",
            "normativ": normativ,
            "rok_mjeseci": rok,
            "izdata_kol": kol,
            "datum_zaduzenja": d_zad,
            "napomena": self.entries["napomena"].get().strip()
        }
        self.destroy()


class LZOApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistem za praćenje LZO i rokova zaduženja v4.0")
        self.root.geometry("1420x820")
        self.root.configure(bg="#f1f5f9")

        # Sakrij glavni prozor dok se korisnik ne prijavi
        self.root.withdraw()

        self.data = []
        self.current_filtered_data = []
        self.undo_stack = []  # Stak za Undo funkcionalnost
        self.sort_directions = {}

        # Inicijalizacija stilova za moderan izgled
        self.setup_styles()

        # Otvori Login Dialog
        LoginWindow(self.root, on_success=self.show_main_app)

    def show_main_app(self):
        self.root.deiconify()
        self.create_widgets()
        self.load_local_db()
        # Snimanje prečice za Undo (Ctrl + Z)
        self.root.bind("<Control-z>", lambda e: self.undo())
        self.root.bind("<Control-Z>", lambda e: self.undo())

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")

        # Tabela (Treeview) Stilizovanje
        self.style.configure("Treeview",
            background="#ffffff",
            foreground="#1e293b",
            rowheight=26,
            fieldbackground="#ffffff",
            font=("Segoe UI", 9)
        )
        self.style.configure("Treeview.Heading",
            background="#334155",
            foreground="#ffffff",
            font=("Segoe UI", 9, "bold"),
            padding=5
        )
        self.style.map("Treeview.Heading", background=[('active', '#475569')])

    def push_undo_state(self):
        """Čuva trenutno stanje podataka za Undo vraćanje koraka."""
        self.undo_stack.append(copy.deepcopy(self.data))
        if len(self.undo_stack) > 30:  # Ograničenje na zadnjih 30 koraka
            self.undo_stack.pop(0)

    def undo(self):
        """Vraća prethodni korak (Undo)."""
        if not self.undo_stack:
            messagebox.showinfo("Undo", "Nema prethodnih koraka za poništavanje.")
            return

        self.data = self.undo_stack.pop()
        self.recalculate_and_refresh()
        self.lbl_status.config(text="↩️ Posljednja akcija je poništena (Undo).", fg="#2563eb")

    def create_widgets(self):
        # --- Zaglavlje Aplikacije ---
        header_frame = tk.Frame(self.root, bg="#1e293b", pady=10, padx=15)
        header_frame.pack(fill=tk.X)

        lbl_title = tk.Label(
            header_frame, text="🛡️ SISTEM ZA PRAĆENJE LIČNE ZAŠTITNE OPREME (LZO)",
            font=("Segoe UI", 13, "bold"), fg="#f8fafc", bg="#1e293b"
        )
        lbl_title.pack(side=tk.LEFT)

        btn_undo = tk.Button(
            header_frame, text="↶ Poništi (Undo)", command=self.undo,
            bg="#475569", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2", padx=10
        )
        btn_undo.pack(side=tk.RIGHT, padx=5)

        # --- Gornji komandni panel ---
        top_frame = tk.Frame(self.root, pady=8, padx=12, bg="#ffffff", bd=1, relief="groove")
        top_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

        btn_import = tk.Button(top_frame, text="📥 Uvoz iz Excel-a", command=self.import_excel, bg="#0284c7", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2", padx=8, pady=4)
        btn_import.pack(side=tk.LEFT, padx=4)

        btn_add = tk.Button(top_frame, text="+ Novo zaduženje", command=self.open_add_dialog, bg="#16a34a", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2", padx=8, pady=4)
        btn_add.pack(side=tk.LEFT, padx=4)

        btn_edit = tk.Button(top_frame, text="✏️ Izmijeni selektovano", command=self.open_edit_dialog, bg="#d97706", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2", padx=8, pady=4)
        btn_edit.pack(side=tk.LEFT, padx=4)

        btn_delete = tk.Button(top_frame, text="🗑️ Obriši selektovano", command=self.delete_selected, bg="#dc2626", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2", padx=8, pady=4)
        btn_delete.pack(side=tk.LEFT, padx=4)

        # Izvještaji
        btn_export_pdf = tk.Button(top_frame, text="📄 Izvezi PDF", command=self.export_pdf_report, bg="#991b1b", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2", padx=8, pady=4)
        btn_export_pdf.pack(side=tk.RIGHT, padx=4)

        btn_export_excel = tk.Button(top_frame, text="📊 Izvezi Excel", command=self.export_excel_report, bg="#15803d", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", cursor="hand2", padx=8, pady=4)
        btn_export_excel.pack(side=tk.RIGHT, padx=4)

        # --- Panel za Pretragu i Napredno Filtriranje ---
        filter_frame = tk.LabelFrame(self.root, text=" Pretraga i Filtriranje ", font=("Segoe UI", 9, "bold"), padx=10, pady=8, bg="#ffffff", fg="#334155")
        filter_frame.pack(fill=tk.X, padx=10, pady=5)

        # 1. Autocomplete / Dynamic Dropdown Pretraga Zaposlenih
        tk.Label(filter_frame, text="Pretraga Zaposlenih:", bg="#ffffff", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=2)
        self.combo_search_emp = ttk.Combobox(filter_frame, width=22, font=("Segoe UI", 9))
        self.combo_search_emp.grid(row=0, column=1, padx=4, pady=2)
        self.combo_search_emp.bind("<KeyRelease>", self.on_emp_search_key)
        self.combo_search_emp.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        # Opšta pretraga (sva polja)
        tk.Label(filter_frame, text="Opšta pretraga:", bg="#ffffff", font=("Segoe UI", 9)).grid(row=0, column=2, sticky="w", padx=(10, 2))
        self.entry_search = ttk.Entry(filter_frame, width=16)
        self.entry_search.grid(row=0, column=3, padx=4, pady=2)
        self.entry_search.bind("<KeyRelease>", lambda e: self.apply_filters())

        # Org jedinica
        tk.Label(filter_frame, text="Org. jedinica:", bg="#ffffff", font=("Segoe UI", 9)).grid(row=0, column=4, sticky="w", padx=(10, 2))
        self.combo_org = ttk.Combobox(filter_frame, state="readonly", width=16)
        self.combo_org.grid(row=0, column=5, padx=4, pady=2)
        self.combo_org.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        # Mjesto / Grad
        tk.Label(filter_frame, text="Grad:", bg="#ffffff", font=("Segoe UI", 9)).grid(row=0, column=6, sticky="w", padx=(10, 2))
        self.combo_city = ttk.Combobox(filter_frame, state="readonly", width=12)
        self.combo_city.grid(row=0, column=7, padx=4, pady=2)
        self.combo_city.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        # Oprema
        tk.Label(filter_frame, text="Oprema:", bg="#ffffff", font=("Segoe UI", 9)).grid(row=0, column=8, sticky="w", padx=(10, 2))
        self.combo_equipment = ttk.Combobox(filter_frame, state="readonly", width=14)
        self.combo_equipment.grid(row=0, column=9, padx=4, pady=2)
        self.combo_equipment.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        # Status
        tk.Label(filter_frame, text="Status:", bg="#ffffff", font=("Segoe UI", 9)).grid(row=0, column=10, sticky="w", padx=(10, 2))
        self.combo_status = ttk.Combobox(filter_frame, state="readonly", width=11, values=["SVI", "ISTEKLO", "USKORO", "VAŽEĆE"])
        self.combo_status.current(0)
        self.combo_status.grid(row=0, column=11, padx=4, pady=2)
        self.combo_status.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        btn_reset = tk.Button(filter_frame, text="Poništi filtere", command=self.reset_filters, font=("Segoe UI", 8), bg="#e2e8f0", relief="flat")
        btn_reset.grid(row=0, column=12, padx=(10, 2))

        # --- Statusna traka sa statistikom ---
        self.lbl_status = tk.Label(self.root, text="Inicijalizacija sistema...", font=("Segoe UI", 9, "bold"), anchor="w", padx=15, pady=6, bg="#e2e8f0", fg="#1e293b")
        self.lbl_status.pack(fill=tk.X)

        # --- Tabela (Treeview) ---
        table_frame = tk.Frame(self.root, bg="#ffffff")
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.cols = (
            "ID", "Zaposleni", "Radno mjesto", "Org. jedinica", "Mjesto",
            "Vel. odjeća", "Vel. obuća", "Oprema", "J.M.", "Normativ",
            "Rok (mj)", "Izdata kol.", "Zaduženo", "Ističe", "Preostalo dana", "Status", "Napomena"
        )

        self.tree = ttk.Treeview(table_frame, columns=self.cols, show="headings", selectmode="extended")

        col_widths = {
            "ID": 30, "Zaposleni": 160, "Radno mjesto": 140, "Org. jedinica": 130, "Mjesto": 95,
            "Vel. odjeća": 70, "Vel. obuća": 70, "Oprema": 160, "J.M.": 50, "Normativ": 60,
            "Rok (mj)": 60, "Izdata kol.": 70, "Zaduženo": 85, "Ističe": 85, "Preostalo dana": 95,
            "Status": 90, "Napomena": 120
        }

        for col in self.cols:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by_column(c))
            center_cols = ["ID", "Vel. odjeća", "Vel. obuća", "J.M.", "Normativ", "Rok (mj)", "Izdata kol.", "Zaduženo", "Ističe", "Preostalo dana", "Status"]
            self.tree.column(col, width=col_widths.get(col, 100), anchor=tk.CENTER if col in center_cols else tk.W)

        self.tree.column("ID", width=0, stretch=False)

        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)

        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # Oznake boja u tabeli
        self.tree.tag_configure("ISTEKLO", background="#fecdd3", foreground="#9f1239")
        self.tree.tag_configure("USKORO", background="#fef08a", foreground="#854d0e")
        self.tree.tag_configure("VAŽEĆE", background="#dcfce7", foreground="#166534")

        self.tree.bind("<Double-1>", self.open_edit_dialog)

    def on_emp_search_key(self, event):
        """Dinamički ažurira opadajući meni sa imenima zaposlenih na osnovu unijetog teksta."""
        typed = self.combo_search_emp.get().strip().lower()
        all_employees = sorted(list(set(row.get("zaposleni", "") for row in self.data if row.get("zaposleni"))))
        
        if typed:
            matches = [emp for emp in all_employees if typed in emp.lower()]
            self.combo_search_emp['values'] = matches
        else:
            self.combo_search_emp['values'] = all_employees

        self.apply_filters()

    def load_local_db(self):
        if os.path.exists(DB_FILE):
            try:
                with open(DB_FILE, "r", encoding="utf-8") as f:
                    self.data = json.load(f)
                self.recalculate_and_refresh()
            except Exception as e:
                messagebox.showerror("Greška", f"Nije moguće učitati bazu: {e}")
        else:
            self.lbl_status.config(text="Lokalna baza nije pronađena. Uvezite podatke iz Excel fajla.", fg="#2563eb")

    def save_local_db(self):
        try:
            with open(DB_FILE, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Greška", f"Greška pri čuvanju baze: {e}")

    def recalculate_and_refresh(self):
        today = datetime.now()
        expired_count = 0
        warning_count = 0
        valid_count = 0

        for item in self.data:
            rok = int(item.get("rok_mjeseci", 12) or 12)
            izdata_kol = int(item.get("izdata_kol", 1) or 1)
            d_zad = item.get("datum_zaduzenja", "")

            if d_zad:
                try:
                    dt_zad = datetime.strptime(d_zad, "%Y-%m-%d")
                    total_months = izdata_kol * rok
                    dt_ist = dt_zad + relativedelta(months=total_months)
                    item["datum_isticanja"] = dt_ist.strftime("%Y-%m-%d")

                    preostalo = (dt_ist - today).days
                    item["preostalo_dana"] = preostalo

                    if preostalo < 0:
                        item["status"] = "ISTEKLO"
                        expired_count += 1
                    elif preostalo <= 30:
                        item["status"] = "USKORO"
                        warning_count += 1
                    else:
                        item["status"] = "VAŽEĆE"
                        valid_count += 1
                except Exception:
                    item["status"] = "GREŠKA"
                    item["preostalo_dana"] = 0
            else:
                item["datum_isticanja"] = ""
                item["preostalo_dana"] = "-"
                item["status"] = "NEZADUŽENO"

        self.save_local_db()
        self.update_filter_dropdowns()
        self.apply_filters()

        total = len(self.current_filtered_data)
        p_val = (valid_count / total * 100) if total > 0 else 0
        p_warn = (warning_count / total * 100) if total > 0 else 0
        p_exp = (expired_count / total * 100) if total > 0 else 0

        msg = f"Ukupno prikazano: {total} | VAŽEĆE: {valid_count} ({p_val:.1f}%) | ISTIČE USKORO (<=30d): {warning_count} ({p_warn:.1f}%) | ISTEKLO: {expired_count} ({p_exp:.1f}%)"
        self.lbl_status.config(text=msg, fg="#dc2626" if expired_count > 0 else "#1e293b")

    def update_filter_dropdowns(self):
        employees = sorted(list(set(row.get("zaposleni", "") for row in self.data if row.get("zaposleni"))))
        orgs = sorted(list(set(row.get("org_jedinica", "") for row in self.data if row.get("org_jedinica"))))
        cities = sorted(list(set(row.get("grad", "") for row in self.data if row.get("grad"))))
        equipments = sorted(list(set(row.get("oprema", "") for row in self.data if row.get("oprema"))))

        self.combo_search_emp['values'] = employees
        self.combo_org['values'] = ["SVE ORG. JEDINICE"] + orgs
        self.combo_city['values'] = ["SVA MJESTA"] + cities
        self.combo_equipment['values'] = ["SVA OPREMA"] + equipments

        if not self.combo_org.get(): self.combo_org.current(0)
        if not self.combo_city.get(): self.combo_city.current(0)
        if not self.combo_equipment.get(): self.combo_equipment.current(0)

    def apply_filters(self):
        emp_txt = self.combo_search_emp.get().strip().lower()
        search_txt = self.entry_search.get().strip().lower()
        selected_org = self.combo_org.get()
        selected_city = self.combo_city.get()
        selected_equip = self.combo_equipment.get()
        selected_status = self.combo_status.get()

        filtered = []
        for row in self.data:
            match_emp = True
            if emp_txt:
                match_emp = emp_txt in str(row.get("zaposleni", "")).lower()

            match_search = True
            if search_txt:
                match_search = (
                    search_txt in str(row.get("zaposleni", "")).lower() or
                    search_txt in str(row.get("rm", "")).lower() or
                    search_txt in str(row.get("org_jedinica", "")).lower() or
                    search_txt in str(row.get("grad", "")).lower() or
                    search_txt in str(row.get("oprema", "")).lower()
                )

            match_org = True
            if selected_org and selected_org != "SVE ORG. JEDINICE":
                match_org = (row.get("org_jedinica") == selected_org)

            match_city = True
            if selected_city and selected_city != "SVA MJESTA":
                match_city = (row.get("grad") == selected_city)

            match_equip = True
            if selected_equip and selected_equip != "SVA OPREMA":
                match_equip = (row.get("oprema") == selected_equip)

            match_status = True
            if selected_status and selected_status != "SVI":
                match_status = (row.get("status") == selected_status)

            if match_emp and match_search and match_org and match_city and match_equip and match_status:
                filtered.append(row)

        self.current_filtered_data = filtered
        self.refresh_table(filtered)

    def reset_filters(self):
        self.combo_search_emp.set("")
        self.entry_search.delete(0, tk.END)
        self.combo_org.current(0)
        self.combo_city.current(0)
        self.combo_equipment.current(0)
        self.combo_status.current(0)
        self.apply_filters()

    def refresh_table(self, dataset):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for row in dataset:
            d_zad = datetime.strptime(row["datum_zaduzenja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if row.get("datum_zaduzenja") else ""
            d_ist = datetime.strptime(row["datum_isticanja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if row.get("datum_isticanja") else ""

            self.tree.insert("", tk.END, values=(
                row.get("id"),
                row.get("zaposleni", ""),
                row.get("rm", ""),
                row.get("org_jedinica", ""),
                row.get("grad", ""),
                row.get("vel_odjeca", ""),
                row.get("vel_obuca", ""),
                row.get("oprema", ""),
                row.get("jm", "KOM"),
                row.get("normativ", 1),
                row.get("rok_mjeseci", 12),
                row.get("izdata_kol", 1),
                d_zad,
                d_ist,
                row.get("preostalo_dana", "-"),
                row.get("status", ""),
                row.get("napomena", "")
            ), tags=(row.get("status", ""),))

    def sort_by_column(self, col):
        ascending = self.sort_directions.get(col, True)
        key_map = {
            "ID": "id", "Zaposleni": "zaposleni", "Radno mjesto": "rm", "Org. jedinica": "org_jedinica",
            "Mjesto": "grad", "Vel. odjeća": "vel_odjeca", "Vel. obuća": "vel_obuca", "Oprema": "oprema",
            "J.M.": "jm", "Normativ": "normativ", "Rok (mj)": "rok_mjeseci", "Izdata kol.": "izdata_kol",
            "Zaduženo": "datum_zaduzenja", "Ističe": "datum_isticanja",
            "Preostalo dana": "preostalo_dana", "Status": "status", "Napomena": "napomena"
        }
        key = key_map.get(col)
        if not key: return

        def sort_val(item):
            val = item.get(key, "")
            if val is None: return ""
            if key in ["id", "normativ", "rok_mjeseci", "izdata_kol"]:
                try: return int(val)
                except: return 0
            if key == "preostalo_dana":
                try: return int(val)
                except: return -99999
            return str(val).lower()

        self.current_filtered_data.sort(key=sort_val, reverse=not ascending)
        self.sort_directions[col] = not ascending

        for c in self.cols:
            if c == col:
                arrow = " ▲" if ascending else " ▼"
                self.tree.heading(c, text=c + arrow)
            else:
                self.tree.heading(c, text=c)

        self.refresh_table(self.current_filtered_data)

    def generate_status_chart(self):
        counts = {"VAŽEĆE": 0, "USKORO": 0, "ISTEKLO": 0}
        for item in self.current_filtered_data:
            st = item.get("status", "VAŽEĆE")
            if st in counts: counts[st] += 1

        labels = list(counts.keys())
        values = list(counts.values())
        chart_colors = ['#16a34a', '#d97706', '#dc2626']

        fig, ax = plt.subplots(figsize=(6, 3))
        bars = ax.bar(labels, values, color=chart_colors, width=0.45)
        ax.set_ylabel('Broj opreme', fontsize=9)
        ax.set_title('Statusni pregled zaduženja LZO opreme', fontsize=10, fontweight='bold')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        for bar in bars:
            height = bar.get_height()
            ax.annotate(f'{height}',
                        xy=(bar.get_x() + bar.get_width() / 2, height),
                        xytext=(0, 3), textcoords="offset points",
                        ha='center', va='bottom', fontweight='bold', fontsize=9)

        plt.tight_layout()
        tmp_img = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
        plt.savefig(tmp_img.name, dpi=150)
        plt.close()
        return tmp_img.name

    def get_summary_percentages(self):
        total = len(self.current_filtered_data)
        if total == 0:
            return "Nema podataka", 0, 0, 0
        
        valid = sum(1 for x in self.current_filtered_data if x.get("status") == "VAŽEĆE")
        warning = sum(1 for x in self.current_filtered_data if x.get("status") == "USKORO")
        expired = sum(1 for x in self.current_filtered_data if x.get("status") == "ISTEKLO")

        p_valid = (valid / total) * 100
        p_warn = (warning / total) * 100
        p_exp = (expired / total) * 100

        summary_text = f"Ukupno: {total} | Važeće: {valid} ({p_valid:.1f}%) | Uskoro ističe: {warning} ({p_warn:.1f}%) | Isteklo: {expired} ({p_exp:.1f}%)"
        return summary_text, p_valid, p_warn, p_exp

    def export_excel_report(self):
        if not self.current_filtered_data:
            messagebox.showwarning("Upozorenje", "Nema podataka za izvoz.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile="Izvjestaj_LZO.xlsx"
        )
        if not save_path: return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "LZO Karton"
            ws.views.sheetView[0].showGridLines = True

            ws.merge_cells("A1:P1")
            title_cell = ws["A1"]
            title_cell.value = "IZVJEŠTAJ ZADUŽENJA LIČNE ZAŠTITNE OPREME (LZO)"
            title_cell.font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")

            summary_text, _, _, _ = self.get_summary_percentages()
            ws.merge_cells("A2:P2")
            sub_cell = ws["A2"]
            sub_cell.value = f"Datum: {datetime.now().strftime('%d.%m.%Y. u %H:%M')} | STATISTIKA: {summary_text}"
            sub_cell.font = Font(name="Calibri", size=10, italic=True, color="333333")
            sub_cell.alignment = Alignment(horizontal="center", vertical="center")

            headers = [
                "Ime i prezime", "Radno mjesto", "Org. jedinica", "Mjesto / Grad",
                "Vel. odjeća", "Vel. obuća", "Oprema", "J.M.", "Normativ", "Rok (mj)",
                "Izdata kol.", "Datum zaduženja", "Datum isticanja", "Preostalo dana", "Status", "Napomena"
            ]

            header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
            header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            thin_border = Border(
                left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
            )

            for col_num, header_title in enumerate(headers, 1):
                cell = ws.cell(row=4, column=col_num)
                cell.value = header_title
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            fill_expired = PatternFill(start_color="FECDD3", end_color="FECDD3", fill_type="solid")
            font_expired = Font(name="Calibri", size=9, color="9F1239", bold=True)
            fill_warning = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
            font_warning = Font(name="Calibri", size=9, color="854D0E", bold=True)
            fill_valid = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
            font_valid = Font(name="Calibri", size=9, color="166534")

            for r_idx, row in enumerate(self.current_filtered_data, start=5):
                d_zad = datetime.strptime(row["datum_zaduzenja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if row.get("datum_zaduzenja") else ""
                d_ist = datetime.strptime(row["datum_isticanja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if row.get("datum_isticanja") else ""

                values = [
                    row.get("zaposleni", ""), row.get("rm", ""), row.get("org_jedinica", ""), row.get("grad", ""),
                    row.get("vel_odjeca", ""), row.get("vel_obuca", ""), row.get("oprema", ""), row.get("jm", "KOM"),
                    row.get("normativ", 1), row.get("rok_mjeseci", 12), row.get("izdata_kol", 1),
                    d_zad, d_ist, row.get("preostalo_dana", "-"), row.get("status", ""), row.get("napomena", "")
                ]

                status_val = row.get("status", "")
                for c_idx, val in enumerate(values, start=1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    cell.font = Font(name="Calibri", size=9)
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="center" if c_idx in [5,6,8,9,10,11,12,13,14,15] else "left", vertical="center")

                    if c_idx == 15:
                        if status_val == "ISTEKLO": cell.fill = fill_expired; cell.font = font_expired
                        elif status_val == "USKORO": cell.fill = fill_warning; cell.font = font_warning
                        elif status_val == "VAŽEĆE": cell.fill = fill_valid; cell.font = font_valid

            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col if cell.row >= 4)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

            chart_img_path = self.generate_status_chart()
            img = OpenpyxlImage(chart_img_path)
            img.width = 450; img.height = 225
            ws.add_image(img, f"B{len(self.current_filtered_data) + 7}")

            wb.save(save_path)
            messagebox.showinfo("Uspjeh", "Excel izvještaj sa procentima i grafikonom je sačuvan!")
        except Exception as e:
            messagebox.showerror("Greška", f"Nije moguće sačuvati Excel fajl: {e}")

    def export_pdf_report(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Greška", "Biblioteka ReportLab nije instalirana ('pip install reportlab').")
            return

        if not self.current_filtered_data:
            messagebox.showwarning("Upozorenje", "Nema podataka za izvoz.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile="Izvjestaj_LZO.pdf"
        )
        if not save_path: return

        try:
            pdf_font = get_pdf_unicode_font()

            doc = SimpleDocTemplate(save_path, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
            styles = getSampleStyleSheet()

            title_style = ParagraphStyle(
                'TitleStyle', parent=styles['Heading1'], fontName=pdf_font,
                fontSize=14, textColor=colors.HexColor('#1E293B'), alignment=1, spaceAfter=4
            )
            sub_style = ParagraphStyle(
                'SubStyle', parent=styles['Normal'], fontName=pdf_font,
                fontSize=9, textColor=colors.HexColor('#475569'), alignment=1, spaceAfter=12
            )
            table_cell_style = ParagraphStyle(
                'TableCell', parent=styles['Normal'], fontName=pdf_font,
                fontSize=7, leading=8
            )
            table_header_style = ParagraphStyle(
                'TableHeader', parent=styles['Normal'], fontName=pdf_font,
                fontSize=7, leading=8, textColor=colors.white
            )

            story = []

            story.append(Paragraph("<b>IZVJEŠTAJ ZADUŽENJA LIČNE ZAŠTITNE OPREME (LZO)</b>", title_style))
            summary_text, _, _, _ = self.get_summary_percentages()
            story.append(Paragraph(f"Datum: {datetime.now().strftime('%d.%m.%Y. u %H:%M')} | STATISTIKA: {summary_text}", sub_style))
            story.append(Spacer(1, 10))

            headers = [
                "Zaposleni", "Radno mjesto", "Org. jed.", "Grad",
                "Odj/Obu", "Oprema", "J.M.", "Norm.", "Rok",
                "Kol.", "Zaduženo", "Ističe", "Preostalo", "Status"
            ]

            table_data = [[Paragraph(f"<b>{h}</b>", table_header_style) for h in headers]]

            for row in self.current_filtered_data:
                d_zad = datetime.strptime(row["datum_zaduzenja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if row.get("datum_zaduzenja") else ""
                d_ist = datetime.strptime(row["datum_isticanja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if row.get("datum_isticanja") else ""

                vel_info = f"{row.get('vel_odjeca','')}/{row.get('vel_obuca','')}"

                row_cells = [
                    Paragraph(str(row.get("zaposleni", "")), table_cell_style),
                    Paragraph(str(row.get("rm", "")), table_cell_style),
                    Paragraph(str(row.get("org_jedinica", "")), table_cell_style),
                    Paragraph(str(row.get("grad", "")), table_cell_style),
                    Paragraph(vel_info, table_cell_style),
                    Paragraph(str(row.get("oprema", "")), table_cell_style),
                    Paragraph(str(row.get("jm", "KOM")), table_cell_style),
                    Paragraph(str(row.get("normativ", 1)), table_cell_style),
                    Paragraph(str(row.get("rok_mjeseci", 12)), table_cell_style),
                    Paragraph(str(row.get("izdata_kol", 1)), table_cell_style),
                    Paragraph(d_zad, table_cell_style),
                    Paragraph(d_ist, table_cell_style),
                    Paragraph(str(row.get("preostalo_dana", "-")), table_cell_style),
                    Paragraph(str(row.get("status", "")), table_cell_style),
                ]
                table_data.append(row_cells)

            pdf_table = Table(table_data, repeatRows=1)
            pdf_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1E293B')),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#F8FAFC')]),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ]))

            story.append(pdf_table)
            story.append(Spacer(1, 15))

            chart_img_path = self.generate_status_chart()
            story.append(ReportLabImage(chart_img_path, width=400, height=200))

            doc.build(story)
            messagebox.showinfo("Uspjeh", "PDF izvještaj je uspješno generisan!")
        except Exception as e:
            messagebox.showerror("Greška", f"Nije moguće generisati PDF izvještaj: {e}")

    def open_add_dialog(self):
        dlg = RecordDialog(self.root, title="Novo zaduženje LZO")
        self.root.wait_window(dlg)
        if dlg.result:
            self.push_undo_state()
            new_id = max([r.get("id", 0) for r in self.data], default=0) + 1
            dlg.result["id"] = new_id
            self.data.append(dlg.result)
            self.recalculate_and_refresh()
            messagebox.showinfo("Uspjeh", "Novo zaduženje je uspješno dodato.")

    def open_edit_dialog(self, event=None):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Upozorenje", "Molimo vas da selektujete stavku za izmjenu.")
            return

        item_vals = self.tree.item(selected[0], "values")
        rec_id = item_vals[0]
        record = next((r for r in self.data if str(r.get("id")) == str(rec_id)), None)

        if not record:
            messagebox.showerror("Greška", "Selektovani zapis nije pronađen u bazi.")
            return

        dlg = RecordDialog(self.root, title="Izmjena zaduženja LZO", record=record)
        self.root.wait_window(dlg)
        if dlg.result:
            self.push_undo_state()
            for idx, r in enumerate(self.data):
                if str(r.get("id")) == str(rec_id):
                    self.data[idx] = dlg.result
                    break
            self.recalculate_and_refresh()
            messagebox.showinfo("Uspjeh", "Zapis je uspješno izmijenjen.")

    def delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Upozorenje", "Molimo vas da selektujete stavku ili više njih za brisanje.")
            return

        if not messagebox.askyesno("Potvrda brisanja", f"Da li ste sigurni da želite obrisati {len(selected)} selektovanih zapisa?"):
            return

        self.push_undo_state()
        ids_to_delete = {str(self.tree.item(s, "values")[0]) for s in selected}
        self.data = [r for r in self.data if str(r.get("id")) not in ids_to_delete]
        self.recalculate_and_refresh()
        messagebox.showinfo("Uspjeh", "Selektovani zapisi su obrisani.")

    def import_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
            self.push_undo_state()

            col_map = {
                "Zaposleni": "zaposleni", "Radno mjesto": "rm", "Org. jedinica": "org_jedinica",
                "Grad": "grad", "Mjesto": "grad", "Vel. odjeća": "vel_odjeca", "Vel. obuća": "vel_obuca",
                "Oprema": "oprema", "J.M.": "jm", "Normativ": "normativ", "Rok (mj)": "rok_mjeseci",
                "Izdata kol.": "izdata_kol", "Zaduženo": "datum_zaduzenja", "Napomena": "napomena"
            }

            next_id = max([r.get("id", 0) for r in self.data], default=0) + 1

            imported_count = 0
            for _, row in df.iterrows():
                rec = {"id": next_id}
                next_id += 1

                for col, key in col_map.items():
                    if col in row and pd.notna(row[col]):
                        rec[key] = str(row[col]).strip()
                    else:
                        rec.setdefault(key, "")

                if rec.get("datum_zaduzenja"):
                    try:
                        rec["datum_zaduzenja"] = pd.to_datetime(rec["datum_zaduzenja"]).strftime("%Y-%m-%d")
                    except Exception:
                        rec["datum_zaduzenja"] = datetime.now().strftime("%Y-%m-%d")

                try: rec["normativ"] = int(float(rec.get("normativ") or 1))
                except: rec["normativ"] = 1

                try: rec["rok_mjeseci"] = int(float(rec.get("rok_mjeseci") or 12))
                except: rec["rok_mjeseci"] = 12

                try: rec["izdata_kol"] = int(float(rec.get("izdata_kol") or 1))
                except: rec["izdata_kol"] = 1

                self.data.append(rec)
                imported_count += 1

            self.recalculate_and_refresh()
            messagebox.showinfo("Uspjeh", f"Uspješno uvezeno {imported_count} zapisa iz Excel fajla!")
        except Exception as e:
            messagebox.showerror("Greška", f"Greška pri uvozu Excel fajla: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = LZOApp(root)
    root.mainloop()
