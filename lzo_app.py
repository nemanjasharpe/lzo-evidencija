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
        """Generiše PNG sliku grafikona za ugradnju u Excel i PDF izvještaje."""
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

            # Naslov i procentualni rezime
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

            # Ugradnja grafikona
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
            # Registracija fonta koji podržava naša slova Č, Š, Ć, Đ, Ž
            pdf_font = get_pdf_unicode_font()

            doc = SimpleDocTemplate(save_path, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
            styles = getSampleStyleSheet()

            title_style = ParagraphStyle(
                'TitleStyle', parent=styles['Heading1'], fontName=pdf_font,
                fontSize=14, textColor=colors.HexColor('#1E293B'), alignment=1, spaceAfter=4
            )
            sub_style = ParagraphStyle(
                'SubStyle', parent=styles['Normal'], fontName=pdf_font,
                fontSize=9, textColor=colors.HexColor('#334155'), alignment=1, spaceAfter=10
            )
            cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontName=pdf_font, fontSize=7, leading=8)
            cell_bold = ParagraphStyle('CellBold', parent=styles['Normal'], fontName=pdf_font, fontSize=7, leading=8)

            elements = []
            elements.append(Paragraph("IZVJEŠTAJ ZADUŽENJA LIČNE ZAŠTITNE OPREME (LZO)", title_style))

            summary_text, _, _, _ = self.get_summary_percentages()
            elements.append(Paragraph(f"Datum generisanja: {datetime.now().strftime('%d.%m.%Y. u %H:%M')} | <b>STATISTIKA:</b> {summary_text}", sub_style))

            # Dodavanje grafikona
            chart_path = self.generate_status_chart()
            elements.append(ReportLabImage(chart_path, width=300, height=150))
            elements.append(Spacer(1, 8))

            headers = ["Zaposleni", "Radno mjesto", "Org. jedinica", "Mjesto", "Oprema", "J.M.", "Norm.", "Rok", "Izd.", "Zaduženo", "Ističe", "Preostalo", "Status"]
            table_data = [[Paragraph(f"<b>{h}</b>", ParagraphStyle('HStyle', parent=cell_bold, textColor=colors.white)) for h in headers]]

            for row in self.current_filtered_data:
                d_zad = datetime.strptime(row["datum_zaduzenja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if row.get("datum_zaduzenja") else ""
                d_ist = datetime.strptime(row["datum_isticanja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if row.get("datum_isticanja") else ""

                r_vals = [
                    Paragraph(row.get("zaposleni", ""), cell_style),
                    Paragraph(row.get("rm", ""), cell_style),
                    Paragraph(row.get("org_jedinica", ""), cell_style),
                    Paragraph(row.get("grad", ""), cell_style),
                    Paragraph(row.get("oprema", ""), cell_style),
                    Paragraph(str(row.get("jm", "KOM")), cell_style),
                    Paragraph(str(row.get("normativ", 1)), cell_style),
                    Paragraph(str(row.get("rok_mjeseci", 12)), cell_style),
                    Paragraph(str(row.get("izdata_kol", 1)), cell_style),
                    Paragraph(d_zad, cell_style),
                    Paragraph(d_ist, cell_style),
                    Paragraph(str(row.get("preostalo_dana", "-")), cell_style),
                    Paragraph(f"<b>{row.get('status', '')}</b>", cell_style)
                ]
                table_data.append(r_vals)

            col_w = [85, 75, 75, 55, 110, 30, 30, 28, 28, 55, 55, 45, 50]
            t = Table(table_data, colWidths=col_w, repeatRows=1)
            
            t_style = [
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E293B')),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
                ('TOPPADDING', (0,0), (-1,-1), 3),
                ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ]

            for idx, r_item in enumerate(self.current_filtered_data, start=1):
                st = r_item.get("status", "")
                if st == "ISTEKLO":
                    t_style.append(('BACKGROUND', (12, idx), (12, idx), colors.HexColor('#FECDD3')))
                elif st == "USKORO":
                    t_style.append(('BACKGROUND', (12, idx), (12, idx), colors.HexColor('#FEF08A')))
                elif st == "VAŽEĆE":
                    t_style.append(('BACKGROUND', (12, idx), (12, idx), colors.HexColor('#DCFCE7')))

            t.setStyle(TableStyle(t_style))
            elements.append(t)

            doc.build(elements)
            messagebox.showinfo("Uspjeh", "PDF izvještaj sa našim slovima, procentima i grafikonom je generisan!")
        except Exception as e:
            messagebox.showerror("Greška", f"Nije moguće generisati PDF fajl: {e}")

    def import_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if not file_path: return

        try:
            df = pd.read_excel(file_path)
            self.push_undo_state()  # Sačuvaj stanje za Undo

            # Mapiranje kolona
            new_items = []
            max_id = max([x.get("id", 0) for x in self.data], default=0)

            for idx, row in df.iterrows():
                max_id += 1
                d_zad = ""
                if pd.notnull(row.get("Datum zaduženja")):
                    try:
                        d_zad = pd.to_datetime(row.get("Datum zaduženja")).strftime("%Y-%m-%d")
                    except: pass

                item = {
                    "id": max_id,
                    "zaposleni": str(row.get("Ime i prezime", "")).strip(),
                    "rm": str(row.get("Radno mjesto", "")).strip(),
                    "org_jedinica": str(row.get("Org. jedinica", "")).strip(),
                    "grad": str(row.get("Mjesto / Grad", "")).strip(),
                    "vel_odjeca": str(row.get("Vel. odjeća", "")).strip(),
                    "vel_obuca": str(row.get("Vel. obuća", "")).strip(),
                    "oprema": str(row.get("Oprema", "")).strip(),
                    "jm": str(row.get("J.M.", "KOM")).strip(),
                    "normativ": int(row.get("Normativ", 1) or 1),
                    "rok_mjeseci": int(row.get("Rok (mj)", 12) or 12),
                    "izdata_kol": int(row.get("Izdata kol.", 1) or 1),
                    "datum_zaduzenja": d_zad,
                    "napomena": str(row.get("Napomena", "")).strip()
                }
                new_items.append(item)

            self.data.extend(new_items)
            self.recalculate_and_refresh()
            messagebox.showinfo("Uspjeh", f"Uspješno uvezeno {len(new_items)} zapisa iz Excel fajla!")
        except Exception as e:
            messagebox.showerror("Greška", f"Nije moguće uvoziti Excel fajl: {e}")

    def delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Upozorenje", "Izaberite jedan ili više zapisa za brisanje.")
            return

        if messagebox.askyesno("Potvrda brisanja", f"Da li ste sigurni da želite obrisati {len(selected)} selektovanih zapisa?"):
            self.push_undo_state()  # Sačuvaj stanje za Undo
            ids_to_delete = [self.tree.item(s)["values"][0] for s in selected]
            self.data = [x for x in self.data if x["id"] not in ids_to_delete]
            self.recalculate_and_refresh()

    def open_add_dialog(self):
        self.show_edit_window(title="Novo zaduženje LZO", item=None)

    def open_edit_dialog(self, event=None):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Upozorenje", "Izaberite zapis iz tabele koji želite izmijeniti.")
            return

        item_id = self.tree.item(selected[0])["values"][0]
        target_item = next((x for x in self.data if x["id"] == item_id), None)
        if target_item:
            self.show_edit_window(title="Izmjena zaduženja LZO", item=target_item)

    def show_edit_window(self, title, item=None):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry("520x640")
        win.configure(bg="#f8fafc")
        win.grab_set()

        fields = [
            ("Ime i Prezime:", "zaposleni"),
            ("Radno Mjesto:", "rm"),
            ("Organizaciona Jedinica:", "org_jedinica"),
            ("Mjesto / Grad:", "grad"),
            ("Veličina Odjeće:", "vel_odjeca"),
            ("Veličina Obuće:", "vel_obuca"),
            ("Naziv Opreme:", "oprema"),
            ("Jedinica Mjere:", "jm"),
            ("Normativ - Količina:", "normativ"),
            ("Rok u mjesecima:", "rok_mjeseci"),
            ("Izdata Količina:", "izdata_kol"),
            ("Datum Zaduženja (GGGG-MM-DD):", "datum_zaduzenja"),
            ("Napomena:", "napomena")
        ]

        entries = {}
        for idx, (label_text, key) in enumerate(fields):
            tk.Label(win, text=label_text, font=("Segoe UI", 9, "bold"), bg="#f8fafc", fg="#334155").grid(row=idx, column=0, sticky="w", padx=20, pady=4)
            entry = ttk.Entry(win, width=32)
            entry.grid(row=idx, column=1, padx=20, pady=4)

            if item:
                entry.insert(0, str(item.get(key, "")))
            entries[key] = entry

        def save():
            zaposleni = entries["zaposleni"].get().strip()
            oprema = entries["oprema"].get().strip()

            if not zaposleni or not oprema:
                messagebox.showerror("Greška", "Ime zaposlenog i naziv opreme su obavezni!", parent=win)
                return

            try:
                normativ = int(entries["normativ"].get().strip() or 1)
                rok = int(entries["rok_mjeseci"].get().strip() or 12)
                izdata_kol = int(entries["izdata_kol"].get().strip() or 1)
            except ValueError:
                messagebox.showerror("Greška", "Normativ, Rok i Izdata količina moraju biti cijeli brojevi!", parent=win)
                return

            d_zad = entries["datum_zaduzenja"].get().strip()
            if d_zad:
                try:
                    datetime.strptime(d_zad, "%Y-%m-%d")
                except ValueError:
                    messagebox.showerror("Greška", "Datum mora biti u formatu GGGG-MM-DD (npr. 2026-05-20)", parent=win)
                    return

            self.push_undo_state()  # Sačuvaj stanje za Undo

            if item:
                item["zaposleni"] = zaposleni
                item["rm"] = entries["rm"].get().strip()
                item["org_jedinica"] = entries["org_jedinica"].get().strip()
                item["grad"] = entries["grad"].get().strip()
                item["vel_odjeca"] = entries["vel_odjeca"].get().strip()
                item["vel_obuca"] = entries["vel_obuca"].get().strip()
                item["oprema"] = oprema
                item["jm"] = entries["jm"].get().strip() or "KOM"
                item["normativ"] = normativ
                item["rok_mjeseci"] = rok
                item["izdata_kol"] = izdata_kol
                item["datum_zaduzenja"] = d_zad
                item["napomena"] = entries["napomena"].get().strip()
            else:
                max_id = max([x.get("id", 0) for x in self.data], default=0) + 1
                new_item = {
                    "id": max_id,
                    "zaposleni": zaposleni,
                    "rm": entries["rm"].get().strip(),
                    "org_jedinica": entries["org_jedinica"].get().strip(),
                    "grad": entries["grad"].get().strip(),
                    "vel_odjeca": entries["vel_odjeca"].get().strip(),
                    "vel_obuca": entries["vel_obuca"].get().strip(),
                    "oprema": oprema,
                    "jm": entries["jm"].get().strip() or "KOM",
                    "normativ": normativ,
                    "rok_mjeseci": rok,
                    "izdata_kol": izdata_kol,
                    "datum_zaduzenja": d_zad,
                    "napomena": entries["napomena"].get().strip()
                }
                self.data.append(new_item)

            self.recalculate_and_refresh()
            win.destroy()

        btn_save = tk.Button(
            win, text="💾 Sačuvaj promjene", command=save,
            bg="#16a34a", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2", pady=6
        )
        btn_save.grid(row=len(fields), column=0, columnspan=2, fill=tk.X, padx=20, pady=15)


if __name__ == "__main__":
    root = tk.Tk()
    app = LZOApp(root)
    root.mainloop()
