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

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as ReportLabImage, KeepTogether
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

DB_FILE = "lzo_baza.json"

def get_cedis_logo_path():
    """Pronalazi lokalnu sliku CEDIS logotipa ako postoji u radnom direktorijumu."""
    for name in ["cedis_logo.png", "logo.png", "cedis_logo.jpg", "logo.jpg", "CEDIS_logo.png", "CEDIS.png"]:
        if os.path.exists(name):
            return name
    return None

def get_unicode_font_name():
    """Registruje font sa podrškom za regionalna slova (č, š, ć, đ, ž) u PDF-u."""
    if not REPORTLAB_AVAILABLE:
        return 'Helvetica'
    
    font_paths = [
        ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
        ("C:/Windows/Fonts/calibri.ttf", "C:/Windows/Fonts/calibrib.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf")
    ]
    
    for regular, bold in font_paths:
        if os.path.exists(regular):
            try:
                pdfmetrics.registerFont(TTFont('CustomUnicode', regular))
                if os.path.exists(bold):
                    pdfmetrics.registerFont(TTFont('CustomUnicode-Bold', bold))
                else:
                    pdfmetrics.registerFont(TTFont('CustomUnicode-Bold', regular))
                return 'CustomUnicode'
            except Exception:
                pass
    return 'Helvetica'

def sort_size_key(size_str):
    """Pomoćna funkcija za prirodno sortiranje veličina (brojevi pa tekstove)."""
    try:
        return (0, float(size_str.replace(',', '.')))
    except ValueError:
        order = {"XS": 1, "S": 2, "M": 3, "L": 4, "XL": 5, "XXL": 6, "3XL": 7, "4XL": 8, "5XL": 9}
        return (1, order.get(size_str.upper(), size_str))

class LZOApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CEDIS - Sistem za praćenje LZO i rokova zaduženja v6.0")
        self.root.geometry("1450x850")
        self.root.configure(bg="#F8FAFC")

        self.data = []
        self.current_filtered_data = []
        self.sort_directions = {}
        self.history = []  # Istorija za Undo (Ctrl+Z)

        self.pdf_font_name = get_unicode_font_name()
        self.setup_styles()
        self.create_widgets()
        self.load_local_db()

        # Prečica za tastaturu Ctrl+Z
        self.root.bind("<Control-z>", lambda e: self.undo())

    def setup_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.style.configure(
            "Treeview",
            background="#FFFFFF",
            foreground="#1E293B",
            rowheight=28,
            fieldbackground="#FFFFFF",
            font=("Segoe UI", 9)
        )
        self.style.configure(
            "Treeview.Heading",
            background="#0F172A",
            foreground="#FFFFFF",
            font=("Segoe UI", 9, "bold"),
            padding=6
        )
        self.style.map("Treeview.Heading", background=[('active', '#1E293B')])
        self.style.map("Treeview", background=[('selected', '#2563EB')], foreground=[('selected', '#FFFFFF')])
        self.style.configure("TCombobox", font=("Segoe UI", 9), padding=4)

    def push_undo_state(self):
        """Pamti trenutno stanje baze prije bilo kakve izmjene."""
        self.history.append(copy.deepcopy(self.data))
        if len(self.history) > 20:
            self.history.pop(0)
        self.btn_undo.config(state=tk.NORMAL)

    def undo(self, event=None):
        """Poništava poslednju akciju (Undo)."""
        if self.history:
            self.data = self.history.pop()
            self.recalculate_and_refresh()
            if not self.history:
                self.btn_undo.config(state=tk.DISABLED)
            self.lbl_status.config(text="↩️ Akcija uspješno poništena (Undo).", fg="#2563EB")
        else:
            messagebox.showinfo("Undo", "Nema prethodnih koraka za poništavanje.")

    def create_widgets(self):
        # Gornji panel sa akcionim dugmadima
        top_frame = tk.Frame(self.root, pady=10, padx=15, bg="#1E293B")
        top_frame.pack(fill=tk.X)

        tk.Label(top_frame, text="⚡ CEDIS LZO Sistem", font=("Segoe UI", 12, "bold"), fg="#F8FAFC", bg="#1E293B").pack(side=tk.LEFT, padx=(0, 15))

        btn_import = tk.Button(top_frame, text="📥 Uvoz iz Excel-a", command=self.import_excel, bg="#2563EB", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2")
        btn_import.pack(side=tk.LEFT, padx=3)

        btn_add = tk.Button(top_frame, text="+ Novo zaduženje", command=self.open_add_dialog, bg="#059669", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2")
        btn_add.pack(side=tk.LEFT, padx=3)

        btn_edit = tk.Button(top_frame, text="✏️ Izmijeni", command=self.open_edit_dialog, bg="#D97706", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2")
        btn_edit.pack(side=tk.LEFT, padx=3)

        btn_delete = tk.Button(top_frame, text="🗑️ Obriši selektovano", command=self.delete_selected, bg="#DC2626", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2")
        btn_delete.pack(side=tk.LEFT, padx=3)

        self.btn_undo = tk.Button(top_frame, text="↩️ Undo", command=self.undo, bg="#475569", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=8, pady=4, cursor="hand2", state=tk.DISABLED)
        self.btn_undo.pack(side=tk.LEFT, padx=(10, 3))

        # Novo dugme za Trebovanje / Specifikaciju opreme
        btn_trebovanje = tk.Button(top_frame, text="📋 Trebovanje i Specifikacija", command=self.open_trebovanje_dialog, bg="#7C3AED", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2")
        btn_trebovanje.pack(side=tk.LEFT, padx=(15, 3))

        btn_export_pdf = tk.Button(top_frame, text="📄 Izvezi PDF (+ %)", command=self.export_pdf_report, bg="#B91C1C", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2")
        btn_export_pdf.pack(side=tk.RIGHT, padx=3)

        btn_export_excel = tk.Button(top_frame, text="📊 Izvezi Excel (+ %)", command=self.export_excel_report, bg="#16A34A", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2")
        btn_export_excel.pack(side=tk.RIGHT, padx=3)

        # Sekcija za filtriranje
        filter_frame = tk.LabelFrame(self.root, text=" Pretraga i Filtriranje ", font=("Segoe UI", 9, "bold"), padx=12, pady=8, bg="#FFFFFF", fg="#1E293B", relief="solid", bd=1)
        filter_frame.pack(fill=tk.X, padx=15, pady=10)

        tk.Label(filter_frame, text="Pretraga (Zaposleni):", bg="#FFFFFF", font=("Segoe UI", 9, "bold"), fg="#334155").grid(row=0, column=0, sticky="w", padx=2)
        
        self.combo_search = ttk.Combobox(filter_frame, width=22, font=("Segoe UI", 9))
        self.combo_search.grid(row=0, column=1, padx=4, pady=2)
        self.combo_search.bind("<KeyRelease>", self.on_search_key_release)
        self.combo_search.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        tk.Label(filter_frame, text="Org. jedinica:", bg="#FFFFFF", font=("Segoe UI", 9), fg="#334155").grid(row=0, column=2, sticky="w", padx=(10, 2))
        self.combo_org = ttk.Combobox(filter_frame, state="readonly", width=16)
        self.combo_org.grid(row=0, column=3, padx=4, pady=2)
        self.combo_org.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        tk.Label(filter_frame, text="Mjesto / Grad:", bg="#FFFFFF", font=("Segoe UI", 9), fg="#334155").grid(row=0, column=4, sticky="w", padx=(10, 2))
        self.combo_city = ttk.Combobox(filter_frame, state="readonly", width=14)
        self.combo_city.grid(row=0, column=5, padx=4, pady=2)
        self.combo_city.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        tk.Label(filter_frame, text="Oprema:", bg="#FFFFFF", font=("Segoe UI", 9), fg="#334155").grid(row=0, column=6, sticky="w", padx=(10, 2))
        self.combo_equipment = ttk.Combobox(filter_frame, state="readonly", width=16)
        self.combo_equipment.grid(row=0, column=7, padx=4, pady=2)
        self.combo_equipment.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        tk.Label(filter_frame, text="Status:", bg="#FFFFFF", font=("Segoe UI", 9), fg="#334155").grid(row=0, column=8, sticky="w", padx=(10, 2))
        self.combo_status = ttk.Combobox(filter_frame, state="readonly", width=12, values=["SVI", "ISTEKLO", "USKORO", "VAŽEĆE"])
        self.combo_status.current(0)
        self.combo_status.grid(row=0, column=9, padx=4, pady=2)
        self.combo_status.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        btn_reset = tk.Button(filter_frame, text="Poništi filtere", command=self.reset_filters, font=("Segoe UI", 8, "bold"), bg="#E2E8F0", fg="#334155", relief="flat", padx=6, pady=2)
        btn_reset.grid(row=0, column=10, padx=(10, 2))

        # Statusna traka
        self.lbl_status = tk.Label(self.root, text="Inicijalizacija sistema...", font=("Segoe UI", 9, "italic"), anchor="w", padx=15, pady=6, bg="#E2E8F0", fg="#1E293B")
        self.lbl_status.pack(fill=tk.X, padx=15, pady=(0, 5))

        # Tabela (Treeview)
        table_frame = tk.Frame(self.root, bg="#F8FAFC")
        table_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        self.cols = (
            "ID", "Zaposleni", "Radno mjesto", "Org. jedinica", "Mjesto",
            "Vel. odjeća", "Vel. obuća", "Oprema", "J.M.", "Normativ",
            "Rok (mj)", "Izdata kol.", "Zaduženo", "Ističe", "Preostalo dana", "Status", "Napomena"
        )

        self.tree = ttk.Treeview(table_frame, columns=self.cols, show="headings", selectmode="extended")

        col_widths = {
            "ID": 30, "Zaposleni": 150, "Radno mjesto": 130, "Org. jedinica": 130, "Mjesto": 90,
            "Vel. odjeća": 70, "Vel. obuća": 70, "Oprema": 150, "J.M.": 50, "Normativ": 60,
            "Rok (mj)": 60, "Izdata kol.": 70, "Zaduženo": 85, "Ističe": 85, "Preostalo dana": 90,
            "Status": 85, "Napomena": 120
        }

        for col in self.cols:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by_column(c))
            center_cols = ["ID", "Vel. odjeća", "Vel. obuća", "J.M.", "Normativ", "Rok (mj)", "Izdata kol.", "Zaduženo", "Ističe", "Preostalo dana", "Status"]
            self.tree.column(col, width=col_widths.get(col, 100), anchor=tk.CENTER if col in center_cols else tk.W)

        self.tree.column("ID", width=0, stretch=False) # Skrivamo ID kolonu

        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)

        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # Stilovi za statuse (Boje redova)
        self.tree.tag_configure("ISTEKLO", background="#FEE2E2", foreground="#991B1B")
        self.tree.tag_configure("USKORO", background="#FEF3C7", foreground="#92400E")
        self.tree.tag_configure("VAŽEĆE", background="#DCFCE7", foreground="#166534")

        self.tree.bind("<Double-1>", self.open_edit_dialog)

    def on_search_key_release(self, event):
        if event.keysym in ["Up", "Down", "Return", "Escape", "Tab"]:
            return

        typed_text = self.combo_search.get().strip().lower()
        if typed_text:
            matching_names = sorted(list(set(
                str(row.get("zaposleni", "")) for row in self.data
                if typed_text in str(row.get("zaposleni", "")).lower()
            )))
            self.combo_search['values'] = matching_names
        else:
            self.combo_search['values'] = sorted(list(set(str(row.get("zaposleni", "")) for row in self.data if row.get("zaposleni"))))

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
            self.lbl_status.config(text="Lokalna baza nije pronađena. Uvezite podatke iz Excel fajla.", fg="#2563EB")

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

        total = len(self.data)
        p_exp = (expired_count / total * 100) if total else 0
        p_warn = (warning_count / total * 100) if total else 0
        p_val = (valid_count / total * 100) if total else 0

        msg = f"Ukupno stavki u bazi: {total} | Prikazano: {len(self.current_filtered_data)} | ISTEKLO: {expired_count} ({p_exp:.1f}%) | USKORO: {warning_count} ({p_warn:.1f}%) | VAŽEĆE: {valid_count} ({p_val:.1f}%)"
        self.lbl_status.config(text=msg, fg="#991B1B" if expired_count > 0 else "#1E293B")

    def update_filter_dropdowns(self):
        employees = sorted(list(set(str(row.get("zaposleni", "")) for row in self.data if row.get("zaposleni"))))
        orgs = sorted(list(set(row.get("org_jedinica", "") for row in self.data if row.get("org_jedinica"))))
        cities = sorted(list(set(row.get("grad", "") for row in self.data if row.get("grad"))))
        equipments = sorted(list(set(row.get("oprema", "") for row in self.data if row.get("oprema"))))

        self.combo_search['values'] = employees
        self.combo_org['values'] = ["SVE ORG. JEDINICE"] + orgs
        self.combo_city['values'] = ["SVA MJESTA"] + cities
        self.combo_equipment['values'] = ["SVA OPREMA"] + equipments

        if not self.combo_org.get(): self.combo_org.current(0)
        if not self.combo_city.get(): self.combo_city.current(0)
        if not self.combo_equipment.get(): self.combo_equipment.current(0)

    def apply_filters(self):
        search_txt = self.combo_search.get().strip().lower()
        selected_org = self.combo_org.get()
        selected_city = self.combo_city.get()
        selected_equip = self.combo_equipment.get()
        selected_status = self.combo_status.get()

        filtered = []
        for row in self.data:
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

            if match_search and match_org and match_city and match_equip and match_status:
                filtered.append(row)

        self.current_filtered_data = filtered
        self.refresh_table(filtered)

    def reset_filters(self):
        self.combo_search.set("")
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

    def import_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)
            self.push_undo_state()

            new_data = []
            max_id = max([r.get("id", 0) for r in self.data], default=0)

            for idx, row in df.iterrows():
                item = {
                    "id": max_id + idx + 1,
                    "zaposleni": str(row.get("Zaposleni", row.get("Ime i prezime", ""))).strip(),
                    "rm": str(row.get("Radno mjesto", "")).strip() if pd.notnull(row.get("Radno mjesto")) else "",
                    "org_jedinica": str(row.get("Org. jedinica", row.get("Organizaciona jedinica", ""))).strip() if pd.notnull(row.get("Org. jedinica")) else "",
                    "grad": str(row.get("Mjesto", row.get("Grad", ""))).strip() if pd.notnull(row.get("Mjesto")) else "",
                    "vel_odjeca": str(row.get("Vel. odjeća", row.get("Veličina odjeće", ""))).strip() if pd.notnull(row.get("Vel. odjeća")) else "",
                    "vel_obuca": str(row.get("Vel. obuća", row.get("Veličina obuće", ""))).strip() if pd.notnull(row.get("Vel. obuća")) else "",
                    "oprema": str(row.get("Oprema", row.get("Naziv opreme", ""))).strip() if pd.notnull(row.get("Oprema")) else "",
                    "jm": str(row.get("J.M.", row.get("Jedinica mjere", "KOM"))).strip() if pd.notnull(row.get("J.M.")) else "KOM",
                    "normativ": int(row.get("Normativ", 1)) if pd.notnull(row.get("Normativ")) else 1,
                    "rok_mjeseci": int(row.get("Rok (mj)", row.get("Rok mjeseci", 12))) if pd.notnull(row.get("Rok (mj)")) else 12,
                    "izdata_kol": int(row.get("Izdata kol.", row.get("Količina", 1))) if pd.notnull(row.get("Izdata kol.")) else 1,
                    "datum_zaduzenja": str(row.get("Datum zaduženja", row.get("Zaduženo", ""))).split(" ")[0] if pd.notnull(row.get("Datum zaduženja")) else "",
                    "napomena": str(row.get("Napomena", "")).strip() if pd.notnull(row.get("Napomena")) else ""
                }
                if item["zaposleni"] and item["oprema"]:
                    new_data.append(item)

            self.data.extend(new_data)
            self.recalculate_and_refresh()
            messagebox.showinfo("Uspjeh", f"Uspješno uvezeno {len(new_data)} novih zapisa iz Excel fajla!")
        except Exception as e:
            messagebox.showerror("Greška", f"Nije moguće uvesti Excel fajl: {e}")

    def open_add_dialog(self):
        self.open_record_dialog(title="Dodaj novo zaduženje LZO")

    def open_edit_dialog(self, event=None):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Upozorenje", "Selektujte red u tabeli koji želite izmijeniti.")
            return

        item_values = self.tree.item(selected_item[0], "values")
        rec_id = int(item_values[0])
        record = next((r for r in self.data if r.get("id") == rec_id), None)

        if record:
            self.open_record_dialog(title="Izmijeni zaduženje LZO", record=record)

    def open_record_dialog(self, title, record=None):
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.geometry("520x620")
        dlg.configure(bg="#F8FAFC")
        dlg.grab_set()

        fields = [
            ("Zaposleni (Ime i Prezime)*:", "zaposleni"),
            ("Radno mjesto:", "rm"),
            ("Organizaciona jedinica:", "org_jedinica"),
            ("Mjesto / Grad:", "grad"),
            ("Veličina odjeće:", "vel_odjeca"),
            ("Veličina obuće:", "vel_obuca"),
            ("Naziv opreme*:", "oprema"),
            ("Jedinica mjere (J.M.):", "jm"),
            ("Normativ (količina):", "normativ"),
            ("Rok trajanja (u mjesecima)*:", "rok_mjeseci"),
            ("Izdata količina*:", "izdata_kol"),
            ("Datum zaduženja (GGGG-MM-DD)*:", "datum_zaduzenja"),
            ("Napomena:", "napomena")
        ]

        entries = {}
        for idx, (label_text, key) in enumerate(fields):
            lbl = tk.Label(dlg, text=label_text, bg="#F8FAFC", font=("Segoe UI", 9, "bold" if "*" in label_text else "normal"), fg="#1E293B")
            lbl.grid(row=idx, column=0, sticky="w", padx=15, pady=4)

            ent = tk.Entry(dlg, font=("Segoe UI", 9), width=32)
            ent.grid(row=idx, column=1, padx=15, pady=4)

            val = str(record.get(key, "")) if record else ""
            if not record and key == "jm": val = "KOM"
            if not record and key == "normativ": val = "1"
            if not record and key == "rok_mjeseci": val = "12"
            if not record and key == "izdata_kol": val = "1"
            if not record and key == "datum_zaduzenja": val = datetime.now().strftime("%Y-%m-%d")

            ent.insert(0, val)
            entries[key] = ent

        def save_record():
            if not entries["zaposleni"].get().strip() or not entries["oprema"].get().strip() or not entries["datum_zaduzenja"].get().strip():
                messagebox.showerror("Greška", "Polja Zaposleni, Oprema i Datum zaduženja su obavezna!")
                return

            try:
                dt_str = entries["datum_zaduzenja"].get().strip()
                datetime.strptime(dt_str, "%Y-%m-%d")
            except ValueError:
                messagebox.showerror("Greška", "Datum zaduženja mora biti u formatu GGGG-MM-DD (npr. 2026-05-15).")
                return

            self.push_undo_state()

            if record:
                rec = record
            else:
                max_id = max([r.get("id", 0) for r in self.data], default=0)
                rec = {"id": max_id + 1}
                self.data.append(rec)

            rec["zaposleni"] = entries["zaposleni"].get().strip()
            rec["rm"] = entries["rm"].get().strip()
            rec["org_jedinica"] = entries["org_jedinica"].get().strip()
            rec["grad"] = entries["grad"].get().strip()
            rec["vel_odjeca"] = entries["vel_odjeca"].get().strip()
            rec["vel_obuca"] = entries["vel_obuca"].get().strip()
            rec["oprema"] = entries["oprema"].get().strip()
            rec["jm"] = entries["jm"].get().strip() or "KOM"
            rec["normativ"] = int(entries["normativ"].get().strip() or 1)
            rec["rok_mjeseci"] = int(entries["rok_mjeseci"].get().strip() or 12)
            rec["izdata_kol"] = int(entries["izdata_kol"].get().strip() or 1)
            rec["datum_zaduzenja"] = dt_str
            rec["napomena"] = entries["napomena"].get().strip()

            dlg.destroy()
            self.recalculate_and_refresh()
            messagebox.showinfo("Uspjeh", "Podaci uspješno sačuvani!")

        btn_save = tk.Button(dlg, text="💾 Sačuvaj podatak", command=save_record, bg="#059669", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", pady=6)
        btn_save.grid(row=len(fields), column=0, columnspan=2, pady=15, padx=15, sticky="ew")

    def delete_selected(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Upozorenje", "Selektujte bar jedan red u tabeli za brisanje.")
            return

        if messagebox.askyesno("Potvrda brisanja", f"Da li ste sigurni da želite obrisati selektovane stavke ({len(selected_items)})?"):
            self.push_undo_state()
            ids_to_del = [int(self.tree.item(item, "values")[0]) for item in selected_items]
            self.data = [r for r in self.data if r.get("id") not in ids_to_del]
            self.recalculate_and_refresh()
            messagebox.showinfo("Uspjeh", "Selektovane stavke su uspješno obrisane.")

    # ------------------- TREBOVANJE / SPECIFIKACIJA OPREME -------------------

    def generate_trebovanje_data(self, target_year, include_expired, org_unit="SVE ORG. JEDINICE"):
        """Filtrira opremu i pravi agregiranu specifikaciju po artiklima, veličinama i količinama."""
        summary = {}
        filtered_items = []

        for item in self.data:
            if org_unit != "SVE ORG. JEDINICE" and item.get("org_jedinica") != org_unit:
                continue

            d_ist = item.get("datum_isticanja", "")
            status = item.get("status", "")

            is_match = False
            if d_ist:
                try:
                    dt_ist = datetime.strptime(d_ist, "%Y-%m-%d")
                    if dt_ist.year == target_year:
                        is_match = True
                    elif include_expired and (status == "ISTEKLO" or dt_ist.year < target_year):
                        is_match = True
                except Exception:
                    pass
            elif include_expired and status == "ISTEKLO":
                is_match = True

            if is_match:
                filtered_items.append(item)
                eq_name = item.get("oprema", "Nedefinisana oprema").strip()
                qty = int(item.get("normativ", 1) or item.get("izdata_kol", 1) or 1)
                jm = item.get("jm", "KOM").strip()

                vel_odj = str(item.get("vel_odjeca", "")).strip()
                vel_obu = str(item.get("vel_obuca", "")).strip()

                if vel_obu and vel_obu != "-" and vel_obu.lower() != "nan":
                    size = vel_obu
                elif vel_odj and vel_odj != "-" and vel_odj.lower() != "nan":
                    size = vel_odj
                else:
                    size = "Standard"

                if eq_name not in summary:
                    summary[eq_name] = {
                        "jm": jm,
                        "sizes": {},
                        "total_qty": 0,
                        "items": []
                    }

                summary[eq_name]["sizes"][size] = summary[eq_name]["sizes"].get(size, 0) + qty
                summary[eq_name]["total_qty"] += qty
                summary[eq_name]["items"].append(item)

        return summary, filtered_items

    def open_trebovanje_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("📋 CEDIS - Izrada Trebovanja i Specifikacije LZO")
        dlg.geometry("980x680")
        dlg.configure(bg="#F8FAFC")
        dlg.grab_set()

        top_ctrl = tk.LabelFrame(dlg, text=" Parametri trebovanja ", font=("Segoe UI", 9, "bold"), bg="#FFFFFF", padx=12, pady=8)
        top_ctrl.pack(fill=tk.X, padx=15, pady=10)

        tk.Label(top_ctrl, text="Za godinu:", bg="#FFFFFF", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=5, sticky="w")
        current_year = datetime.now().year
        year_values = [str(y) for y in range(current_year - 2, current_year + 6)]
        combo_year = ttk.Combobox(top_ctrl, values=year_values, state="readonly", width=8)
        combo_year.set(str(current_year))
        combo_year.grid(row=0, column=1, padx=5)

        mode_var = tk.StringVar(value="YEAR_ONLY")
        rb1 = tk.Radiobutton(top_ctrl, text="Samo oprema koja ističe u izabranoj godini", variable=mode_var, value="YEAR_ONLY", bg="#FFFFFF", font=("Segoe UI", 9))
        rb1.grid(row=0, column=2, padx=10)

        rb2 = tk.Radiobutton(top_ctrl, text="Izabrana godina + SVA prethodno istekla oprema", variable=mode_var, value="WITH_EXPIRED", bg="#FFFFFF", font=("Segoe UI", 9, "bold"), fg="#B91C1C")
        rb2.grid(row=0, column=3, padx=10)

        tk.Label(top_ctrl, text="Org. jedinica:", bg="#FFFFFF", font=("Segoe UI", 9)).grid(row=1, column=0, padx=5, pady=5, sticky="w")
        orgs = sorted(list(set(r.get("org_jedinica", "") for r in self.data if r.get("org_jedinica"))))
        combo_org = ttk.Combobox(top_ctrl, values=["SVE ORG. JEDINICE"] + orgs, state="readonly", width=25)
        combo_org.current(0)
        combo_org.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky="w")

        # Prikaz specifikacije u tabeli
        tree_frame = tk.Frame(dlg, bg="#F8FAFC")
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        cols_spec = ("Oprema", "J.M.", "Specifikacija po veličinama", "Ukupno količina")
        tree_spec = ttk.Treeview(tree_frame, columns=cols_spec, show="headings")
        tree_spec.heading("Oprema", text="Naziv opreme / Artikal")
        tree_spec.heading("J.M.", text="J.M.")
        tree_spec.heading("Specifikacija po veličinama", text="Razrađene veličine i količine")
        tree_spec.heading("Ukupno količina", text="Ukupna količina za trebovanje")

        tree_spec.column("Oprema", width=220, anchor=tk.W)
        tree_spec.column("J.M.", width=60, anchor=tk.CENTER)
        tree_spec.column("Specifikacija po veličinama", width=480, anchor=tk.W)
        tree_spec.column("Ukupno količina", width=140, anchor=tk.CENTER)

        sc_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree_spec.yview)
        tree_spec.configure(yscroll=sc_y.set)
        sc_y.pack(side=tk.RIGHT, fill=tk.Y)
        tree_spec.pack(fill=tk.BOTH, expand=True)

        lbl_summary_info = tk.Label(dlg, text="", font=("Segoe UI", 9, "bold"), bg="#F8FAFC", fg="#1E293B", anchor="w")
        lbl_summary_info.pack(fill=tk.X, padx=15, pady=4)

        def refresh_trebovanje_preview():
            for item in tree_spec.get_children():
                tree_spec.delete(item)

            yr = int(combo_year.get())
            inc_exp = (mode_var.get() == "WITH_EXPIRED")
            org_sel = combo_org.get()

            summary, filtered = self.generate_trebovanje_data(yr, inc_exp, org_sel)

            grand_total_qty = 0
            for eq_name, details in sorted(summary.items()):
                sizes_dict = details["sizes"]
                sorted_sizes = sorted(sizes_dict.keys(), key=sort_size_key)
                size_str = ", ".join([f"vel. {s}: {sizes_dict[s]} {details['jm']}" for s in sorted_sizes])

                tree_spec.insert("", tk.END, values=(
                    eq_name,
                    details["jm"],
                    size_str,
                    f"{details['total_qty']} {details['jm']}"
                ))
                grand_total_qty += details["total_qty"]

            lbl_summary_info.config(text=f"📊 Ukupno artikala za trebovanje: {len(summary)} vrst(a) | Ukupna količina opreme: {grand_total_qty} kom/par | Obuhvaćeno radnika: {len(filtered)}")

        btn_calc = tk.Button(top_ctrl, text="🔄 Osvježi prikaz", command=refresh_trebovanje_preview, bg="#2563EB", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=2)
        btn_calc.grid(row=1, column=3, padx=10, pady=5)

        # Dugmad za izvoz Trebovanja
        bot_frame = tk.Frame(dlg, bg="#F8FAFC", pady=10)
        bot_frame.pack(fill=tk.X, padx=15)

        def export_treb_excel():
            yr = int(combo_year.get())
            inc_exp = (mode_var.get() == "WITH_EXPIRED")
            org_sel = combo_org.get()
            summary, filtered = self.generate_trebovanje_data(yr, inc_exp, org_sel)
            self.export_trebovanje_excel_file(yr, inc_exp, org_sel, summary, filtered)

        def export_treb_pdf():
            yr = int(combo_year.get())
            inc_exp = (mode_var.get() == "WITH_EXPIRED")
            org_sel = combo_org.get()
            summary, filtered = self.generate_trebovanje_data(yr, inc_exp, org_sel)
            self.export_trebovanje_pdf_file(yr, inc_exp, org_sel, summary, filtered)

        btn_exp_ex = tk.Button(bot_frame, text="📊 Izvezi Trebovanje u Excel", command=export_treb_excel, bg="#16A34A", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", padx=15, pady=6)
        btn_exp_ex.pack(side=tk.LEFT, padx=5)

        btn_exp_pdf = tk.Button(bot_frame, text="📄 Izvezi Trebovanje u PDF", command=export_treb_pdf, bg="#B91C1C", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", padx=15, pady=6)
        btn_exp_pdf.pack(side=tk.LEFT, padx=5)

        refresh_trebovanje_preview()

    def export_trebovanje_excel_file(self, year, include_expired, org_unit, summary, filtered):
        if not summary:
            messagebox.showwarning("Upozorenje", "Nema podataka za trebovanje po izabranim kriterijumima.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=f"Trebovanje_LZO_{year}_CEDIS.xlsx"
        )
        if not save_path: return

        try:
            wb = openpyxl.Workbook()
            
            # List 1: Sumarna specifikacija po artiklima i veličinama
            ws1 = wb.active
            ws1.title = "Sumarna Specifikacija"
            ws1.views.sheetView[0].showGridLines = True

            # CEDIS Logo if available
            logo_path = get_cedis_logo_path()
            if logo_path:
                try:
                    img = OpenpyxlImage(logo_path)
                    img.width = 130
                    img.height = 45
                    ws1.add_image(img, "A1")
                except Exception:
                    pass

            ws1.merge_cells("C1:F1")
            t_cell = ws1["C1"]
            t_cell.value = "CRNOGORSKI ELEKTRODISTRIBUTIVNI SISTEM - CEDIS"
            t_cell.font = Font(name="Calibri", size=13, bold=True, color="0F172A")

            ws1.merge_cells("C2:F2")
            sub_cell = ws1["C2"]
            title_mode = f"SPECIFIKACIJA TREBOVANJA LZO ZA {year}. GODINU" + (" (SA ISTEKLIM ROKOVIMA)" if include_expired else "")
            sub_cell.value = title_mode
            sub_cell.font = Font(name="Calibri", size=11, bold=True, color="2563EB")

            ws1.cell(row=3, column=1, value=f"Organizaciona jedinica: {org_unit} | Datum generisanja: {datetime.now().strftime('%d.%m.%Y. u %H:%M')}").font = Font(size=9, italic=True)

            headers = ["R.br.", "Naziv opreme / Artikal", "J.M.", "Specifikacija po veličinama i količinama", "Ukupno za trebovanje"]
            header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
            header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
            thin_border = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))

            for c_idx, h_text in enumerate(headers, 1):
                cell = ws1.cell(row=5, column=c_idx, value=h_text)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            r_idx = 6
            grand_total = 0
            for idx, (eq_name, details) in enumerate(sorted(summary.items()), 1):
                sizes_dict = details["sizes"]
                sorted_sizes = sorted(sizes_dict.keys(), key=sort_size_key)
                size_str = "; ".join([f"vel. {s}: {sizes_dict[s]} {details['jm']}" for s in sorted_sizes])

                ws1.cell(row=r_idx, column=1, value=idx).alignment = Alignment(horizontal="center")
                ws1.cell(row=r_idx, column=2, value=eq_name)
                ws1.cell(row=r_idx, column=3, value=details["jm"]).alignment = Alignment(horizontal="center")
                ws1.cell(row=r_idx, column=4, value=size_str)
                ws1.cell(row=r_idx, column=5, value=details["total_qty"]).alignment = Alignment(horizontal="center")

                for c in range(1, 6):
                    ws1.cell(row=r_idx, column=c).border = thin_border
                    ws1.cell(row=r_idx, column=c).font = Font(name="Calibri", size=9.5)

                grand_total += details["total_qty"]
                r_idx += 1

            # Ukupan red
            ws1.cell(row=r_idx, column=2, value="UKUPNO SVE STAVKE:").font = Font(name="Calibri", size=10, bold=True)
            tot_cell = ws1.cell(row=r_idx, column=5, value=grand_total)
            tot_cell.font = Font(name="Calibri", size=10, bold=True, color="1E293B")
            tot_cell.alignment = Alignment(horizontal="center")
            tot_cell.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

            for col in ws1.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws1.column_dimensions[col_letter].width = max(max_len + 3, 12)

            # List 2: Pojedinačna specifikacija po zaposlenima
            ws2 = wb.create_sheet(title="Pojedinačno po zaposlenima")
            ws2.views.sheetView[0].showGridLines = True

            headers_det = ["R.br.", "Ime i prezime", "Radno mjesto", "Org. jedinica", "Mjesto", "Oprema", "Vel. odjeća", "Vel. obuća", "Količina", "J.M.", "Ističe / Isteklo"]
            for c_idx, h_text in enumerate(headers_det, 1):
                cell = ws2.cell(row=1, column=c_idx, value=h_text)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            for r_i, item in enumerate(filtered, 2):
                d_ist = datetime.strptime(item["datum_isticanja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if item.get("datum_isticanja") else ""
                vals = [
                    r_i - 1, item.get("zaposleni", ""), item.get("rm", ""), item.get("org_jedinica", ""), item.get("grad", ""),
                    item.get("oprema", ""), item.get("vel_odjeca", ""), item.get("vel_obuca", ""),
                    item.get("normativ", 1) or item.get("izdata_kol", 1), item.get("jm", "KOM"), d_ist
                ]
                for c_i, val in enumerate(vals, 1):
                    cell = ws2.cell(row=r_i, column=c_i, value=val)
                    cell.border = thin_border
                    cell.font = Font(name="Calibri", size=9)
                    if c_i in [1, 7, 8, 9, 10, 11]:
                        cell.alignment = Alignment(horizontal="center")

            for col in ws2.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws2.column_dimensions[col_letter].width = max(max_len + 3, 10)

            wb.save(save_path)
            messagebox.showinfo("Uspjeh", "Specifikacija trebovanja je uspješno sačuvana u Excel fajl!")
        except Exception as e:
            messagebox.showerror("Greška", f"Nije moguće sačuvati Trebovanje u Excel: {e}")

    def export_trebovanje_pdf_file(self, year, include_expired, org_unit, summary, filtered):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Greška", "ReportLab nije instaliran za PDF izvoz.")
            return

        if not summary:
            messagebox.showwarning("Upozorenje", "Nema podataka za trebovanje po izabranim kriterijumima.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile=f"Trebovanje_LZO_{year}_CEDIS.pdf"
        )
        if not save_path: return

        try:
            doc = SimpleDocTemplate(save_path, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
            styles = getSampleStyleSheet()

            font_name = self.pdf_font_name
            font_bold = 'CustomUnicode-Bold' if font_name == 'CustomUnicode' else 'Helvetica-Bold'

            title_style = ParagraphStyle('TStyle', parent=styles['Heading1'], fontName=font_bold, fontSize=13, textColor=colors.HexColor('#0F172A'), alignment=0)
            sub_style = ParagraphStyle('SStyle', parent=styles['Normal'], fontName=font_bold, fontSize=10, textColor=colors.HexColor('#2563EB'), alignment=0)
            meta_style = ParagraphStyle('MStyle', parent=styles['Normal'], fontName=font_name, fontSize=8, textColor=colors.HexColor('#475569'), alignment=0)
            cell_style = ParagraphStyle('CStyle', parent=styles['Normal'], fontName=font_name, fontSize=8, leading=9)
            cell_bold = ParagraphStyle('CBStyle', parent=styles['Normal'], fontName=font_bold, fontSize=8, leading=9)

            elements = []

            # Zaglavlje sa CEDIS logotipom
            logo_path = get_cedis_logo_path()
            title_text = "CRNOGORSKI ELEKTRODISTRIBUTIVNI SISTEM\nSLUŽBA ZA ZZNR I LABORATORIJSKA ISPITIVANJA"
            title_p = Paragraph(f"<b>{title_text.replace(chr(10), '<br/>')}</b>", title_style)
            sub_title_p = Paragraph(f"<b>SPECIFIKACIJA TREBOVANJA OPREME ZA {year}. GODINU</b>" + (" (SA ISTEKLIM ROKOVIMA)" if include_expired else ""), sub_style)
            meta_p = Paragraph(f"Organizaciona jedinica: <b>{org_unit}</b> | Datum izrade: <b>{datetime.now().strftime('%d.%m.%Y. u %H:%M')}</b>", meta_style)

            if logo_path:
                try:
                    img_logo = ReportLabImage(logo_path, width=120, height=45)
                    header_table = Table([[img_logo, [title_p, Spacer(1, 3), sub_title_p, meta_p]]], colWidths=[130, 670])
                    header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
                    elements.append(header_table)
                except Exception:
                    elements.append(title_p); elements.append(sub_style); elements.append(meta_p)
            else:
                elements.append(title_p); elements.append(sub_title_p); elements.append(meta_p)

            elements.append(Spacer(1, 12))

            # Tabela 1: Sumarna specifikacija po artiklima i veličinama
            elements.append(Paragraph("<b>1. Sumarna specifikacija trebovanja po artiklima i veličinama:</b>", ParagraphStyle('SecStyle', parent=styles['Normal'], fontName=font_bold, fontSize=10, textColor=colors.HexColor('#0F172A'))))
            elements.append(Spacer(1, 5))

            headers_t1 = ["R.br.", "Naziv opreme / Artikal", "J.M.", "Razrađene veličine i količine za trebovanje", "Ukupno"]
            table_data_1 = [[Paragraph(f"<b>{h}</b>", ParagraphStyle('H1', parent=cell_bold, textColor=colors.white)) for h in headers_t1]]

            grand_total = 0
            for idx, (eq_name, details) in enumerate(sorted(summary.items()), 1):
                sizes_dict = details["sizes"]
                sorted_sizes = sorted(sizes_dict.keys(), key=sort_size_key)
                size_str = "; ".join([f"vel. <b>{s}</b>: {sizes_dict[s]} {details['jm']}" for s in sorted_sizes])

                table_data_1.append([
                    Paragraph(str(idx), cell_style),
                    Paragraph(eq_name, cell_bold),
                    Paragraph(details["jm"], cell_style),
                    Paragraph(size_str, cell_style),
                    Paragraph(f"<b>{details['total_qty']}</b>", cell_bold)
                ])
                grand_total += details["total_qty"]

            table_data_1.append([
                Paragraph("", cell_style),
                Paragraph("<b>UKUPNO SVE STAVKE:</b>", cell_bold),
                Paragraph("", cell_style),
                Paragraph("", cell_style),
                Paragraph(f"<b>{grand_total}</b>", cell_bold)
            ])

            t1 = Table(table_data_1, colWidths=[35, 210, 45, 430, 80])
            t1.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('ALIGN', (0,0), (0,-1), 'CENTER'),
                ('ALIGN', (2,0), (2,-1), 'CENTER'),
                ('ALIGN', (4,0), (4,-1), 'CENTER'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
                ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#F1F5F9')),
            ]))
            elements.append(t1)

            elements.append(Spacer(1, 15))

            # Potpisi na dnu
            sig_data = [
                [Paragraph("<b>Specifikaciju izradio:</b>", cell_style), Paragraph("<b>Odobrio rukovodilac:</b>", cell_style)],
                [Spacer(1, 25), Spacer(1, 25)],
                [Paragraph("___________________________", cell_style), Paragraph("___________________________", cell_style)]
            ]
            sig_table = Table(sig_data, colWidths=[400, 400])
            elements.append(KeepTogether([sig_table]))

            doc.build(elements)
            messagebox.showinfo("Uspjeh", "Specifikacija trebovanja je uspješno sačuvana u PDF fajl!")
        except Exception as e:
            messagebox.showerror("Greška", f"Nije moguće generisati PDF Trebovanje: {e}")

    # ------------------- GLAVNI IZVEŠTAJI (EXCEL & PDF) -------------------

    def generate_status_chart(self):
        counts = {"VAŽEĆE": 0, "USKORO": 0, "ISTEKLO": 0}
        for item in self.current_filtered_data:
            st = item.get("status", "VAŽEĆE")
            if st in counts: counts[st] += 1

        total = sum(counts.values())
        labels = list(counts.keys())
        values = list(counts.values())
        chart_colors = ['#16A34A', '#D97706', '#DC2626']

        fig, ax = plt.subplots(figsize=(6.5, 3.0))
        bars = ax.bar(labels, values, color=chart_colors, width=0.45)
        ax.set_ylabel('Broj stavki', fontsize=8)
        ax.set_title('Statusni pregled LZO opreme CEDIS', fontsize=9, fontweight='bold')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        for bar in bars:
            height = bar.get_height()
            percentage = (height / total * 100) if total > 0 else 0
            ax.annotate(f'{height}\n({percentage:.1f}%)',
                        xy=(bar.get_x() + bar.get_width() / 2, height),
                        xytext=(0, 3), textcoords="offset points",
                        ha='center', va='bottom', fontweight='bold', fontsize=8)

        plt.tight_layout()
        tmp_img = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
        plt.savefig(tmp_img.name, dpi=150)
        plt.close()
        return tmp_img.name

    def export_excel_report(self):
        if not self.current_filtered_data:
            messagebox.showwarning("Upozorenje", "Nema podataka za izvoz.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile="Izvjestaj_LZO_CEDIS.xlsx"
        )
        if not save_path: return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "LZO Karton"
            ws.views.sheetView[0].showGridLines = True

            total = len(self.current_filtered_data)
            exp = sum(1 for r in self.current_filtered_data if r.get("status") == "ISTEKLO")
            warn = sum(1 for r in self.current_filtered_data if r.get("status") == "USKORO")
            val = sum(1 for r in self.current_filtered_data if r.get("status") == "VAŽEĆE")

            p_exp = (exp / total * 100) if total else 0
            p_warn = (warn / total * 100) if total else 0
            p_val = (val / total * 100) if total else 0

            # Umetanje logoa ako postoji
            logo_path = get_cedis_logo_path()
            if logo_path:
                try:
                    img_logo = OpenpyxlImage(logo_path)
                    img_logo.width = 120
                    img_logo.height = 40
                    ws.add_image(img_logo, "A1")
                except Exception:
                    pass

            ws.merge_cells("C1:P1")
            title_cell = ws["C1"]
            title_cell.value = "CRNOGORSKI ELEKTRODISTRIBUTIVNI SISTEM - IZVJEŠTAJ ZADUŽENJA LZO"
            title_cell.font = Font(name="Calibri", size=13, bold=True, color="0F172A")
            title_cell.alignment = Alignment(horizontal="left", vertical="center")

            ws.merge_cells("C2:P2")
            sub_cell = ws["C2"]
            sub_cell.value = f"Datum: {datetime.now().strftime('%d.%m.%Y. u %H:%M')} | Ukupno stavki: {total} | Važeće: {val} ({p_val:.1f}%) | Uskoro ističe: {warn} ({p_warn:.1f}%) | Isteklo: {exp} ({p_exp:.1f}%)"
            sub_cell.font = Font(name="Calibri", size=9.5, italic=True, color="475569")
            sub_cell.alignment = Alignment(horizontal="left", vertical="center")

            headers = [
                "Ime i prezime", "Radno mjesto", "Org. jedinica", "Mjesto / Grad",
                "Vel. odjeća", "Vel. obuća", "Oprema", "J.M.", "Normativ", "Rok (mj)",
                "Izdata kol.", "Datum zaduženja", "Datum isticanja", "Preostalo dana", "Status", "Napomena"
            ]

            header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
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
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = thin_border

            fill_expired = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            font_expired = Font(name="Calibri", size=9, color="9C0006", bold=True)
            fill_warning = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            font_warning = Font(name="Calibri", size=9, color="9C6500", bold=True)
            fill_valid = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            font_valid = Font(name="Calibri", size=9, color="006100")
            font_regular = Font(name="Calibri", size=9)

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
                for c_idx, val_item in enumerate(values, start=1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=val_item)
                    cell.font = font_regular
                    cell.border = thin_border

                    if c_idx in [5, 6, 8, 9, 10, 11, 12, 13, 14, 15]:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

                    if c_idx == 15:
                        if status_val == "ISTEKLO":
                            cell.fill = fill_expired; cell.font = font_expired
                        elif status_val == "USKORO":
                            cell.fill = fill_warning; cell.font = font_warning
                        elif status_val == "VAŽEĆE":
                            cell.fill = fill_valid; cell.font = font_valid

            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.row < 4: continue
                    val_str = str(cell.value or '')
                    if len(val_str) > max_len: max_len = len(val_str)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

            chart_img_path = self.generate_status_chart()
            img = OpenpyxlImage(chart_img_path)
            img.width = 450
            img.height = 210
            chart_row = len(self.current_filtered_data) + 7
            ws.add_image(img, f"B{chart_row}")

            wb.save(save_path)
            messagebox.showinfo("Uspjeh", "Excel izvještaj CEDIS sa procentima i dijagramom je uspešno sačuvan!")
        except Exception as e:
            messagebox.showerror("Greška", f"Nije moguće sačuvati Excel fajl: {e}")

    def export_pdf_report(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Greška", "Biblioteka ReportLab nije instalirana.")
            return

        if not self.current_filtered_data:
            messagebox.showwarning("Upozorenje", "Nema podataka za izvoz.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            initialfile="Izvjestaj_LZO_CEDIS.pdf"
        )
        if not save_path: return

        try:
            doc = SimpleDocTemplate(save_path, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
            styles = getSampleStyleSheet()

            total = len(self.current_filtered_data)
            exp = sum(1 for r in self.current_filtered_data if r.get("status") == "ISTEKLO")
            warn = sum(1 for r in self.current_filtered_data if r.get("status") == "USKORO")
            val = sum(1 for r in self.current_filtered_data if r.get("status") == "VAŽEĆE")

            p_exp = (exp / total * 100) if total else 0
            p_warn = (warn / total * 100) if total else 0
            p_val = (val / total * 100) if total else 0

            font_name = self.pdf_font_name
            font_bold = 'CustomUnicode-Bold' if font_name == 'CustomUnicode' else 'Helvetica-Bold'

            title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName=font_bold, fontSize=13, textColor=colors.HexColor('#0F172A'), alignment=0)
            sub_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontName=font_name, fontSize=8.5, textColor=colors.HexColor('#475569'), alignment=0)
            cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontName=font_name, fontSize=7, leading=8)
            cell_bold = ParagraphStyle('CellBold', parent=styles['Normal'], fontName=font_bold, fontSize=7, leading=8)

            elements = []

            # CEDIS Logo Zaglavlje
            logo_path = get_cedis_logo_path()
            title_text = "CRNOGORSKI ELEKTRODISTRIBUTIVNI SISTEM - CEDIS\nIZVJEŠTAJ ZADUŽENJA LIČNE ZAŠTITNE OPREME"
            title_p = Paragraph(f"<b>{title_text.replace(chr(10), '<br/>')}</b>", title_style)
            sub_text = f"Datum: {datetime.now().strftime('%d.%m.%Y. u %H:%M')} | Ukupno: {total} | Važeće: {val} ({p_val:.1f}%) | Uskoro ističe: {warn} ({p_warn:.1f}%) | Isteklo: {exp} ({p_exp:.1f}%)"
            sub_p = Paragraph(sub_text, sub_style)

            if logo_path:
                try:
                    img_logo = ReportLabImage(logo_path, width=120, height=42)
                    hdr_table = Table([[img_logo, [title_p, Spacer(1, 4), sub_p]]], colWidths=[130, 670])
                    hdr_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
                    elements.append(hdr_table)
                except Exception:
                    elements.append(title_p); elements.append(sub_p)
            else:
                elements.append(title_p); elements.append(sub_p)

            elements.append(Spacer(1, 10))

            chart_path = self.generate_status_chart()
            elements.append(ReportLabImage(chart_path, width=300, height=140))
            elements.append(Spacer(1, 10))

            headers = ["Zaposleni", "Radno mjesto", "Org. jedinica", "Mjesto", "Oprema", "J.M.", "Norm.", "Rok", "Izd.", "Zaduženo", "Ističe", "Preostalo", "Status"]
            table_data = [[Paragraph(f"<b>{h}</b>", ParagraphStyle('HStyle', parent=cell_bold, textColor=colors.white)) for h in headers]]

            for row in self.current_filtered_data:
                d_zad = datetime.strptime(row["datum_zaduzenja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if row.get("datum_zaduzenja") else ""
                d_ist = datetime.strptime(row["datum_isticanja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if row.get("datum_isticanja") else ""

                st = row.get("status", "")
                st_color = "#166534" if st == "VAŽEĆE" else ("#92400E" if st == "USKORO" else "#991B1B")

                table_data.append([
                    Paragraph(row.get("zaposleni", ""), cell_bold),
                    Paragraph(row.get("rm", ""), cell_style),
                    Paragraph(row.get("org_jedinica", ""), cell_style),
                    Paragraph(row.get("grad", ""), cell_style),
                    Paragraph(row.get("oprema", ""), cell_style),
                    Paragraph(row.get("jm", "KOM"), cell_style),
                    Paragraph(str(row.get("normativ", 1)), cell_style),
                    Paragraph(str(row.get("rok_mjeseci", 12)), cell_style),
                    Paragraph(str(row.get("izdata_kol", 1)), cell_style),
                    Paragraph(d_zad, cell_style),
                    Paragraph(d_ist, cell_style),
                    Paragraph(str(row.get("preostalo_dana", "-")), cell_style),
                    Paragraph(f"<font color='{st_color}'><b>{st}</b></font>", cell_bold)
                ])

            t = Table(table_data, colWidths=[90, 80, 80, 55, 95, 30, 35, 30, 30, 55, 55, 50, 55])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')])
            ]))

            elements.append(t)
            doc.build(elements)
            messagebox.showinfo("Uspjeh", "PDF izvještaj CEDIS sa dijagramom je sačuvan!")
        except Exception as e:
            messagebox.showerror("Greška", f"Nije moguće generisati PDF izvještaj: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = LZOApp(root)
    root.mainloop()
