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
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as ReportLabImage
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# Pokušaj uvoza PyMuPDF za konverziju PDF logotipa u PNG
try:
    import fitz 
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False

DB_FILE = "lzo_baza.json"
LOGO_PDF_FILE = "cedis_logo.pdf"
LOGO_PNG_FILE = "cedis_logo.png"

def ensure_logo_png():
    """Provjerava i konvertuje PDF logo u PNG format ako je potrebno."""
    if os.path.exists(LOGO_PNG_FILE):
        return LOGO_PNG_FILE
    
    if os.path.exists(LOGO_PDF_FILE) and FITZ_AVAILABLE:
        try:
            doc = fitz.open(LOGO_PDF_FILE)
            page = doc[0]
            pix = page.get_pixmap(dpi=150)
            pix.save(LOGO_PNG_FILE)
            return LOGO_PNG_FILE
        except Exception:
            pass
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

class LZOApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistem za praćenje LZO i rokova zaduženja v5.0")
        self.root.geometry("1420x820")
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
        if len(self.history) > 20: # Ograničavamo na 20 koraka
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

        tk.Label(top_frame, text="🛡️ LZO Sistem", font=("Segoe UI", 12, "bold"), fg="#F8FAFC", bg="#1E293B").pack(side=tk.LEFT, padx=(0, 15))

        btn_import = tk.Button(top_frame, text="📥 Uvoz iz Excel-a", command=self.import_excel, bg="#2563EB", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2")
        btn_import.pack(side=tk.LEFT, padx=4)

        btn_add = tk.Button(top_frame, text="+ Novo zaduženje", command=self.open_add_dialog, bg="#059669", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2")
        btn_add.pack(side=tk.LEFT, padx=4)

        btn_edit = tk.Button(top_frame, text="✏️ Izmijeni", command=self.open_edit_dialog, bg="#D97706", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2")
        btn_edit.pack(side=tk.LEFT, padx=4)

        btn_delete = tk.Button(top_frame, text="🗑️ Obriši selektovano", command=self.delete_selected, bg="#DC2626", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2")
        btn_delete.pack(side=tk.LEFT, padx=4)

        self.btn_undo = tk.Button(top_frame, text="↩️ Poništi (Undo)", command=self.undo, bg="#475569", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2", state=tk.DISABLED)
        self.btn_undo.pack(side=tk.LEFT, padx=(15, 4))

        btn_export_pdf = tk.Button(top_frame, text="📄 Izvezi PDF (+ %)", command=self.export_pdf_report, bg="#B91C1C", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2")
        btn_export_pdf.pack(side=tk.RIGHT, padx=4)

        btn_export_excel = tk.Button(top_frame, text="📊 Izvezi Excel (+ %)", command=self.export_excel_report, bg="#16A34A", fg="white", font=("Segoe UI", 9, "bold"), relief="flat", padx=10, pady=4, cursor="hand2")
        btn_export_excel.pack(side=tk.RIGHT, padx=4)

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

    def import_excel(self):
        """Pomoćna metoda za uvoz podataka iz Excel fajla."""
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_path: return
        try:
            df = pd.read_excel(file_path)
            self.push_undo_state()
            new_data = []
            for idx, row in df.iterrows():
                item = {
                    "id": idx + 1,
                    "zaposleni": str(row.get("Zaposleni", "")),
                    "rm": str(row.get("Radno mjesto", "")),
                    "org_jedinica": str(row.get("Org. jedinica", "")),
                    "grad": str(row.get("Mjesto", "")),
                    "vel_odjeca": str(row.get("Vel. odjeća", "")),
                    "vel_obuca": str(row.get("Vel. obuća", "")),
                    "oprema": str(row.get("Oprema", "")),
                    "jm": str(row.get("J.M.", "KOM")),
                    "normativ": row.get("Normativ", 1),
                    "rok_mjeseci": row.get("Rok (mj)", 12),
                    "izdata_kol": row.get("Izdata kol.", 1),
                    "datum_zaduzenja": str(row.get("Zaduženo", ""))[:10] if pd.notnull(row.get("Zaduženo")) else "",
                    "napomena": str(row.get("Napomena", "")) if pd.notnull(row.get("Napomena")) else ""
                }
                new_data.append(item)
            self.data = new_data
            self.recalculate_and_refresh()
            messagebox.showinfo("Uspjeh", "Podaci su uspješno uvezeni!")
        except Exception as e:
            messagebox.showerror("Greška", f"Greška pri uvozu Excel fajla: {e}")

    def open_add_dialog(self):
        pass # Implementacija po potrebi

    def open_edit_dialog(self, event=None):
        pass # Implementacija po potrebi

    def delete_selected(self):
        pass # Implementacija po potrebi

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

        msg = f"Ukupno: {total} | Prikazano: {len(self.current_filtered_data)} | ISTEKLO: {expired_count} ({p_exp:.1f}%) | USKORO: {warning_count} ({p_warn:.1f}%) | VAŽEĆE: {valid_count} ({p_val:.1f}%)"
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

    def generate_status_chart(self):
        """Generiše linijski/stubičasti dijagram statusa opreme sa procentima."""
        counts = {"VAŽEĆE": 0, "USKORO": 0, "ISTEKLO": 0}
        for item in self.current_filtered_data:
            st = item.get("status", "VAŽEĆE")
            if st in counts:
                counts[st] += 1

        total = sum(counts.values())
        labels = list(counts.keys())
        values = list(counts.values())
        chart_colors = ['#16A34A', '#D97706', '#DC2626']

        fig, ax = plt.subplots(figsize=(6.5, 3.2))
        bars = ax.bar(labels, values, color=chart_colors, width=0.45)
        ax.set_ylabel('Broj stavki', fontsize=9)
        ax.set_title('Statusni pregled LZO opreme sa procentima', fontsize=10, fontweight='bold')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        for bar in bars:
            height = bar.get_height()
            percentage = (height / total * 100) if total > 0 else 0
            ax.annotate(f'{height}\n({percentage:.1f}%)',
                        xy=(bar.get_x() + bar.get_width() / 2, height),
                        xytext=(0, 3),
                        textcoords="offset points",
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
            initialfile="Izvjestaj_LZO.xlsx"
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

            ws.merge_cells("A1:P1")
            title_cell = ws["A1"]
            title_cell.value = "IZVJEŠTAJ ZADUŽENJA LIČNE ZAŠTITNE OPREME (LZO)"
            title_cell.font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")

            ws.merge_cells("A2:P2")
            sub_cell = ws["A2"]
            sub_cell.value = f"Datum generisanja: {datetime.now().strftime('%d.%m.%Y. u %H:%M')} | Ukupno stavki: {total} | Važeće: {val} ({p_val:.1f}%) | Uskoro ističe: {warn} ({p_warn:.1f}%) | Isteklo: {exp} ({p_exp:.1f}%)"
            sub_cell.font = Font(name="Calibri", size=10, italic=True, color="595959")
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
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = thin_border

            fill_expired = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            font_expired = Font(name="Calibri", size=9, color="9C0006", bold=True)
            fill_warning = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            font_warning = Font(name="Calibri", size=9, color="9C6500", bold=True)
            fill_valid = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            font_valid = Font(name="Calibri", size=9, color="006100")
            font_regular = Font(name="Calibri", size=9)

            start_row = 5
            for r_idx, row in enumerate(self.current_filtered_data, start=start_row):
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
            img.height = 225
            chart_row = len(self.current_filtered_data) + 7
            ws.add_image(img, f"B{chart_row}")

            wb.save(save_path)
            messagebox.showinfo("Uspjeh", "Excel izvještaj sa procentima i grafikonom je sačuvan!")
        except Exception as e:
            messagebox.showerror("Greška", f"Nije moguće sačuvati Excel fajl: {e}")

    def export_pdf_report(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Greška", "Biblioteka ReportLab nije instalirana. Instalirajte je preko 'pip install reportlab'.")
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

            title_style = ParagraphStyle(
                'TitleStyle', parent=styles['Heading1'], fontName=font_bold,
                fontSize=13, textColor=colors.HexColor('#1E293B'), alignment=0, spaceAfter=2
            )
            sub_style = ParagraphStyle(
                'SubStyle', parent=styles['Normal'], fontName=font_name,
                fontSize=8, textColor=colors.HexColor('#475569'), alignment=0, spaceAfter=4
            )
            cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontName=font_name, fontSize=7, leading=8)
            cell_bold = ParagraphStyle('CellBold', parent=styles['Normal'], fontName=font_bold, fontSize=7, leading=8)

            elements = []

            # Pristup logotipu
            logo_img_path = ensure_logo_png()
            
            title_text = "IZVJEŠTAJ ZADUŽENJA LIČNE ZAŠTITNE OPREME (LZO)"
            sub_text = f"Datum: {datetime.now().strftime('%d.%m.%Y. u %H:%M')} | Ukupno: {total} | Važeće: {val} ({p_val:.1f}%) | Uskoro ističe: {warn} ({p_warn:.1f}%) | Isteklo: {exp} ({p_exp:.1f}%)"

            if logo_img_path and os.path.exists(logo_img_path):
                # Formiramo zaglavlje sa logotipom na lijevoj strani
                logo_img = ReportLabImage(logo_img_path, width=110, height=40)
                text_block = [
                    Paragraph(title_text, title_style),
                    Spacer(1, 3),
                    Paragraph(sub_text, sub_style)
                ]
                header_table = Table([[logo_img, text_block]], colWidths=[120, 680])
                header_table.setStyle(TableStyle([
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('ALIGN', (0,0), (0,0), 'LEFT'),
                    ('LEFTPADDING', (0,0), (-1,-1), 0),
                    ('RIGHTPADDING', (0,0), (-1,-1), 0),
                ]))
                elements.append(header_table)
            else:
                elements.append(Paragraph(title_text, title_style))
                elements.append(Paragraph(sub_text, sub_style))

            elements.append(Spacer(1, 8))

            # Grafikon statusa
            chart_path = self.generate_status_chart()
            elements.append(ReportLabImage(chart_path, width=300, height=140))
            elements.append(Spacer(1, 8))

            # Tabela sa podacima
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
                    Paragraph(row.get("status", ""), cell_bold)
                ]
                table_data.append(r_vals)

            col_widths = [100, 90, 85, 60, 110, 35, 35, 35, 35, 55, 55, 50, 55]
            pdf_table = Table(table_data, colWidths=col_widths, repeatRows=1)
            
            t_style = [
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E293B')),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
                ('TOPPADDING', (0,0), (-1,-1), 3),
                ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ]

            # Oboji statusne ćelije u zavisnosti od vrijednosti
            for idx, r in enumerate(self.current_filtered_data, start=1):
                st = r.get("status", "")
                if st == "ISTEKLO":
                    t_style.append(('BACKGROUND', (12, idx), (12, idx), colors.HexColor('#FEE2E2')))
                    t_style.append(('TEXTCOLOR', (12, idx), (12, idx), colors.HexColor('#991B1B')))
                elif st == "USKORO":
                    t_style.append(('BACKGROUND', (12, idx), (12, idx), colors.HexColor('#FEF3C7')))
                    t_style.append(('TEXTCOLOR', (12, idx), (12, idx), colors.HexColor('#92400E')))
                elif st == "VAŽEĆE":
                    t_style.append(('BACKGROUND', (12, idx), (12, idx), colors.HexColor('#DCFCE7')))
                    t_style.append(('TEXTCOLOR', (12, idx), (12, idx), colors.HexColor('#166534')))

            pdf_table.setStyle(TableStyle(t_style))
            elements.append(pdf_table)

            doc.build(elements)
            messagebox.showinfo("Uspjeh", "PDF izvještaj sa CEDIS logotipom je uspješno generisan!")
        except Exception as e:
            messagebox.showerror("Greška", f"Nije moguće sačuvati PDF fajl: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = LZOApp(root)
    root.mainloop()
