import os
import json
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
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

DB_FILE = "lzo_baza.json"

class LZOApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistem za praćenje LZO i rokova zaduženja v4.0")
        self.root.geometry("1380x780")

        self.data = []
        self.current_filtered_data = []
        self.sort_directions = {}
        
        self.create_widgets()
        self.load_local_db()

    def create_widgets(self):
        # --- Gornji komandni panel ---
        top_frame = tk.Frame(self.root, pady=8, padx=10, bg="#f4f4f4")
        top_frame.pack(fill=tk.X)

        btn_import = tk.Button(top_frame, text="Uvoz iz Excel-a", command=self.import_excel, bg="#2b5797", fg="white", font=("Arial", 9, "bold"))
        btn_import.pack(side=tk.LEFT, padx=5)

        btn_add = tk.Button(top_frame, text="+ Novo zaduženje", command=self.open_add_dialog, bg="#107c41", fg="white", font=("Arial", 9, "bold"))
        btn_add.pack(side=tk.LEFT, padx=5)

        btn_edit = tk.Button(top_frame, text="✏️ Izmijeni selektovano", command=self.open_edit_dialog, bg="#d97706", fg="white", font=("Arial", 9, "bold"))
        btn_edit.pack(side=tk.LEFT, padx=5)

        btn_delete = tk.Button(top_frame, text="Obriši selektovano (Višestruko)", command=self.delete_selected, bg="#a80000", fg="white", font=("Arial", 9))
        btn_delete.pack(side=tk.LEFT, padx=5)

        # Dugmad za izvoz izvještaja
        btn_export_pdf = tk.Button(top_frame, text="Izvezi PDF (+ Grafikon)", command=self.export_pdf_report, bg="#b91c1c", fg="white", font=("Arial", 9, "bold"))
        btn_export_pdf.pack(side=tk.RIGHT, padx=5)

        btn_export_excel = tk.Button(top_frame, text="Izvezi Excel (+ Grafikon)", command=self.export_excel_report, bg="#008a00", fg="white", font=("Arial", 9, "bold"))
        btn_export_excel.pack(side=tk.RIGHT, padx=5)

        # --- Panel za Pretragu i Napredno Filtriranje ---
        filter_frame = tk.LabelFrame(self.root, text=" Pretraga i Filtriranje ", font=("Arial", 9, "bold"), padx=10, pady=6, bg="#f9f9f9")
        filter_frame.pack(fill=tk.X, padx=10, pady=5)

        # Pretraga
        tk.Label(filter_frame, text="Pretraga:", bg="#f9f9f9", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky="w", padx=2)
        self.entry_search = tk.Entry(filter_frame, width=20, font=("Arial", 9))
        self.entry_search.grid(row=0, column=1, padx=4, pady=2)
        self.entry_search.bind("<KeyRelease>", lambda e: self.apply_filters())

        # Org jedinica
        tk.Label(filter_frame, text="Org. jedinica:", bg="#f9f9f9", font=("Arial", 9)).grid(row=0, column=2, sticky="w", padx=(10, 2))
        self.combo_org = ttk.Combobox(filter_frame, state="readonly", width=16)
        self.combo_org.grid(row=0, column=3, padx=4, pady=2)
        self.combo_org.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        # Mjesto / Grad
        tk.Label(filter_frame, text="Mjesto / Grad:", bg="#f9f9f9", font=("Arial", 9)).grid(row=0, column=4, sticky="w", padx=(10, 2))
        self.combo_city = ttk.Combobox(filter_frame, state="readonly", width=14)
        self.combo_city.grid(row=0, column=5, padx=4, pady=2)
        self.combo_city.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        # Oprema
        tk.Label(filter_frame, text="Oprema:", bg="#f9f9f9", font=("Arial", 9)).grid(row=0, column=6, sticky="w", padx=(10, 2))
        self.combo_equipment = ttk.Combobox(filter_frame, state="readonly", width=16)
        self.combo_equipment.grid(row=0, column=7, padx=4, pady=2)
        self.combo_equipment.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        # Status
        tk.Label(filter_frame, text="Status:", bg="#f9f9f9", font=("Arial", 9)).grid(row=0, column=8, sticky="w", padx=(10, 2))
        self.combo_status = ttk.Combobox(filter_frame, state="readonly", width=12, values=["SVI", "ISTEKLO", "USKORO", "VAŽEĆE"])
        self.combo_status.current(0)
        self.combo_status.grid(row=0, column=9, padx=4, pady=2)
        self.combo_status.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        btn_reset = tk.Button(filter_frame, text="Poništi filtere", command=self.reset_filters, font=("Arial", 8))
        btn_reset.grid(row=0, column=10, padx=(10, 2))

        # --- Statusna traka ---
        self.lbl_status = tk.Label(self.root, text="Inicijalizacija sistema...", font=("Arial", 9, "italic"), anchor="w", padx=15, pady=4)
        self.lbl_status.pack(fill=tk.X)

        # --- Tabela (Treeview) ---
        table_frame = tk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.cols = (
            "ID", "Zaposleni", "Radno mjesto", "Org. jedinica", "Mjesto",
            "Vel. odjeća", "Vel. obuća", "Oprema", "J.M.", "Normativ",
            "Rok (mj)", "Izdata kol.", "Zaduženo", "Ističe", "Preostalo dana", "Status", "Napomena", "Akcija"
        )
        
        # Omogućena višestruka selekcija (selectmode="extended")
        self.tree = ttk.Treeview(table_frame, columns=self.cols, show="headings", selectmode="extended")

        col_widths = {
            "ID": 30, "Zaposleni": 150, "Radno mjesto": 130, "Org. jedinica": 130, "Mjesto": 90,
            "Vel. odjeća": 70, "Vel. obuća": 70, "Oprema": 150, "J.M.": 50, "Normativ": 60,
            "Rok (mj)": 60, "Izdata kol.": 70, "Zaduženo": 85, "Ističe": 85, "Preostalo dana": 90,
            "Status": 85, "Napomena": 120, "Akcija": 70
        }

        for col in self.cols:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by_column(c))
            center_cols = ["ID", "Vel. odjeća", "Vel. obuća", "J.M.", "Normativ", "Rok (mj)", "Izdata kol.", "Zaduženo", "Ističe", "Preostalo dana", "Status", "Akcija"]
            self.tree.column(col, width=col_widths.get(col, 100), anchor=tk.CENTER if col in center_cols else tk.W)

        self.tree.column("ID", width=0, stretch=False)

        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)

        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)

        self.tree.tag_configure("ISTEKLO", background="#ffc7ce", foreground="#9c0006")
        self.tree.tag_configure("USKORO", background="#ffeb9c", foreground="#9c6500")
        self.tree.tag_configure("VAŽEĆE", background="#c6efce", foreground="#006100")

        self.tree.bind("<Double-1>", self.open_edit_dialog)

    def load_local_db(self):
        if os.path.exists(DB_FILE):
            try:
                with open(DB_FILE, "r", encoding="utf-8") as f:
                    self.data = json.load(f)
                self.recalculate_and_refresh()
            except Exception as e:
                messagebox.showerror("Greška", f"Nije moguće učitati bazu: {e}")
        else:
            self.lbl_status.config(text="Lokalna baza nije pronađena. Uvezite podatke iz Excel fajla.", fg="blue")

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

        for item in self.data:
            rok = int(item.get("rok_mjeseci", 12) or 12)
            izdata_kol = int(item.get("izdata_kol", 1) or 1)
            d_zad = item.get("datum_zaduzenja", "")

            if d_zad:
                try:
                    dt_zad = datetime.strptime(d_zad, "%Y-%m-%d")
                    # Proračun: izdata količina * rok u mjesecima
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

        msg = f"Ukupno u bazi: {len(self.data)} | Prikazano: {len(self.current_filtered_data)} | ISTEKLO: {expired_count} | ISTIČE USKORO (<=30 dana): {warning_count}"
        self.lbl_status.config(text=msg, fg="#9c0006" if expired_count > 0 else "black")

    def update_filter_dropdowns(self):
        orgs = sorted(list(set(row.get("org_jedinica", "") for row in self.data if row.get("org_jedinica"))))
        cities = sorted(list(set(row.get("grad", "") for row in self.data if row.get("grad"))))
        equipments = sorted(list(set(row.get("oprema", "") for row in self.data if row.get("oprema"))))

        self.combo_org['values'] = ["SVE ORG. JEDINICE"] + orgs
        self.combo_city['values'] = ["SVA MJESTA"] + cities
        self.combo_equipment['values'] = ["SVA OPREMA"] + equipments

        if not self.combo_org.get(): self.combo_org.current(0)
        if not self.combo_city.get(): self.combo_city.current(0)
        if not self.combo_equipment.get(): self.combo_equipment.current(0)

    def apply_filters(self):
        search_txt = self.entry_search.get().strip().lower()
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
                row.get("napomena", ""),
                "✏️ Izmijeni"
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
        """Generiše privremenu PNG sliku grafikona za ugradnju u Excel i PDF."""
        counts = {"VAŽEĆE": 0, "USKORO": 0, "ISTEKLO": 0}
        for item in self.current_filtered_data:
            st = item.get("status", "VAŽEĆE")
            if st in counts:
                counts[st] += 1

        labels = list(counts.keys())
        values = list(counts.values())
        chart_colors = ['#2e7d32', '#f57f17', '#c62828']

        fig, ax = plt.subplots(figsize=(6, 3))
        bars = ax.bar(labels, values, color=chart_colors, width=0.45)
        ax.set_ylabel('Broj stavki', fontsize=9)
        ax.set_title('Statusni pregled zadužene LZO opreme', fontsize=10, fontweight='bold')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        for bar in bars:
            height = bar.get_height()
            ax.annotate(f'{height}',
                        xy=(bar.get_x() + bar.get_width() / 2, height),
                        xytext=(0, 3),
                        textcoords="offset points",
                        ha='center', va='bottom', fontweight='bold', fontsize=9)

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

            # Zaglavlje
            ws.merge_cells("A1:O1")
            title_cell = ws["A1"]
            title_cell.value = "IZVJEŠTAJ ZADUŽENJA LIČNE ZAŠTITNE OPREME (LZO)"
            title_cell.font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")

            ws.merge_cells("A2:O2")
            sub_cell = ws["A2"]
            sub_cell.value = f"Datum generisanja: {datetime.now().strftime('%d.%m.%Y. u %H:%M')} | Ukupno stavki: {len(self.current_filtered_data)}"
            sub_cell.font = Font(name="Calibri", size=10, italic=True, color="595959")
            sub_cell.alignment = Alignment(horizontal="center", vertical="center")

            headers = [
                "Ime i prezime", "Radno mjesto", "Org. jedinica", "Mjesto / Grad",
                "Vel. odjeća", "Vel. obuća", "Oprema", "J.M.", "Normativ", "Rok (mj)",
                "Izdata kol.", "Datum zaduženja", "Datum isticanja", "Preostalo dana", "Status", "Napomena"
            ]

            header_fill = PatternFill(start_color="2B5797", end_color="2B5797", fill_type="solid")
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
                for c_idx, val in enumerate(values, start=1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
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

            # Ugradnja grafikona
            chart_img_path = self.generate_status_chart()
            img = OpenpyxlImage(chart_img_path)
            img.width = 450
            img.height = 225
            chart_row = len(self.current_filtered_data) + 7
            ws.add_image(img, f"B{chart_row}")

            wb.save(save_path)
            messagebox.showinfo("Uspjeh", "Excel izvještaj sa grafikonom je uspješno sačuvan!")
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

            title_style = ParagraphStyle(
                'TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold',
                fontSize=14, textColor=colors.HexColor('#1F4E78'), alignment=1, spaceAfter=4
            )
            sub_style = ParagraphStyle(
                'SubStyle', parent=styles['Normal'], fontName='Helvetica-Oblique',
                fontSize=9, textColor=colors.HexColor('#595959'), alignment=1, spaceAfter=12
            )
            cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=7, leading=8)
            cell_bold = ParagraphStyle('CellBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7, leading=8)

            elements = []
            elements.append(Paragraph("IZVJEŠTAJ ZADUŽENJA LIČNE ZAŠTITNE OPREME (LZO)", title_style))
            elements.append(Paragraph(f"Datum generisanja: {datetime.now().strftime('%d.%m.%Y. u %H:%M')} | Ukupno stavki: {len(self.current_filtered_data)}", sub_style))

            # Dodavanje grafikona
            chart_path = self.generate_status_chart()
            elements.append(ReportLabImage(chart_path, width=320, height=160))
            elements.append(Spacer(1, 10))

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
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#2B5797')),
                ('ALIGN', (0,0), (-1,-1), 'LEFT'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D9D9D9')),
                ('TOPPADDING', (0,0), (-1,-1), 3),
                ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ]

            # Bojenje statusa u tabeli
            for idx, r_item in enumerate(self.current_filtered_data, start=1):
                st = r_item.get("status", "")
                if st == "ISTEKLO":
                    t_style.append(('BACKGROUND', (12, idx), (12, idx), colors.HexColor('#FFC7CE')))
                elif st == "USKORO":
                    t_style.append(('BACKGROUND', (12, idx), (12, idx), colors.HexColor('#FFEB9C')))
                elif st == "VAŽEĆE":
                    t_style.append(('BACKGROUND', (12, idx), (12, idx), colors.HexColor('#C6EFCE')))

            t.setStyle(TableStyle(t_style))
            elements.append(t)

            doc.build(elements)
            messagebox.showinfo("Uspjeh", "PDF izvještaj sa grafikonom je uspješno generisan!")
        except Exception as e:
            messagebox.showerror("Greška", f"Nije moguće generisati PDF fajl: {e}")

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
        win.geometry("500x620")
        win.grab_set()

        fields = [
            ("Ime i Prezime (A):", "zaposleni"),
            ("Radno Mjesto (B):", "rm"),
            ("Organizaciona Jedinica (C):", "org_jedinica"),
            ("Mjesto / Grad (D):", "grad"),
            ("Veličina Odjeće (E):", "vel_odjeca"),
            ("Veličina Obuće (F):", "vel_obuca"),
            ("Naziv Opreme (G):", "oprema"),
            ("Jedinica Mjere (H):", "jm"),
            ("Normativ - Količina (I):", "normativ"),
            ("Rok u mjesecima (J):", "rok_mjeseci"),
            ("Izdata Količina (K):", "izdata_kol"),
            ("Datum Zaduženja (L) (GGGG-MM-DD):", "datum_zaduzenja"),
            ("Napomena (O):", "napomena")
        ]

        entries = {}
        for idx, (label_text, key) in enumerate(fields):
            tk.Label(win, text=label_text, font=("Arial", 9, "bold")).grid(row=idx, column=0, sticky="w", padx=15, pady=3)
            entry = tk.Entry(win, width=32)
            entry.grid(row=idx, column=1, padx=15, pady=3)
            
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
                new_id = max([x["id"] for x in self.data], default=0) + 1
                self.data.append({
                    "id": new_id,
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
                })

            win.destroy()
            self.recalculate_and_refresh()
            messagebox.showinfo("Uspeh", "Podaci su sačuvani!")

        btn_save = tk.Button(win, text="Sačuvaj promjene", command=save, bg="#107c41", fg="white", font=("Arial", 10, "bold"), pady=5)
        btn_save.grid(row=len(fields), column=0, columnspan=2, pady=15)

    def delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Upozorenje", "Izaberite jednu ili više stavki iz tabele koje želite obrisati.")
            return

        count = len(selected)
        if messagebox.askyesno("Potvrda višestrukog brisanja", f"Da li ste sigurni da želite obrisati selektovane stavke ({count} zapis/a)?"):
            selected_ids = set()
            for item in selected:
                vals = self.tree.item(item)["values"]
                if vals:
                    selected_ids.add(vals[0])

            self.data = [x for x in self.data if x["id"] not in selected_ids]
            self.recalculate_and_refresh()
            messagebox.showinfo("Uspjeh", f"Uspješno obrisano {len(selected_ids)} stavki.")

    def import_excel(self):
        if self.data:
            if not messagebox.askyesno("Potvrda", "Već postoje podaci u bazi. Uvozom iz Excel-a ćete zamijeniti trenutne podatke. Nastaviti?"):
                return

        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_path: return

        try:
            xls = pd.ExcelFile(file_path)
            new_data = []
            idx = 1

            for sheet in xls.sheet_names:
                if sheet == 'NAPOMENA': continue
                df = pd.read_excel(file_path, sheet_name=sheet)
                
                # FFill hijerarhijskih kolona
                emp_cols = [df.columns[0], df.columns[1], df.columns[2], df.columns[3]]
                df[emp_cols] = df[emp_cols].ffill()

                for _, r in df.iterrows():
                    zaposleni = str(r.iloc[0]).strip() if pd.notnull(r.iloc[0]) else ""
                    if not zaposleni or zaposleni.lower() == "nan": continue

                    d_zad = ""
                    # Kolona L (indeks 11) je Datum zaduženja
                    if len(r) > 11 and pd.notnull(r.iloc[11]):
                        try:
                            d_zad = pd.to_datetime(r.iloc[11]).strftime("%Y-%m-%d")
                        except: pass

                    def parse_int(val, default=1):
                        try:
                            if pd.notnull(val) and str(val).strip().isdigit():
                                return int(val)
                        except: pass
                        return default

                    new_data.append({
                        "id": idx,
                        "zaposleni": zaposleni,                                                  # Kolona A (0)
                        "rm": str(r.iloc[1]).strip() if pd.notnull(r.iloc[1]) else "",          # Kolona B (1)
                        "org_jedinica": str(r.iloc[2]).strip() if pd.notnull(r.iloc[2]) else "",# Kolona C (2)
                        "grad": str(r.iloc[3]).strip() if pd.notnull(r.iloc[3]) else "",         # Kolona D (3)
                        "vel_odjeca": str(r.iloc[4]).strip() if pd.notnull(r.iloc[4]) else "",   # Kolona E (4)
                        "vel_obuca": str(r.iloc[5]).strip() if pd.notnull(r.iloc[5]) else "",    # Kolona F (5)
                        "oprema": str(r.iloc[6]).strip() if pd.notnull(r.iloc[6]) else "",       # Kolona G (6)
                        "jm": str(r.iloc[7]).strip() if pd.notnull(r.iloc[7]) else "KOM",        # Kolona H (7)
                        "normativ": parse_int(r.iloc[8], 1),                                     # Kolona I (8)
                        "rok_mjeseci": parse_int(r.iloc[9], 12),                                 # Kolona J (9)
                        "izdata_kol": parse_int(r.iloc[10], 1),                                  # Kolona K (10)
                        "datum_zaduzenja": d_zad,                                               # Kolona L (11)
                        "napomena": str(r.iloc[14]).strip() if len(r) > 14 and pd.notnull(r.iloc[14]) and str(r.iloc[14]).lower() != "nan" else "" # Kolona O (14)
                    })
                    idx += 1

            self.data = new_data
            self.recalculate_and_refresh()
            messagebox.showinfo("Uspjeh", "Podaci su uspješno uvezeni i proračunati prema zadatoj strukturi kolona!")
        except Exception as e:
            messagebox.showerror("Greška pri uvozu", f"Nije moguće pročitati Excel fajl: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = LZOApp(root)
    root.mainloop()
