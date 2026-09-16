import os
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DB_FILE = "lzo_baza.json"

class LZOApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistem za praćenje LZO i rokova zaduženja v3.0")
        self.root.geometry("1280x750")

        self.data = []
        self.current_filtered_data = []
        self.sort_directions = {}
        self.create_widgets()
        self.load_local_db()

    def create_widgets(self):
        # --- Gornji komandni panel ---
        top_frame = tk.Frame(self.root, pady=8, padx=10, bg="#f4f4f4")
        top_frame.pack(fill=tk.X)

        btn_import = tk.Button(top_frame, text="Inicijalni uvoz (Excel)", command=self.import_excel, bg="#2b5797", fg="white", font=("Arial", 9, "bold"))
        btn_import.pack(side=tk.LEFT, padx=5)

        btn_add = tk.Button(top_frame, text="+ Novo zaduženje", command=self.open_add_dialog, bg="#107c41", fg="white", font=("Arial", 9, "bold"))
        btn_add.pack(side=tk.LEFT, padx=5)

        btn_delete = tk.Button(top_frame, text="Obriši selektovano", command=self.delete_selected, bg="#a80000", fg="white", font=("Arial", 9))
        btn_delete.pack(side=tk.LEFT, padx=5)

        btn_export = tk.Button(top_frame, text="Izvezi filtrirani karton (Excel)", command=self.export_excel_report, bg="#008a00", fg="white", font=("Arial", 9, "bold"))
        btn_export.pack(side=tk.RIGHT, padx=5)

        # --- Panel za Pretragu i Napredno Filtriranje ---
        filter_frame = tk.LabelFrame(self.root, text=" Pretraga i Filtriranje ", font=("Arial", 9, "bold"), padx=10, pady=6, bg="#f9f9f9")
        filter_frame.pack(fill=tk.X, padx=10, pady=5)

        # Live pretraga po imenu zaposlenog ili službi
        tk.Label(filter_frame, text="Pretraga (Zaposleni / Služba):", bg="#f9f9f9", font=("Arial", 9)).grid(row=0, column=0, sticky="w", padx=5)
        self.entry_search = tk.Entry(filter_frame, width=22, font=("Arial", 9))
        self.entry_search.grid(row=0, column=1, padx=5, pady=2)
        self.entry_search.bind("<KeyRelease>", lambda e: self.apply_filters())

        # Filter po mjestu / gradu
        tk.Label(filter_frame, text="Mjesto / Grad:", bg="#f9f9f9", font=("Arial", 9)).grid(row=0, column=2, sticky="w", padx=(15, 5))
        self.combo_city = ttk.Combobox(filter_frame, state="readonly", width=18)
        self.combo_city.grid(row=0, column=3, padx=5, pady=2)
        self.combo_city.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        # Filter po vrsti opreme
        tk.Label(filter_frame, text="Vrsta opreme:", bg="#f9f9f9", font=("Arial", 9)).grid(row=0, column=4, sticky="w", padx=(15, 5))
        self.combo_equipment = ttk.Combobox(filter_frame, state="readonly", width=22)
        self.combo_equipment.grid(row=0, column=5, padx=5, pady=2)
        self.combo_equipment.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        # Filter po statusu
        tk.Label(filter_frame, text="Status:", bg="#f9f9f9", font=("Arial", 9)).grid(row=0, column=6, sticky="w", padx=(15, 5))
        self.combo_status = ttk.Combobox(filter_frame, state="readonly", width=14, values=["SVI", "ISTEKLO", "USKORO", "VAŽEĆE"])
        self.combo_status.current(0)
        self.combo_status.grid(row=0, column=7, padx=5, pady=2)
        self.combo_status.bind("<<ComboboxSelected>>", lambda e: self.apply_filters())

        btn_reset = tk.Button(filter_frame, text="Poništi filtere", command=self.reset_filters, font=("Arial", 8))
        btn_reset.grid(row=0, column=8, padx=(15, 5))

        # --- Statusna traka ---
        self.lbl_status = tk.Label(self.root, text="Inicijalizacija sistema...", font=("Arial", 9, "italic"), anchor="w", padx=15, pady=4)
        self.lbl_status.pack(fill=tk.X)

        # --- Tabela (Treeview) ---
        table_frame = tk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.cols = ("ID", "Zaposleni", "Radno mjesto", "Mjesto", "Oprema", "Količina", "J.M.", "Rok (mj)", "Zaduženo", "Ističe", "Preostalo dana", "Status", "Napomena")
        self.tree = ttk.Treeview(table_frame, columns=self.cols, show="headings")

        col_widths = {"ID": 30, "Zaposleni": 160, "Radno mjesto": 150, "Mjesto": 90, "Oprema": 170, "Količina": 55, "J.M.": 45, "Rok (mj)": 60, "Zaduženo": 85, "Ističe": 85, "Preostalo dana": 95, "Status": 85, "Napomena": 130}
        
        for col in self.cols:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by_column(c))
            self.tree.column(col, width=col_widths.get(col, 100), anchor=tk.CENTER if col in ["ID", "Količina", "J.M.", "Rok (mj)", "Zaduženo", "Ističe", "Preostalo dana", "Status"] else tk.W)

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
            d_zad = item.get("datum_zaduzenja", "")
            
            if d_zad:
                try:
                    dt_zad = datetime.strptime(d_zad, "%Y-%m-%d")
                    dt_ist = dt_zad + timedelta(days=rok*30.4375)
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
        cities = sorted(list(set(row["grad"] for row in self.data if row.get("grad"))))
        equipments = sorted(list(set(row["oprema"] for row in self.data if row.get("oprema"))))
        
        self.combo_city['values'] = ["SVA MJESTA"] + cities
        self.combo_equipment['values'] = ["SVA OPREMA"] + equipments
        
        if not self.combo_city.get(): self.combo_city.current(0)
        if not self.combo_equipment.get(): self.combo_equipment.current(0)

    def apply_filters(self):
        search_txt = self.entry_search.get().strip().lower()
        selected_city = self.combo_city.get()
        selected_equip = self.combo_equipment.get()
        selected_status = self.combo_status.get()

        filtered = []
        for row in self.data:
            match_search = True
            if search_txt:
                match_search = (search_txt in row.get("zaposleni", "").lower() or
                                search_txt in row.get("rm", "").lower())

            match_city = True
            if selected_city and selected_city != "SVA MJESTA":
                match_city = (row.get("grad") == selected_city)

            match_equip = True
            if selected_equip and selected_equip != "SVA OPREMA":
                match_equip = (row.get("oprema") == selected_equip)

            match_status = True
            if selected_status and selected_status != "SVI":
                match_status = (row.get("status") == selected_status)

            if match_search and match_city and match_equip and match_status:
                filtered.append(row)

        self.current_filtered_data = filtered
        self.refresh_table(filtered)

    def reset_filters(self):
        self.entry_search.delete(0, tk.END)
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
                row.get("grad", ""),
                row.get("oprema", ""),
                row.get("kolicina", 1),
                row.get("jm", "KOM"),
                row.get("rok_mjeseci", 12),
                d_zad,
                d_ist,
                row.get("preostalo_dana", "-"),
                row.get("status", ""),
                row.get("napomena", "")
            ), tags=(row.get("status", ""),))

    def sort_by_column(self, col):
        ascending = self.sort_directions.get(col, True)
        
        key_map = {
            "ID": "id", "Zaposleni": "zaposleni", "Radno mjesto": "rm", "Mjesto": "grad",
            "Oprema": "oprema", "Količina": "kolicina", "J.M.": "jm", "Rok (mj)": "rok_mjeseci",
            "Zaduženo": "datum_zaduzenja", "Ističe": "datum_isticanja",
            "Preostalo dana": "preostalo_dana", "Status": "status", "Napomena": "napomena"
        }
        
        key = key_map.get(col)
        if not key: return

        def sort_val(item):
            val = item.get(key, "")
            if val is None: return ""
            if key in ["id", "kolicina", "rok_mjeseci"]:
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

    def export_excel_report(self):
        if not hasattr(self, 'current_filtered_data') or not self.current_filtered_data:
            messagebox.showwarning("Upozorenje", "Nema podataka za izvoz.")
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile="Izvjestaj_LZO.xlsx"
        )
        if not save_path:
            return

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "LZO Karton"
            ws.views.sheetView[0].showGridLines = True

            # Naslovni blok
            ws.merge_cells("A1:L1")
            title_cell = ws["A1"]
            title_cell.value = "IZVJEŠTAJ ZADUŽENJA LIČNE ZAŠTITNE OPREME (LZO)"
            title_cell.font = Font(name="Calibri", size=14, bold=True, color="1F4E78")
            title_cell.alignment = Alignment(horizontal="center", vertical="center")

            ws.merge_cells("A2:L2")
            sub_cell = ws["A2"]
            sub_cell.value = "Datum generisanja: " + datetime.now().strftime('%d.%m.%Y. u %H:%M') + " | Ukupno stavki: " + str(len(self.current_filtered_data))
            sub_cell.font = Font(name="Calibri", size=10, italic=True, color="595959")
            sub_cell.alignment = Alignment(horizontal="center", vertical="center")

            ws.row_dimensions[1].height = 25
            ws.row_dimensions[2].height = 18
            ws.row_dimensions[4].height = 24

            # Zaglavlje tabele
            headers = ["Ime i prezime", "Radno mjesto / Služba", "Mjesto / Grad", "Oprema", "Količina", "J.M.", "Rok (mj)", "Datum zaduženja", "Datum isticanja", "Preostalo dana", "Status", "Napomena"]
            
            header_fill = PatternFill(start_color="2B5797", end_color="2B5797", fill_type="solid")
            header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
            thin_border = Border(
                left=Side(style='thin', color='D9D9D9'),
                right=Side(style='thin', color='D9D9D9'),
                top=Side(style='thin', color='D9D9D9'),
                bottom=Side(style='thin', color='D9D9D9')
            )

            for col_num, header_title in enumerate(headers, 1):
                cell = ws.cell(row=4, column=col_num)
                cell.value = header_title
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = thin_border

            # Stilovi za statusne boje
            fill_expired = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            font_expired = Font(name="Calibri", size=10, color="9C0006", bold=True)
            
            fill_warning = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
            font_warning = Font(name="Calibri", size=10, color="9C6500", bold=True)
            
            fill_valid = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            font_valid = Font(name="Calibri", size=10, color="006100")

            font_regular = Font(name="Calibri", size=10)

            # Popunjavanje redova
            start_row = 5
            for r_idx, row in enumerate(self.current_filtered_data, start=start_row):
                ws.row_dimensions[r_idx].height = 20
                
                d_zad = datetime.strptime(row["datum_zaduzenja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if row.get("datum_zaduzenja") else ""
                d_ist = datetime.strptime(row["datum_isticanja"], "%Y-%m-%d").strftime("%d.%m.%Y.") if row.get("datum_isticanja") else ""

                values = [
                    row.get("zaposleni", ""),
                    row.get("rm", ""),
                    row.get("grad", ""),
                    row.get("oprema", ""),
                    row.get("kolicina", 1),
                    row.get("jm", "KOM"),
                    row.get("rok_mjeseci", 12),
                    d_zad,
                    d_ist,
                    row.get("preostalo_dana", "-"),
                    row.get("status", ""),
                    row.get("napomena", "")
                ]

                status_val = row.get("status", "")

                for c_idx, val in enumerate(values, start=1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    cell.font = font_regular
                    cell.border = thin_border

                    # Poravnanje po sredini za numeričke/datum/status kolone
                    if c_idx in [5, 6, 7, 8, 9, 10, 11]:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        cell.alignment = Alignment(horizontal="left", vertical="center")

                    # Bojenje statusne kolone
                    if c_idx == 11:
                        if status_val == "ISTEKLO":
                            cell.fill = fill_expired
                            cell.font = font_expired
                        elif status_val == "USKORO":
                            cell.fill = fill_warning
                            cell.font = font_warning
                        elif status_val == "VAŽEĆE":
                            cell.fill = fill_valid
                            cell.font = font_valid

            # Auto-fit širina kolona
            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for cell in col:
                    if cell.row < 4: continue
                    val_str = str(cell.value or '')
                    if len(val_str) > max_len:
                        max_len = len(val_str)
                ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

            wb.save(save_path)
            messagebox.showinfo("Uspjeh", "Formatirani izvještaj je uspješno sačuvan!")
        except Exception as e:
            messagebox.showerror("Greška pri izvozu", f"Nije moguće sačuvati Excel fajl: {e}")

    def open_add_dialog(self):
        self.show_edit_window(title="Novo zaduženje", item=None)

    def open_edit_dialog(self, event):
        selected = self.tree.selection()
        if not selected: return
        item_id = self.tree.item(selected[0])["values"][0]
        
        target_item = next((x for x in self.data if x["id"] == item_id), None)
        if target_item:
            self.show_edit_window(title="Izmjena zaduženja", item=target_item)

    def show_edit_window(self, title, item=None):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.geometry("450x520")
        win.grab_set()

        fields = [
            ("Ime i Prezime:", "zaposleni"),
            ("Radno Mjesto / Služba:", "rm"),
            ("Grad / Mjesto:", "grad"),
            ("Naziv opreme:", "oprema"),
            ("Količina:", "kolicina"),
            ("Jedinica mjere:", "jm"),
            ("Rok (u mjesecima):", "rok_mjeseci"),
            ("Datum zaduženja (GGGG-MM-DD):", "datum_zaduzenja"),
            ("Napomena:", "napomena")
        ]

        entries = {}
        for idx, (label_text, key) in enumerate(fields):
            tk.Label(win, text=label_text, font=("Arial", 9, "bold")).grid(row=idx, column=0, sticky="w", padx=15, pady=5)
            entry = tk.Entry(win, width=32)
            entry.grid(row=idx, column=1, padx=15, pady=5)
            
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
                rok = int(entries["rok_mjeseci"].get().strip() or 12)
                kol = int(entries["kolicina"].get().strip() or 1)
            except ValueError:
                messagebox.showerror("Greška", "Količina i Rok moraju biti brojevi!", parent=win)
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
                item["grad"] = entries["grad"].get().strip()
                item["oprema"] = oprema
                item["kolicina"] = kol
                item["jm"] = entries["jm"].get().strip()
                item["rok_mjeseci"] = rok
                item["datum_zaduzenja"] = d_zad
                item["napomena"] = entries["napomena"].get().strip()
            else:
                new_id = max([x["id"] for x in self.data], default=0) + 1
                self.data.append({
                    "id": new_id,
                    "zaposleni": zaposleni,
                    "rm": entries["rm"].get().strip(),
                    "grad": entries["grad"].get().strip(),
                    "oprema": oprema,
                    "kolicina": kol,
                    "jm": entries["jm"].get().strip(),
                    "rok_mjeseci": rok,
                    "datum_zaduzenja": d_zad,
                    "napomena": entries["napomena"].get().strip()
                })

            win.destroy()
            self.recalculate_and_refresh()
            messagebox.showinfo("Uspeh", "Podaci su sačuvani!")

        btn_save = tk.Button(win, text="Sačuvaj promjene", command=save, bg="#107c41", fg="white", font=("Arial", 10, "bold"), pady=5)
        btn_save.grid(row=len(fields), column=0, columnspan=2, pady=20)

    def delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Upozorenje", "Izaberite zapis iz tabele koji želite obrisati.")
            return

        item_id = self.tree.item(selected[0])["values"][0]
        if messagebox.askyesno("Potvrda", "Da li ste sigurni da želite obrisati izabrano zaduženje?"):
            self.data = [x for x in self.data if x["id"] != item_id]
            self.recalculate_and_refresh()

    def import_excel(self):
        if self.data:
            if not messagebox.askyesno("Potvrda", "Već postoje podaci u bazi. Uvozom iz Excel-a ćete zamijeniti trenutne podatke. Nastaviti?"):
                return

        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_path:
            return

        try:
            xls = pd.ExcelFile(file_path)
            new_data = []
            idx = 1

            for sheet in xls.sheet_names:
                if sheet == 'NAPOMENA': continue
                df = pd.read_excel(file_path, sheet_name=sheet)
                
                emp_cols = [df.columns[0], df.columns[1], df.columns[2]]
                df[emp_cols] = df[emp_cols].ffill()

                for _, r in df.iterrows():
                    zaposleni = str(r.iloc[0]).strip() if pd.notnull(r.iloc[0]) else ""
                    if not zaposleni or zaposleni == "nan": continue

                    d_zad = ""
                    if pd.notnull(r.iloc[9]):
                        try:
                            d_zad = pd.to_datetime(r.iloc[9]).strftime("%Y-%m-%d")
                        except: pass

                    new_data.append({
                        "id": idx,
                        "zaposleni": zaposleni,
                        "rm": str(r.iloc[1]).strip() if pd.notnull(r.iloc[1]) else "",
                        "grad": str(r.iloc[2]).strip() if pd.notnull(r.iloc[2]) else "",
                        "vel_odjeca": str(r.iloc[3]).strip() if pd.notnull(r.iloc[3]) else "",
                        "vel_obuca": str(r.iloc[4]).strip() if pd.notnull(r.iloc[4]) else "",
                        "oprema": str(r.iloc[5]).strip() if pd.notnull(r.iloc[5]) else "",
                        "jm": str(r.iloc[6]).strip() if pd.notnull(r.iloc[6]) else "KOM",
                        "kolicina": int(r.iloc[7]) if pd.notnull(r.iloc[7]) and str(r.iloc[7]).isdigit() else 1,
                        "rok_mjeseci": int(r.iloc[8]) if pd.notnull(r.iloc[8]) and str(r.iloc[8]).isdigit() else 12,
                        "datum_zaduzenja": d_zad,
                        "napomena": str(r.iloc[12]).strip() if len(r) > 12 and pd.notnull(r.iloc[12]) and str(r.iloc[12]) != "nan" else ""
                    })
                    idx += 1

            self.data = new_data
            self.recalculate_and_refresh()
            messagebox.showinfo("Uspjeh", "Podaci su uspješno uvezeni i sačuvani u lokalnu bazu!")
        except Exception as e:
            messagebox.showerror("Greška pri uvozu", f"Nije moguće pročitati Excel fajl: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = LZOApp(root)
    root.mainloop()
