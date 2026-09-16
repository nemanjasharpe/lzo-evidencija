import os
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime, timedelta

DB_FILE = "lzo_baza.json"

class LZOApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistem za praćenje LZO i rokova zaduženja v2.0")
        self.root.geometry("1200x700")

        self.data = []
        self.create_widgets()
        
        # Automatsko učitavanje lokalne baze pri pokretanju
        self.load_local_db()

    def create_widgets(self):
        # --- Gornji kontrolni panel ---
        top_frame = tk.Frame(self.root, pady=10, padx=10, bg="#f4f4f4")
        top_frame.pack(fill=tk.X)

        btn_import = tk.Button(top_frame, text="Inicijalni uvoz iz Excel-a", command=self.import_excel, bg="#2b5797", fg="white", font=("Arial", 9, "bold"))
        btn_import.pack(side=tk.LEFT, padx=5)

        btn_add = tk.Button(top_frame, text="+ Novo zaduženje", command=self.open_add_dialog, bg="#107c41", fg="white", font=("Arial", 9, "bold"))
        btn_add.pack(side=tk.LEFT, padx=5)

        btn_delete = tk.Button(top_frame, text="Stoši selektovano", command=self.delete_selected, bg="#a80000", fg="white", font=("Arial", 9))
        btn_delete.pack(side=tk.LEFT, padx=5)

        tk.Label(top_frame, text="Filtriraj po radniku:", bg="#f4f4f4", font=("Arial", 10)).pack(side=tk.LEFT, padx=(20, 5))
        self.combo_worker = ttk.Combobox(top_frame, state="readonly", width=28)
        self.combo_worker.pack(side=tk.LEFT, padx=5)
        self.combo_worker.bind("<<ComboboxSelected>>", self.filter_by_worker)

        btn_reset = tk.Button(top_frame, text="Prikaži sve", command=self.reset_filter)
        btn_reset.pack(side=tk.LEFT, padx=5)

        btn_export = tk.Button(top_frame, text="Izvezi karton radnika", command=self.export_worker_report, bg="#008a00", fg="white", font=("Arial", 9, "bold"))
        btn_export.pack(side=tk.RIGHT, padx=5)

        # --- Statusna traka ---
        self.lbl_status = tk.Label(self.root, text="Inicijalizacija sistema...", font=("Arial", 10, "italic"), anchor="w", padx=15, py=4)
        self.lbl_status.pack(fill=tk.X)

        # --- Tabela ---
        table_frame = tk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        cols = ("ID", "Zaposleni", "Radno mjesto", "Mjesto", "Oprema", "Količina", "J.M.", "Rok (mj)", "Zaduženo", "Ističe", "Preostalo dana", "Status", "Napomena")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings")

        col_widths = {"ID": 40, "Zaposleni": 150, "Radno mjesto": 140, "Mjesto": 80, "Oprema": 160, "Količina": 55, "J.M.": 45, "Rok (mj)": 60, "Zaduženo": 85, "Ističe": 85, "Preostalo dana": 90, "Status": 85, "Napomena": 120}
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=col_widths.get(col, 100), anchor=tk.CENTER if col in ["ID", "Količina", "J.M.", "Rok (mj)", "Zaduženo", "Ističe", "Preostalo dana", "Status"] else tk.W)

        # Sakrij ID kolonu vizuelno ali je zadrži za identifikaciju
        self.tree.column("ID", width=0, stretch=False)

        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscroll=scrollbar_y.set, xscroll=scrollbar_x.set)

        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # Oznake za obojene redove
        self.tree.tag_configure("ISTEKLO", background="#ffc7ce", foreground="#9c0006")
        self.tree.tag_configure("USKORO", background="#ffeb9c", foreground="#9c6500")
        self.tree.tag_configure("VAŽEĆE", background="#c6efce", foreground="#006100")

        # Event za dvoklik - Izmjena zapisa
        self.tree.bind("<Double-1>", self.open_edit_dialog)

    # --- Baza podataka i proračuni ---
    def load_local_db(self):
        if os.path.exists(DB_FILE):
            try:
                with open(DB_FILE, "r", encoding="utf-8") as f:
                    self.data = json.load(f)
                self.recalculate_and_refresh()
            except Exception as e:
                messagebox.showerror("Greška", f"Nije moguće učitati bazu: {e}")
        else:
            self.lbl_status.config(text="Lokalna baza nije pronađena. Možete učitati podatak iz Excel fajla.", fg="blue")

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
                    # Izračunavanje datuma isticanja na osnovu zadatog roka u mjesecima
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
        self.populate_worker_combo()
        self.refresh_table(self.data)

        msg = f"Baza učitana ({len(self.data)} stavki) | ISTEKLO: {expired_count} | ISTIČE USKORO (<=30 dana): {warning_count}"
        self.lbl_status.config(text=msg, fg="#9c0006" if expired_count > 0 else "black")

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

    def populate_worker_combo(self):
        workers = sorted(list(set(row["zaposleni"] for row in self.data if row.get("zaposleni"))))
        self.combo_worker['values'] = workers

    # --- Inicijalni uvoz iz Excel-a ---
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
                
                # Popuni prazne spojene ćelije za radnike
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
            messagebox.showerror("Greška pri uvozu", f"Nije moguće pročitati Excel fajl:\n{e}")

    # --- Dijalozi za dodavanje / izmjenu ---
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
            ("Radno Mjesto:", "rm"),
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
            
            # Ako vršimo izmjenu, popuni vrijednosti
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

            if item: # Update
                item["zaposleni"] = zaposleni
                item["rm"] = entries["rm"].get().strip()
                item["grad"] = entries["grad"].get().strip()
                item["oprema"] = oprema
                item["kolicina"] = kol
                item["jm"] = entries["jm"].get().strip()
                item["rok_mjeseci"] = rok
                item["datum_zaduzenja"] = d_zad
                item["napomena"] = entries["napomena"].get().strip()
            else: # Insert
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

    # --- Filtriranje i Izvoz ---
    def filter_by_worker(self, event=None):
        worker = self.combo_worker.get()
        if worker:
            filtered = [x for x in self.data if x["zaposleni"] == worker]
            self.refresh_table(filtered)

    def reset_filter(self):
        self.combo_worker.set("")
        self.refresh_table(self.data)

    def export_worker_report(self):
        worker = self.combo_worker.get()
        if not worker:
            messagebox.showwarning("Upozorenje", "Izaberite radnika iz padajućeg menija za izvoz kartona.")
            return

        filtered = [x for x in self.data if x["zaposleni"] == worker]
        df_exp = pd.DataFrame(filtered)

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=f"Karton_LZO_{worker.replace(' ', '_')}.xlsx"
        )
        if save_path:
            try:
                cols_map = {
                    "zaposleni": "Ime i prezime", "rm": "Radno mjesto", "grad": "Mjesto",
                    "oprema": "Oprema", "kolicina": "Količina", "jm": "J.M.",
                    "rok_mjeseci": "Rok (mjeseci)", "datum_zaduzenja": "Datum zaduženja",
                    "datum_isticanja": "Datum isticanja", "status": "Status", "napomena": "Napomena"
                }
                df_exp = df_exp[list(cols_map.keys())].rename(columns=cols_map)
                df_exp.to_excel(save_path, index=False)
                messagebox.showinfo("Uspjeh", f"Karton za radnika {worker} je sačuvan!")
            except Exception as e:
                messagebox.showerror("Greška", f"Nije moguće izvesti fajl: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = LZOApp(root)
    root.mainloop()
