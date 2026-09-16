import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime

class LZOApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistem za praćenje LZO i rokova zaduženja")
        self.root.geometry("1200x680")

        self.df_all = pd.DataFrame()
        self.df_current = pd.DataFrame()

        self.setup_ui()

    def setup_ui(self):
        # Kontrolni panel na vrhu
        top_frame = tk.LabelFrame(self.root, text=" Upravljanje podacima ", font=("Arial", 10, "bold"), padx=10, pady=10)
        top_frame.pack(fill=tk.X, padx=10, pady=5)

        btn_load = tk.Button(top_frame, text="📁 Učitaj Excel fajl", command=self.load_excel, bg="#0056b3", fg="white", font=("Arial", 9, "bold"))
        btn_load.pack(side=tk.LEFT, padx=5)

        btn_alert = tk.Button(top_frame, text="⚠️ Provjeri istekle/kritične rokove", command=self.check_alerts, bg="#d9534f", fg="white", font=("Arial", 9, "bold"))
        btn_alert.pack(side=tk.LEFT, padx=5)

        tk.Label(top_frame, text="Zaposleni:", font=("Arial", 9)).pack(side=tk.LEFT, padx=(20, 5))
        self.combo_emp = ttk.Combobox(top_frame, state="readonly", width=30)
        self.combo_emp.pack(side=tk.LEFT, padx=5)
        self.combo_emp.bind("<<ComboboxSelected>>", self.filter_by_employee)

        btn_export = tk.Button(top_frame, text="📄 Izvezi karton zaposlenog (Excel)", command=self.export_employee_report, bg="#28a745", fg="white", font=("Arial", 9, "bold"))
        btn_export.pack(side=tk.LEFT, padx=5)

        btn_reset = tk.Button(top_frame, text="🔄 Prikaži sve", command=self.reset_view)
        btn_reset.pack(side=tk.LEFT, padx=5)

        # Tabela za prikaz (Treeview)
        tree_frame = tk.Frame(self.root)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.tree = ttk.Treeview(tree_frame, show="headings")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # Formatiranje boja za upozorenja
        self.tree.tag_configure('expired', background='#f8d7da')  # Crvena za istekle
        self.tree.tag_configure('warning', background='#fff3cd')  # Žuta za rokove < 30 dana

    def load_excel(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not filepath:
            return

        try:
            excel_file = pd.ExcelFile(filepath)
            frames = []
            
            # Učitavanje svih sheet-ova
            for sheet_name in excel_file.sheet_names:
                df_sheet = pd.read_excel(filepath, sheet_name=sheet_name)
                # Ujednačavanje naziva kolona po pozicijama A-M
                if df_sheet.shape[1] >= 13:
                    df_sheet = df_sheet.iloc[:, :13]
                    df_sheet.columns = [
                        "Ime i prezime", "Radno mjesto", "Grad", "Veličina odjeće",
                        "Veličina obuće", "Oprema", "Jedinica mjere", "Količina",
                        "Rok (mjeseci)", "Datum zadnjeg zaduženja", "Datum isticanja",
                        "Preostalo dana", "Napomena"
                    ]
                    frames.append(df_sheet)

            if frames:
                self.df_all = pd.concat(frames, ignore_index=True)
                self.df_all.dropna(subset=["Ime i prezime", "Oprema"], inplace=True)
                
                # Izračunavanje preostalih dana u odnosu na današnji datum
                self.df_all["Datum isticanja"] = pd.to_datetime(self.df_all["Datum isticanja"], errors='coerce')
                today = pd.Timestamp.now().normalize()
                self.df_all["Preostalo dana"] = (self.df_all["Datum isticanja"] - today).dt.days

                self.df_current = self.df_all.copy()
                self.update_combobox()
                self.display_data(self.df_current)
                messagebox.showinfo("Uspjeh", f"Uspješno učitano {len(self.df_all)} zapisa iz svih radnih listova.")
        except Exception as e:
            messagebox.showerror("Greška", f"Greška prilikom otvaranja fajla:\n{str(e)}")

    def update_combobox(self):
        employees = sorted(list(self.df_all["Ime i prezime"].dropna().unique()))
        self.combo_emp["values"] = employees

    def display_data(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)

        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=110, anchor=tk.CENTER)

        for _, row in df.iterrows():
            vals = list(row)
            # Formatiranje datuma za lepši prikaz
            if pd.notnull(vals[9]): vals[9] = pd.to_datetime(vals[9]).strftime('%d.%m.%Y.')
            if pd.notnull(vals[10]): vals[10] = pd.to_datetime(vals[10]).strftime('%d.%m.%Y.')
            
            days_left = row["Preostalo dana"]
            tag = ""
            if pd.notnull(days_left):
                if days_left < 0:
                    tag = "expired"
                elif days_left <= 30:
                    tag = "warning"

            self.tree.insert("", tk.END, values=vals, tags=(tag,))

    def check_alerts(self):
        if self.df_all.empty:
            return
        critical_df = self.df_all[self.df_all["Preostalo dana"] <= 30]
        self.display_data(critical_df)
        
        expired_count = len(self.df_all[self.df_all["Preostalo dana"] < 0])
        warning_count = len(self.df_all[(self.df_all["Preostalo dana"] >= 0) & (self.df_all["Preostalo dana"] <= 30)])
        
        messagebox.showwarning(
            "Izvještaj o rokovima", 
            f"Istekli rokovi: {expired_count} kom/pari\nRokovi koji ističu u narednih 30 dana: {warning_count} kom/pari"
        )

    def filter_by_employee(self, event=None):
        emp_name = self.combo_emp.get()
        if emp_name and not self.df_all.empty:
            filtered_df = self.df_all[self.df_all["Ime i prezime"] == emp_name]
            self.df_current = filtered_df
            self.display_data(filtered_df)

    def reset_view(self):
        if not self.df_all.empty:
            self.df_current = self.df_all.copy()
            self.combo_emp.set('')
            self.display_data(self.df_all)

    def export_employee_report(self):
        emp_name = self.combo_emp.get()
        if not emp_name:
            messagebox.showwarning("Upozorenje", "Molimo izaberite zaposlenog iz padajućeg menija!")
            return

        emp_df = self.df_all[self.df_all["Ime i prezime"] == emp_name]
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"Karton_LZO_{emp_name.replace(' ', '_')}.xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if save_path:
            emp_df.to_excel(save_path, index=False)
            messagebox.showinfo("Uspjeh", f"Karton zaduženja za {emp_name} je uspješno sačuvan.")

if __name__ == "__main__":
    root = tk.Tk()
    app = LZOApp(root)
    root.mainloop()
