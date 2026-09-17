import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

# ReportLab moduli za PDF izvještaje
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ==========================================
# POMOĆNA FUNKCIJA ZA PYINSTALLER RESURSE
# ==========================================
def resource_path(relative_path):
    """Pronađe apsolutnu putanju do resursa (radi i u razvoju i unutar .exe)"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ==========================================
# 1. UNDO / REDO MEHANIZAM (MEMENTO PATTERN)
# ==========================================
class UndoManager:
    def __init__(self, max_depth=30):
        self.undo_stack = []
        self.redo_stack = []
        self.max_depth = max_depth

    def push_state(self, state_data):
        """Spasava trenutno stanje u stak"""
        self.undo_stack.append(pd.DataFrame(state_data).copy(deep=True))
        if len(self.undo_stack) > self.max_depth:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def undo(self, current_state):
        if not self.undo_stack:
            return None
        self.redo_stack.append(pd.DataFrame(current_state).copy(deep=True))
        return self.undo_stack.pop()

    def redo(self, current_state):
        if not self.redo_stack:
            return None
        self.undo_stack.append(pd.DataFrame(current_state).copy(deep=True))
        return self.redo_stack.pop()

# ==========================================
# 2. DROPDOWN PRETRAGA SA SUBSTRING MATCH-OM
# ==========================================
class SearchableCombobox(ttk.Combobox):
    def set_completion_list(self, completion_list):
        self._completion_list = sorted(completion_list)
        self['values'] = self._completion_list
        self.bind('<KeyRelease>', self._handle_keyrelease)

    def _handle_keyrelease(self, event):
        if event.keysym in ('BackSpace', 'Left', 'Right', 'Up', 'Down', 'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Return'):
            return
        
        value = self.get().lower()
        if value == '':
            self['values'] = self._completion_list
        else:
            # Prikazuje sve unose koji sadrže unesena slova u imenu/prezimenu
            hits = [item for item in self._completion_list if value in item.lower()]
            self['values'] = hits
            
        self.event_generate('<Down>')

# ==========================================
# 3. GENERISANJE PDF IZVJEŠTAJA (PROCENTI, ENKLITIKE, LOGO)
# ==========================================
def register_latin_fonts():
    """Registruje Arial font sa podrškom za Č, Š, Ć, Đ, Ž"""
    try:
        font_path = "C:\\Windows\\Fonts\\arial.ttf"
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont('ArialCustom', font_path))
            return 'ArialCustom'
    except Exception:
        pass
    return 'Helvetica'

def generate_pdf_report(filename, data_df):
    doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    story = []
    
    font_name = register_latin_fonts()
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontName=font_name,
        fontSize=15,
        textColor=colors.HexColor('#0F172A'),
        spaceAfter=10
    )
    
    text_style = ParagraphStyle(
        'CustomText',
        parent=styles['Normal'],
        fontName=font_name,
        fontSize=10,
        textColor=colors.HexColor('#334155')
    )

    # Učitavanje logotipa (preko resource_path za .exe)
    logo_file = resource_path("cedis_logo.png")
    if os.path.exists(logo_file):
        logo = RLImage(logo_file, width=130, height=45)
        logo.hAlign = 'LEFT'
        story.append(logo)
        story.append(Spacer(1, 10))

    story.append(Paragraph("<b>IZVJEŠTAJ O LIČNOJ ZAŠTITNOJ OPREMI (LZO) - REGION 4</b>", title_style))
    story.append(Spacer(1, 10))

    # Izračunavanje statistike i procenata
    total = len(data_df)
    if total > 0:
        u_roku = len(data_df[data_df['Status'] == 'U roku'])
        u_isteku = len(data_df[data_df['Status'] == 'U isteku'])
        isteklo = len(data_df[data_df['Status'] == 'Isteklo'])
        
        p_roku = (u_roku / total) * 100
        p_isteku = (u_isteku / total) * 100
        p_isteklo = (isteklo / total) * 100
    else:
        u_roku = u_isteku = isteklo = p_roku = p_isteku = p_isteklo = 0

    stats_text = f"""
    <b>Statistika zadužene opreme:</b><br/>
    • <b>Ukupno zaduženja:</b> {total}<br/>
    • <b>U roku:</b> {u_roku} ({p_roku:.1f}%)<br/>
    • <b>U isteku (uskoro ističe):</b> {u_isteku} ({p_isteku:.1f}%)<br/>
    • <b>Isteklo:</b> {isteklo} ({p_isteklo:.1f}%)
    """
    story.append(Paragraph(stats_text, text_style))
    story.append(Spacer(1, 15))

    # Tabela sa prikazom naših slova (Č, Š, Ć, Đ, Ž)
    table_data = [["Zaposleni", "Oprema", "Datum zaduženja", "Status"]]
    for _, row in data_df.iterrows():
        table_data.append([
            Paragraph(str(row.get('Zaposleni', '')), text_style),
            Paragraph(str(row.get('Oprema', '')), text_style),
            Paragraph(str(row.get('Datum', '')), text_style),
            Paragraph(str(row.get('Status', '')), text_style)
        ])

    t = Table(table_data, colWidths=[150, 150, 100, 100])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0284C7')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#CBD5E1')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t)
    doc.build(story)

# ==========================================
# 4 & 5. LOGOVANJE I MODERAN INTERFEJS
# ==========================================
class LZOApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CEDIS - Evidencija LZO Opreme")
        self.geometry("980x620")
        self.configure(bg="#F8FAFC")
        
        self.undo_mgr = UndoManager()
        self.data = pd.DataFrame(columns=["Zaposleni", "Oprema", "Datum", "Status"])

        self._apply_styles()
        self.withdraw()  # Sakriva glavni prozor dok prijava ne uspije
        self.show_login_screen()

    def _apply_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure('TFrame', background='#F8FAFC')
        style.configure('TLabelframe', background='#F8FAFC', foreground='#1E293B', font=('Segoe UI', 10, 'bold'))
        style.configure('TLabelframe.Label', background='#F8FAFC', foreground='#1E293B')
        style.configure('TLabel', background='#F8FAFC', foreground='#334155', font=('Segoe UI', 10))
        style.configure('Header.TLabel', font=('Segoe UI', 13, 'bold'), foreground='#0F172A')
        
        style.configure('Primary.TButton', font=('Segoe UI', 9, 'bold'), background='#0284C7', foreground='white', borderwidth=0)
        style.map('Primary.TButton', background=[('active', '#0369A1')])
        
        style.configure('Secondary.TButton', font=('Segoe UI', 9), background='#64748B', foreground='white', borderwidth=0)
        style.map('Secondary.TButton', background=[('active', '#475569')])

        style.configure('Treeview', font=('Segoe UI', 9), rowheight=28, background='white', fieldbackground='white')
        style.configure('Treeview.Heading', font=('Segoe UI', 10, 'bold'), background='#E2E8F0', foreground='#0F172A')

    def show_login_screen(self):
        login_win = tk.Toplevel(self)
        login_win.title("CEDIS - Prijava na sistem")
        login_win.geometry("380x300")
        login_win.configure(bg="#1E293B")
        login_win.resizable(False, False)
        login_win.grab_set()

        tk.Label(login_win, text="CEDIS LZO SISTEM", font=('Segoe UI', 14, 'bold'), bg="#1E293B", fg="white").pack(pady=20)

        form_frame = tk.Frame(login_win, bg="#1E293B")
        form_frame.pack(pady=5)

        tk.Label(form_frame, text="Korisničko ime:", bg="#1E293B", fg="#94A3B8", font=('Segoe UI', 10)).grid(row=0, column=0, sticky="w", pady=6, padx=5)
        user_entry = ttk.Entry(form_frame, width=22)
        user_entry.grid(row=0, column=1, pady=6)
        user_entry.focus()

        tk.Label(form_frame, text="Lozinka:", bg="#1E293B", fg="#94A3B8", font=('Segoe UI', 10)).grid(row=1, column=0, sticky="w", pady=6, padx=5)
        pass_entry = ttk.Entry(form_frame, show="*", width=22)
        pass_entry.grid(row=1, column=1, pady=6)

        def check_credentials():
            if user_entry.get().strip() == "REGION 4" and pass_entry.get().strip() == "12345678":
                login_win.destroy()
                self.deiconify()
                self.build_main_ui()
            else:
                messagebox.showerror("Greška pri prijavi", "Netačno korisničko ime ili lozinka!", parent=login_win)

        login_win.bind('<Return>', lambda e: check_credentials())
        
        tk.Button(
            login_win, text="PRIJAVI SE", font=('Segoe UI', 10, 'bold'), bg="#0284C7", fg="white", 
            bd=0, padx=20, pady=8, activebackground="#0369A1", activeforeground="white", command=check_credentials
        ).pack(pady=20)

    def build_main_ui(self):
        self.bind('<Control-z>', lambda e: self.perform_undo())
        self.bind('<Control-y>', lambda e: self.perform_redo())

        # Zaglavlje
        header = ttk.Frame(self)
        header.pack(fill='x', padx=20, pady=15)
        
        ttk.Label(header, text="Sistem za upravljanje LZO opremom - Region 4", style='Header.TLabel').pack(side='left')
        
        btn_box = ttk.Frame(header)
        btn_box.pack(side='right')
        ttk.Button(btn_box, text="↶ Undo (Ctrl+Z)", style='Secondary.TButton', command=self.perform_undo).pack(side='left', padx=3)
        ttk.Button(btn_box, text="↷ Redo (Ctrl+Y)", style='Secondary.TButton', command=self.perform_redo).pack(side='left', padx=3)

        # Pretraga
        search_frame = ttk.LabelFrame(self, text=" Pretraga i filtri ")
        search_frame.pack(fill='x', padx=20, pady=10)

        ttk.Label(search_frame, text="Zaposleni:").pack(side='left', padx=10, pady=12)
        
        self.search_combo = SearchableCombobox(search_frame, width=35)
        self.search_combo.pack(side='left', padx=5, pady=12)

        ttk.Button(search_frame, text="Generiši PDF Izvještaj", style='Primary.TButton', command=self.export_pdf).pack(side='right', padx=10, pady=12)

        # Tabela
        table_frame = ttk.Frame(self)
        table_frame.pack(fill='both', expand=True, padx=20, pady=10)

        columns = ("Zaposleni", "Oprema", "Datum", "Status")
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings')
        
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor='w', width=180)

        self.tree.pack(fill='both', expand=True)
        self._load_sample_data()

    def _load_sample_data(self):
        initial_data = [
            {"Zaposleni": "Željko Čurović", "Oprema": "Zaštitni šlem", "Datum": "2024-01-15", "Status": "U roku"},
            {"Zaposleni": "Šćepan Đurović", "Oprema": "Rukavice 10kV", "Datum": "2023-05-10", "Status": "Isteklo"},
            {"Zaposleni": "Miloš Žižić", "Oprema": "Zaštitne cipele S3", "Datum": "2023-11-01", "Status": "U isteku"}
        ]
        self.data = pd.DataFrame(initial_data)
        self.search_combo.set_completion_list(self.data["Zaposleni"].unique().tolist())
        self.refresh_tree()

    def refresh_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for _, row in self.data.iterrows():
            self.tree.insert("", "end", values=(row["Zaposleni"], row["Oprema"], row["Datum"], row["Status"]))

    def perform_undo(self):
        prev = self.undo_mgr.undo(self.data)
        if prev is not None:
            self.data = prev
            self.refresh_tree()

    def perform_redo(self):
        nxt = self.undo_mgr.redo(self.data)
        if nxt is not None:
            self.data = nxt
            self.refresh_tree()

    def export_pdf(self):
        try:
            generate_pdf_report("LZO_Izvjestaj_Region4.pdf", self.data)
            messagebox.showinfo("Uspjeh", "PDF izvještaj je uspješno izgenerisan!")
        except Exception as e:
            messagebox.showerror("Greška", f"Greška pri izradi PDF izvještaja:\n{e}")

if __name__ == "__main__":
    app = LZOApp()
    app.mainloop()
