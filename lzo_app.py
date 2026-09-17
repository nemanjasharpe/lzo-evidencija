import streamlit as st
import pandas as pd
import datetime
from dateutil.relativedelta import relativedelta
import io
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ---------------------------------------------------------
# 1. PODEŠAVANJE STRANICE & LOGO
# ---------------------------------------------------------
st.set_page_config(page_title="CEDIS - Evidencija LZO", layout="wide", page_icon="🛡️")

# Inicijalizacija Session State
if "data" not in st.session_state:
    st.session_state.data = pd.DataFrame()

# ---------------------------------------------------------
# 2. LOGIKA ZA PRORAČUN DATUMA (KOLONE M I N)
# ---------------------------------------------------------
def recalculate_lzo(df):
    if df.empty:
        return df

    today = pd.to_datetime(datetime.date.today())
    d_isteka, d_razlika, statusi = [], [], []

    for _, row in df.iterrows():
        try:
            dt_zad = pd.to_datetime(row['Datum zaduženja'])
            if pd.isna(dt_zad):
                raise ValueError("Nedostaje datum")

            izdata_kol = float(row['Izdata količina']) if pd.notna(row['Izdata količina']) else 1.0
            rok_mjeseci = float(row['Rok (mjeseci)']) if pd.notna(row['Rok (mjeseci)']) else 12.0

            total_months = int(izdata_kol * rok_mjeseci)
            dt_ist = dt_zad + relativedelta(months=total_months)
            razlika_dana = (dt_ist - today).days

            d_isteka.append(dt_ist.strftime('%Y-%m-%d'))
            d_razlika.append(razlika_dana)

            if razlika_dana < 0:
                statusi.append("Isteklo")
            elif razlika_dana <= 100:
                statusi.append("Pripremiti nabavku (≤100d)")
            else:
                statusi.append("U roku")
        except Exception:
            d_isteka.append("")
            d_razlika.append(None)
            statusi.append("Nedostaje datum")

    df['Datum isticanja'] = d_isteka
    df['Razlika (dani)'] = d_razlika
    df['Status'] = statusi
    return df

# ---------------------------------------------------------
# 3. GENERISANJE PDF IZVJEŠTAJA
# ---------------------------------------------------------
def generate_pdf_report(df_filtered, logo_bytes=None):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(A4),
        rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20
    )
    story = []
    styles = getSampleStyleSheet()

    # Zaglavlje i Logo
    header_data = []
    if logo_bytes:
        try:
            logo_img = Image(io.BytesIO(logo_bytes), width=120, height=45)
            header_data.append([logo_img, Paragraph("<b>CEDIS d.o.o. Podgorica</b><br/>Služba za ZZNR i laboratorijska ispitivanja", styles['Heading3'])])
        except Exception:
            header_data.append(["CEDIS", Paragraph("<b>CEDIS d.o.o. Podgorica</b><br/>Služba za ZZNR i laboratorijska ispitivanja", styles['Heading3'])])
    else:
        header_data.append(["CEDIS LOGO", Paragraph("<b>CEDIS d.o.o. Podgorica</b><br/>Služba za ZZNR i laboratorijska ispitivanja", styles['Heading3'])])

    header_table = Table(header_data, colWidths=[130, 650])
    header_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
    story.append(header_table)
    story.append(Spacer(1, 10))

    story.append(Paragraph("<b>IZVJEŠTAJ O STANJU LIČNE ZAŠTITNE OPREME (Obrazac 8)</b>", styles['Title']))
    story.append(Spacer(1, 10))

    # Generisanje dijagrama
    fig, ax = plt.subplots(figsize=(6, 2.5))
    counts = df_filtered['Status'].value_counts()
    colors_map = {'Isteklo': '#d32f2f', 'Pripremiti nabavku (≤100d)': '#ffa000', 'U roku': '#388e3c'}
    bar_colors = [colors_map.get(x, '#757575') for x in counts.index]
    
    ax.bar(counts.index, counts.values, color=bar_colors)
    ax.set_ylabel("Broj komada/pari")
    ax.set_title("Pregled statusa zadužene LZO opreme")
    plt.tight_layout()
    
    chart_img_buf = io.BytesIO()
    plt.savefig(chart_img_buf, format='png', dpi=120)
    plt.close(fig)
    chart_img_buf.seek(0)
    
    story.append(Image(chart_img_buf, width=400, height=160))
    story.append(Spacer(1, 15))

    # Tabela sa podacima
    cols = ['Zaposleni', 'Radno mjesto', 'Grad', 'Naziv opreme', 'Datum zaduženja', 'Datum isticanja', 'Status', 'NAPOMENA']
    df_export = df_filtered[cols].copy().fillna('')
    
    table_data = [cols] + df_export.values.tolist()
    t = Table(table_data, colWidths=[110, 140, 60, 150, 75, 75, 80, 100])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0288d1")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
        ('FONTSIZE', (0,0), (-1,-1), 7),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    story.append(t)
    doc.build(story)
    buffer.seek(0)
    return buffer

# ---------------------------------------------------------
# 4. GLAVNI INTERFEJS
# ---------------------------------------------------------
st.title("🛡️ Evidencija Lične Zaštitne Opreme (Obrazac 8)")

# Bočna traka - Učitavanje fajla
with st.sidebar:
    st.header("⚙️ Postavke i Učitavanje")
    uploaded_file = st.file_uploader("Učitajte Excel fajl (Obrazac 8)", type=["xlsx", "xls"])
    cedis_logo = st.file_uploader("Učitajte CEDIS Logo (PNG/JPG)", type=["png", "jpg", "jpeg"])

if uploaded_file and st.session_state.data.empty:
    df_raw = pd.read_excel(uploaded_file, sheet_name='1. Održavanje')
    df_raw.rename(columns={
        'RM': 'Radno mjesto',
        'Konfekcijski broj - veličina': 'Veličina obuće',
        'Unnamed: 5': 'Veličina odjeće',
        'Naziv sredstva/opreme': 'Naziv opreme',
        'Količina (Normativ)': 'Normativ',
        'Izdata količina': 'Izdata količina',
        'NAPOMENA': 'NAPOMENA'
    }, inplace=True)

    # Popunjavanje objedinjenih polja (ffill)
    fill_cols = ['Zaposleni', 'Radno mjesto', 'Organizaciona jedinica', 'Grad', 'Veličina obuće', 'Veličina odjeće']
    df_raw[fill_cols] = df_raw[fill_cols].ffill()
    df_raw.dropna(subset=['Naziv opreme'], inplace=True)
    st.session_state.data = recalculate_lzo(df_raw)

if not st.session_state.data.empty:
    df = st.session_state.data

    # --- FILTRIRANJE I PRETRAGA ---
    st.subheader("🔍 Pretraga i Filtriranje")
    c1, c2, c3, c4, c5 = st.columns(5)
    
    with c1:
        search_query = st.text_input("Tekstualna pretraga", placeholder="Pretraži...")
    with c2:
        grad_filter = st.multiselect("Grad", options=sorted(df['Grad'].dropna().unique()))
    with c3:
        org_filter = st.multiselect("Org. jedinica", options=sorted(df['Organizaciona jedinica'].dropna().unique()))
    with c4:
        oprema_filter = st.multiselect("Naziv opreme", options=sorted(df['Naziv opreme'].dropna().unique()))
    with c5:
        status_filter = st.multiselect("Status", options=sorted(df['Status'].dropna().unique()))

    # Primjena filtera
    df_filtered = df.copy()
    if search_query:
        df_filtered = df_filtered[df_filtered.astype(str).apply(lambda row: row.str.lower().str.contains(search_query.lower()).any(), axis=1)]
    if grad_filter:
        df_filtered = df_filtered[df_filtered['Grad'].isin(grad_filter)]
    if org_filter:
        df_filtered = df_filtered[df_filtered['Organizaciona jedinica'].isin(org_filter)]
    if oprema_filter:
        df_filtered = df_filtered[df_filtered['Naziv opreme'].isin(oprema_filter)]
    if status_filter:
        df_filtered = df_filtered[df_filtered['Status'].isin(status_filter)]

    # --- PREDSTAVLJANJE KROZ DATA EDITOR (OMOGUĆAVA DIREKTNU IZMJENU I BRISANJE) ---
    st.subheader("📋 Tabela LZO Opreme")
    st.caption("Možete mijenjati podatke direktno u tabeli ili koristiti opciju višestrukog brisanja redova.")

    # Tabela za brisanje i izmjenu
    edited_df = st.data_editor(
        df_filtered,
        num_rows="dynamic",
        use_container_width=True,
        key="editor"
    )

    if st.button("💾 Sačuvaj izmjene u tabeli"):
        st.session_state.data = recalculate_lzo(edited_df)
        st.success("Izmjene uspješno sačuvane i ponovo proračunate!")
        st.rerun()

    # --- PREDMETNA IZMJENA POMOĆU FORME (POJEDINAČNA "OLOVKA" REŽIM) ---
    with st.expander("✏️ Pojedinačna detaljna izmjena stavke"):
        selected_index = st.number_input("Izaberite Redni Broj (Index) za izmjenu", min_value=0, max_value=len(df)-1, step=1)
        if selected_index < len(df):
            row = df.iloc[selected_index]
            with st.form("edit_form"):
                e_zap = st.text_input("Zaposleni", value=str(row['Zaposleni']))
                e_rm = st.text_input("Radno mjesto", value=str(row['Radno mjesto']))
                e_grad = st.text_input("Grad", value=str(row['Grad']))
                e_opr = st.text_input("Naziv opreme", value=str(row['Naziv opreme']))
                e_kol = st.number_input("Izdata količina", value=float(row['Izdata količina']) if pd.notna(row['Izdata količina']) else 1.0)
                e_rok = st.number_input("Rok (mjeseci)", value=float(row['Rok (mjeseci)']) if pd.notna(row['Rok (mjeseci)']) else 12.0)
                e_dat = st.text_input("Datum zaduženja (YYYY-MM-DD)", value=str(row['Datum zaduženja']))
                e_nap = st.text_area("NAPOMENA", value=str(row['NAPOMENA']) if pd.notna(row['NAPOMENA']) else "")
                
                if st.form_submit_button("Ažuriraj stavku"):
                    st.session_state.data.at[selected_index, 'Zaposleni'] = e_zap
                    st.session_state.data.at[selected_index, 'Radno mjesto'] = e_rm
                    st.session_state.data.at[selected_index, 'Grad'] = e_grad
                    st.session_state.data.at[selected_index, 'Naziv opreme'] = e_opr
                    st.session_state.data.at[selected_index, 'Izdata količina'] = e_kol
                    st.session_state.data.at[selected_index, 'Rok (mjeseci)'] = e_rok
                    st.session_state.data.at[selected_index, 'Datum zaduženja'] = e_dat
                    st.session_state.data.at[selected_index, 'NAPOMENA'] = e_nap
                    st.session_state.data = recalculate_lzo(st.session_state.data)
                    st.success("Stavka uspješno ažurirana!")
                    st.rerun()

    # --- IZVOZ IZVJEŠTAJA ---
    st.subheader("📤 Izvoz Izvještaja")
    col_e1, col_e2 = st.columns(2)

    # Excel izvoz
    with col_e1:
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
            df_filtered.to_excel(writer, sheet_name='LZO Izvještaj', index=False)
        excel_buffer.seek(0)
        
        st.download_button(
            label="📥 Preuzmi Excel Izvještaj",
            data=excel_buffer,
            file_name=f"LZO_Izvjestaj_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # PDF izvoz
    with col_e2:
        logo_data = cedis_logo.getvalue() if cedis_logo else None
        pdf_buffer = generate_pdf_report(df_filtered, logo_data)
        
        st.download_button(
            label="📄 Preuzmi PDF Izvještaj (sa Grafikonom i CEDIS Logom)",
            data=pdf_buffer,
            file_name=f"LZO_Izvjestaj_{datetime.date.today()}.pdf",
            mime="application/pdf"
        )
else:
    st.info("Molimo vas da učitate Excel fajl `Evidencija 8. - LZO R4 test_3.xlsx` u bočnoj traci.")
