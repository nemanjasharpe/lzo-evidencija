import streamlit as st
import pandas as pd
import datetime
from dateutil.relativedelta import relativedelta
import plotly.express as px
import io

# Podešavanje stranice
st.set_page_config(page_title="CEDIS - Evidencija ZZO", layout="wide")

# --- POMOĆNE FUNKCIJE ZA PRORAČUN ---
def izracunaj_rokove(df):
    """Obračunava Datum isticanja i Preostalo dana u odnosu na današnji datum."""
    danas = pd.to_datetime(datetime.date.today())
    
    datum_isticanja = []
    dani_razlika = []
    statusi = []

    for _, row in df.iterrows():
        try:
            dt_zaduzivanje = pd.to_datetime(row["Datum zaduženja"])
            izdata_kol = float(row["Količina izdata"])
            rok_mjeseci = float(row["Rok (mjeseci)"])
            
            # Množenje izdate količine sa rokom u mjesecima
            ukupno_mjeseci = int(izdata_kol * rok_mjeseci)
            dt_isteka = dt_zaduzivanje + relativedelta(months=ukupno_mjeseci)
            razlika = (dt_isteka - danas).days
            
            datum_isticanja.append(dt_isteka.strftime("%Y-%m-%d"))
            dani_razlika.append(razlika)
            
            if razlika < 0:
                statusi.append("Isteklo")
            elif razlika <= 30:
                statusi.append("Ističe uskoro (≤30 dana)")
            else:
                statusi.append("U roku")
        except Exception:
            datum_isticanja.append(None)
            dani_razlika.append(None)
            statusi.append("Nepoznato")

    df["Datum isticanja"] = datum_isticanja
    df["Razlika (dana)"] = dani_razlika
    df["Status"] = statusi
    return df

# Initialize Session State Data
if "data" not in st.session_state:
    # Inicijalna struktura tabele prema specifikaciji (A - O)
    cols = [
        "Selektovano", "Ime i prezime", "Radno mjesto", "Organizaciona jedinica", "Grad",
        "Veličina odjeće", "Veličina obuće", "Naziv opreme", "Jedinica mjere",
        "Količina po normativu", "Rok (mjeseci)", "Količina izdata",
        "Datum zaduženja", "Datum isticanja", "Razlika (dana)", "Napomena", "Status"
    ]
    # Ogledni podaci
    initial_data = [
        [False, "Marko Marković", "Elektromonter", "Sektor za održavanje", "Podgorica", "L", "43", "Zaštitno odijelo", "komadi", 1, 12, 1, "2025-09-01", "", 0, "Redovna podjela", ""],
        [False, "Ivan Ivanović", "Inženjer ispitivanja", "Služba za ZZNR i lab. ispitivanja", "Herceg Novi", "XL", "45", "Zaštitne rukavice 10kV", "pari", 2, 6, 1, "2026-01-15", "", 0, "Ograničene zalihe", ""],
    ]
    df_init = pd.DataFrame(initial_data, columns=cols)
    st.session_state.data = izracunaj_rokove(df_init)

# --- ZAGLAVLJE SA LOGO-OM CEDIS-A ---
col_logo, col_title = st.columns([1, 4])
with col_logo:
    # Prikaz CEDIS logotipa (možete zamijeniti putanju sa lokalnom slikom 'cedis_logo.png')
    st.image("https://cedis.me/wp-content/uploads/2017/03/cedis-logo.png", width=180, use_container_width=False)
with col_title:
    st.title("Sistem za Evidenciju i Upravljanje ZZO - CEDIS")
    st.caption("Praćenje zaduženja, normativa i rokova trajanja zaštitne opreme")

st.markdown("---")

# --- BOČNA TRAKA (UVOZ EXCEL-A / NOVI UNOS) ---
with st.sidebar:
    st.header("⚙️ Upravljanje podacima")
    
    # Uvoz iz Excela
    uploaded_file = st.file_uploader("Učitaj Excel tabelu", type=["xlsx", "xls"])
    if uploaded_file is not None:
        try:
            df_uploaded = pd.read_excel(uploaded_file)
            df_uploaded["Selektovano"] = False
            st.session_state.data = izracunaj_rokove(df_uploaded)
            st.success("Tabela uspješno učitana!")
        except Exception as e:
            st.error(f"Greška pri učitavanju fajla: {e}")

    st.markdown("---")
    st.subheader("➕ Dodaj novog zaposlenog / opremu")
    with st.form("forma_novi_unos", clear_on_submit=True):
        ime = st.text_input("Ime i prezime")
        rm = st.text_input("Radno mjesto")
        org = st.text_input("Organizaciona jedinica")
        grad = st.text_input("Grad")
        v_odj = st.text_input("Veličina odjeće")
        v_obu = st.text_input("Veličina obuće")
        oprema = st.text_input("Naziv opreme")
        jm = st.selectbox("Jedinica mjere", ["komadi", "pari"])
        normativ = st.number_input("Količina po normativu", min_value=1, value=1)
        rok = st.number_input("Rok (u mjesecima)", min_value=1, value=12)
        izdata = st.number_input("Izdata količina", min_value=1, value=1)
        dt_zad = st.date_input("Datum zaduženja", datetime.date.today())
        napomena = st.text_area("Napomena", "")
        
        submitted = st.form_submit_button("Sačuvaj unos")
        if submitted and ime:
            novi_red = {
                "Selektovano": False,
                "Ime i prezime": ime,
                "Radno mjesto": rm,
                "Organizaciona jedinica": org,
                "Grad": grad,
                "Veličina odjeće": v_odj,
                "Veličina obuće": v_obu,
                "Naziv opreme": oprema,
                "Jedinica mjere": jm,
                "Količina po normativu": normativ,
                "Rok (mjeseci)": rok,
                "Količina izdata": izdata,
                "Datum zaduženja": str(dt_zad),
                "Datum isticanja": "",
                "Razlika (dana)": 0,
                "Napomena": napomena,
                "Status": ""
            }
            st.session_state.data = pd.concat([st.session_state.data, pd.DataFrame([novi_red])], ignore_index=True)
            st.session_state.data = izracunaj_rokove(st.session_state.data)
            st.success("Podatak uspješno dodat!")
            st.rerun()

# --- PRETRAGA I FILTRIRANJE ---
st.subheader("🔍 Pretraga i filtriranje podataka")
f_col1, f_col2, f_col3, f_col4 = st.columns(4)

with f_col1:
    search_text = st.text_input("Pretraga (Ime, Oprema, RM):", "")
with f_col2:
    org_filter = st.multiselect("Organizaciona jedinica:", options=st.session_state.data["Organizaciona jedinica"].unique())
with f_col3:
    grad_filter = st.multiselect("Grad:", options=st.session_state.data["Grad"].unique())
with f_col4:
    status_filter = st.multiselect("Status opreme:", options=["Isteklo", "Ističe uskoro (≤30 dana)", "U roku"])

# Primjena filtera
df_filtered = st.session_state.data.copy()

if search_text:
    df_filtered = df_filtered[
        df_filtered["Ime i prezime"].astype(str).str.contains(search_text, case=False) |
        df_filtered["Naziv opreme"].astype(str).str.contains(search_text, case=False) |
        df_filtered["Radno mjesto"].astype(str).str.contains(search_text, case=False)
    ]
if org_filter:
    df_filtered = df_filtered[df_filtered["Organizaciona jedinica"].isin(org_filter)]
if grad_filter:
    df_filtered = df_filtered[df_filtered["Grad"].isin(grad_filter)]
if status_filter:
    df_filtered = df_filtered[df_filtered["Status"].isin(status_filter)]

# --- ANALITIKA I GRAFIKON STATUSI OPREME ---
st.markdown("### 📊 Status i analitika opreme")
col_chart, col_stats = st.columns([2, 1])

status_counts = df_filtered["Status"].value_counts().reset_index()
status_counts.columns = ["Status", "Broj stavki"]

with col_chart:
    fig = px.pie(
        status_counts, 
        values="Broj stavki", 
        names="Status", 
        title="Distribucija statusa opreme",
        color="Status",
        color_discrete_map={
            "Isteklo": "#EF553B", 
            "Ističe uskoro (≤30 dana)": "#FECB52", 
            "U roku": "#00CC96"
        },
        hole=0.4
    )
    st.plotly_chart(fig, use_container_width=True)

with col_stats:
    st.metric("Ukupno filtriranih unosa", len(df_filtered))
    st.metric("Istekla oprema", len(df_filtered[df_filtered["Status"] == "Isteklo"]))
    st.metric("Ističe u narednih 30 dana", len(df_filtered[df_filtered["Status"] == "Ističe uskoro (≤30 dana)"]))

# --- Tabela i Mogućnost Višestrukog Brisanja / Izmjene ---
st.markdown("### 📋 Evidenciona tabela opreme")

# DUGMAD ZA AKCIJE (Višestruko brisanje)
btn_col1, btn_col2 = st.columns([2, 8])
with btn_col1:
    if st.button("🗑️ Izbriši selektovane stavke", type="primary"):
        st.session_state.data = st.session_state.data[~st.session_state.data["Selektovano"]].reset_index(drop=True)
        st.session_state.data = izracunaj_rokove(st.session_state.data)
        st.success("Odabrane stavke su uklonjene.")
        st.rerun()

# Interaktivni prikaz tabele sa selekcijom
edited_df = st.data_editor(
    df_filtered,
    column_config={
        "Selektovano": st.column_config.CheckboxColumn("Izbor", default=False),
        "Status": st.column_config.TextColumn("Status", disabled=True),
        "Datum isticanja": st.column_config.TextColumn("Datum isticanja", disabled=True),
        "Razlika (dana)": st.column_config.NumberColumn("Razlika (dana)", disabled=True)
    },
    disabled=["Datum isticanja", "Razlika (dana)", "Status"],
    hide_index=True,
    use_container_width=True
)

# Sinhronizacija izmjena u tabeli nazad u session state
if not edited_df.equals(df_filtered):
    st.session_state.data.update(edited_df)
    st.session_state.data = izracunaj_rokove(st.session_state.data)

# --- OPCIJA ZA IZMJENU POJEDINAČNE STAVKE (✏️ Olovka / Modalni prozor) ---
st.markdown("### ✏️ Izmjena pojedinačnih stavki")
izbor_indeks = st.selectbox(
    "Odaberite red za detaljnu izmjenu podataka (Ime - Oprema):",
    options=st.session_state.data.index,
    format_func=lambda x: f"Red {x+1}: {st.session_state.data.loc[x, 'Ime i prezime']} - {st.session_state.data.loc[x, 'Naziv opreme']}"
)

if izbor_indeks is not None:
    with st.expander(f"📝 Izmijeni podatke za stavku br. {izbor_indeks + 1}", expanded=False):
        row = st.session_state.data.loc[izbor_indeks]
        with st.form(f"edit_form_{izbor_indeks}"):
            e_col1, e_col2, e_col3 = st.columns(3)
            with e_col1:
                e_ime = st.text_input("Ime i prezime", value=row["Ime i prezime"])
                e_rm = st.text_input("Radno mjesto", value=row["Radno mjesto"])
                e_org = st.text_input("Organizaciona jedinica", value=row["Organizaciona jedinica"])
                e_grad = st.text_input("Grad", value=row["Grad"])
                e_v_odj = st.text_input("Veličina odjeće", value=str(row["Veličina odjeće"]))
            with e_col2:
                e_v_obu = st.text_input("Veličina obuće", value=str(row["Veličina obuće"]))
                e_oprema = st.text_input("Naziv opreme", value=row["Naziv opreme"])
                e_jm = st.selectbox("Jedinica mjere", ["komadi", "pari"], index=0 if row["Jedinica mjere"] == "komadi" else 1)
                e_norm = st.number_input("Količina po normativu", value=int(row["Količina po normativu"]))
                e_rok = st.number_input("Rok (mjeseci)", value=int(row["Rok (mjeseci)"]))
            with e_col3:
                e_izdata = st.number_input("Izdata količina", value=int(row["Količina izdata"]))
                e_dt = st.date_input("Datum zaduženja", value=pd.to_datetime(row["Datum zaduženja"]).date())
                e_napomena = st.text_area("Napomena", value=str(row["Napomena"]))

            if st.form_submit_button("Sačuvaj izmjene stavke"):
                st.session_state.data.loc[izbor_indeks, "Ime i prezime"] = e_ime
                st.session_state.data.loc[izbor_indeks, "Radno mjesto"] = e_rm
                st.session_state.data.loc[izbor_indeks, "Organizaciona jedinica"] = e_org
                st.session_state.data.loc[izbor_indeks, "Grad"] = e_grad
                st.session_state.data.loc[izbor_indeks, "Veličina odjeće"] = e_v_odj
                st.session_state.data.loc[izbor_indeks, "Veličina obuće"] = e_v_obu
                st.session_state.data.loc[izbor_indeks, "Naziv opreme"] = e_oprema
                st.session_state.data.loc[izbor_indeks, "Jedinica mjere"] = e_jm
                st.session_state.data.loc[izbor_indeks, "Količina po normativu"] = e_norm
                st.session_state.data.loc[izbor_indeks, "Rok (mjeseci)"] = e_rok
                st.session_state.data.loc[izbor_indeks, "Količina izdata"] = e_izdata
                st.session_state.data.loc[izbor_indeks, "Datum zaduženja"] = str(e_dt)
                st.session_state.data.loc[izbor_indeks, "Napomena"] = e_napomena
                
                st.session_state.data = izracunaj_rokove(st.session_state.data)
                st.success("Ažurirano!")
                st.rerun()

# --- IZVOZ IZVJEŠTAJA (EXCEL & PDF) ---
st.markdown("---")
st.subheader("📥 Izvoz izvještaja")

exp_col1, exp_col2 = st.columns(2)

# Excel Export
with exp_col1:
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        df_filtered.drop(columns=["Selektovano"]).to_excel(writer, index=False, sheet_name="ZZO Evidencija")
    
    st.download_button(
        label="📊 Preuzmi Excel izvještaj (.xlsx)",
        data=excel_buffer.getvalue(),
        file_name=f"CEDIS_ZZO_Izvjestaj_{datetime.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# PDF Export (Koristi HTML/CSS konverziju za čist i pregledan prikaz)
with exp_col2:
    html_content = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; font-size: 11px; }}
            h2 {{ color: #003366; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
            th, td {{ border: 1px solid #ddd; padding: 6px; text-align: left; }}
            th {{ background-color: #f2f2f2; color: #333; }}
            .isteklo {{ background-color: #ffcccc; }}
            .uskoro {{ background-color: #fff3cd; }}
        </style>
    </head>
    <body>
        <h2>Crnogorski elektrodistributivni sistem (CEDIS)</h2>
        <h3>Izvještaj o zaduženju zaštitne opreme na dan: {datetime.date.today().strftime('%d.%m.%Y')}</h3>
        <p><b>Ukupno evidentiranih stavki u izvještaju:</b> {len(df_filtered)}</p>
        <table>
            <tr>
                <th>Zaposleni</th><th>Radno mjesto</th><th>Org. jedinica</th><th>Grad</th>
                <th>Oprema</th><th>Izdata kol.</th><th>Datum zad.</th><th>Datum isteka</th><th>Status</th>
            </tr>
            {"".join([f"<tr class='{'isteklo' if r['Status']=='Isteklo' else ('uskoro' if r['Status']=='Ističe uskoro (≤30 dana)' else '')}'><td>{r['Ime i prezime']}</td><td>{r['Radno mjesto']}</td><td>{r['Organizaciona jedinica']}</td><td>{r['Grad']}</td><td>{r['Naziv opreme']}</td><td>{r['Količina izdata']} {r['Jedinica mjere']}</td><td>{r['Datum zaduženja']}</td><td>{r['Datum isticanja']}</td><td>{r['Status']}</td></tr>" for _, r in df_filtered.iterrows()])}
        </table>
    </body>
    </html>
    """
    
    st.download_button(
        label="📄 Preuzmi PDF/HTML Izvještaj (.html / .pdf)",
        data=html_content,
        file_name=f"CEDIS_ZZO_Izvjestaj_{datetime.date.today()}.html",
        mime="text/html"
    )
