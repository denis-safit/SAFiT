import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE PAGINA E VERSIONE ---
APP_VERSION = "1.0.03"
st.set_page_config(page_title=f"Safit - Portale Avanzamento {APP_VERSION}", layout="wide")

st.markdown(f"""
    <style>
    .main {{ background-color: #fcfcfc; }}
    .stApp {{ margin-top: -30px; }}
    .status-row {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap; 
        padding: 12px 15px;
        border-radius: 8px;
        margin-bottom: 8px;
        gap: 10px;
        font-size: 14px;
    }}
    .delay-row {{ background-color: #fff8e1; border-left: 6px solid #ffc107; color: #5d4037; }}
    .on-time-row {{ background-color: #f1f8e9; border-left: 6px solid #4caf50; color: #1b5e20; }}
    .client-delay-row {{ background-color: #e3f2fd; border-left: 6px solid #2196f3; color: #0d47a1; }}
    .stExpander div {{ height: auto !important; min-height: min-content !important; }}
    .version-tag {{ font-size: 10px; color: #999; text-align: right; }}
    </style>
    """, unsafe_allow_html=True)

# --- 2. GESTIONE UTENTI ---
@st.cache_data
def load_users():
    file_u = 'utenti.xlsx'
    if os.path.exists(file_u):
        try:
            df_u = pd.read_excel(file_u)
            df_u.columns = [str(c).strip() for c in df_u.columns]
            return df_u.set_index('username')[['password', 'cliente_arca']].T.to_dict('list')
        except:
            return {'safit_admin': ['admin2026', 'TUTTI']}
    return {'safit_admin': ['admin2026', 'TUTTI']}

USER_DB = load_users()

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if not st.session_state["authenticated"]:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=300)
            st.title("Accesso Area Riservata")
            user = st.text_input("Username").strip()
            pw = st.text_input("Password", type="password").strip()
            if st.button("Accedi"):
                if user in USER_DB and str(USER_DB[user][0]) == pw:
                    st.session_state["authenticated"] = True
                    st.session_state["user_type"] = USER_DB[user][1]
                    st.session_state["username"] = user
                    st.rerun()
                else: st.error("Username o Password errati")
        return False
    return True

if not check_password(): st.stop()

# --- 3. FUNZIONI TECNICHE ---
def aggiungi_giorni_lavorativi(data_inizio, giorni):
    data_corrente = data_inizio
    while giorni > 0:
        data_corrente += timedelta(days=1)
        if data_corrente.weekday() < 5: giorni -= 1
    return data_corrente

def pulisci_numero(serie):
    return pd.to_numeric(serie.astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').fillna(0)

def standardizza_codice(serie):
    return serie.astype(str).str.replace(r'\s+', '', regex=True).str.strip().str.upper()

@st.cache_data
def load_data():
    try:
        # 1. Caricamento ARCA
        df = pd.read_excel('righe_Ordini_ARCA.xlsx', sheet_name='Foglio1', skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        if 'Articolo C' in df.columns:
            df['Articolo C'] = standardizza_codice(df['Articolo C'])
        for col in ['Cliente Fornitore CD', 'Articolo C', 'Articolo D', 'Data']:
            if col in df.columns: df[col] = df[col].ffill()
        df['Data_Consegna'] = pd.to_datetime(df['Data'], errors='coerce')
        col_res = 'Qta Residua' if 'Qta Residua' in df.columns else 'Qta Doc'
        df['Qta_Effettiva'] = pd.to_numeric(df[col_res], errors='coerce').fillna(0)
        df = df[df['Qta_Effettiva'] > 0]
        
        # 2. Caricamento ACCESS (Logica per Posizione Colonne)
        file_tech = 'Avanzamento_access.xlsx'
        if os.path.exists(file_tech):
            # Carichiamo TUTTO il file senza saltare righe per vedere cosa c'è dentro
            df_tech = pd.read_excel(file_tech)
            
            # Se il file ha almeno 2 colonne, proviamo a forzare i nomi
            if df_tech.shape[1] >= 2:
                # Cerchiamo la riga dove iniziano i dati (spesso c'è sporcizia sopra)
                # Proviamo a rinominare le colonne in base a quello che troviamo
                df_tech.columns = [str(c).strip() for c in df_tech.columns]
                
                # Se non troviamo 'Codice', prendiamo la prima colonna che somiglia a un codice
                col_codice = 'Codice' if 'Codice' in df_tech.columns else df_tech.columns[0]
                col_gia = 'Gia' if 'Gia' in df_tech.columns else (df_tech.columns[1] if df_tech.shape[1] > 1 else None)
                
                df_tech = df_tech.rename(columns={col_codice: 'Art_Key', col_gia: 'Gia'})
                df_tech['Art_Key'] = standardizza_codice(df_tech['Art_Key'])
                
                # Pulizia numeri
                for c in ['Gia', 'Acq', 'Lan', 'GRZ', 'TMP', 'RWI', 'TRS']:
                    if c in df_tech.columns: df_tech[c] = pulisci_numero(df_tech[c])
                
                # Somma lavorazioni
                lavoraz_cols = [c for c in ['Acq', 'Lan', 'GRZ', 'TMP', 'RWI', 'TRS'] if c in df_tech.columns]
                df_tech['Lavorazione_Totale'] = df_tech[lavoraz_cols].sum(axis=1) if lavoraz_cols else 0
                
                # Merge
                df = pd.merge(df, df_tech[['Art_Key', 'Gia', 'Lavorazione_Totale']], 
                              left_on='Articolo C', right_on='Art_Key', how='left')
            else:
                st.error("Il file Access sembra quasi vuoto (meno di 2 colonne)")
        else:
            st.error("File Avanzamento_access.xlsx NON TROVATO sul server!")
            
        return df
    except Exception as e:
        st.error(f"Errore caricamento: {e}")
        return pd.DataFrame()

data = load_data()
oggi_dt = datetime.now()

# --- 4. DIAGNOSTICA DI EMERGENZA ---
if not data.empty:
    with st.expander("🔍 DIAGNOSTICA DI EMERGENZA"):
        st.write("File caricati correttamente. Colonne presenti:", data.columns.tolist())
        match_trouvati = data['Gia'].notna().sum() if 'Gia' in data.columns else 0
        st.write(f"Righe con giacenza collegata: {match_trouvati} su {len(data)}")
        if match_trouvati == 0:
            st.warning("⚠️ ATTENZIONE: Nessun codice articolo di Arca corrisponde a quelli di Access.")
            st.write("Esempio codici Arca:", data['Articolo C'].head(3).tolist())

# --- 5. SIDEBAR ---
with st.sidebar:
    if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
    st.write(f"Utente: **{st.session_state.get('username')}**")
    st.markdown(f'<p class="version-tag">Versione: {APP_VERSION}</p>', unsafe_allow_html=True)
    
    if st.session_state["user_type"] == "TUTTI":
        c_list = sorted([str(x) for x in data['Cliente Fornitore CD'].unique()]) if not data.empty else []
        sel_cli = st.selectbox("👤 Seleziona Cliente:", c_list)
    else: sel_cli = st.session_state["user_type"]
    
    filtro_label = st.radio("Stato ordini:", ["Mostra tutto", "Solo Disponibili", "In Lavorazione", "In Ritardo"])
    if st.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

# --- 6. INTERFACCIA CENTRALE ---
st.title("Portale Avanzamento Produzione")

if not data.empty:
    df_cli = data[data['Cliente Fornitore CD'] == sel_cli].copy()
    if not df_cli.empty:
        articoli_view = sorted([str(x) for x in df_cli['Articolo C'].unique()])
        sel_art = st.selectbox("🔍 Cerca Prodotto:", ["Tutti i prodotti"] + articoli_view)
        lista_art = articoli_view if sel_art == "Tutti i prodotti" else [sel_art]

        for art in lista_art:
            df_art = df_cli[df_cli['Articolo C'] == art].sort_values('Data_Consegna')
            desc = df_art['Articolo D'].iloc[0] if 'Articolo D' in df_art.columns else ""
            
            st_gia = float(df_art['Gia'].iloc[0]) if 'Gia' in df_art.columns and pd.notnull(df_art['Gia'].iloc[0]) else 0.0
            st_lav = float(df_art['Lavorazione_Totale'].iloc[0]) if 'Lavorazione_Totale' in df_art.columns and pd.notnull(df_art['Lavorazione_Totale'].iloc[0]) else 0.0

            righe_mostra = []
            for _, row in df_art.iterrows():
                qta = float(row['Qta_Effettiva'])
                req_date = row['Data_Consegna']
                
                if st_gia >= qta:
                    st_gia -= qta
                    eta, nota, cat = oggi_dt, "Pronto", "DISP"
                elif (st_gia + st_lav) >= qta:
                    st_gia = 0
                    eta, nota, cat = aggiungi_giorni_lavorativi(oggi_dt, 10), "In Lavorazione", "LAV"
                else:
                    eta, nota, cat = aggiungi_giorni_lavorativi(oggi_dt, 25), "Nuova Produzione", "PROD"

                rit = (pd.notnull(req_date) and eta.date() > req_date.date())
                css = "on-time-row" if cat == "DISP" and not rit else ("client-delay-row" if cat == "DISP" else "delay-row")

                ok = (filtro_label == "Mostra tutto") or (filtro_label == "Solo Disponibili" and cat == "DISP") or (filtro_label == "In Lavorazione" and cat == "LAV") or (filtro_label == "In Ritardo" and rit)
                if ok: righe_mostra.append({'css': css, 'date': req_date, 'qta': qta, 'eta': eta, 'nota': nota})

            if righe_mostra:
                with st.expander(f"📦 {art} — {desc} | Residuo: {df_art['Qta_Effettiva'].sum():,.0f}"):
                    for r in righe_mostra:
                        st.markdown(f'<div class="status-row {r["css"]}"><span><b>Consegna:</b> {r["date"].strftime("%d/%m/%Y") if pd.notnull(r["date"]) else "N/D"} | <b>Q.tà:</b> {r["qta"]:,.0f}</span><span><b>Stima:</b> {r["eta"].strftime("%d/%m/%Y")} ({r["nota"]})</span></div>', unsafe_allow_html=True)
else:
    st.warning("Nessun dato caricato. Controlla i file Excel.")
