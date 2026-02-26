import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE PAGINA E VERSIONE ---
APP_VERSION = "1.0.02"
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
    /* COLORI LOGICA SAFIT */
    .on-time-row {{ background-color: #f1f8e9; border-left: 6px solid #4caf50; color: #1b5e20; }} /* VERDE: Pronto in tempo */
    .client-delay-row {{ background-color: #e3f2fd; border-left: 6px solid #2196f3; color: #0d47a1; }} /* AZZURRO: Disponibile ma non ritirato */
    .delay-row {{ background-color: #fff8e1; border-left: 6px solid #ffc107; color: #5d4037; }} /* GIALLO: In Lavorazione/Nuova Prod */
    .prod-delay-row {{ background-color: #ffebee; border-left: 6px solid #f44336; color: #b71c1c; }} /* ROSSO: Ritardo Produzione Grave */
    
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
            user = st.text_input("Username")
            pw = st.text_input("Password", type="password")
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

@st.cache_data
def load_data():
    try:
        df = pd.read_excel('righe_Ordini_ARCA.xlsx', sheet_name='Foglio1', skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        for col in ['Cliente Fornitore CD', 'Articolo C', 'Articolo D', 'Data']:
            if col in df.columns: df[col] = df[col].ffill()
        df['Data_Consegna'] = pd.to_datetime(df['Data'], errors='coerce')
        col_res = 'Qta Residua' if 'Qta Residua' in df.columns else 'Qta Doc'
        df['Qta_Effettiva'] = pd.to_numeric(df[col_res], errors='coerce').fillna(0)
        df = df[df['Qta_Effettiva'] > 0]
        
        if os.path.exists('Avanzamento_access.xlsx'):
            df_tech = pd.read_excel('Avanzamento_access.xlsx', skiprows=1) 
            df_tech.columns = [str(c).strip() for c in df_tech.columns]
            if 'Codice' in df_tech.columns:
                df_tech = df_tech.rename(columns={'Codice': 'Art_Key'})
                campi_num = ['Gia', 'Acq', 'Lan', 'Grz', 'Tmp', 'Rwi', 'Trs']
                for c in campi_num:
                    if c in df_tech.columns: df_tech[c] = pulisci_numero(df_tech[c])
                
                df_tech['Lavorazione_Totale'] = df_tech.get('Acq', 0) + df_tech.get('Lan', 0) + \
                                               df_tech.get('Grz', 0) + df_tech.get('Tmp', 0) + \
                                               df_tech.get('Rwi', 0) + df_tech.get('Trs', 0)
                
                df = pd.merge(df, df_tech[['Art_Key', 'Gia', 'Lavorazione_Totale']], 
                              left_on='Articolo C', right_on='Art_Key', how='left')
        return df
    except: return pd.DataFrame()

data = load_data()
oggi_dt = datetime.now()

# --- 4. SIDEBAR ---
with st.sidebar:
    if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
    st.write(f"Utente: **{st.session_state['username']}**")
    st.markdown(f'<p class="version-tag">Versione: {APP_VERSION}</p>', unsafe_allow_html=True)
    
    if st.session_state["user_type"] == "TUTTI":
        clienti_list = sorted([str(x) for x in data['Cliente Fornitore CD'].unique()])
        sel_cli = st.selectbox("👤 Seleziona Cliente:", clienti_list)
    else: sel_cli = st.session_state["user_type"]
    
    filtro_label = st.radio("Filtra per stato:", ["Mostra tutto", "Solo Disponibili", "In Lavorazione", "In Ritardo"])
    if st.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

# --- 5. LOGICA CENTRALE ---
st.title("🚜 Portale Avanzamento Produzione")
df_cli = data[data['Cliente Fornitore CD'] == sel_cli].copy()

if not df_cli.empty:
    articoli_dis = sorted([str(x) for x in df_cli['Articolo C'].unique()])
    sel_art = st.selectbox("🔍 Cerca Codice Prodotto:", ["Tutti i prodotti"] + articoli_dis)
    articoli_view = articoli_dis if sel_art == "Tutti i prodotti" else [sel_art]

    for art in articoli_view:
        df_art = df_cli[df_cli['Articolo C'] == art].sort_values('Data_Consegna')
        desc = df_art['Articolo D'].iloc[0] if 'Articolo D' in df_art.columns else ""
        
        st_gia = float(df_art['Gia'].iloc[0]) if 'Gia' in df_art.columns and pd.notnull(df_art['Gia'].iloc[0]) else 0.0
        st_lav = float(df_art['Lavorazione_Totale'].iloc[0]) if 'Lavorazione_Totale' in df_art.columns and pd.notnull(df_art['Lavorazione_Totale'].iloc[0]) else 0.0

        righe_mostra = []
        for _, row in df_art.iterrows():
            qta = float(row['Qta_Effettiva'])
            req_date = row['Data_Consegna']
            
            # --- DETERMINAZIONE STATO ---
            if st_gia >= qta:
                st_gia -= qta
                eta, nota, cat_core = oggi_dt, "Pronto", "DISPONIBILE"
            elif (st_gia + st_lav) >= qta:
                st_gia = 0
                eta, nota, cat_core = aggiungi_giorni_lavorativi(oggi_dt, 10), "In Lavorazione", "LAVORAZIONE"
            else:
                eta, nota, cat_core = aggiungi_giorni_lavorativi(oggi_dt, 25), "Nuova Produzione", "LAVORAZIONE"

            is_ritardo = (pd.notnull(req_date) and eta.date() > req_date.date())
            
            # --- ASSEGNAZIONE COLORI E TESTI SPECIFICI ---
            if cat_core == "DISPONIBILE":
                if is_ritardo:
                    css, nota_display = "client-delay-row", "Pronto (Ritardo Ritiro)"
                else:
                    css, nota_display = "on-time-row", "Pronto"
            else: # LAVORAZIONE o NUOVA PROD
                if is_ritardo:
                    css, nota_display = "prod-delay-row", f"{nota} (In Ritardo)"
                else:
                    css, nota_display = "delay-row", nota

            # --- FILTRO ---
            passa = False
            if filtro_label == "Mostra tutto": passa = True
            elif filtro_label == "Solo Disponibili" and cat_core == "DISPONIBILE": passa = True
            elif filtro_label == "In Lavorazione" and cat_core == "LAVORAZIONE": passa = True
            elif filtro_label == "In Ritardo" and is_ritardo: passa = True

            if passa:
                righe_mostra.append({'css': css, 'date': req_date, 'qta': qta, 'eta': eta, 'nota': nota_display})

        if righe_mostra:
            with st.expander(f"📦 {art} — {desc} | Residuo: {df_art['Qta_Effettiva'].sum():,.0f}"):
                for r in righe_mostra:
                    st.markdown(f'<div class="status-row {r["css"]}"><span><b>Consegna:</b> {r["date"].strftime("%d/%m/%Y") if pd.notnull(r["date"]) else "N.D."} | <b>Q.tà:</b> {r["qta"]:,.0f}</span><span><b>Stima:</b> {r["eta"].strftime("%d/%m/%Y")} ({r["nota"]})</span></div>', unsafe_allow_html=True)
