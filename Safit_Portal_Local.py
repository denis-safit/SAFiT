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
    .on-time-row {{ background-color: #f1f8e9; border-left: 6px solid #4caf50; color: #1b5e20; }} 
    .client-delay-row {{ background-color: #e3f2fd; border-left: 6px solid #2196f3; color: #0d47a1; }} 
    .delay-row {{ background-color: #fff8e1; border-left: 6px solid #ffc107; color: #5d4037; }} 
    .prod-delay-row {{ background-color: #ffebee; border-left: 6px solid #f44336; color: #b71c1c; }} 
    .stExpander div {{ height: auto !important; min-height: min-content !important; }}
    .version-tag {{ font-size: 10px; color: #999; text-align: right; }}
    </style>
    """, unsafe_allow_html=True)

# --- 2. GESTIONE UTENTI (Versione Protetta) ---
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
# st.write("Utenti caricati:", list(USER_DB.keys())) # Rimuovi # per debug

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if not st.session_state["authenticated"]:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=300)
            st.title("Accesso Area Riservata")
            u_in = st.text_input("Username").strip()
            p_in = st.text_input("Password", type="password").strip()
            if st.button("Accedi"):
                for db_u, info in USER_DB.items():
                    if str(db_u).strip() == u_in and str(info[0]).strip() == p_in:
                        st.session_state["authenticated"] = True
                        st.session_state["user_type"] = info[1]
                        st.session_state["username"] = u_in
                        st.rerun()
                st.error("Username o Password errati")
        return False
    return True

if not check_password(): st.stop()

# --- 3. FUNZIONI TECNICHE ---
def aggiungi_giorni_lavorativi(start_date, days):
    curr = start_date
    while days > 0:
        curr += timedelta(days=1)
        if curr.weekday() < 5: days -= 1
    return curr

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
        c_res = 'Qta Residua' if 'Qta Residua' in df.columns else 'Qta Doc'
        df['Qta_Effettiva'] = pd.to_numeric(df[c_res], errors='coerce').fillna(0)
        df = df[df['Qta_Effettiva'] > 0]
        
        if os.path.exists('Avanzamento_access.xlsx'):
            df_t = pd.read_excel('Avanzamento_access.xlsx', skiprows=1) 
            df_t.columns = [str(c).strip() for c in df_t.columns]
            if 'Codice' in df_t.columns:
                df_t = df_t.rename(columns={'Codice': 'Art_Key'})
                for c in ['Gia', 'Acq', 'Lan', 'Grz', 'Tmp', 'Rwi', 'Trs']:
                    if c in df_t.columns: df_t[c] = pulisci_numero(df_t[c])
                df_t['Lav_Tot'] = df_t.get('Acq', 0) + df_t.get('Lan', 0) + df_t.get('Grz', 0) + \
                                 df_t.get('Tmp', 0) + df_t.get('Rwi', 0) + df_t.get('Trs', 0)
                df = pd.merge(df, df_t[['Art_Key', 'Gia', 'Lav_Tot']], left_on='Articolo C', right_on='Art_Key', how='left')
        return df
    except: return pd.DataFrame()

data = load_data()
oggi = datetime.now()

# --- 4. SIDEBAR ---
with st.sidebar:
    if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
    if st.session_state.get("username"):
        st.write(f"Utente: **{st.session_state['username']}**")
    st.markdown(f'<p class="version-tag">Versione: {APP_VERSION}</p>', unsafe_allow_html=True)
    
    if st.session_state["user_type"] == "TUTTI":
        c_list = sorted([str(x) for x in data['Cliente Fornitore CD'].unique()])
        sel_cli = st.selectbox("👤 Seleziona Cliente:", c_list)
    else: sel_cli = st.session_state["user_type"]
    
    f_label = st.radio("Filtra per stato:", ["Mostra tutto", "Solo Disponibili", "In Lavorazione", "In Ritardo"])
    if st.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

# --- 5. LOGICA CENTRALE ---
st.title("Portale Avanzamento Produzione")
df_cli = data[data['Cliente Fornitore CD'] == sel_cli].copy()

if not df_cli.empty:
    articoli = sorted([str(x) for x in df_cli['Articolo C'].unique()])
    sel_art = st.selectbox("🔍 Cerca Codice Prodotto:", ["Tutti i prodotti"] + articoli)
    a_view = articoli if sel_art == "Tutti i prodotti" else [sel_art]

    for art in a_view:
        df_art = df_cli[df_cli['Articolo C'] == art].sort_values('Data_Consegna')
        desc = df_art['Articolo D'].iloc[0] if 'Articolo D' in df_art.columns else ""
        gia = float(df_art['Gia'].iloc[0]) if 'Gia' in df_art.columns and pd.notnull(df_art['Gia'].iloc[0]) else 0.0
        lav = float(df_art['Lav_Tot'].iloc[0]) if 'Lav_Tot' in df_art.columns and pd.notnull(df_art['Lav_Tot'].iloc[0]) else 0.0

        mostra = []
        for _, row in df_art.iterrows():
            q = float(row['Qta_Effettiva'])
            d = row['Data_Consegna']
            
            if gia >= q:
                gia -= q
                eta, nota, cat = oggi, "Pronto", "DISP"
            elif (gia + lav) >= q:
                gia = 0
                eta, nota, cat = aggiungi_giorni_lavorativi(oggi, 10), "In Lavorazione", "LAV"
            else:
                eta, nota, cat = aggiungi_giorni_lavorativi(oggi, 25), "Nuova Produzione", "PROD"

            rit = (pd.notnull(d) and eta.date() > d.date())
            
            if cat == "DISP":
                css, n_disp = ("client-delay-row", "Pronto (Ritardo Ritiro)") if rit else ("on-time-row", "Pronto")
            else:
                css, n_disp = ("prod-delay-row", f"{nota} (In Ritardo)") if rit else ("delay-row", nota)

            # Filtro Corretto
            ok = (f_label == "Mostra tutto") or \
                 (f_label == "Solo Disponibili" and cat == "DISP") or \
                 (f_label == "In Lavorazione" and (cat == "LAV" or cat == "PROD") and not rit) or \
                 (f_label == "In Ritardo" and rit)

            if ok:
                mostra.append({'css': css, 'd': d, 'q': q, 'eta': eta, 'n': n_disp})

        if mostra:
            with st.expander(f"📦 {art} — {desc} | Residuo: {df_art['Qta_Effettiva'].sum():,.0f}"):
                for r in mostra:
                    st.markdown(f'<div class="status-row {r["css"]}"><span><b>Consegna:</b> {r["d"].strftime("%d/%m/%Y") if pd.notnull(r["d"]) else "N.D."} | <b>Q.tà:</b> {r["q"]:,.0f}</span><span><b>Stima:</b> {r["eta"].strftime("%d/%m/%Y")} ({r["n"]})</span></div>', unsafe_allow_html=True)
