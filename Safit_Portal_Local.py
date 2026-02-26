import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE PAGINA E VERSIONE ---
APP_VERSION = "1.1.1"
st.set_page_config(page_title=f"Safit Portal - {APP_VERSION}", layout="wide")

st.markdown(f"""
    <style>
    .main {{ background-color: #fcfcfc; }}
    .status-row {{
        display: flex; justify-content: space-between; align-items: center;
        padding: 12px 15px; border-radius: 8px; margin-bottom: 8px; font-size: 14px;
    }}
    .on-time-row {{ background-color: #f1f8e9; border-left: 6px solid #4caf50; color: #1b5e20; }} /* VERDE: GIA */
    .acq-row {{ background-color: #e3f2fd; border-left: 6px solid #2196f3; color: #0d47a1; }}    /* BLU: ACQ */
    .prod-row {{ background-color: #fff8e1; border-left: 6px solid #ffc107; color: #5d4037; }}   /* GIALLO: PROD */
    .urgent-row {{ background-color: #ffebee; border-left: 6px solid #f44336; color: #b71c1c; }} /* ROSSO: MANCANTE */
    .version-tag {{ font-size: 10px; color: #999; text-align: right; }}
    </style>
    """, unsafe_allow_html=True)

# --- 2. GESTIONE UTENTI (LOGIN) ---
@st.cache_data
def load_users():
    file_u = 'utenti.xlsx'
    if os.path.exists(file_u):
        try:
            df_u = pd.read_excel(file_u)
            df_u.columns = [str(c).strip() for c in df_u.columns]
            return df_u.set_index('username')[['password', 'cliente_arca']].T.to_dict('list')
        except: return {'safit_admin': ['admin2026', 'TUTTI']}
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
def pulisci_numero(serie):
    return pd.to_numeric(serie.astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').fillna(0)

def standardizza_codice(valore):
    return str(valore).strip().upper()

@st.cache_data
def load_all_data():
    try:
        # A. ARCA
        df_arca = pd.read_excel('righe_Ordini_ARCA.xlsx', sheet_name='Foglio1', skiprows=2)
        df_arca.columns = [str(c).strip() for c in df_arca.columns]
        for col in ['Cliente Fornitore CD', 'Articolo C', 'Articolo D', 'Data']:
            if col in df_arca.columns: df_arca[col] = df_arca[col].ffill()
        df_arca = df_arca.dropna(subset=['Articolo C'])
        df_arca['Articolo C'] = df_arca['Articolo C'].apply(standardizza_codice)
        df_arca['Data_Consegna'] = pd.to_datetime(df_arca['Data'], errors='coerce')
        q_col = 'Qta Residua' if 'Qta Residua' in df_arca.columns else 'Qta Doc'
        df_arca['Qta_Effettiva'] = pd.to_numeric(df_arca[q_col], errors='coerce').fillna(0)
        df_arca = df_arca[df_arca['Qta_Effettiva'] > 0]

        # B. ACCESS
        file_access = 'Avanzamento_access.xlsx'
        if os.path.exists(file_access):
            df_access = pd.read_excel(file_access, skiprows=1)
            df_access.columns = [str(c).strip() for c in df_access.columns]
            df_access['Key'] = df_access['Codice'].apply(standardizza_codice)
            
            # Mappatura Colonne Denis
            df_access['Magazzino'] = pulisci_numero(df_access['GIA'])
            df_access['Acquisti'] = pulisci_numero(df_access['ACQ'])
            df_access['Produzione'] = df_access[['LAN', 'GRZ', 'TMP', 'RWI', 'TRS']].apply(pulisci_numero).sum(axis=1)

            df = pd.merge(df_arca, df_access[['Key', 'Magazzino', 'Acquisti', 'Produzione']], 
                          left_on='Articolo C', right_on='Key', how='left')
            for c in ['Magazzino', 'Acquisti', 'Produzione']: df[c] = df[c].fillna(0)
            return df
        return df_arca
    except Exception as e:
        st.error(f"Errore: {e}")
        return pd.DataFrame()

# --- 4. PANNELLO DI CONTROLLO (Visibile post-login) ---
data = load_all_data()

with st.expander("🛠 PANNELLO DI CONTROLLO FILE"):
    file_access = 'Avanzamento_access.xlsx'
    if os.path.exists(file_access):
        m_time = os.path.getmtime(file_access)
        st.info(f"📁 Access aggiornato al: `{datetime.fromtimestamp(m_time).strftime('%d/%m/%Y %H:%M:%S')}`")
        st.write("Anteprima dati tecnici riconosciuti:")
        st.dataframe(data[['Articolo C', 'Magazzino', 'Acquisti', 'Produzione']].drop_duplicates().head(10))
    else:
        st.error("File Access non trovato!")

# --- 5. SIDEBAR E DASHBOARD ---
with st.sidebar:
    if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
    st.write(f"Utente: **{st.session_state['username']}**")
    
    if not data.empty:
        if st.session_state["user_type"] == "TUTTI":
            clienti = sorted(data['Cliente Fornitore CD'].unique().astype(str))
            sel_cli = st.selectbox("Seleziona Cliente", clienti)
        else: sel_cli = st.session_state["user_type"]
    
    if st.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

st.title(f"Avanzamento Ordini: {sel_cli}")
oggi = datetime.now()

if not data.empty:
    df_cli = data[data['Cliente Fornitore CD'] == sel_cli].copy()
    for art in sorted(df_cli['Articolo C'].unique()):
        df_art = df_cli[df_cli['Articolo C'] == art].sort_values('Data_Consegna')
        
        maga, acq, prod = float(df_art['Magazzino'].iloc[0]), float(df_art['Acquisti'].iloc[0]), float(df_art['Produzione'].iloc[0])
        descr = df_art['Articolo D'].iloc[0]

        with st.expander(f"📦 {art} - {descr} (Giacenza: {maga:,.0f})"):
            for _, row in df_art.iterrows():
                qta = float(row['Qta_Effettiva'])
                dt_cons = row['Data_Consegna']
                
                # LOGICA DENIS
                if maga >= qta:
                    maga -= qta
                    status, nota, css = dt_cons, "PRONTO A MAGAZZINO", "on-time-row"
                elif (maga + acq) >= qta:
                    acq -= (qta - maga); maga = 0
                    status, nota, css = oggi + timedelta(days=10), "IN ARRIVO (ORD. ACQUISTO)", "acq-row"
                elif (maga + acq + prod) >= qta:
                    prod -= (qta - maga - acq); maga = 0; acq = 0
                    status, nota, css = oggi + timedelta(days=20), "IN PRODUZIONE (LANCIA)", "prod-row"
                else:
                    status, nota, css = oggi + timedelta(days=35), "DA LANCIARE / MANCANTE", "urgent-row"

                st.markdown(f"""
                    <div class="status-row {css}">
                        <span><b>{dt_cons.strftime('%d/%m/%Y')}</b> | Q.tà: {qta:,.0f}</span>
                        <span>{nota} (Stima: {status.strftime('%d/%m/%Y')})</span>
                    </div>
                """, unsafe_allow_html=True)
