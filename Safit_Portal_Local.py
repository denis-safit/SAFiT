import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE PAGINA ---
APP_VERSION = "1.1.2"
st.set_page_config(page_title=f"Safit Portal - {APP_VERSION}", layout="wide")

st.markdown(f"""
    <style>
    .status-row {{ display: flex; justify-content: space-between; padding: 12px; border-radius: 8px; margin-bottom: 8px; font-size: 14px; }}
    .on-time-row {{ background-color: #f1f8e9; border-left: 6px solid #4caf50; color: #1b5e20; }}
    .acq-row {{ background-color: #e3f2fd; border-left: 6px solid #2196f3; color: #0d47a1; }}
    .prod-row {{ background-color: #fff8e1; border-left: 6px solid #ffc107; color: #5d4037; }}
    .urgent-row {{ background-color: #ffebee; border-left: 6px solid #f44336; color: #b71c1c; }}
    </style>
    """, unsafe_allow_html=True)

# --- 2. LOGIN ---
if "authenticated" not in st.session_state: st.session_state["authenticated"] = False
if not st.session_state["authenticated"]:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=300)
        st.title("Area Riservata Safit")
        user = st.text_input("Username")
        pw = st.text_input("Password", type="password")
        if st.button("Accedi"):
            if user == "safit_admin" and pw == "admin2026": # Sostituisci con i tuoi dati
                st.session_state["authenticated"] = True
                st.rerun()
            else: st.error("Credenziali errate")
    st.stop()

# --- 3. CARICAMENTO DATI ---
def pulisci_num(serie):
    return pd.to_numeric(serie.astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').fillna(0)

@st.cache_data
def load_data():
    try:
        # ARCA
        df_a = pd.read_excel('righe_Ordini_ARCA.xlsx', sheet_name='Foglio1', skiprows=2)
        df_a.columns = [str(c).strip() for c in df_a.columns]
        for c in ['Cliente Fornitore CD', 'Articolo C', 'Articolo D', 'Data']:
            if c in df_a.columns: df_a[c] = df_a[c].ffill()
        df_a = df_a.dropna(subset=['Articolo C'])
        df_a['Art_Key'] = df_a['Articolo C'].astype(str).str.strip().upper()
        df_a['Data_C'] = pd.to_datetime(df_a['Data'], errors='coerce')
        q_col = 'Qta Residua' if 'Qta Residua' in df_a.columns else 'Qta Doc'
        df_a['Qta'] = pulisci_num(df_a[q_col])
        df_a = df_a[df_a['Qta'] > 0]

        # ACCESS
        file_acc = 'Avanzamento_access.xlsx'
        if os.path.exists(file_acc):
            df_t = pd.read_excel(file_acc, skiprows=1)
            df_t.columns = [str(c).strip() for c in df_t.columns]
            df_t['Key'] = df_t['Codice'].astype(str).str.strip().upper()
            
            # Mappatura Colonne Denis
            df_t['Mag'] = pulisci_num(df_t['Gia'])
            df_t['Acquisti'] = pulisci_num(df_t['Acq'])
            df_t['Produz'] = df_t[['Lan', 'GRZ', 'TMP', 'RWI', 'TRS']].apply(pulisci_num).sum(axis=1)

            df_merged = pd.merge(df_a, df_t[['Key', 'Mag', 'Acquisti', 'Produz']], 
                                 left_on='Art_Key', right_on='Key', how='left')
            return df_merged.fillna(0)
        return df_a
    except Exception as e:
        st.error(f"Errore caricamento: {e}")
        return pd.DataFrame()

# --- 4. DASHBOARD ---
data = load_data()
oggi = datetime.now()

with st.sidebar:
    if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
    if not data.empty:
        clilist = sorted(data['Cliente Fornitore CD'].unique().astype(str))
        sel_cli = st.selectbox("Cliente", clilist)
    if st.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

st.title(f"Piano Avanzamento: {sel_cli}")

if not data.empty:
    df_cli = data[data['Cliente Fornitore CD'] == sel_cli]
    for art in sorted(df_cli['Articolo C'].unique()):
        df_art = df_cli[df_cli['Articolo C'] == art].sort_values('Data_C')
        m, a, p = float(df_art['Mag'].iloc[0]), float(df_art['Acquisti'].iloc[0]), float(df_art['Produz'].iloc[0])
        
        with st.expander(f"📦 {art} (Giacenza: {m:,.0f})"):
            for _, row in df_art.iterrows():
                q = float(row['Qta'])
                d = row['Data_C']
                
                if m >= q:
                    m -= q; s, n, c = d, "PRONTO (GIA)", "on-time-row"
                elif (m + a) >= q:
                    a -= (q - m); m = 0; s, n, c = oggi + timedelta(days=10), "ARRIVO (ACQ)", "acq-row"
                elif (m + a + p) >= q:
                    p -= (q - m - a); m = 0; a = 0; s, n, c = oggi + timedelta(days=20), "PRODUZIONE", "prod-row"
                else:
                    s, n, c = oggi + timedelta(days=35), "DA LANCIARE", "urgent-row"

                st.markdown(f'<div class="status-row {c}"><span><b>{d.strftime("%d/%m/%Y")}</b> | Qta: {q:,.0f}</span><span>{n} (Est: {s.strftime("%d/%m/%Y")})</span></div>', unsafe_allow_html=True)
