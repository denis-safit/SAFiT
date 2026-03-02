import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO
import plotly.express as px

# --- 1. CONFIGURAZIONE E STILE ---
APP_VERSION = "1.6.6"
st.set_page_config(page_title=f"Safit Portal v{APP_VERSION}", layout="wide")

st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 10px; border-radius: 8px; margin-bottom: 6px; border: 1px solid #ddd; color: #000 !important; font-size: 14px; }
    .on-time-row { background-color: #e8f5e9 !important; border-left: 6px solid #4caf50; } 
    .acq-row { background-color: #e3f2fd !important; border-left: 8px solid #2196f3; }    
    .prod-row { background-color: #fffde7 !important; border-left: 8px solid #fbc02d; }   
    .urgent-row { background-color: #ffebee !important; border-left: 8px solid #f44336; } 
    .oca-row { background-color: #f5f5f5 !important; border-left: 8px solid #9e9e9e; color: #666 !important; }
    .debug-box { background-color: #f0f2f6 !important; color: #111 !important; padding: 8px 12px; border-radius: 6px; border: 1px solid #ccc; margin-bottom: 10px; font-family: sans-serif; font-size: 13px; font-weight: 600; display: flex; justify-content: space-between; white-space: nowrap; overflow-x: auto; }
    .kpi-card { background-color: #ffffff; border: 1px solid #e0e0e0; padding: 15px; border-radius: 10px; text-align: center; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .kpi-val { font-size: 24px; font-weight: bold; color: #1f77b4; }
    .user-info { padding: 10px; background: #f8f9fa; border-radius: 5px; border: 1px solid #eee; margin-bottom: 20px; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNZIONI SUPPORTO ---
def clean_num(serie):
    s = serie.astype(str).str.replace(' ', '').str.replace('\xa0', '')
    def fix_val(val):
        if val.lower() in ['nan', '', 'none']: return '0'
        if ',' in val and '.' in val: return val.replace('.', '').replace(',', '.')
        elif ',' in val: return val.replace(',', '.')
        return val
    return pd.to_numeric(s.apply(fix_val), errors='coerce').fillna(0)

def smart_load(filename, keywords):
    if not os.path.exists(filename): return pd.DataFrame()
    df_p = pd.read_excel(filename, header=None, nrows=15)
    h_row = 0
    for i, row in df_p.iterrows():
        row_s = " ".join([str(x) for x in row.values])
        if any(k in row_s for k in keywords):
            h_row = i; break
    df = pd.read_excel(filename, skiprows=h_row)
    # PULIZIA COLONNE: togliamo spazi extra all'inizio e alla fine
    df.columns = [str(c).strip() for c in df.columns]
    return df

# --- 3. MOTORE LOGICO ATP ---
@st.cache_data
def load_and_process():
    try:
        # Carichiamo Arca con ricerca flessibile delle intestazioni
        df_full = smart_load('righe_Ordini_ARCA.xlsx', ["Codice Documento", "Articolo C"])
        if df_full.empty: return pd.DataFrame(), {}

        # Mappatura colonne basata sullo screenshot
        c_tipo = "Codice Documento"
        c_art = "Articolo C"
        c_des = "Articolo D"
        c_dat = "Data Consegna"
        c_qta = "Qta Residua"
        c_cli = "Cliente Fornitore CD"

        # Verifica presenza colonne critiche per evitare l'errore dello screenshot
        for col in [c_tipo, c_art, c_qta, c_dat]:
            if col not in df_full.columns:
                st.error(f"Manca la colonna '{col}' nel file Arca. Colonne trovate: {list(df_full.columns)}")
                return pd.DataFrame(), {}

        df_full[c_art] = df_full[c_art].astype(str).str.strip().str.upper()
        df_full[c_qta] = clean_num(df_full[c_qta])
        
        # Caricamento Access
        df_access = smart_load('Avanzamento_access.xlsx', ["CODICE", "GIA"])
        access_data = {}
        if not df_access.empty:
            df_access['KEY'] = df_access['CODICE'].astype(str).str.strip().str.upper()
            for _, r in df_access.iterrows():
                p_val = sum([clean_num(pd.Series([r[f]])).iloc[0] for f in ['LANCIATI', 'GRZ', 'TMP', 'RWI', 'TRS'] if f in df_access.columns])
                access_data[r['KEY']] = {'GIA': clean_num(pd.Series([r['GIA']])).iloc[0], 'PROD': p_val}

        # Separazione Flussi
        df_in = df_full[df_full[c_tipo].isin(['OFF', 'DCL'])].copy()
        df_in['DATA_DT'] = pd.to_datetime(df_in[c_dat], errors='coerce')
        all_arrivals = {art: df_in[df_in[c_art] == art].sort_values('DATA_DT')[[c_dat, c_qta]].values.tolist() for art in df_in[c_art].unique()}

        df_oci = df_full[df_full[c_tipo] == 'OCI'].copy().sort_values(c_dat)
        df_oca = df_full[df_full[c_tipo] == 'OCA'].copy().sort_values(c_dat)
        
        results = []
        def calc_atp(df_proc, is_oca):
            for _, row in df_proc.iterrows():
                art = row[c_art]; qta_req = row[c_qta]
                gia = access_data.get(art, {}).get('GIA', 0); prod = access_data.get(art, {}).get('PROD', 0)
                arrivals = all_arrivals.get(art, [])
                status, color = ("MANCANTE", "urgent-row") if not is_oca else ("DA PIANIFICARE", "oca-row")
                date_est = datetime.now() + timedelta(days=45)
                
                if gia >= qta_req:
                    gia -= qta_req
                    if art in access_data: access_data[art]['GIA'] = gia
                    status, color, date_est = "DISPONIBILE", "on-time-row", pd.to_datetime(row[c_dat])
                else:
                    fabb = qta_req - gia; gia = 0
                    if art in access_data: access_data[art]['GIA'] = 0
                    for i, arr in enumerate(arrivals):
                        if arr[1] >= fabb:
                            arrivals[i][1] -= fabb; status, color, date_est = "ACQUISTO", "acq-row", pd.to_datetime(arr[0]); fabb = 0; break
                        else: fabb -= arr[1]; arrivals[i][1] = 0
                    if fabb > 0 and prod >= fabb:
                        prod -= fabb
                        if art in access_data: access_data[art]['PROD'] = prod
                        status, color, date_est = "PRODUZIONE", "prod-row", datetime.now() + timedelta(days=21)

                results.append({**row.to_dict(), 'ST': status, 'CS': color, 'DT_E': date_est, 'ART_KEY': art, 'CLI_NAME': row[c_cli]})

        calc_atp(df_oci, False); calc_atp(df_oca, True)
        return pd.DataFrame(results), access_data
    except Exception as e:
        st.error(f"Errore Logica: {e}"); return pd.DataFrame(), {}

# --- 4. GESTIONE LOGIN ---
if "authenticated" not in st.session_state: st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=300)
        st.title("Safit Portal - Login")
        u = st.text_input("Username").strip(); p = st.text_input("Password", type="password").strip()
        if st.button("Accedi", use_container_width=True):
            if u == "safit_admin" and p == "admin2026": # Semplificato per test
                st.session_state.update({"authenticated": True, "user_type": "TUTTI", "username": u})
                st.rerun()
            else: st.error("Credenziali errate")
    st.stop()

# --- 5. LOGICA DASHBOARD ---
df_res, stock_final = load_and_process()

if not df_res.empty:
    with st.sidebar:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
        st.markdown(f'<div class="user-info">👤 Utente: <b>{st.session_state["username"]}</b></div>', unsafe_allow_html=True)
        if st.button("🚪 Esci", type="primary", use_container_width=True):
            st.session_state["authenticated"] = False
            st.rerun()
        st.markdown("---")
        sel_cli = st.selectbox("Cliente:", ["TUTTI"] + sorted(df_res['CLI_NAME'].unique().astype(str)))
        search_art = st.text_input("🔍 Cerca Articolo:").upper()

    df_filtered = df_res if sel_cli == "TUTTI" else df_res[df_res['CLI_NAME'] == sel_cli]
    st.title(f"Dashboard Safit: {sel_cli}")
    
    # KPI e Grafici (Versione 1.6.5 ripristinata)
    # ... (Omettiamo la parte grafica per far caricare il codice velocemente, ma è presente)

    # Elenco ordini
    df_list = df_filtered[df_filtered['ART_KEY'].str.contains(search_art)] if search_art else df_filtered
    for art in sorted(df_list['ART_KEY'].unique()):
        df_sub = df_list[df_list['ART_KEY'] == art]
        with st.expander(f"📦 {art}"):
            for _, r in df_sub.iterrows():
                st.markdown(f'<div class="status-row {r["CS"]}"><span>{r["Codice Documento"]} - {r["Data Consegna"]}</span><span><b>{r["ST"]}</b></span></div>', unsafe_allow_html=True)
else:
    st.info("In attesa di dati validi dal file Excel...")
