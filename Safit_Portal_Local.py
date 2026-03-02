import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO
import plotly.express as px

# --- 1. CONFIGURAZIONE ---
APP_VERSION = "1.6.9-Turbo"
st.set_page_config(page_title=f"Safit Portal v{APP_VERSION}", layout="wide")

st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 10px; border-radius: 8px; margin-bottom: 6px; border: 1px solid #ddd; color: #000 !important; font-size: 14px; }
    .on-time-row { background-color: #e8f5e9 !important; border-left: 6px solid #4caf50; } 
    .acq-row { background-color: #e3f2fd !important; border-left: 6px solid #2196f3; }    
    .prod-row { background-color: #fffde7 !important; border-left: 8px solid #fbc02d; }   
    .urgent-row { background-color: #ffebee !important; border-left: 8px solid #f44336; } 
    .oca-row { background-color: #f5f5f5 !important; border-left: 8px solid #9e9e9e; color: #666 !important; }
    .debug-box { background-color: #f0f2f6 !important; color: #111 !important; padding: 8px 12px; border-radius: 6px; border: 1px solid #ccc; margin-bottom: 10px; font-family: sans-serif; font-size: 13px; font-weight: 600; display: flex; justify-content: space-between; }
    .kpi-card { background-color: #ffffff; border: 1px solid #e0e0e0; padding: 15px; border-radius: 10px; text-align: center; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .kpi-val { font-size: 24px; font-weight: bold; color: #1f77b4; }
    .user-info { padding: 10px; background: #f8f9fa; border-radius: 5px; border: 1px solid #eee; margin-bottom: 20px; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNZIONI VELOCI ---
def fmt_n(val):
    try: return f"{int(round(float(val), 0)):,}".replace(",", ".")
    except: return "0"

def clean_num(serie):
    return pd.to_numeric(serie.astype(str).str.replace(r'[^\d,.-]', '', regex=True).str.replace(',', '.'), errors='coerce').fillna(0)

def fast_load(filename, key_col):
    if not os.path.exists(filename): return pd.DataFrame()
    # Legge tutto il file per sicurezza
    df = pd.read_excel(filename)
    # Cerca la riga dove iniziano davvero i dati (salta loghi o intestazioni vuote)
    for i in range(min(len(df), 20)):
        if key_col in df.iloc[i].astype(str).values:
            df = pd.read_excel(filename, skiprows=i+1)
            break
    df.columns = [str(c).strip() for c in df.columns]
    # Riempie i valori mancanti dovuti ai raggruppamenti (importante!)
    df = df.ffill() 
    return df.dropna(subset=[key_col])

# --- 3. MOTORE ATP ---
@st.cache_data(ttl=600) # Cache di 10 minuti per velocità
def load_and_process():
    try:
        df_full = fast_load('righe_Ordini_ARCA.xlsx', "Articolo C")
        if df_full.empty: return pd.DataFrame(), {}

        c_tipo, c_art, c_des, c_qta, c_dat, c_cli = "Codice Documento", "Articolo C", "Articolo D", "Qta Residua", "Data Consegna", "Cliente Fornitore CD"
        df_full[c_qta] = clean_num(df_full[c_qta])
        
        # Access
        df_acc = fast_load('Avanzamento_access.xlsx', "CODICE")
        stock = {}
        if not df_acc.empty:
            for _, r in df_acc.iterrows():
                p = sum([clean_num(pd.Series([r.get(f, 0)])).iloc[0] for f in ['LANCIATI', 'GRZ', 'TMP', 'RWI', 'TRS']])
                stock[str(r['CODICE']).strip().upper()] = {'GIA': clean_num(pd.Series([r.get('GIA', 0)])).iloc[0], 'PROD': p}

        # Elaborazione code
        arrivals = {}
        df_in = df_full[df_full[c_tipo].isin(['OFF', 'DCL'])].copy()
        for art, g in df_in.groupby(c_art):
            arrivals[art] = g.sort_values(c_dat)[[c_dat, c_qta]].values.tolist()

        final = []
        # Ordiniamo per data per processare correttamente lo stock
        df_orders = df_full[df_full[c_tipo].isin(['OCI', 'OCA'])].sort_values([c_dat, c_tipo], ascending=[True, False])

        for _, row in df_orders.iterrows():
            art = str(row[c_art]).strip().upper()
            qta = row[c_qta]
            s = stock.get(art, {'GIA':0, 'PROD':0})
            arr = arrivals.get(art, [])
            
            is_oca = (row[c_tipo] == 'OCA')
            st_val, color, dt_e = ("MANCANTE", "urgent-row", datetime.now() + timedelta(days=45)) if not is_oca else ("DA PIANIFICARE", "oca-row", row[c_dat])

            if s['GIA'] >= qta:
                s['GIA'] -= qta; st_val, color, dt_e = "DISPONIBILE", "on-time-row", row[c_dat]
            else:
                fabb = qta - s['GIA']; s['GIA'] = 0
                for a in arr:
                    if a[1] >= fabb:
                        a[1] -= fabb; st_val, color, dt_e = "ACQUISTO", "acq-row", a[0]; fabb = 0; break
                    else: fabb -= a[1]; a[1] = 0
                if fabb > 0 and s['PROD'] >= fabb:
                    s['PROD'] -= fabb; st_val, color, dt_e = "PRODUZIONE", "prod-row", datetime.now() + timedelta(days=21); fabb = 0

            final.append({**row.to_dict(), 'ST': st_val, 'CS': color, 'DT_E': dt_e, 'ART_KEY': art, 'CLI_NAME': str(row[c_cli])})
        
        return pd.DataFrame(final), stock
    except Exception as e:
        st.error(f"Errore: {e}"); return pd.DataFrame(), {}

# --- 4. INTERFACCIA ---
if "authenticated" not in st.session_state: st.session_state["authenticated"] = False
if not st.session_state["authenticated"]:
    # ... (Login semplificato per brevità, usa safit_admin / admin2026)
    u = st.text_input("User"); p = st.text_input("Pwd", type="password")
    if st.button("Login"):
        if u == "safit_admin" and p == "admin2026": 
            st.session_state.update({"authenticated": True, "username": u, "user_type": "TUTTI"})
            st.rerun()
    st.stop()

df_res, stock_final = load_and_process()

if not df_res.empty:
    with st.sidebar:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
        st.markdown(f"👤 **{st.session_state['username']}**")
        if st.button("Esci"): st.session_state.authenticated = False; st.rerun()
        sel_cli = st.selectbox("Cliente:", ["TUTTI"] + sorted(df_res['CLI_NAME'].unique().tolist()))
        search = st.text_input("Cerca Articolo:").upper()

    df_f = df_res if sel_cli == "TUTTI" else df_res[df_res['CLI_NAME'] == sel_cli]
    
    # KPI
    st.title(f"Safit Portal: {sel_cli}")
    c1, c2, c3 = st.columns(3)
    c1.metric("Pezzi Totali", fmt_n(df_f['Qta Residua'].sum()))
    c2.metric("Articoli Unici", len(df_f['ART_KEY'].unique()))
    c3.metric("Clienti", len(df_f['CLI_NAME'].unique()))

    # LISTA
    df_disp = df_f[df_f['ART_KEY'].str.contains(search)] if search else df_f
    for art, g in df_disp.groupby('ART_KEY'):
        with st.expander(f"📦 {art} - {g['Articolo D'].iloc[0]} ({len(g)} ordini)"):
            for _, r in g.iterrows():
                st.markdown(f'<div class="status-row {r["CS"]}"><span>{r["Codice Documento"]} | {r["CLI_NAME"]} | Qta: {fmt_n(r["Qta Residua"])}</span><span><b>{r["ST"]}</b></span></div>', unsafe_allow_html=True)
