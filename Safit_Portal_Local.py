import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO
import plotly.express as px

# --- 1. CONFIGURAZIONE E STILE ---
APP_VERSION = "3.1.0-V1.4-Logic-Restore"
st.set_page_config(page_title=f"Safit Portal v{APP_VERSION}", layout="wide")

st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 10px; border-radius: 8px; margin-bottom: 6px; border: 1px solid #ddd; color: #000 !important; font-size: 14px; }
    .on-time-row { background-color: #e8f5e9 !important; border-left: 6px solid #4caf50; } 
    .acq-row { background-color: #e3f2fd !important; border-left: 8px solid #2196f3; }    
    .prod-row { background-color: #fffde7 !important; border-left: 8px solid #fbc02d; }   
    .urgent-row { background-color: #ffebee !important; border-left: 8px solid #f44336; } 
    .oca-row { background-color: #f5f5f5 !important; border-left: 8px solid #9e9e9e; color: #666 !important; }
    .debug-box { background-color: #f8f9fa !important; color: #333 !important; padding: 12px; border-radius: 8px; border: 1px dotted #bbb; margin-bottom: 10px; display: flex; justify-content: space-around; font-size: 14px; }
    .kpi-card { background-color: #ffffff; border: 1px solid #e0e0e0; padding: 15px; border-radius: 10px; text-align: center; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .kpi-val { font-size: 24px; font-weight: bold; color: #1f77b4; }
    .user-info { padding: 10px; background: #f8f9fa; border-radius: 5px; border: 1px solid #eee; margin-bottom: 20px; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. DATABASE UTENTI ---
USER_DB = {
    "safit_admin": ["admin2026", "TUTTI"],
    "denis": ["denis2026", "TUTTI"],
    "cliente_test": ["safit2026", "CLIENTE_ARCA_PROVA"]
}

# --- 3. FUNZIONI TECNICHE ---
def clean_num(serie):
    s = serie.astype(str).str.replace(' ', '').str.replace('\xa0', '')
    def fix_val(val):
        if val.lower() in ['nan', '', 'none']: return '0'
        if ',' in val and '.' in val: return val.replace('.', '').replace(',', '.')
        elif ',' in val: return val.replace(',', '.')
        return val
    return pd.to_numeric(s.apply(fix_val), errors='coerce').fillna(0)

def smart_load(filename, key_col):
    if not os.path.exists(filename): return pd.DataFrame()
    df_p = pd.read_excel(filename, header=None, nrows=20)
    h_row = 0
    for i, row in df_p.iterrows():
        if key_col in row.astype(str).values:
            h_row = i; break
    df = pd.read_excel(filename, skiprows=h_row)
    df.columns = [str(c).strip() for c in df.columns]
    if "CODICE" not in key_col: df = df.ffill() 
    return df

# --- 4. MOTORE DI CALCOLO (LOGICA V1.4) ---
@st.cache_data(ttl=300)
def load_and_process():
    try:
        # 1. Carico Ordini Arca
        df_arca = smart_load('righe_Ordini_ARCA.xlsx', "Articolo C")
        if df_arca.empty: return pd.DataFrame(), {}
        
        c_tipo, c_art, c_qta, c_dat, c_cli = "Codice Documento", "Articolo C", "Qta Residua", "Data Consegna", "Cliente Fornitore CD"
        df_arca[c_dat] = pd.to_datetime(df_arca[c_dat], errors='coerce')
        df_arca[c_qta] = clean_num(df_arca[c_qta])
        df_arca = df_arca.dropna(subset=[c_dat, c_art])

        # 2. Carico Scorte Access
        df_acc = smart_load('Avanzamento_access.xlsx', "CODICE")
        stock_map = {}
        if not df_acc.empty:
            df_acc.columns = [str(c).strip().upper() for c in df_acc.columns]
            for _, r in df_acc.iterrows():
                art_code = str(r['CODICE']).strip().upper()
                # Calcolo GIA, ACQ e PROD come da v1.4
                gia = clean_num(pd.Series([r.get('GIA', 0)])).iloc[0]
                acq = clean_num(pd.Series([r.get('INACQ', 0)])).iloc[0]
                prod = 0
                for f in ['LANCIATI', 'GRZ', 'TMP', 'RWI', 'TRS', 'ACQ', 'TRF']: # Unione campi prod v1.4 + Denis
                    prod += clean_num(pd.Series([r.get(f, 0)])).iloc[0]
                
                stock_map[art_code] = {'GIA': gia, 'ACQ': acq, 'PROD': prod}

        # 3. Calcolo Copertura Sequenziale (Logica v1.4)
        df_orders = df_arca[df_arca[c_tipo].isin(['OCI', 'OCA'])].sort_values([c_art, c_dat])
        final_results = []
        oggi = datetime.now()

        for art, group in df_orders.groupby(c_art):
            s = stock_map.get(str(art).upper(), {'GIA': 0, 'ACQ': 0, 'PROD': 0})
            m, a, p = float(s['GIA']), float(s['ACQ']), float(s['PROD'])
            
            for _, row in group.iterrows():
                q = float(row[c_qta])
                # Logica a cascata v1.4
                if m >= q:
                    m -= q; s_v, c_v, d_v = "DISPONIBILE", "on-time-row", row[c_dat]
                elif (m + a) >= q:
                    a -= (q - m); m = 0; s_v, c_v, d_v = "ACQUISTO", "acq-row", oggi + timedelta(days=12)
                elif (m + a + p) >= q:
                    p -= (q - m - a); m = 0; a = 0; s_v, c_v, d_v = "PRODUZIONE", "prod-row", oggi + timedelta(days=22)
                else:
                    s_v, c_v, d_v = "MANCANTE", "urgent-row", oggi + timedelta(days=40)

                res = row.to_dict()
                res.update({'ST': s_v, 'CS': c_v, 'DT_EXP': d_v, 'ART_KEY': art, 'CLI_NAME': str(row[c_cli])})
                final_results.append(res)
        
        return pd.DataFrame(final_results), stock_map
    except Exception as e:
        st.error(f"Errore: {e}"); return pd.DataFrame(), {}

# --- 5. LOGICA ACCESSO ---
if "auth" not in st.session_state: st.session_state.auth = False
if not st.session_state.auth:
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=250)
        u = st.text_input("Utente"); p = st.text_input("Password", type="password")
        if st.button("Entra", use_container_width=True):
            if u in USER_DB and USER_DB[u][0] == p:
                st.session_state.auth, st.session_state.user, st.session_state.permesso = True, u, USER_DB[u][1]
                st.rerun()
            else: st.error("Credenziali non valide")
    st.stop()

# --- 6. UI E KPI ---
df_res, stock_raw = load_and_process()

if not df_res.empty:
    with st.sidebar:
        st.markdown(f'<div class="user-info">👤 <b>{st.session_state.user}</b></div>', unsafe_allow_html=True)
        sel_cli = st.selectbox("Cliente:", ["TUTTI"] + sorted(df_res['CLI_NAME'].unique().tolist())) if st.session_state.permesso == "TUTTI" else st.session_state.permesso
        stati_possibili = ["DISPONIBILE", "ACQUISTO", "PRODUZIONE", "MANCANTE"]
        sel_stati = [s for s in stati_possibili if st.checkbox(s, value=True, key=f"k_{s}")]
        search = st.text_input("🔍 Cerca Articolo:").upper()

    # Filtro Finale
    df_f = df_res[df_res['CLI_NAME'] == sel_cli] if sel_cli != "TUTTI" else df_res.copy()
    df_f = df_f[df_f['ST'].isin(sel_stati)]
    if search: df_f = df_f[df_f['ART_KEY'].str.contains(search)]

    st.title("Pannello Controllo Consegne Safit")
    
    # KPI CORRETTI
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f'<div class="kpi-card"><div>PEZZI FILTRATI</div><div class="kpi-val">{int(df_f["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div>PRONTI</div><div class="kpi-val" style="color:#4caf50">{int(df_f[df_f["ST"]=="DISPONIBILE"]["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card"><div>IN ACQUISTO</div><div class="kpi-val" style="color:#2196f3">{int(df_f[df_f["ST"]=="ACQUISTO"]["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi-card"><div>MANCANTI</div><div class="kpi-val" style="color:#f44336">{int(df_f[df_f["ST"]=="MANCANTE"]["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)

    st.markdown("---")
    for art, g in df_f.groupby('ART_KEY'):
        with st.expander(f"📦 {art} - {g['Articolo D'].iloc[0]}"):
            s_info = stock_raw.get(art, {'GIA': 0, 'ACQ': 0, 'PROD': 0})
            st.markdown(f'<div class="debug-box"><span>📦 GIA: <b>{int(s_info["GIA"])}</b></span><span>🚚 ACQ: <b>{int(s_info["ACQ"])}</b></span><span>⚙️ PROD: <b>{int(s_info["PROD"])}</b></span></div>', unsafe_allow_html=True)
            for _, r in g.iterrows():
                st.markdown(f'<div class="status-row {r["CS"]}"><span>📅 {r["DT_EXP"].strftime("%d/%m/%Y")} | Q: {int(r["Qta Residua"])} | {r["CLI_NAME"]}</span><span><b>{r["ST"]}</b></span></div>', unsafe_allow_html=True)
