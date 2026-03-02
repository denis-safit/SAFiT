import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO
import plotly.express as px

# --- 1. CONFIGURAZIONE ---
APP_VERSION = "2.1.0-Logistics-Fix"
st.set_page_config(page_title=f"Safit Portal v{APP_VERSION}", layout="wide")

st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 10px; border-radius: 8px; margin-bottom: 6px; border: 1px solid #ddd; color: #000 !important; font-size: 14px; }
    .on-time-row { background-color: #e8f5e9 !important; border-left: 6px solid #4caf50; } 
    .acq-row { background-color: #e3f2fd !important; border-left: 8px solid #2196f3; }    
    .prod-row { background-color: #fffde7 !important; border-left: 8px solid #fbc02d; }   
    .urgent-row { background-color: #ffebee !important; border-left: 8px solid #f44336; } 
    .oca-row { background-color: #f5f5f5 !important; border-left: 8px solid #9e9e9e; color: #666 !important; }
    .debug-box { background-color: #f0f2f6 !important; color: #111 !important; padding: 8px 12px; border-radius: 6px; border: 1px solid #ccc; margin-bottom: 10px; font-family: sans-serif; font-size: 13px; font-weight: 600; display: flex; justify-content: space-between; }
    .kpi-card { background-color: #ffffff; border: 1px solid #e0e0e0; padding: 15px; border-radius: 10px; text-align: center; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .kpi-val { font-size: 22px; font-weight: bold; color: #1f77b4; }
    .user-info { padding: 10px; background: #f8f9fa; border-radius: 5px; border: 1px solid #eee; margin-bottom: 20px; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. UTENTI ---
USER_DB = {
    "safit_admin": ["admin2026", "TUTTI"],
    "denis": ["denis2026", "TUTTI"]
}

# --- 3. UTILITY ---
def clean_num(serie):
    s = serie.astype(str).str.replace(' ', '').str.replace('\xa0', '')
    def fix_val(val):
        if val.lower() in ['nan', '', 'none']: return '0'
        return val.replace('.', '').replace(',', '.') if ',' in val else val
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
    df = df.ffill() 
    return df

# --- 4. MOTORE ATP AGGIORNATO ---
@st.cache_data(ttl=300)
def load_and_process():
    try:
        # 1. Carico Arca
        df_full = smart_load('righe_Ordini_ARCA.xlsx', "Articolo C")
        if df_full.empty: return pd.DataFrame(), {}
        
        c_tipo, c_art, c_qta, c_dat, c_cli = "Codice Documento", "Articolo C", "Qta Residua", "Data Consegna", "Cliente Fornitore CD"
        df_full[c_dat] = pd.to_datetime(df_full[c_dat], errors='coerce')
        df_full[c_qta] = clean_num(df_full[c_qta])
        df_full = df_full.dropna(subset=[c_dat, c_art])

        # 2. Carico Access (Stock + Produzione)
        df_acc = smart_load('Avanzamento_access.xlsx', "CODICE")
        stock = {}
        if not df_acc.empty:
            for _, r in df_acc.iterrows():
                # Sommo i campi produzione chiesti da Denis: ACQ, TMP, RWI, TRF (o TRF se TRS era errore)
                # Uso r.get con default 0 per evitare errori se manca una colonna
                prod_val = sum([clean_num(pd.Series([r.get(f, 0)])).iloc[0] for f in ['ACQ', 'TMP', 'RWI', 'TRF', 'TRS']])
                stock[str(r['CODICE']).strip().upper()] = {
                    'GIA': clean_num(pd.Series([r.get('GIA', 0)])).iloc[0], 
                    'PROD': prod_val
                }

        # 3. Estraggo Arrivi da Acquisti (OFF / OFR) da Arca
        df_arrivals = df_full[df_full[c_tipo].isin(['OFF', 'OFR'])].copy()
        # Creiamo un dizionario di liste [data, qta] per articolo
        arca_arrivals = {}
        for art, g in df_arrivals.groupby(c_art):
            arca_arrivals[str(art).strip().upper()] = g.sort_values(c_dat)[[c_dat, c_qta]].values.tolist()

        # 4. Processo Ordini (OCI / OCA)
        df_orders = df_full[df_full[c_tipo].isin(['OCI', 'OCA'])].sort_values([c_dat, c_tipo], ascending=[True, False])
        
        final = []
        for _, row in df_orders.iterrows():
            art = str(row[c_art]).strip().upper()
            qta_fabb = row[c_qta]
            s = stock.get(art, {'GIA': 0, 'PROD': 0})
            arr_list = arca_arrivals.get(art, [])
            
            st_v, col, dt_e = ("MANCANTE", "urgent-row", row[c_dat] + timedelta(days=45))
            if row[c_tipo] == 'OCA': st_v, col = "DA PIANIFICARE", "oca-row"

            # STEP A: Giacenza Fisica
            if s['GIA'] >= qta_fabb:
                s['GIA'] -= qta_fabb
                st_v, col, dt_e = "DISPONIBILE", "on-time-row", row[c_dat]
            else:
                qta_fabb -= s['GIA']
                s['GIA'] = 0
                
                # STEP B: Arrivi da Acquisti (OFF/OFR Arca)
                trovato_acquisto = False
                for a in arr_list:
                    if a[1] > 0:
                        if a[1] >= qta_fabb:
                            a[1] -= qta_fabb
                            st_v, col, dt_e = "ACQUISTO", "acq-row", a[0]
                            qta_fabb = 0
                            trovato_acquisto = True
                            break
                        else:
                            qta_fabb -= a[1]
                            a[1] = 0
                
                # STEP C: Avanzamento Produzione (Access)
                if qta_fabb > 0 and s['PROD'] >= qta_fabb:
                    s['PROD'] -= qta_fabb
                    st_v, col, dt_e = "PRODUZIONE", "prod-row", datetime.now() + timedelta(days=15)
                    qta_fabb = 0

            res = row.to_dict()
            res.update({'ST': st_v, 'CS': col, 'DT_EXP': dt_e, 'ART_KEY': art, 'CLI_NAME': str(row[c_cli])})
            final.append(res)
        
        return pd.DataFrame(final), stock
    except Exception as e:
        st.error(f"Errore: {e}"); return pd.DataFrame(), {}

# --- LOGIN E UI (Standard come approvato) ---
if "authenticated" not in st.session_state: st.session_state["authenticated"] = False
if not st.session_state["authenticated"]:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=300)
        st.title("Safit Portal - Login")
        u = st.text_input("User"); p = st.text_input("Pass", type="password")
        if st.button("Accedi"):
            if u in USER_DB and USER_DB[u][0] == p:
                st.session_state.update({"authenticated": True, "username": u, "user_type": USER_DB[u][1]})
                st.rerun()
    st.stop()

df_res, stock_final = load_and_process()
if not df_res.empty:
    with st.sidebar:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
        st.markdown(f'<div class="user-info">👤 <b>{st.session_state["username"]}</b></div>', unsafe_allow_html=True)
        if st.button("🚪 Esci"): st.session_state.authenticated = False; st.rerun()
        st.markdown("---")
        # Filtri
        if st.session_state["user_type"] == "TUTTI":
            sel_cli = st.selectbox("Cliente:", ["TUTTI"] + sorted(df_res['CLI_NAME'].unique().tolist()))
        else:
            sel_cli = st.session_state["user_type"]
        
        sel_st = st.multiselect("Stato:", df_res['ST'].unique(), default=df_res['ST'].unique())
        search = st.text_input("🔍 Articolo:").upper()

    # Applicazione Filtri
    df_f = df_res[df_res['CLI_NAME'] == sel_cli] if sel_cli != "TUTTI" else df_res.copy()
    df_f = df_f[df_f['ST'].isin(sel_st)]
    if search: df_f = df_f[df_f['ART_KEY'].str.contains(search)]

    # Dashboard
    st.title(f"Piano Consegne Safit")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Pezzi Totali", f"{int(df_f['Qta Residua'].sum()):,}".replace(",", "."))
    k2.metric("Disponibile %", f"{(len(df_f[df_f['ST']=='DISPONIBILE'])/len(df_f)*100):.1f}%" if len(df_f)>0 else "0%")
    
    # Grafici
    c1, c2 = st.columns(2)
    with c1: st.plotly_chart(px.pie(df_f, values='Qta Residua', names='ST', color='ST', color_discrete_map={'DISPONIBILE':'#4caf50','ACQUISTO':'#2196f3','PRODUZIONE':'#fbc02d','MANCANTE':'#f44336','DA PIANIFICARE':'#9e9e9e'}), use_container_width=True)
    with c2:
        df_f['Famiglia'] = df_f['Articolo D'].apply(lambda x: " ".join(str(x).split()[:2]).upper())
        st.plotly_chart(px.pie(df_f.groupby('Famiglia')['Qta Residua'].sum().reset_index().head(10), values='Qta Residua', names='Famiglia', hole=0.4), use_container_width=True)

    # Lista
    for art, g in df_f.groupby('ART_KEY'):
        with st.expander(f"📦 {art} - {g['Articolo D'].iloc[0]}"):
            info = stock_final.get(art, {'GIA': 0, 'PROD': 0})
            st.markdown(f'<div class="debug-box"><span>GIA: {info["GIA"]}</span><span>AVANZAMENTO (Access): {info["PROD"]}</span></div>', unsafe_allow_html=True)
            for _, r in g.iterrows():
                tag = "📋 [PREV]" if r['Codice Documento'] == "OCA" else "🛒 [ORD]"
                # Se è in ACQUISTO, mostriamo la data reale del file Arca
                d_str = r['DT_EXP'].strftime("%d/%m/%Y")
                st.markdown(f'<div class="status-row {r["CS"]}"><span>{tag} 📅 {d_str} | Q: {int(r["Qta Residua"])} | {r["CLI_NAME"]}</span><span><b>{r["ST"]}</b></span></div>', unsafe_allow_html=True)
