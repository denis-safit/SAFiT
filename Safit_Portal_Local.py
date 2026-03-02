import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO
import plotly.express as px

# --- 1. CONFIGURAZIONE E STILE ---
APP_VERSION = "1.6.2"
st.set_page_config(page_title=f"Safit Portal v{APP_VERSION}", layout="wide")

st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 10px; border-radius: 8px; margin-bottom: 6px; border: 1px solid #ddd; color: #000 !important; font-size: 14px; }
    .on-time-row { background-color: #e8f5e9 !important; border-left: 6px solid #4caf50; } 
    .acq-row { background-color: #e3f2fd !important; border-left: 8px solid #2196f3; }    
    .prod-row { background-color: #fffde7 !important; border-left: 8px solid #fbc02d; }   
    .urgent-row { background-color: #ffebee !important; border-left: 8px solid #f44336; } 
    .debug-box { background-color: #f0f2f6 !important; color: #111 !important; padding: 8px 12px; border-radius: 6px; border: 1px solid #ccc; margin-bottom: 10px; font-family: sans-serif; font-size: 13px; font-weight: 600; display: flex; justify-content: space-between; white-space: nowrap; overflow-x: auto; }
    .kpi-card { background-color: #ffffff; border: 1px solid #e0e0e0; padding: 15px; border-radius: 10px; text-align: center; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .kpi-val { font-size: 24px; font-weight: bold; color: #1f77b4; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNZIONI SUPPORTO ---
def fmt_n(val):
    try: return f"{int(round(float(val), 0)):,}".replace(",", ".")
    except: return "0"

def clean_num(serie):
    s = serie.astype(str).str.replace(' ', '').str.replace('\xa0', '')
    def fix_val(val):
        if val.lower() in ['nan', '', 'none']: return '0'
        if ',' in val and '.' in val: return val.replace('.', '').replace(',', '.')
        elif ',' in val: return val.replace(',', '.')
        return val
    return pd.to_numeric(s.apply(fix_val), errors='coerce').fillna(0)

def smart_load(filename):
    if not os.path.exists(filename): return pd.DataFrame()
    df_p = pd.read_excel(filename, header=None, nrows=15)
    h_row = 0
    for i, row in df_p.iterrows():
        row_s = " ".join([str(x) for x in row.values])
        if any(k in row_s for k in ["Codice Documento", "Articolo C", "Cliente Fornitore"]):
            h_row = i; break
    df = pd.read_excel(filename, skiprows=h_row)
    df.columns = [str(c).strip() for c in df.columns]
    return df

# --- 3. MOTORE LOGICO ATP UNIFICATO ---
@st.cache_data
def load_and_process():
    try:
        # A. Caricamento unico file ARCA
        df_full = smart_load('righe_Ordini_ARCA.xlsx')
        if df_full.empty: return pd.DataFrame()

        # Pulizia colonne basata sullo screenshot
        c_tipo = "Codice Documento"
        c_art = "Articolo C"
        c_des = "Articolo D"
        c_qta = "Qta Residua"
        c_dat = "Data Consegna"
        c_cli = "Cliente Fornitore CD"

        df_full[c_art] = df_full[c_art].astype(str).str.strip().str.upper()
        df_full[c_qta] = clean_num(df_full[c_qta])
        
        # B. Caricamento Access (Giacenza + Produzione)
        df_access = smart_load('Avanzamento_access.xlsx')
        access_data = {}
        if not df_access.empty:
            df_access['KEY'] = df_access['CODICE'].astype(str).str.strip().str.upper()
            for _, r in df_access.iterrows():
                p_val = 0
                for f in ['LANCIATI', 'GRZ', 'TMP', 'RWI', 'TRS']:
                    if f in df_access.columns: p_val += clean_num(pd.Series([r[f]])).iloc[0]
                access_data[r['KEY']] = {'GIA': clean_num(pd.Series([r['GIA']])).iloc[0], 'PROD': p_val}

        # C. Preparazione Arrivi Reali (OFF e DCL)
        df_in = df_full[df_full[c_tipo].isin(['OFF', 'DCL'])].copy()
        df_in['DATA_DT'] = pd.to_datetime(df_in[c_dat], errors='coerce')
        all_arrivals = {}
        for art in df_in[c_art].unique():
            all_arrivals[art] = df_in[df_in[c_art] == art].sort_values('DATA_DT')[[c_dat, c_qta]].values.tolist()

        # D. Elaborazione Ordini Cliente (OCI / OCA)
        df_oci = df_full[df_full[c_tipo].isin(['OCI', 'OCA'])].copy().sort_values(c_dat)
        
        results = []
        for _, row in df_oci.iterrows():
            art = row[c_art]
            qta_req = row[c_qta]
            
            # Parametri articolo
            gia = access_data.get(art, {}).get('GIA', 0)
            prod = access_data.get(art, {}).get('PROD', 0)
            arrivals = all_arrivals.get(art, [])

            status, color, date_est = "MANCANTE", "urgent-row", datetime.now() + timedelta(days=45)

            # 1. Coperto da Giacenza?
            if gia >= qta_req:
                gia -= qta_req
                if art in access_data: access_data[art]['GIA'] = gia
                status, color, date_est = "DISPONIBILE", "on-time-row", pd.to_datetime(row[c_dat])
            
            # 2. Coperto da Arrivi Reali (OFF/DCL)?
            else:
                fabbisogno = qta_req - gia
                gia = 0
                if art in access_data: access_data[art]['GIA'] = 0
                
                trovato_acq = False
                for i, arr in enumerate(arrivals):
                    d_arr, q_arr = arr[0], arr[1]
                    if q_arr >= fabbisogno:
                        arrivals[i][1] = q_arr - fabbisogno
                        status, color, date_est = "ACQUISTO", "acq-row", pd.to_datetime(d_arr)
                        trovato_acq = True
                        break
                    else:
                        fabbisogno -= q_arr
                        arrivals[i][1] = 0
                
                # 3. Coperto da Produzione Interna (Access)?
                if not trovato_acq and prod >= fabbisogno:
                    prod -= fabbisogno
                    if art in access_data: access_data[art]['PROD'] = prod
                    status, color, date_est = "PRODUZIONE", "prod-row", datetime.now() + timedelta(days=21)

            res = row.to_dict()
            res.update({'ST': status, 'CS': color, 'DT_E': date_est, 'ART_KEY': art, 'CLI_NAME': row[c_cli]})
            results.append(res)

        return pd.DataFrame(results), access_data
    except Exception as e:
        st.error(f"Errore: {e}")
        return pd.DataFrame(), {}

# --- 4. GESTIONE UTENTI E INTERFACCIA ---
@st.cache_data
def get_user_db():
    if os.path.exists('utenti.xlsx'):
        try:
            df_u = pd.read_excel('utenti.xlsx')
            return df_u.set_index('username')[['password', 'cliente_arca']].T.to_dict('list')
        except: pass
    return {'safit_admin': ['admin2026', 'TUTTI']}

USER_DB = get_user_db()
if "authenticated" not in st.session_state: st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=300)
        st.title("Safit Portal - Login")
        u = st.text_input("Username").strip()
        p = st.text_input("Password", type="password").strip()
        if st.button("Accedi"):
            if u in USER_DB and str(USER_DB[u][0]) == p:
                st.session_state.update({"authenticated": True, "user_type": USER_DB[u][1], "username": u})
                st.rerun()
            else: st.error("Credenziali errate")
    st.stop()

# --- 5. LOGICA DASHBOARD ---
df_res, stock_final = load_and_process()

if not df_res.empty:
    if st.session_state["user_type"] == "TUTTI":
        with st.sidebar:
            st.markdown("### 🛠️ Amministrazione")
            sel_cli = st.selectbox("Seleziona Cliente:", ["TUTTI I CLIENTI"] + sorted(df_res['CLI_NAME'].unique().astype(str)))
    else:
        sel_cli = st.session_state["user_type"]

    df_filtered = df_res if sel_cli == "TUTTI I CLIENTI" else df_res[df_res['CLI_NAME'] == sel_cli]

    # DASHBOARD ADMIN
    if st.session_state["user_type"] == "TUTTI":
        st.title(f"Dashboard Analitica: {sel_cli}")
        tot_q = df_filtered['Qta Residua'].sum()
        def get_perc(s): return (df_filtered[df_filtered['ST'] == s]['Qta Residua'].sum() / tot_q * 100) if tot_q > 0 else 0

        k1, k2, k3, k4 = st.columns(4)
        with k1: st.metric("Pezzi Ordinati", fmt_n(tot_q))
        with k2: st.metric("Pronti %", f"{get_perc('DISPONIBILE'):.1f}%")
        with k3: st.metric("In Prod %", f"{get_perc('PRODUZIONE'):.1f}%")
        with k4: st.metric("Mancanti %", f"{get_perc('MANCANTE'):.1f}%")

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            fig1 = px.pie(df_filtered, values='Qta Residua', names='ST', color='ST', 
                         color_discrete_map={'DISPONIBILE':'#4caf50','ACQUISTO':'#2196f3','PRODUZIONE':'#fbc02d','MANCANTE':'#f44336'})
            st.plotly_chart(fig1, use_container_width=True)
        with col_g2:
            df_filtered['Famiglia'] = df_filtered['Articolo D'].apply(lambda x: " ".join(str(x).split()[:2]).upper())
            fig2 = px.pie(df_filtered.groupby('Famiglia')['Qta Residua'].sum().reset_index().sort_values('Qta Residua', ascending=False).head(10), values='Qta Residua', names='Famiglia', hole=0.4)
            st.plotly_chart(fig2, use_container_width=True)

    # VISUALIZZAZIONE LISTA
    st.markdown("---")
    with st.sidebar:
        search_art = st.text_input("🔍 Filtra per Codice Articolo:").upper()
    
    df_list = df_filtered[df_filtered['ART_KEY'].str.contains(search_art)] if search_art else df_filtered

    for art in sorted(df_list['ART_KEY'].unique()):
        df_sub = df_list[df_list['ART_KEY'] == art]
        with st.expander(f"📦 {art} - {df_sub['Articolo D'].iloc[0]}"):
            # Recupero info stock aggiornate per l'articolo
            info = stock_final.get(art, {'GIA':0, 'PROD':0})
            st.markdown(f'''<div class="debug-box">
                <span>GIA Attuale: <b>{fmt_n(info['GIA'])}</b></span>
                <span>PROD Attuale: <b>{fmt_n(info['PROD'])}</b></span>
            </div>''', unsafe_allow_html=True)
            for _, r in df_sub.iterrows():
                st.markdown(f"""<div class="status-row {r['CS']}">
                    <span>📅 <b>{pd.to_datetime(r['Data Consegna']).strftime('%d/%m/%Y')}</b> | Q: {fmt_n(r['Qta Residua'])}</span>
                    <span><b>{r['ST']}</b> ({pd.to_datetime(r['DT_E']).strftime('%d/%m/%Y')})</span>
                </div>""", unsafe_allow_html=True)
else:
    st.warning("Dati non trovati nel file righe_Ordini_ARCA.xlsx")
