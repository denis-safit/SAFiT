import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE E STILE ---
APP_VERSION = "1.1.8"
st.set_page_config(page_title=f"Safit Portal v{APP_VERSION}", layout="wide")

st.markdown(f"""
    <style>
    .status-row {{
        display: flex; justify-content: space-between; align-items: center;
        padding: 12px 18px; border-radius: 10px; margin-bottom: 8px; font-size: 14px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }}
    .on-time-row {{ background-color: #e8f5e9; border-left: 8px solid #4caf50; color: #2e7d32; }} /* GIA */
    .acq-row {{ background-color: #e3f2fd; border-left: 8px solid #2196f3; color: #1565c0; }}    /* ACQ */
    .prod-row {{ background-color: #fffde7; border-left: 8px solid #fbc02d; color: #f57f17; }}   /* PROD */
    .urgent-row {{ background-color: #ffebee; border-left: 8px solid #f44336; color: #c62828; }} /* MANCANTE */
    </style>
    """, unsafe_allow_html=True)

# --- 2. GESTIONE UTENTI ---
@st.cache_data
def load_users():
    if os.path.exists('utenti.xlsx'):
        try:
            df_u = pd.read_excel('utenti.xlsx')
            df_u.columns = [str(c).strip() for c in df_u.columns]
            return df_u.set_index('username')[['password', 'cliente_arca']].T.to_dict('list')
        except: pass
    return {'safit_admin': ['admin2026', 'TUTTI']}

USER_DB = load_users()

if "authenticated" not in st.session_state: st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=300)
        u = st.text_input("Username").strip(); p = st.text_input("Password", type="password").strip()
        if st.button("Accedi"):
            if u in USER_DB and str(USER_DB[u][0]) == p:
                st.session_state.update({"authenticated": True, "user_type": USER_DB[u][1], "username": u})
                st.rerun()
            else: st.error("Accesso negato")
    st.stop()

# --- 3. FUNZIONI TECNICHE ---
def clean_num(serie):
    return pd.to_numeric(serie.astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').fillna(0)

def find_col(df, target):
    for c in df.columns:
        if c.strip().upper() == target.strip().upper(): return c
    return None

# --- 4. CARICAMENTO DATI (GESTIONE PIVOT) ---
@st.cache_data
def load_data_pivot():
    try:
        # ARCA
        df_a = pd.read_excel('righe_Ordini_ARCA.xlsx', sheet_name='Foglio1', skiprows=2)
        df_a.columns = [str(c).strip() for c in df_a.columns]
        
        # Identificazione Colonne
        c_cli = find_col(df_a, 'Cliente Fornitore CD')
        c_art = find_col(df_a, 'Articolo C')
        c_des = find_col(df_a, 'Articolo D')
        c_dat = find_col(df_a, 'Data')
        c_qta = find_col(df_a, 'Qta Residua') or find_col(df_a, 'Qta Doc')

        if not c_art:
            st.error(f"Colonna 'Articolo C' non trovata! Colonne: {list(df_a.columns)}")
            return pd.DataFrame()

        # LOGICA PIVOT: Riempimento celle vuote
        for col in [c_cli, c_art, c_des]:
            if col: df_a[col] = df_a[col].ffill()
        
        df_a = df_a.dropna(subset=[c_art])
        df_a['Art_Key'] = df_a[c_art].astype(str).str.strip().upper()
        df_a['Data_Dt'] = pd.to_datetime(df_a[c_dat], errors='coerce')
        df_a['Qta_Res'] = clean_num(df_a[c_qta])
        df_a = df_a[df_a['Qta_Res'] > 0]

        # ACCESS
        if os.path.exists('Avanzamento_access.xlsx'):
            df_t = pd.read_excel('Avanzamento_access.xlsx', skiprows=1)
            df_t.columns = [str(c).strip() for c in df_t.columns]
            
            df_t['Key_Acc'] = df_t[find_col(df_t, 'Codice')].astype(str).str.strip().upper()
            df_t['GIA_TOT'] = clean_num(df_t[find_col(df_t, 'Gia')])
            df_t['ACQ_TOT'] = clean_num(df_t[find_col(df_t, 'Acq')])
            
            # Somma Produzione
            p_cols = ['Lan', 'GRZ', 'TMP', 'RWI', 'TRS']
            df_t['PROD_TOT'] = 0
            for p in p_cols:
                cp = find_col(df_t, p)
                if cp: df_t['PROD_TOT'] += clean_num(df_t[cp])

            df_f = pd.merge(df_a, df_t[['Key_Acc', 'GIA_TOT', 'ACQ_TOT', 'PROD_TOT']], 
                            left_on='Art_Key', right_on='Key_Acc', how='left')
            return df_f.fillna(0)
        return df_a
    except Exception as e:
        st.error(f"Errore: {e}"); return pd.DataFrame()

# --- 5. LOGICA APPLICATIVA ---
data = load_data_pivot()
sel_cli = "Seleziona..." # Protezione NameError

if not data.empty:
    with st.sidebar:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
        st.write(f"Utente: {st.session_state['username']}")
        c_cli_n = find_col(data, 'Cliente Fornitore CD')
        if st.session_state["user_type"] == "TUTTI":
            clist = sorted(data[c_cli_n].unique().astype(str))
            sel_cli = st.selectbox("Cliente:", clist)
        else: sel_cli = st.session_state["user_type"]
        if st.button("Logout"): st.session_state["authenticated"] = False; st.rerun()

    st.title(f"Avanzamento: {sel_cli}")
    df_c = data[data[c_cli_n] == sel_cli].copy()
    oggi = datetime.now()
    
    # CALCOLO A CASCATA PER ARTICOLO
    for art in sorted(df_c[find_col(df_c, 'Articolo C')].unique()):
        df_art = df_c[df_c[find_col(df_c, 'Articolo C')] == art].sort_values('Data_Dt')
        
        m_stock = float(df_art['GIA_TOT'].iloc[0])
        a_stock = float(df_art['ACQ_TOT'].iloc[0])
        p_stock = float(df_art['PROD_TOT'].iloc[0])
        descr = df_art[find_col(df_art, 'Articolo D')].iloc[0]

        with st.expander(f"📦 {art} - {descr} (GIA: {m_stock:,.0f})"):
            for _, row in df_art.iterrows():
                q = float(row['Qta_Res'])
                dt = row['Data_Dt']
                
                if m_stock >= q:
                    m_stock -= q; msg, css, est = "PRONTO (GIA)", "on-time-row", dt
                elif (m_stock + a_stock) >= q:
                    a_stock -= (q-m_stock); m_stock = 0; msg, css, est = "ACQUISTO", "acq-row", oggi+timedelta(12)
                elif (m_stock + a_stock + p_stock) >= q:
                    p_stock -= (q-m_stock-a_stock); m_stock=0; a_stock=0; msg, css, est = "PRODUZIONE", "prod-row", oggi+timedelta(22)
                else:
                    msg, css, est = "MANCANTE", "urgent-row", oggi+timedelta(40)

                st.markdown(f'<div class="status-row {css}"><span>📅 <b>{dt.strftime("%d/%m/%Y")}</b> | Qta: {q:,.0f}</span><span>{msg} ({est.strftime("%d/%m/%Y")})</span></div>', unsafe_allow_html=True)
