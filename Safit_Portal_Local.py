import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE E STILE ---
APP_VERSION = "1.1.9"
st.set_page_config(page_title=f"Safit Portal v{APP_VERSION}", layout="wide")

st.markdown(f"""
    <style>
    .status-row {{
        display: flex; justify-content: space-between; align-items: center;
        padding: 12px 18px; border-radius: 10px; margin-bottom: 8px; font-size: 14px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }}
    .on-time-row {{ background-color: #e8f5e9; border-left: 8px solid #4caf50; color: #2e7d32; }} 
    .acq-row {{ background-color: #e3f2fd; border-left: 8px solid #2196f3; color: #1565c0; }}    
    .prod-row {{ background-color: #fffde7; border-left: 8px solid #fbc02d; color: #f57f17; }}   
    .urgent-row {{ background-color: #ffebee; border-left: 8px solid #f44336; color: #c62828; }} 
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

if "authenticated" not in st.session_state: 
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=300)
        st.title("Login Safit Portal")
        u = st.text_input("Username").strip()
        p = st.text_input("Password", type="password").strip()
        if st.button("Accedi"):
            if u in USER_DB and str(USER_DB[u][0]) == p:
                st.session_state.update({"authenticated": True, "user_type": USER_DB[u][1], "username": u})
                st.rerun()
            else: st.error("Credenziali errate")
    st.stop()

# --- 3. FUNZIONI TECNICHE ---
def clean_num(serie):
    return pd.to_numeric(serie.astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').fillna(0)

def find_col(df, targets):
    """Cerca una colonna tra una lista di nomi possibili"""
    for c in df.columns:
        if any(t.upper() in str(c).upper() for t in targets):
            return c
    return None

def smart_load_excel(filename, sheet=0):
    """Trova automaticamente la riga di intestazione corretta"""
    # Carica le prime 20 righe per cercare i titoli
    df_preview = pd.read_excel(filename, sheet_name=sheet, header=None, nrows=20)
    header_row = 0
    for i, row in df_preview.iterrows():
        # Se troviamo 'Articolo' o 'Cliente' in una riga, quella è la nostra testata
        row_str = " ".join([str(x) for x in row.values])
        if "Articolo" in row_str or "Cliente" in row_str:
            header_row = i
            break
    
    # Ricarica il file dalla riga giusta
    df = pd.read_excel(filename, sheet_name=sheet, skiprows=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    return df

# --- 4. CARICAMENTO DATI ---
@st.cache_data
def load_full_data():
    try:
        # ARCA
        file_arca = 'righe_Ordini_ARCA.xlsx'
        df_a = smart_load_excel(file_arca)
        
        c_cli = find_col(df_a, ['Cliente Fornitore', 'CD'])
        c_art = find_col(df_a, ['Articolo C', 'Cod. Art'])
        c_des = find_col(df_a, ['Articolo D', 'Descriz'])
        c_dat = find_col(df_a, ['Data'])
        c_qta = find_col(df_a, ['Qta Residua', 'Qta Doc', 'Residua'])

        if not c_art:
            st.error(f"Impossibile trovare la colonna Articolo! Colonne rilevate: {list(df_a.columns)}")
            return pd.DataFrame()

        # Gestione Pivot (Riempimento vuoti)
        for col in [c_cli, c_art, c_des]:
            if col: df_a[col] = df_a[col].ffill()
        
        df_a = df_a.dropna(subset=[c_art])
        df_a['Art_Key'] = df_a[c_art].astype(str).str.strip().upper()
        df_a['Data_Dt'] = pd.to_datetime(df_a[c_dat], errors='coerce')
        df_a['Qta_Res'] = clean_num(df_a[c_qta])
        df_a = df_a[df_a['Qta_Res'] > 0]

        # ACCESS
        file_acc = 'Avanzamento_access.xlsx'
        if os.path.exists(file_acc):
            df_t = smart_load_excel(file_acc)
            c_code = find_col(df_t, ['Codice', 'Articolo'])
            if c_code:
                df_t['Key_Acc'] = df_t[c_code].astype(str).str.strip().upper()
                df_t['GIA_TOT'] = clean_num(df_t[find_col(df_t, ['Gia'])])
                df_t['ACQ_TOT'] = clean_num(df_t[find_col(df_t, ['Acq'])])
                
                # Somma fasi produzione
                p_cols = ['Lan', 'GRZ', 'TMP', 'RWI', 'TRS']
                df_t['PROD_TOT'] = 0
                for p in p_cols:
                    cp = find_col(df_t, [p])
                    if cp: df_t['PROD_TOT'] += clean_num(df_t[cp])

                df_f = pd.merge(df_a, df_t[['Key_Acc', 'GIA_TOT', 'ACQ_TOT', 'PROD_TOT']], 
                                left_on='Art_Key', right_on='Key_Acc', how='left')
                return df_f.fillna(0)
        return df_a
    except Exception as e:
        st.error(f"Errore caricamento: {e}"); return pd.DataFrame()

# --- 5. INTERFACCIA ---
data = load_full_data()
sel_cli = "Seleziona..."

if not data.empty:
    with st.sidebar:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
        st.write(f"👤 Utente: **{st.session_state['username']}**")
        c_cli_n = find_col(data, ['Cliente Fornitore'])
        if st.session_state["user_type"] == "TUTTI":
            clist = sorted(data[c_cli_n].unique().astype(str))
            sel_cli = st.selectbox("Cliente:", clist)
        else: sel_cli = st.session_state["user_type"]
        if st.button("Logout"): st.session_state["authenticated"] = False; st.rerun()

    st.title(f"Avanzamento: {sel_cli}")
    df_c = data[data[c_cli_n] == sel_cli].copy()
    oggi = datetime.now()
    
    c_art_n = find_col(df_c, ['Articolo C'])
    c_des_n = find_col(df_c, ['Articolo D'])

    for art in sorted(df_c[c_art_n].unique()):
        df_art = df_c[df_c[c_art_n] == art].sort_values('Data_Dt')
        m_stock = float(df_art['GIA_TOT'].iloc[0])
        a_stock = float(df_art['ACQ_TOT'].iloc[0])
        p_stock = float(df_art['PROD_TOT'].iloc[0])
        descr = df_art[c_des_n].iloc[0]

        with st.expander(f"📦 {art} - {descr} (Disp: {m_stock:,.0f})"):
            for _, row in df_art.iterrows():
                q = float(row['Qta_Res'])
                dt = row['Data_Dt']
                
                if m_stock >= q:
                    m_stock -= q; msg, css, est = "PRONTO (GIA)", "on-time-row", dt
                elif (m_stock + a_stock) >= q:
                    a_stock -= (q-m_stock); m_stock = 0; msg, css, est = "IN ARRIVO (ACQ)", "acq-row", oggi+timedelta(12)
                elif (m_stock + a_stock + p_stock) >= q:
                    p_stock -= (q-m_stock-a_stock); m_stock=0; a_stock=0; msg, css, est = "PRODUZIONE", "prod-row", oggi+timedelta(22)
                else:
                    msg, css, est = "DA LANCIARE", "urgent-row", oggi+timedelta(40)

                st.markdown(f'<div class="status-row {css}"><span>📅 <b>{dt.strftime("%d/%m/%Y")}</b> | Qta: {q:,.0f}</span><span>{msg} (Est: {est.strftime("%d/%m/%Y")})</span></div>', unsafe_allow_html=True)
