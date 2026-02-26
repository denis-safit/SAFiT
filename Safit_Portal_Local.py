import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE PAGINA ---
APP_VERSION = "1.0.11"
st.set_page_config(page_title=f"Safit - Portale Avanzamento {APP_VERSION}", layout="wide")

st.markdown(f"""
    <style>
    .main {{ background-color: #fcfcfc; }}
    .stApp {{ margin-top: -30px; }}
    
    /* Intestazione riga: Codice a sinistra, Pillola a destra */
    .custom-header {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 10px;
        background: #f0f2f6;
        border: 1px solid #ccc;
        border-radius: 8px;
        margin-bottom: 2px;
    }}

    /* Pillola grande richiesta */
    .pill-bg {{
        background-color: #ddd;
        border-radius: 12px;
        width: 160px;
        height: 24px;
        border: 1px solid #bbb;
        position: relative;
    }}
    .pill-fill {{
        background-color: #4caf50;
        height: 100%;
        border-radius: 12px;
    }}
    .pill-text {{
        position: absolute;
        top: 0; left: 0; width: 100%;
        font-size: 14px;
        font-weight: bold;
        line-height: 22px;
        text-align: center;
        color: #000;
    }}

    .status-row {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 10px 15px;
        border-radius: 8px;
        margin-bottom: 5px;
    }}
    
    /* COLORI LOGICA 1.0.02 (PULITI) */
    .on-time-row {{ background-color: #f1f8e9; border-left: 6px solid #4caf50; color: #1b5e20; }}
    .client-delay-row {{ background-color: #e3f2fd; border-left: 6px solid #2196f3; color: #0d47a1; }}
    .delay-row {{ background-color: #fff8e1; border-left: 6px solid #ffc107; color: #5d4037; }}
    .prod-delay-row {{ background-color: #ffebee; border-left: 6px solid #f44336; color: #b71c1c; }}
    </style>
    """, unsafe_allow_html=True)

# --- 2. GESTIONE UTENTI (RIPRISTINO LOGO E LOGICA) ---
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
    if "authenticated" not in st.session_state: st.session_state["authenticated"] = False
    if not st.session_state["authenticated"]:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=300)
            st.title("Accesso Area Riservata")
            user, pw = st.text_input("Username"), st.text_input("Password", type="password")
            if st.button("Accedi"):
                if user in USER_DB and str(USER_DB[user][0]) == pw:
                    st.session_state["authenticated"], st.session_state["user_type"], st.session_state["username"] = True, USER_DB[user][1], user
                    st.rerun()
                else: st.error("Username o Password errati")
        return False
    return True

if not check_password(): st.stop()

# --- 3. CARICAMENTO DATI ---
def aggiungi_giorni_lavorativi(data_inizio, giorni):
    data_corrente = data_inizio
    while giorni > 0:
        data_corrente += timedelta(days=1)
        if data_corrente.weekday() < 5: giorni -= 1
    return data_corrente

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
        col_res = 'Qta Residua' if 'Qta Residua' in df.columns else 'Qta Doc'
        df['Qta_Effettiva'] = pd.to_numeric(df[col_res], errors='coerce').fillna(0)
        df = df[df['Qta_Effettiva'] > 0]
        
        if os.path.exists('Avanzamento_access.xlsx'):
            df_tech = pd.read_excel('Avanzamento_access.xlsx', skiprows=1) 
            df_tech.columns = [str(c).strip() for c in df_tech.columns]
            if 'Codice' in df_tech.columns:
                df_tech = df_tech.rename(columns={'Codice': 'Art_Key'})
                for c in ['Gia', 'Acq', 'Lan', 'Grz', 'Tmp', 'Rwi', 'Trs']:
                    if c in df_tech.columns: df_tech[c] = pulisci_numero(df_tech[c])
                df = pd.merge(df, df_tech, left_on='Articolo C', right_on='Art_Key', how='left')
        return df
    except: return pd.DataFrame()

data = load_data()
oggi_dt = datetime.now()

# --- 4. SIDEBAR (LOGO E FILTRI) ---
with st.sidebar:
    if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
    st.write(f"Utente: **{st.session_state['username']}**")
    if st.session_state["user_type"] == "TUTTI":
        clienti_list = sorted([str(x) for x in data['Cliente Fornitore CD'].unique()])
        sel_cli = st.selectbox("👤 Seleziona Cliente:", clienti_list)
    else: sel_cli = st.session_state["user_type"]
    
    filtro_label = st.radio("Filtra per stato:", ["Mostra tutto", "Solo Disponibili", "In Lavorazione", "In Ritardo"])
    if st.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

# --- 5. VISUALIZZAZIONE ---
st.title("🚜 Portale Avanzamento Produzione")
df_cli = data[data['Cliente Fornitore CD'] == sel_cli].copy()

if not df_cli.empty:
    articoli_view = sorted([str(x) for x in df_cli['Articolo C'].unique()])
    
    for art in articoli_view:
        df_art = df_cli[df_cli['Articolo C'] == art].sort_values('Data_Consegna')
        desc = df_art['Articolo D'].iloc[0]
        row_t = df_art.iloc[0]
        
        # Dati tecnici per calcolo logico
        st_gia = float(row_t.get('Gia', 0))
        st_trs, st_rwi, st_acq, st_tmp, st_lan, st_grz = [float(row_t.get(k, 0)) for k in ['Trs','Rwi','Acq','Tmp','Lan','Grz']]

        righe_mostra = []
        for _, row in df_art.iterrows():
            qta, req_date = float(row['Qta_Effettiva']), row['Data_Consegna']
            
            # CATEGORIA CORE (FISSA 1.0.02)
            cat_core = "LAVORAZIONE"
            if st_gia >= qta:
                cat_core = "DISPONIBILE"
                st_gia -= qta
            
            eta = oggi_dt if cat_core == "DISPONIBILE" else (aggiungi_giorni_lavorativi(oggi_dt, 10) if (st_trs+st_rwi+st_acq+st_tmp+st_lan) > 0 else aggiungi_giorni_lavorativi(oggi_dt, 25))
            is_ritardo = (pd.notnull(req_date) and eta.date() > req_date.date())

            # FILTRO (FISSO 1.0.02)
            passa = False
            if filtro_label == "Mostra tutto": passa = True
            elif filtro_label == "Solo Disponibili" and cat_core == "DISPONIBILE": passa = True
            elif filtro_label == "In Lavorazione" and cat_core == "LAVORAZIONE": passa = True
            elif filtro_label == "In Ritardo" and is_ritardo: passa = True

            if passa:
                # Calcolo percentuale solo per la pillola
                pct = 10
                if cat_core == "DISPONIBILE": pct = 100
                elif st_trs > 0: pct = 75
                elif st_rwi > 0 or st_acq > 0: pct = 60
                elif st_tmp > 0 or st_lan > 0: pct = 50
                elif st_grz > 0: pct = 30
                
                css = "on-time-row" if (cat_core == "DISPONIBILE" and not is_ritardo) else ("prod-delay-row" if is_ritardo else "delay-row")
                nota = "Pronto" if cat_core == "DISPONIBILE" else "In Lavorazione"
                righe_mostra.append({'css': css, 'date': req_date, 'qta': qta, 'eta': eta, 'nota': nota, 'pct': pct})

        if righe_mostra:
            p = righe_mostra[0]['pct']
            # TITOLO ARTICOLO + PILLOLA
            st.markdown(f"""
                <div class="custom-header">
                    <span>📦 <b>{art}</b> — <small>{desc}</small></span>
                    <div class="pill-bg">
                        <div class="pill-fill" style="width: {p}%;"></div>
                        <div class="pill-text">{p}%</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            
            # Expander pulito senza scritte
            with st.expander(""):
                for r in righe_mostra:
                    st.markdown(f'<div class="status-row {r["css"]}"><span><b>Consegna:</b> {r["date"].strftime("%d/%m/%Y") if pd.notnull(r["date"]) else "N.D."} | <b>Q.tà:</b> {r["qta"]:,.0f}</span><span><b>Stima:</b> {r["eta"].strftime("%d/%m/%Y")} ({r["nota"]})</span></div>', unsafe_allow_html=True)
            st.markdown("---")
