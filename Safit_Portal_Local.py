import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE PAGINA E VERSIONE ---
APP_VERSION = "1.0.08"
st.set_page_config(page_title=f"Safit - Portale Avanzamento {APP_VERSION}", layout="wide")

st.markdown(f"""
    <style>
    .main {{ background-color: #fcfcfc; }}
    .stApp {{ margin-top: -30px; }}
    
    /* Pillola Ingrandita (Doppia dimensione) */
    .pill-container {{
        display: flex;
        align-items: center;
        gap: 10px;
    }}
    .pill-bg {{
        background-color: #ddd;
        border-radius: 12px;
        width: 160px; /* Raddoppiata */
        height: 24px; /* Ingrandita */
        border: 1px solid #bbb;
        position: relative;
        flex-shrink: 0;
    }}
    .pill-fill {{
        background-color: #4caf50;
        height: 100%;
        border-radius: 12px;
    }}
    .pill-text {{
        position: absolute;
        top: 0; left: 0; width: 100%;
        font-size: 14px; /* Più leggibile */
        font-weight: bold;
        line-height: 22px;
        text-align: center;
        color: #000;
    }}

    .status-row {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap; 
        padding: 10px 15px;
        border-radius: 8px;
        margin-bottom: 5px;
        gap: 10px;
        font-size: 13px;
    }}
    
    /* COLORI LOGICA SAFIT (BLINDATI v1.0.02) */
    .on-time-row {{ background-color: #f1f8e9; border-left: 6px solid #4caf50; color: #1b5e20; }}
    .client-delay-row {{ background-color: #e3f2fd; border-left: 6px solid #2196f3; color: #0d47a1; }}
    .delay-row {{ background-color: #fff8e1; border-left: 6px solid #ffc107; color: #5d4037; }}
    .prod-delay-row {{ background-color: #ffebee; border-left: 6px solid #f44336; color: #b71c1c; }}
    
    .version-tag {{ font-size: 10px; color: #999; text-align: right; }}
    </style>
    """, unsafe_allow_html=True)

# --- 2. GESTIONE UTENTI ---
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

# --- 3. FUNZIONI TECNICHE ---
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
                df = pd.merge(df, df_tech[['Art_Key', 'Gia', 'Acq', 'Lan', 'Grz', 'Tmp', 'Rwi', 'Trs']], left_on='Articolo C', right_on='Art_Key', how='left')
        return df
    except: return pd.DataFrame()

data = load_data()
oggi_dt = datetime.now()

# --- 4. SIDEBAR ---
with st.sidebar:
    if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
    st.write(f"Utente: **{st.session_state['username']}**")
    st.markdown(f'<p class="version-tag">Versione: {APP_VERSION}</p>', unsafe_allow_html=True)
    if st.session_state["user_type"] == "TUTTI":
        clienti_list = sorted([str(x) for x in data['Cliente Fornitore CD'].unique()])
        sel_cli = st.selectbox("👤 Seleziona Cliente:", clienti_list)
    else: sel_cli = st.session_state["user_type"]
    filtro_label = st.radio("Filtra per stato:", ["Mostra tutto", "Solo Disponibili", "In Lavorazione", "In Ritardo"])
    if st.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

# --- 5. LOGICA CENTRALE ---
st.title("🚜 Portale Avanzamento Produzione")
df_cli = data[data['Cliente Fornitore CD'] == sel_cli].copy()

if not df_cli.empty:
    articoli_dis = sorted([str(x) for x in df_cli['Articolo C'].unique()])
    sel_art = st.selectbox("🔍 Cerca Codice Prodotto:", ["Tutti i prodotti"] + articoli_dis)
    articoli_view = articoli_dis if sel_art == "Tutti i prodotti" else [sel_art]

    for art in articoli_view:
        df_art = df_cli[df_cli['Articolo C'] == art].sort_values('Data_Consegna')
        desc = df_art['Articolo D'].iloc[0] if 'Articolo D' in df_art.columns else ""
        row_t = df_art.iloc[0]
        st_gia, st_acq, st_lan, st_grz, st_tmp, st_rwi, st_trs = [float(row_t.get(k, 0)) for k in ['Gia', 'Acq', 'Lan', 'Grz', 'Tmp', 'Rwi', 'Trs']]
        
        righe_mostra = []
        for _, row in df_art.iterrows():
            qta, req_date = float(row['Qta_Effettiva']), row['Data_Consegna']
            settimane = ((req_date - oggi_dt).days / 7) if pd.notnull(req_date) else 99

            # LOGICA STEP (Percentuale)
            pct, cat_core = 10, "LAVORAZIONE"
            if st_gia >= qta: pct, cat_core = 100, "DISPONIBILE"; st_gia -= qta
            elif st_trs > 0: pct = 75
            elif st_rwi > 0 or st_acq > 0: pct = 60
            elif st_tmp > 0 or st_lan > 0: pct = 50
            elif st_grz > 0: pct = 30
            elif settimane <= 4: pct = 20
            
            # STIMA CONSEGNA E RITARDO
            eta = oggi_dt if pct == 100 else (aggiungi_giorni_lavorativi(oggi_dt, 10) if pct >= 50 else aggiungi_giorni_lavorativi(oggi_dt, 25))
            is_ritardo = (pd.notnull(req_date) and eta.date() > req_date.date())
            
            # FILTRO BLINDATO v1.0.02
            passa = False
            if filtro_label == "Mostra tutto": passa = True
            elif filtro_label == "Solo Disponibili" and cat_core == "DISPONIBILE": passa = True
            elif filtro_label == "In Lavorazione" and cat_core == "LAVORAZIONE": passa = True
            elif filtro_label == "In Ritardo" and is_ritardo: passa = True

            if passa:
                if cat_core == "DISPONIBILE":
                    css, nota = ("client-delay-row", "Ritardo Ritiro") if is_ritardo else ("on-time-row", "Pronto")
                else:
                    css, nota = ("prod-delay-row", "In Ritardo") if is_ritardo else ("delay-row", "In Lavorazione")
                righe_mostra.append({'css': css, 'date': req_date, 'qta': qta, 'eta': eta, 'nota': nota, 'pct': pct})

        if righe_mostra:
            current_pct = righe_mostra[0]['pct']
            # Header Expander: Codice + Descrizione + Pillola Grande
            header_label = f"📦 {art} — {desc} | {current_pct}%"
            
            with st.expander(header_label):
                # Visualizzazione della Pillola Grafica Grande come prima cosa dentro l'expander
                st.markdown(f'''
                    <div class="pill-container">
                        <span style="font-size: 14px; font-weight: bold;">Avanzamento:</span>
                        <div class="pill-bg">
                            <div class="pill-fill" style="width: {current_pct}%;"></div>
                            <div class="pill-text">{current_pct}%</div>
                        </div>
                    </div>
                    <hr style="margin: 10px 0;">
                ''', unsafe_allow_html=True)
                
                for r in righe_mostra:
                    st.markdown(f'''
                        <div class="status-row {r["css"]}">
                            <span><b>Consegna:</b> {r["date"].strftime("%d/%m/%Y") if pd.notnull(r["date"]) else "N.D."} | <b>Q.tà:</b> {r["qta"]:,.0f}</span>
                            <span><b>Stima:</b> {r["eta"].strftime("%d/%m/%Y")} ({r["nota"]})</span>
                        </div>
                    ''', unsafe_allow_html=True)
