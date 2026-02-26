import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE PAGINA E VERSIONE ---
APP_VERSION = "1.0.03"
st.set_page_config(page_title=f"Safit - Portale Avanzamento {APP_VERSION}", layout="wide")

st.markdown(f"""
    <style>
    .main {{ background-color: #fcfcfc; }}
    .stApp {{ margin-top: -30px; }}
    .status-row {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap; 
        padding: 12px 15px;
        border-radius: 8px;
        margin-bottom: 5px;
        gap: 10px;
        font-size: 14px;
    }}
    /* COLORI LOGICA SAFIT */
    .on-time-row {{ background-color: #f1f8e9; border-left: 6px solid #4caf50; color: #1b5e20; }}
    .client-delay-row {{ background-color: #e3f2fd; border-left: 6px solid #2196f3; color: #0d47a1; }}
    .delay-row {{ background-color: #fff8e1; border-left: 6px solid #ffc107; color: #5d4037; }}
    .prod-delay-row {{ background-color: #ffebee; border-left: 6px solid #f44336; color: #b71c1c; }}
    
    /* STILE PROGRESS BAR */
    .progress-container {{
        width: 100%;
        background-color: #e0e0e0;
        border-radius: 5px;
        margin-bottom: 15px;
        height: 12px;
        overflow: hidden;
    }}
    .progress-bar {{
        height: 100%;
        background-color: #4caf50;
        text-align: center;
        line-height: 12px;
        color: white;
        font-size: 10px;
        transition: width 0.5s;
    }}
    .step-text {{ font-size: 12px; color: #666; font-style: italic; margin-bottom: 5px; display: block; }}
    
    .stExpander div {{ height: auto !important; min-height: min-content !important; }}
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
            user = st.text_input("Username")
            pw = st.text_input("Password", type="password")
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
                # Mappatura campi Access come da tabella Denis
                for c in ['Gia', 'Acq', 'Lan', 'Grz', 'Tmp', 'Rwi', 'Trs']:
                    if c in df_tech.columns: df_tech[c] = pulisci_numero(df_tech[c])
                
                df = pd.merge(df, df_tech[['Art_Key', 'Gia', 'Acq', 'Lan', 'Grz', 'Tmp', 'Rwi', 'Trs']], 
                              left_on='Articolo C', right_on='Art_Key', how='left')
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
        
        # Dati tecnici per calcolo step
        row_t = df_art.iloc[0]
        st_gia, st_acq, st_lan, st_grz, st_tmp, st_rwi, st_trs = [float(row_t.get(k, 0)) for k in ['Gia', 'Acq', 'Lan', 'Grz', 'Tmp', 'Rwi', 'Trs']]
        
        righe_mostra = []
        for _, row in df_art.iterrows():
            qta, req_date = float(row['Qta_Effettiva']), row['Data_Consegna']
            
            # Calcolo Settimane Residue
            settimane = ((req_date - oggi_dt).days / 7) if pd.notnull(req_date) else 99

            # --- DETERMINAZIONE STEP (Logica Denis) ---
            pct, step_nome = 10, "Conferma Ordine Inviata" # Step 0 base
            
            if st_gia >= qta: pct, step_nome = 100, "Disponibile"; st_gia -= qta
            elif st_trs > 0: pct, step_nome = 75, "In preparazione (Transito)"
            elif st_rwi > 0: pct, step_nome = 60, "Lavorazione esterna fornitore"
            elif st_acq > 0: pct, step_nome = 60, "Merce in acquisto"
            elif st_tmp > 0 or st_lan > 0: pct, step_nome = 50, "Semilavorati disponibili"
            elif st_grz > 0: pct, step_nome = 30, "Grezzo disponibile"
            # Se nessuna giacenza tecnica, usiamo il tempo residuo
            elif settimane <= 1: pct, step_nome = 75, "In preparazione"
            elif settimane <= 4: pct, step_nome = 20, "Materia prima disponibile"
            
            # --- COLORAZIONE E STATO ---
            eta = oggi_dt if pct == 100 else (aggiungi_giorni_lavorativi(oggi_dt, 10) if pct >= 50 else aggiungi_giorni_lavorativi(oggi_dt, 25))
            is_rit = (pd.notnull(req_date) and eta.date() > req_date.date())
            
            if pct == 100:
                css, nota = ("client-delay-row", "Ritardo Ritiro") if is_rit else ("on-time-row", "Pronto")
            else:
                css, nota = ("prod-delay-row", f"{step_nome} (In Ritardo)") if is_rit else ("delay-row", step_nome)

            # Filtro
            passa = (filtro_label == "Mostra tutto") or (filtro_label == "Solo Disponibili" and pct == 100) or (filtro_label == "In Lavorazione" and pct < 100) or (filtro_label == "In Ritardo" and is_rit)
            
            if passa:
                righe_mostra.append({'css': css, 'date': req_date, 'qta': qta, 'eta': eta, 'nota': nota, 'pct': pct})

        if righe_mostra:
            with st.expander(f"📦 {art} — {desc}"):
                for r in righe_mostra:
                    st.markdown(f'<span class="step-text">{r["nota"]}</span>', unsafe_allow_html=True)
                    st.markdown(f'<div class="progress-container"><div class="progress-bar" style="width: {r["pct"]}%;">{r["pct"]}%</div></div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="status-row {r["css"]}"><span><b>Consegna:</b> {r["date"].strftime("%d/%m/%Y") if pd.notnull(r["date"]) else "N.D."} | <b>Q.tà:</b> {r["qta"]:,.0f}</span><span><b>Stima:</b> {r["eta"].strftime("%d/%m/%Y")}</span></div>', unsafe_allow_html=True)
