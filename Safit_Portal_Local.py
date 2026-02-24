import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE PAGINA E STILE SAFIT ---
st.set_page_config(page_title="Safit - Portale Avanzamento", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #fcfcfc; }
    .stApp { margin-top: -30px; }
    .delay-row { background-color: #fff8e1; border-left: 6px solid #ffc107; padding: 15px; border-radius: 8px; margin-bottom: 12px; }
    .on-time-row { background-color: #f1f8e9; border-left: 6px solid #4caf50; padding: 15px; border-radius: 8px; margin-bottom: 12px; }
    .client-delay-row { background-color: #e3f2fd; border-left: 6px solid #2196f3; padding: 15px; border-radius: 8px; margin-bottom: 12px; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. GESTIONE UTENTI (LETTURA EXCEL) ---
@st.cache_data
def load_users():
    # Cerca il file utenti.xlsx nella cartella corrente
    file_u = 'utenti.xlsx'
    if os.path.exists(file_u):
        try:
            df_u = pd.read_excel(file_u)
            df_u.columns = [str(c).strip() for c in df_u.columns]
            # Converte in dizionario: {username: [password, cliente_arca]}
            return df_u.set_index('username')[['password', 'cliente_arca']].T.to_dict('list')
        except Exception as e:
            st.error(f"Errore caricamento utenti: {e}")
            return {'safit_admin': ['admin2026', 'TUTTI']}
    else:
        # Fallback se il file manca durante lo sviluppo
        return {'safit_admin': ['admin2026', 'TUTTI']}

USER_DB = load_users()

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    
    if not st.session_state["authenticated"]:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if os.path.exists('Logo SAFIT.JPG'):
                st.image('Logo SAFIT.JPG', width=300)
            st.title("Accesso Area Riservata")
            user = st.text_input("Username")
            pw = st.text_input("Password", type="password")
            if st.button("Accedi"):
                # Controllo credenziali dal file Excel
                if user in USER_DB and str(USER_DB[user][0]) == pw:
                    st.session_state["authenticated"] = True
                    st.session_state["user_type"] = USER_DB[user][1]
                    st.session_state["username"] = user
                    st.rerun()
                else:
                    st.error("Username o Password errati")
        return False
    return True

if not check_password():
    st.stop()

# --- 3. FUNZIONI TECNICHE (CALCOLI E PULIZIA) ---
def aggiungi_giorni_lavorativi(data_inizio, giorni):
    data_corrente = data_inizio
    while giorni > 0:
        data_corrente += timedelta(days=1)
        if data_corrente.weekday() < 5: 
            giorni -= 1
    return data_corrente

def pulisci_numero(serie):
    """Trasforma stringhe come '2.501,00' in numeri decimali"""
    return pd.to_numeric(
        serie.astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False), 
        errors='coerce'
    ).fillna(0)

@st.cache_data
def load_data():
st.write(f"DEBUG: Clienti trovati nel file: {df['Cliente Fornitore CD'].unique()}")
    try:
        # Cerchiamo il file ignorando maiuscole/minuscole nella cartella
        files = os.listdir('.')
        file_arca = next((f for f in files if f.lower() == 'righe_ordini_arca.xlsx'), None)
        file_access = next((f for f in files if f.lower() == 'avanzamento_access.xlsx'), None)

        if not file_arca:
            st.error(f"❌ File degli ordini non trovato! File presenti: {files}")
            return pd.DataFrame()

        # Caricamento ARCA
        df = pd.read_excel(file_arca, sheet_name='Foglio1', skiprows=2, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]
        
        # Riempimento celle vuote per trascinamento (ffill)
        for col in ['Cliente Fornitore CD', 'Articolo C', 'Articolo D', 'Data', 'Documento']:
            if col in df.columns: df[col] = df[col].ffill()
            
        df['Data_Consegna'] = pd.to_datetime(df['Data'], errors='coerce')
        col_res = 'Qta Residua' if 'Qta Residua' in df.columns else 'Qta Doc'
        df['Qta_Effettiva'] = pd.to_numeric(df[col_res], errors='coerce').fillna(0)
        df = df[df['Qta_Effettiva'] > 0]
        
        # Caricamento Access
        if file_access:
            df_tech = pd.read_excel(file_access, skiprows=1, engine='openpyxl') 
            df_tech.columns = [str(c).strip() for c in df_tech.columns]
            if 'Codice' in df_tech.columns:
                df_tech = df_tech.rename(columns={'Codice': 'Art_Key'})
                df_tech['Art_Key'] = df_tech['Art_Key'].astype(str).str.strip()
                df_tech['Gia'] = pulisci_numero(df_tech['Gia'])
                df_tech['Acq'] = pulisci_numero(df_tech['Acq'])
                df = pd.merge(df, df_tech[['Art_Key', 'Acq', 'Gia']], left_on='Articolo C', right_on='Art_Key', how='left')
        return df
    except Exception as e:
        st.error(f"Errore tecnico nel caricamento: {e}")
        return pd.DataFrame()

data = load_data()
oggi_dt = datetime.now()

# --- 4. SIDEBAR E FILTRI ---
with st.sidebar:
    if os.path.exists('Logo SAFIT.JPG'):
        st.image('Logo SAFIT.JPG', use_container_width=True)
    
    st.write(f"Utente: **{st.session_state['username']}**")
    
    # Se ADMIN, può scegliere il cliente. Se CLIENTE, è bloccato
    if st.session_state["user_type"] == "TUTTI":
        clienti_list = sorted([str(x) for x in data['Cliente Fornitore CD'].unique()])
        sel_cli = st.selectbox("👤 Seleziona Cliente:", clienti_list)
    else:
        sel_cli = st.session_state["user_type"]
        st.info(f"Vista riservata cliente:\n{sel_cli}")
    
    filtro_label = st.radio("Stato ordini:", ["Mostra tutto", "Solo Disponibili", "In Lavorazione", "In Ritardo"])
    
    if st.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

# --- 5. INTERFACCIA CENTRALE ---
st.title("🚜 Portale Avanzamento Produzione")
df_cli = data[data['Cliente Fornitore CD'] == sel_cli].copy()

if not df_cli.empty:
    articoli_dis = sorted([str(x) for x in df_cli['Articolo C'].unique()])
    sel_art = st.selectbox("🔍 Cerca Codice Prodotto:", ["Tutti i prodotti"] + articoli_dis)
    articoli_view = articoli_dis if sel_art == "Tutti i prodotti" else [sel_art]

    for art in articoli_view:
        df_art = df_cli[df_cli['Articolo C'] == art].sort_values('Data_Consegna')
        desc = df_art['Articolo D'].iloc[0] if 'Articolo D' in df_art.columns else ""
        
        # Inizializza stock per calcolo ACQ
        st_gia = float(df_art['Gia'].iloc[0]) if 'Gia' in df_art.columns and pd.notnull(df_art['Gia'].iloc[0]) else 0.0
        st_acq = float(df_art['Acq'].iloc[0]) if 'Acq' in df_art.columns and pd.notnull(df_art['Acq'].iloc[0]) else 0.0

        righe_mostra = []
        for _, row in df_art.iterrows():
            qta = float(row['Qta_Effettiva'])
            req_date = row['Data_Consegna']
            
            # Logica di calcolo ACQ ed erosione stock
            if st_gia >= qta:
                st_gia -= qta
                eta, nota, cat = oggi_dt, "Pronto", "Solo Disponibili"
                css = "on-time-row"
            elif (st_gia + st_acq) >= qta:
                rimanente = qta - st_gia
                st_gia, st_acq = 0, st_acq - rimanente
                eta, nota, cat = aggiungi_giorni_lavorativi(oggi_dt, 10), "In Lavorazione", "In Lavorazione"
                css = "on-time-row"
            else:
                eta, nota, cat = aggiungi_giorni_lavorativi(oggi_dt, 25), "Nuova Produzione", "In Lavorazione"
                css = "delay-row"

            # Gestione ritardi rispetto alla data richiesta
            if pd.notnull(req_date) and eta.date() > req_date.date():
                if cat != "Solo Disponibili":
                    cat = "In Ritardo"
                    css = "delay-row"
                else:
                    css = "client-delay-row"

            if filtro_label == "Mostra tutto" or filtro_label == cat:
                righe_mostra.append({'css': css, 'date': req_date, 'qta': qta, 'eta': eta, 'nota': nota})

        if righe_mostra:
            with st.expander(f"📦 {art} — {desc} | Residuo: {df_art['Qta_Effettiva'].sum():,.0f}"):
                for r in righe_mostra:
                    st.markdown(f'<div class="{r["css"]}"><b>Consegna: {r["date"].strftime("%d/%m/%Y") if pd.notnull(r["date"]) else "N.D."}</b> | Q.tà: {r["qta"]:,.0f} <span style="float:right;">Stima: {r["eta"].strftime("%d/%m/%Y")} ({r["nota"]})</span></div>', unsafe_allow_html=True)
else:
    st.warning("Nessun ordine in sospeso trovato per questo cliente.")


