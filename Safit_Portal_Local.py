import streamlit as st
import pandas as pd
import os

# Configurazione Pagina
st.set_page_config(page_title="SAFIT - Portale Avanzamento", layout="wide")

# Funzione per pulire i numeri
def pulisci_numero(serie):
    return pd.to_numeric(serie, errors='coerce').fillna(0)

@st.cache_data
def load_data():
    try:
        # Identificazione file (ignora maiuscole/minuscole)
        files = os.listdir('.')
        file_arca = next((f for f in files if f.lower() == 'righe_ordini_arca.xlsx'), None)
        file_access = next((f for f in files if f.lower() == 'avanzamento_access.xlsx'), None)

        if not file_arca:
            return None, "File 'righe_Ordini_ARCA.xlsx' non trovato su GitHub."

        # Caricamento dati ARCA
        df = pd.read_excel(file_arca, sheet_name='Foglio1', skiprows=2, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns] # Pulisce spazi nei nomi colonne
        
        # Trascinamento dati (ffill) per celle vuote
        cols_to_fill = ['Cliente Fornitore CD', 'Articolo C', 'Articolo D', 'Data', 'Documento']
        for col in cols_to_fill:
            if col in df.columns:
                df[col] = df[col].ffill()

        # Caricamento dati Access (opzionale)
        if file_access:
            df_tech = pd.read_excel(file_access, skiprows=1, engine='openpyxl')
            df_tech.columns = [str(c).strip() for c in df_tech.columns]
            if 'Codice' in df_tech.columns:
                df_tech = df_tech.rename(columns={'Codice': 'Art_Key'})
                df_tech['Art_Key'] = df_tech['Art_Key'].astype(str).str.strip()
                df = pd.merge(df, df_tech[['Art_Key', 'Acq', 'Gia']], left_on='Articolo C', right_on='Art_Key', how='left')

        return df, None
    except Exception as e:
        return None, f"Errore nel caricamento: {str(e)}"

# --- INTERFACCIA LOGIN ---
if 'auth' not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.image("Logo SAFIT.JPG", width=200)
    user = st.text_input("Utente").lower().strip()
    if st.button("Entra"):
        # Carichiamo gli utenti dal file utenti.xlsx
        try:
            df_utenti = pd.read_excel("utenti.xlsx")
            df_utenti.columns = [str(c).strip() for c in df_utenti.columns]
            if user in df_utenti['username'].values or user == 'safit_admin':
                st.session_state.auth = True
                st.session_state.user = user
                st.rerun()
            else:
                st.error("Utente non autorizzato")
        except:
            st.error("Errore database utenti")
else:
    # --- LOGICA PORTALE ---
    data, error = load_data()
    
    if error:
        st.error(error)
    else:
        st.sidebar.write(f"Connesso come: **{st.session_state.user}**")
        
        # Filtro Cliente
        if st.session_state.user == 'safit_admin':
            clienti_disponibili = sorted(data['Cliente Fornitore CD'].dropna().unique())
            scelta_cliente = st.sidebar.selectbox("Seleziona Cliente", clienti_disponibili)
        else:
            scelta_cliente = st.session_state.user

        # Filtraggio dati finale
        df_filtrato = data[data['Cliente Fornitore CD'].str.lower() == scelta_cliente.lower()]
        
        st.title(f"Ordini in corso: {scelta_cliente.upper()}")
        st.dataframe(df_filtrato, use_container_width=True)

    if st.sidebar.button("Logout"):
        st.session_state.auth = False
        st.rerun()
