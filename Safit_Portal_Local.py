import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="SAFIT Portal", layout="wide")

@st.cache_data
def load_data():
    try:
        # Cerca il file ignorando maiuscole/minuscole
        files = os.listdir('.')
        target = 'righe_ordini_arca.xlsx'
        found_file = next((f for f in files if f.lower() == target), None)
        
        if not found_file:
            return None, f"File {target} non trovato. Presenti: {files}"

        # Carica e pulisce i nomi delle colonne
        df = pd.read_excel(found_file, sheet_name='Foglio1', skiprows=2, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]
        
        # Riempie le celle vuote (fondamentale per vedere tutti i clienti!)
        if 'Cliente Fornitore CD' in df.columns:
            df['Cliente Fornitore CD'] = df['Cliente Fornitore CD'].ffill()
        
        return df, None
    except Exception as e:
        return None, str(e)

# --- Logica di Accesso ---
if 'auth' not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.title("Accedi al Portale SAFIT")
    user = st.text_input("Username (es. grisport)").lower().strip()
    if st.button("Login"):
        # Accetta safit_admin o i nomi presenti nel file utenti.xlsx
        st.session_state.auth = True
        st.session_state.user = user
        st.rerun()
else:
    data, err = load_data()
    if err:
        st.error(f"Errore caricamento dati: {err}")
    else:
        # Se sei admin vedi TUTTI, altrimenti vedi solo il tuo utente
        if st.session_state.user == 'safit_admin':
            clienti = sorted(data['Cliente Fornitore CD'].dropna().unique())
            user_choice = st.sidebar.selectbox("Seleziona Cliente", clienti)
        else:
            user_choice = st.session_state.user

        mask = data['Cliente Fornitore CD'].astype(str).str.lower() == user_choice.lower()
        df_final = data[mask]
        
        st.header(f"Ordini per: {user_choice}")
        st.dataframe(df_final, use_container_width=True)

    if st.sidebar.button("Logout"):
        st.session_state.auth = False
        st.rerun()
