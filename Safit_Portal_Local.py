import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="SAFIT Portal", layout="wide")

@st.cache_data
def load_data():
    try:
        files = os.listdir('.')
        target = 'righe_ordini_arca.xlsx'
        found_file = next((f for f in files if f.lower() == target), None)
        
        if not found_file:
            return None, f"File {target} non trovato su GitHub."

        df = pd.read_excel(found_file, sheet_name='Foglio1', skiprows=2, engine='openpyxl')
        df.columns = [str(c).strip() for c in df.columns]
        
        # Corregge il problema dei nomi clienti mancanti nelle righe successive
        if 'Cliente Fornitore CD' in df.columns:
            df['Cliente Fornitore CD'] = df['Cliente Fornitore CD'].ffill()
        
        return df, None
    except Exception as e:
        return None, str(e)

# --- SISTEMA DI ACCESSO SICURO ---
if 'auth' not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    st.image("Logo SAFIT.JPG", width=200)
    st.title("Accesso Area Riservata SAFIT")
    
    with st.form("login_form"):
        user = st.text_input("Username").lower().strip()
        password = st.text_input("Password", type="password")
        submit = st.form_submit_button("Accedi")
        
        if submit:
            try:
                # Legge il database utenti
                df_utenti = pd.read_excel("utenti.xlsx")
                df_utenti.columns = [str(c).strip() for c in df_utenti.columns]
                
                # Verifica credenziali
                match = df_utenti[(df_utenti['username'].astype(str).str.lower() == user) & 
                                  (df_utenti['password'].astype(str) == password)]
                
                if not match.empty:
                    st.session_state.auth = True
                    st.session_state.user = user
                    st.rerun()
                else:
                    st.error("Username o Password errati")
            except Exception as e:
                st.error(f"Errore nel database utenti: {e}")
else:
    # --- PORTALE ATTIVO ---
    data, err = load_data()
    if err:
        st.error(f"Errore caricamento dati: {err}")
    else:
        st.sidebar.write(f"Utente: **{st.session_state.user}**")
        
        # Logica Admin vs Cliente
        if st.session_state.user == 'safit_admin':
            clienti = sorted(data['Cliente Fornitore CD'].dropna().unique())
            user_choice = st.sidebar.selectbox("Seleziona Cliente", clienti)
        else:
            user_choice = st.session_state.user

        # Mostra i dati
        mask = data['Cliente Fornitore CD'].astype(str).str.lower() == user_choice.lower()
        df_final = data[mask]
        
        st.header(f"Prospetto Ordini: {user_choice.upper()}")
        st.dataframe(df_final, use_container_width=True)

    if st.sidebar.button("Logout"):
        st.session_state.auth = False
        st.rerun()
