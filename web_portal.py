import streamlit as st
import pandas as pd
import streamlit_authenticator as stauth

# --- CONFIGURAZIONE PAGINA ---
st.set_page_config(page_title="Safit - Tracking Produzione", layout="wide")

# --- SISTEMA DI LOGIN (Semplificato) ---
names = ['Cliente Annic', 'Denis Venezian']
usernames = ['annic', 'admin']
passwords = ['annic2024', 'safit2024'] # In produzione usa hash!

authenticator = stauth.Authenticate(names, usernames, passwords, 'safit_cookie', 'safit_key', cookie_expiry_days=30)
name, authentication_status, username = authenticator.login('Login', 'main')

if authentication_status:
    authenticator.logout('Logout', 'sidebar')
    st.title(f"Benvenuto nel portale Safit, {name}")

    # --- CARICAMENTO DATI ---
    @st.cache_data
    def load_data():
        # Qui uniamo i due file che abbiamo preparato
        df_arca = pd.read_excel('righe_Ordini_ARCA.xlsx', skiprows=1).ffill()
        # Supponiamo di avere l'export di Access chiamato 'Avanzamento_Produzione.xlsx'
        df_access = pd.read_excel('Avanzamento_Produzione.xlsx').ffill()
        
        # Join dei dati
        return pd.merge(df_arca, df_access, on='Commessa', how='left')

    df = load_data()

    # --- FILTRO CLIENTE ---
    # Se è un cliente, vede solo i suoi dati. Se è admin, vede tutto.
    if username != 'admin':
        # Filtra per una colonna 'CodiceCliente' o 'RagioneSociale'
        df = df[df['Cliente'].str.contains(name.split()[1], case=False)]

    # --- INTERFACCIA TRACKING (Stile "Pacco Amazon") ---
    st.subheader("Stato Avanzamento Ordini")
    
    for index, row in df.iterrows():
        with st.expander(f"📦 Ordine: {row['Commessa']} - Articolo: {row['Articolo']}"):
            col1, col2, col3 = st.columns(3)
            col1.metric("Fase Attuale", row['Fase'])
            col2.metric("Stato", row['Stato'])
            col3.metric("Consegna Prevista", str(row['Data']))
            
            # Barra di progresso dinamica basata sulla fase
            fasi = ['Taglio', 'Piegatura', 'Saldatura', 'Verniciatura', 'Imballo']
            try:
                progresso = (fasi.index(row['Fase']) + 1) / len(fasi)
                st.progress(progresso)
            except:
                st.progress(0)

elif authentication_status == False:
    st.error('Username o Password errati')
elif authentication_status == None:
    st.warning('Inserire username e password')
