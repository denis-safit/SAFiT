import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE PAGINA ---
APP_VERSION = "1.0.06"
st.set_page_config(page_title=f"Safit - Portale Avanzamento {APP_VERSION}", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #fcfcfc; }
    .status-row {
        display: flex; justify-content: space-between; align-items: center;
        padding: 12px 15px; border-radius: 8px; margin-bottom: 8px; font-size: 14px;
    }
    .delay-row { background-color: #fff8e1; border-left: 6px solid #ffc107; color: #5d4037; }
    .on-time-row { background-color: #f1f8e9; border-left: 6px solid #4caf50; color: #1b5e20; }
    .client-delay-row { background-color: #e3f2fd; border-left: 6px solid #2196f3; color: #0d47a1; }
    .version-tag { font-size: 10px; color: #999; text-align: right; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNZIONI TECNICHE ---
def pulisci_numero(serie):
    return pd.to_numeric(serie.astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').fillna(0)

def standardizza_codice(valore):
    return str(valore).strip().upper()

@st.cache_data
def load_data():
    try:
        # A. CARICAMENTO ARCA
        df = pd.read_excel('righe_Ordini_ARCA.xlsx', sheet_name='Foglio1', skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        for col in ['Cliente Fornitore CD', 'Articolo C', 'Articolo D', 'Data']:
            if col in df.columns: df[col] = df[col].ffill()
        
        df = df.dropna(subset=['Articolo C'])
        df['Articolo C'] = df['Articolo C'].apply(standardizza_codice)
        
        df['Data_Consegna'] = pd.to_datetime(df['Data'], errors='coerce')
        col_res = 'Qta Residua' if 'Qta Residua' in df.columns else 'Qta Doc'
        df['Qta_Effettiva'] = pd.to_numeric(df[col_res], errors='coerce').fillna(0)
        df = df[df['Qta_Effettiva'] > 0]
        
        # B. CARICAMENTO ACCESS
        file_tech = 'Avanzamento_access.xlsx'
        if os.path.exists(file_tech):
            # Carichiamo saltando 1 riga (come visto negli screenshot)
            df_tech = pd.read_excel(file_tech, skiprows=1)
            df_tech.columns = [str(c).strip() for c in df_tech.columns]
            
            # --- RICERCA INTELLIGENTE COLONNE ---
            # Cerchiamo la colonna del Codice
            col_cod = next((c for c in df_tech.columns if c.upper() in ['CODICE', 'ARTICOLO', 'COD.ART']), df_tech.columns[0])
            
            # Cerchiamo la colonna della Giacenza (Gia, Gia., Giacenza, ecc)
            col_gia = next((c for c in df_tech.columns if c.upper().startswith('GIA')), None)
            
            if col_gia:
                df_tech['Art_Key_Match'] = df_tech[col_cod].apply(standardizza_codice)
                df_tech['Gia_Pulita'] = pulisci_numero(df_tech[col_gia])
                
                # Altre colonne produzione (Maiuscole/Minuscole)
                for c in ['Acq', 'Lan', 'GRZ', 'TMP', 'RWI', 'TRS']:
                    # Cerchiamo la colonna ignorando il maiuscolo
                    col_found = next((col for col in df_tech.columns if col.upper() == c.upper()), None)
                    if col_found:
                        df_tech[c] = pulisci_numero(df_tech[col_found])
                    else:
                        df_tech[c] = 0
                
                df_tech['Lavorazione_Totale'] = df_tech[['Acq', 'Lan', 'GRZ', 'TMP', 'RWI', 'TRS']].sum(axis=1)
                
                # MERGE
                df = pd.merge(df, df_tech[['Art_Key_Match', 'Gia_Pulita', 'Lavorazione_Totale']], 
                              left_on='Articolo C', right_on='Art_Key_Match', how='left')
                
                df['Gia'] = df['Gia_Pulita'].fillna(0)
                df['Lavorazione_Totale'] = df['Lavorazione_Totale'].fillna(0)
            else:
                st.error("❌ Non trovo nessuna colonna che inizia per 'Gia' nel file Access!")
        
        return df
    except Exception as e:
        st.error(f"Errore caricamento: {e}")
        return pd.DataFrame()

# --- 3. LOGICA ACCESSO ---
if "authenticated" not in st.session_state: st.session_state["authenticated"] = False
if not st.session_state["authenticated"]:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("Accesso Safit")
        u = st.text_input("Username")
        p = st.text_input("Password", type="password")
        if st.button("Entra"):
            st.session_state["authenticated"] = True
            st.rerun()
    st.stop()

# --- 4. DASHBOARD ---
data = load_data()
oggi_dt = datetime.now()

with st.sidebar:
    st.image('Logo SAFIT.JPG', use_container_width=True) if os.path.exists('Logo SAFIT.JPG') else st.title("SAFIT")
    if not data.empty:
        c_list = sorted(data['Cliente Fornitore CD'].unique().astype(str))
        sel_cli = st.selectbox("Seleziona Cliente:", c_list)
    if st.button("Esci"):
        st.session_state["authenticated"] = False
        st.rerun()

if not data.empty:
    df_cli = data[data['Cliente Fornitore CD'] == sel_cli].copy()
    
    st.title(f"Ordini Cliente: {sel_cli}")
    
    # BOX DI VERIFICA
    with st.expander("🔍 VERIFICA GIACENZE CARICATE (Denis, controlla qui)"):
        st.write("Dati estratti dal file Access per questo cliente:")
        st.dataframe(df_cli[['Articolo C', 'Gia', 'Lavorazione_Totale']].drop_duplicates())

    articoli = sorted(df_cli['Articolo C'].unique())
    for art in articoli:
        df_art = df_cli[df_cli['Articolo C'] == art].sort_values('Data_Consegna')
        
        st_gia = float(df_art['Gia'].iloc[0])
        st_lav = float(df_art['Lavorazione_Totale'].iloc[0])
        desc = df_art['Articolo D'].iloc[0]

        with st.expander(f"📦 {art} — {desc} (Disp. Magazzino: {st_gia:,.0f})"):
            for _, row in df_art.iterrows():
                qta = float(row['Qta_Effettiva'])
                req_date = row['Data_Consegna']
                
                # LOGICA DISPONIBILITÀ
                if st_gia >= qta:
                    st_gia -= qta
                    eta, nota, css = oggi_dt, "Disponibile (Pronto)", "on-time-row"
                elif (st_gia + st_lav) >= qta:
                    st_gia = 0
                    eta, nota, css = oggi_dt + timedelta(days=10), "In Lavorazione", "delay-row"
                else:
                    eta, nota, css = oggi_dt + timedelta(days=25), "Pianificare Produzione", "delay-row"
                
                # Controllo Ritardo Cliente
                if eta.date() > req_date.date() and css == "on-time-row":
                    css = "client-delay-row"

                st.markdown(f"""
                    <div class="status-row {css}">
                        <span>📅 Consegna: {req_date.strftime('%d/%m/%Y')} | <b>Q.tà: {qta:,.0f}</b></span>
                        <span>📦 {nota} (Previsto: {eta.strftime('%d/%m/%Y')})</span>
                    </div>
                """, unsafe_allow_html=True)
else:
    st.warning("Nessun dato caricato. Controlla i file Excel nella cartella.")
