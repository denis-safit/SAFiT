import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE ---
APP_VERSION = "1.0.08"
st.set_page_config(page_title=f"Safit Check - {APP_VERSION}", layout="wide")

st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 10px; border-radius: 5px; margin-bottom: 5px; }
    .delay-row { background-color: #fff8e1; border-left: 5px solid #ffc107; }
    .on-time-row { background-color: #f1f8e9; border-left: 5px solid #4caf50; }
    .client-delay-row { background-color: #e3f2fd; border-left: 5px solid #2196f3; }
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
        # A. ARCA
        df = pd.read_excel('righe_Ordini_ARCA.xlsx', sheet_name='Foglio1', skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        for col in ['Cliente Fornitore CD', 'Articolo C', 'Articolo D', 'Data']:
            if col in df.columns: df[col] = df[col].ffill()
        df = df.dropna(subset=['Articolo C'])
        df['Articolo C'] = df['Articolo C'].apply(standardizza_codice)
        df['Data_Consegna'] = pd.to_datetime(df['Data'], errors='coerce')
        df['Qta_Effettiva'] = pd.to_numeric(df['Qta Residua'], errors='coerce').fillna(pd.to_numeric(df['Qta Doc'], errors='coerce')).fillna(0)
        
        # B. ACCESS
        file_tech = 'Avanzamento_access.xlsx'
        if os.path.exists(file_tech):
            df_tech = pd.read_excel(file_tech, skiprows=1)
            df_tech.columns = [str(c).strip() for c in df_tech.columns]
            
            # Identificazione dinamica
            col_cod = next((c for c in df_tech.columns if 'CODICE' in c.upper()), df_tech.columns[0])
            col_gia = next((c for c in df_tech.columns if 'GIA' in c.upper()), df_tech.columns[1])
            
            df_tech['Art_Key_Match'] = df_tech[col_cod].apply(standardizza_codice)
            df_tech['Gia_Val'] = pulisci_numero(df_tech[col_gia])
            
            for c in ['Acq', 'Lan', 'GRZ', 'TMP', 'RWI', 'TRS']:
                col_f = next((col for col in df_tech.columns if col.upper() == c.upper()), None)
                df_tech[c] = pulisci_numero(df_tech[col_f]) if col_f else 0
            
            df_tech['Lav_Tot'] = df_tech[['Acq', 'Lan', 'GRZ', 'TMP', 'RWI', 'TRS']].sum(axis=1)
            
            df = pd.merge(df, df_tech[['Art_Key_Match', 'Gia_Val', 'Lav_Tot']], 
                          left_on='Articolo C', right_on='Art_Key_Match', how='left')
            df['Gia'] = df['Gia_Val'].fillna(0)
            df['Lavorazione_Totale'] = df['Lav_Tot'].fillna(0)
            
        return df
    except Exception as e:
        st.error(f"Errore caricamento: {e}")
        return pd.DataFrame()

# --- 3. LOGIN ---
if "authenticated" not in st.session_state: st.session_state["authenticated"] = False
if not st.session_state["authenticated"]:
    st.title("Safit Login")
    u = st.text_input("Username")
    p = st.text_input("Password", type="password")
    if st.button("Entra"):
        st.session_state["authenticated"] = True
        st.rerun()
    st.stop()

# --- 4. PANNELLO DI CONTROLLO FILE (Solo per te) ---
st.title("🛠 Pannello di Controllo Dati")
file_access = 'Avanzamento_access.xlsx'

if os.path.exists(file_access):
    m_time = os.path.getmtime(file_access)
    dt_m = datetime.fromtimestamp(m_time)
    st.info(f"📁 **File Access rilevato!** Ultima modifica: `{dt_m.strftime('%d/%m/%Y %H:%M:%S')}`")
    
    # Mostriamo cosa c'è dentro veramente
    with st.expander("👀 Anteprima grezza del file Access (Clicca per aprire)"):
        test_df = pd.read_excel(file_access, nrows=5)
        st.write("Prime 5 righe trovate nel file:")
        st.dataframe(test_df)
else:
    st.error("❌ IL FILE 'Avanzamento_access.xlsx' NON ESISTE NELLA CARTELLA!")

st.divider()

# --- 5. DASHBOARD ---
data = load_data()
if not data.empty:
    with st.sidebar:
        c_list = sorted(data['Cliente Fornitore CD'].unique().astype(str))
        sel_cli = st.selectbox("Seleziona Cliente:", c_list)
        if st.button("Esci"):
            st.session_state["authenticated"] = False
            st.rerun()

    df_cli = data[data['Cliente Fornitore CD'] == sel_cli].copy()
    
    # Calcolo match
    match_ok = df_cli[df_cli['Art_Key_Match'].notna()].shape[0]
    st.success(f"✅ Trovata giacenza per {match_ok} righe su {len(df_cli)} totali del cliente {sel_cli}")

    articoli = sorted(df_cli['Articolo C'].unique())
    for art in articoli:
        df_art = df_cli[df_cli['Articolo C'] == art].sort_values('Data_Consegna')
        st_gia = float(df_art['Gia'].iloc[0])
        st_lav = float(df_art['Lavorazione_Totale'].iloc[0])
        
        with st.expander(f"📦 {art} (Giacenza: {st_gia:,.0f})"):
            for _, row in df_art.iterrows():
                qta = float(row['Qta_Effettiva'])
                # Logica semplificata per test
                if st_gia >= qta:
                    st_gia -= qta
                    st.info(f"Pronto - Consegna: {row['Data_Consegna'].strftime('%d/%m/%Y')} | Qta: {qta}")
                else:
                    st.warning(f"Da produrre - Consegna: {row['Data_Consegna'].strftime('%d/%m/%Y')} | Qta: {qta}")
