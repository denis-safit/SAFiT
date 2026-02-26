import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE ---
APP_VERSION = "1.1.0"
st.set_page_config(page_title=f"Safit Avanzamento - {APP_VERSION}", layout="wide")

# CSS personalizzato per i colori Safit
st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 12px; border-radius: 8px; margin-bottom: 8px; }
    .on-time-row { background-color: #f1f8e9; border-left: 6px solid #4caf50; color: #1b5e20; } /* VERDE: PRONTO */
    .acq-row { background-color: #e3f2fd; border-left: 6px solid #2196f3; color: #0d47a1; }    /* BLU: IN ARRIVO ACQUISTO */
    .prod-row { background-color: #fff8e1; border-left: 6px solid #ffc107; color: #5d4037; }   /* GIALLO: IN PRODUZIONE */
    .urgent-row { background-color: #ffebee; border-left: 6px solid #f44336; color: #b71c1c; } /* ROSSO: MANCANTE TOTALE */
    </style>
    """, unsafe_allow_html=True)

def pulisci_numero(serie):
    return pd.to_numeric(serie.astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').fillna(0)

@st.cache_data
def load_data():
    try:
        # A. ARCA
        df_arca = pd.read_excel('righe_Ordini_ARCA.xlsx', sheet_name='Foglio1', skiprows=2)
        df_arca.columns = [str(c).strip() for c in df_arca.columns]
        for col in ['Cliente Fornitore CD', 'Articolo C', 'Articolo D', 'Data']:
            if col in df_arca.columns: df_arca[col] = df_arca[col].ffill()
        df_arca = df_arca.dropna(subset=['Articolo C'])
        df_arca['Articolo C'] = df_arca['Articolo C'].astype(str).str.strip().str.upper()
        df_arca['Data_Consegna'] = pd.to_datetime(df_arca['Data'], errors='coerce')
        q_col = 'Qta Residua' if 'Qta Residua' in df_arca.columns else 'Qta Doc'
        df_arca['Qta_Effettiva'] = pd.to_numeric(df_arca[q_col], errors='coerce').fillna(0)
        df_arca = df_arca[df_arca['Qta_Effettiva'] > 0]

        # B. ACCESS
        file_access = 'Avanzamento_access.xlsx'
        if os.path.exists(file_access):
            # Carichiamo saltando la prima riga se necessario
            df_access = pd.read_excel(file_access, skiprows=1)
            df_access.columns = [str(c).strip() for c in df_access.columns]
            
            # Mappatura secondo le tue istruzioni
            df_access['Key'] = df_access['Codice'].astype(str).str.strip().str.upper()
            df_access['Magazzino'] = pulisci_numero(df_access['GIA'])
            df_access['Acquisti'] = pulisci_numero(df_access['ACQ'])
            # Produzione = Somma di LAN, GRZ, TMP, RWI, TRS
            df_access['Produzione'] = df_access[['LAN', 'GRZ', 'TMP', 'RWI', 'TRS']].apply(pulisci_numero).sum(axis=1)

            # Unione
            df = pd.merge(df_arca, df_access[['Key', 'Magazzino', 'Acquisti', 'Produzione']], 
                          left_on='Articolo C', right_on='Key', how='left')
            
            # Riempiamo i vuoti per gli articoli non trovati in Access
            for c in ['Magazzino', 'Acquisti', 'Produzione']:
                df[c] = df[c].fillna(0)
            return df
        return df_arca
    except Exception as e:
        st.error(f"Errore caricamento: {e}")
        return pd.DataFrame()

# --- INTERFACCIA ---
if "authenticated" not in st.session_state: st.session_state["authenticated"] = False
if not st.session_state["authenticated"]:
    st.title("Safit Login")
    if st.button("Entra"): st.session_state["authenticated"] = True; st.rerun()
    st.stop()

data = load_data()
if not data.empty:
    with st.sidebar:
        st.image('Logo SAFIT.JPG', width=200) if os.path.exists('Logo SAFIT.JPG') else None
        clienti = sorted(data['Cliente Fornitore CD'].unique().astype(str))
        sel_cli = st.selectbox("Seleziona Cliente", clienti)
        if st.button("Esci"): st.session_state["authenticated"] = False; st.rerun()

    st.title(f"Piano Consegne: {sel_cli}")
    df_cli = data[data['Cliente Fornitore CD'] == sel_cli].copy()
    oggi = datetime.now()

    for art in sorted(df_cli['Articolo C'].unique()):
        df_art = df_cli[df_cli['Articolo C'] == art].sort_values('Data_Consegna')
        
        # Scorte iniziali
        maga = float(df_art['Magazzino'].iloc[0])
        acq = float(df_art['Acquisti'].iloc[0])
        prod = float(df_art['Produzione'].iloc[0])
        descr = df_art['Articolo D'].iloc[0]

        with st.expander(f"📦 {art} - {descr} (Giacenza: {maga:,.0f})"):
            for _, row in df_art.iterrows():
                qta = float(row['Qta_Effettiva'])
                dt_cons = row['Data_Consegna']
                
                # LOGICA DI ASSEGNAZIONE
                if maga >= qta:
                    maga -= qta
                    status, nota, css = dt_cons, "PRONTO A MAGAZZINO", "on-time-row"
                elif (maga + acq) >= qta:
                    # Se usiamo gli acquisti, ipotizziamo 10gg per l'arrivo
                    acq -= (qta - maga); maga = 0
                    status, nota, css = oggi + timedelta(days=10), "IN ARRIVO (ACQUISTI)", "acq-row"
                elif (maga + acq + prod) >= qta:
                    # Se usiamo la produzione, ipotizziamo 20gg
                    prod -= (qta - maga - acq); maga = 0; acq = 0
                    status, nota, css = oggi + timedelta(days=20), "IN PRODUZIONE", "prod-row"
                else:
                    # Nulla di disponibile o lanciato
                    status, nota, css = oggi + timedelta(days=35), "DA LANCIARE / MANCANTE", "urgent-row"

                st.markdown(f"""
                    <div class="status-row {css}">
                        <span><b>{dt_cons.strftime('%d/%m/%Y')}</b> | Q.tà: {qta:,.0f}</span>
                        <span>{nota} (Stima: {status.strftime('%d/%m/%Y')})</span>
                    </div>
                """, unsafe_allow_html=True)
