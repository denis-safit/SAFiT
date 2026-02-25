import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
import io

# --- 1. CONFIGURAZIONE PAGINA ---
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

def aggiungi_giorni_lavorativi(data_inizio, giorni):
    data_corrente = data_inizio
    while giorni > 0:
        data_corrente += timedelta(days=1)
        if data_corrente.weekday() < 5: 
            giorni -= 1
    return data_corrente

@st.cache_data
def load_data():
    try:
        # Caricamento Ordini ARCA
        df = pd.read_excel('righe_Ordini_ARCA.xlsx', sheet_name='Foglio1', skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.replace(to_replace=r'\(vuoto\)', value=pd.NA, regex=True)
        for col in ['Cliente Fornitore CD', 'Articolo C', 'Articolo D', 'Data', 'Documento']:
            if col in df.columns: df[col] = df[col].ffill()
        df['Data_Consegna'] = pd.to_datetime(df['Data'], errors='coerce')
        col_res = 'Qta Residua' if 'Qta Residua' in df.columns else 'Qta Doc'
        df['Qta_Effettiva'] = pd.to_numeric(df[col_res], errors='coerce').fillna(0)
        df = df[df['Qta_Effettiva'] > 0]
        
        # Caricamento Access ottimizzato per la tua tabella
        if os.path.exists('Avanzamento_access.xlsx'):
            # Saltiamo la riga 1 ("ONE_PRODUZIONE") e leggiamo le intestazioni alla riga 2
            df_tech = pd.read_excel('Avanzamento_access.xlsx', skiprows=1) 
            df_tech.columns = [str(c).strip() for c in df_tech.columns]
            
            # Verifichiamo la presenza delle colonne necessarie
            if 'Codice' in df_tech.columns:
                df_tech = df_tech.rename(columns={'Codice': 'Art_Key'})
                df_tech['Art_Key'] = df_tech['Art_Key'].astype(str).str.strip()
                df['Articolo C'] = df['Articolo C'].astype(str).str.strip()
                
                # Convertiamo Gia e Acq in numeri
                for c in ['Gia', 'Acq']:
                    if c in df_tech.columns: 
                        df_tech[c] = pd.to_numeric(df_tech[c], errors='coerce').fillna(0)
                
                # Unione dati
                df = pd.merge(df, df_tech[['Art_Key', 'Acq', 'Gia']], left_on='Articolo C', right_on='Art_Key', how='left')
            else:
                st.error("Errore: Colonna 'Codice' non trovata nella riga di intestazione del file Access.")
        return df
    except Exception as e:
        st.error(f"Errore tecnico caricamento: {e}")
        return pd.DataFrame()

data = load_data()

# --- 2. INTERFACCIA UTENTE ---
if not data.empty:
    with st.sidebar:
        # LOGO SAFIT UFFICIALE
        if os.path.exists('Logo SAFIT.JPG'):
            st.image('Logo SAFIT.JPG', use_container_width=True)
        else:
            st.title("SAFIT S.r.l.")
        
        st.divider()
        st.header("Filtri Avanzamento")
        clienti = sorted([str(x) for x in data['Cliente Fornitore CD'].unique()])
        sel_cli = st.selectbox("👤 Seleziona Cliente:", clienti)
        filtro_label = st.radio("Stato ordini:", ["Mostra tutto", "Solo Disponibili", "In Lavorazione", "In Ritardo"])
        st.divider()

    st.title("Portale Stato Avanzamento Produzione")
    st.caption(f"Aggiornamento live: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    df_cli = data[data['Cliente Fornitore CD'] == sel_cli].copy()
    articoli = sorted([str(x) for x in df_cli['Articolo C'].unique()])
    sel_art = st.selectbox("📦 Cerca Prodotto:", ["Tutti i prodotti"] + articoli)
    articoli_view = articoli if sel_art == "Tutti i prodotti" else [sel_art]
    
    oggi = datetime.now()
    report_dettagliato = [] 

    for art in articoli_view:
        df_art = df_cli[df_cli['Articolo C'] == art].sort_values('Data_Consegna')
        desc = df_art['Articolo D'].iloc[0] if 'Articolo D' in df_art.columns else ""
        
        # LOGICA QUANTITÀ
        stock_gia = float(df_art['Gia'].iloc[0]) if 'Gia' in df_art.columns and pd.notnull(df_art['Gia'].iloc[0]) else 0.0
        stock_acq = float(df_art['Acq'].iloc[0]) if 'Acq' in df_art.columns and pd.notnull(df_art['Acq'].iloc[0]) else 0.0

        righe_mostra = []
        for _, row in df_art.iterrows():
            qta = float(row['Qta_Effettiva'])
            req_date = row['Data_Consegna']
            doc = row['Documento'] if 'Documento' in row else "N.D."
            
            # Algoritmo di assegnazione stock
            if stock_gia >= qta:
                stock_gia -= qta
                eta, nota, pronto = oggi, "Pronto", True
            elif (stock_gia + stock_acq) >= qta:
                rimanente = qta - stock_gia
                stock_gia, stock_acq = 0, stock_acq - rimanente
                eta, nota, pronto = aggiungi_giorni_lavorativi(oggi, 10), "In Lavorazione", False
            else:
                eta, nota, pronto = aggiungi_giorni_lavorativi(oggi, 25), "Nuova Produzione", False

            # Colori e categorie
            css, cat = "on-time-row", "In Lavorazione"
            if pronto:
                cat = "Solo Disponibili"
                if pd.notnull(req_date) and oggi.date() > req_date.date(): css = "client-delay-row"
            elif pd.notnull(req_date) and eta.date() > req_date.date():
                cat = "In Ritardo"
                css = "delay-row"

            if filtro_label == "Mostra tutto" or filtro_label == cat:
                righe_mostra.append({'css': css, 'date': req_date, 'qta': qta, 'eta': eta, 'nota': nota})
                report_dettagliato.append({
                    'Codice Articolo': art, 'Descrizione': desc, 'N. Ordine': doc,
                    'Q.tà Residua': qta, 'Consegna Richiesta': req_date,
                    'Stima Consegna Safit': eta.strftime('%d/%m/%Y'), 'Note': nota
                })

        if righe_mostra:
            with st.expander(f"📦 {art} — {desc} | Residuo: {df_art['Qta_Effettiva'].sum():,.0f}"):
                for r in righe_mostra:
                    st.markdown(f'<div class="{r["css"]}"><b>Consegna: {r["date"].strftime("%d/%m/%Y") if pd.notnull(r["date"]) else "N.D."}</b> &nbsp;&nbsp;&nbsp; Q.tà: {r["qta"]:,.0f} <span style="float:right;"><small>Stima Safit: {r["eta"].strftime("%d/%m/%Y")} ({r["nota"]})</small></span></div>', unsafe_allow_html=True)

    # --- 3. DOWNLOAD EXCEL ---
    if report_dettagliato:
        df_excel = pd.DataFrame(report_dettagliato).sort_values(by='Consegna Richiesta')
        df_excel['Consegna Richiesta'] = df_excel['Consegna Richiesta'].dt.strftime('%d/%m/%Y')
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_excel.to_excel(writer, index=False, sheet_name='Dettaglio_Safit')
        st.sidebar.download_button("💾 Scarica Report Excel", output.getvalue(), f"Report_Safit_{sel_cli}.xlsx")
