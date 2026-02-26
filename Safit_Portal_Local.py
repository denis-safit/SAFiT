import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE ---
APP_VERSION = "1.0.14"
st.set_page_config(page_title=f"Safit - Portale Avanzamento {APP_VERSION}", layout="wide")

st.markdown(f"""
    <style>
    .custom-header {{
        display: flex; justify-content: space-between; align-items: center;
        padding: 12px; background: #f8f9fa; border: 1px solid #ddd; border-radius: 8px; margin-bottom: 4px;
    }}
    .pill-bg {{
        background-color: #eee; border-radius: 12px; width: 160px; height: 26px;
        border: 1px solid #ccc; position: relative;
    }}
    .pill-fill {{ background-color: #28a745; height: 100%; border-radius: 12px; }}
    .pill-text {{
        position: absolute; top: 0; width: 100%; font-size: 14px;
        font-weight: bold; line-height: 24px; text-align: center; color: #000;
    }}
    .status-row {{
        display: flex; justify-content: space-between; padding: 10px;
        border-radius: 6px; margin-bottom: 4px; border-left: 6px solid;
    }}
    .on-time-row {{ background-color: #f1f8e9; border-color: #4caf50; color: #1b5e20; }}
    .delay-row {{ background-color: #fff8e1; border-color: #ffc107; color: #5d4037; }}
    .prod-delay-row {{ background-color: #ffebee; border-color: #f44336; color: #b71c1c; }}
    </style>
    """, unsafe_allow_html=True)

# --- 2. CARICAMENTO DATI ---
@st.cache_data
def load_data():
    try:
        # Ordini
        df = pd.read_excel('righe_Ordini_ARCA.xlsx', sheet_name='Foglio1', skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        for col in ['Cliente Fornitore CD', 'Articolo C', 'Articolo D', 'Data']:
            if col in df.columns: df[col] = df[col].ffill()
        
        df['Articolo C'] = df['Articolo C'].astype(str).str.strip().str.upper()
        df['Data_Consegna'] = pd.to_datetime(df['Data'], errors='coerce')
        q_col = 'Qta Residua' if 'Qta Residua' in df.columns else 'Qta Doc'
        df['Qta_Effettiva'] = pd.to_numeric(df[q_col], errors='coerce').fillna(0)
        df = df[df['Qta_Effettiva'] > 0]

        # Magazzino
        if os.path.exists('Avanzamento_access.xlsx'):
            df_t = pd.read_excel('Avanzamento_access.xlsx', skiprows=1)
            df_t.columns = [str(c).strip() for c in df_t.columns]
            c_col = 'Codice' if 'Codice' in df_t.columns else 'Art_Key'
            
            df_t[c_col] = df_t[c_col].astype(str).str.strip().str.upper()
            for c in ['Gia','Trs','Rwi','Acq','Tmp','Lan','Grz']:
                if c in df_t.columns:
                    df_t[c] = pd.to_numeric(df_t[c].astype(str).str.replace('.','').str.replace(',','.'), errors='coerce').fillna(0)
            
            df = df.merge(df_t[[c_col,'Gia','Trs','Rwi','Acq','Tmp','Lan','Grz']], left_on='Articolo C', right_on=c_col, how='left')
        return df
    except: return pd.DataFrame()

data = load_data()
oggi = datetime.now()

# --- 3. SIDEBAR ---
with st.sidebar:
    if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG')
    filtro = st.radio("Stato:", ["Mostra tutto", "Solo Disponibili", "In Lavorazione", "In Ritardo"])

# --- 4. LOGICA ---
df_cli = data.copy() # Qui andrebbe il filtro cliente
if not df_cli.empty:
    for art in sorted(df_cli['Articolo C'].unique()):
        df_art = df_cli[df_cli['Articolo C'] == art].sort_values('Data_Consegna')
        st_gia = float(df_art.iloc[0].get('Gia', 0))
        
        mostra_art = []
        for _, r in df_art.iterrows():
            # Logica Disponibilità
            if st_gia >= r['Qta_Effettiva']:
                cat, p, stima = "DISPONIBILE", 100, oggi
                st_gia -= r['Qta_Effettiva']
            else:
                cat, p, stima = "LAVORAZIONE", 20, oggi + timedelta(days=20)
            
            ritardo = (pd.notnull(r['Data_Consegna']) and stima.date() > r['Data_Consegna'].date())
            
            # Applicazione Filtri
            keep = False
            if filtro == "Mostra tutto": keep = True
            elif filtro == "Solo Disponibili" and cat == "DISPONIBILE": keep = True
            elif filtro == "In Lavorazione" and cat == "LAVORAZIONE": keep = True
            elif filtro == "In Ritardo" and ritardo: keep = True
            
            if keep:
                css = "on-time-row" if (cat=="DISPONIBILE" and not ritardo) else ("prod-delay-row" if ritardo else "delay-row")
                mostra_art.append({'css': css, 'r': r, 'p': p, 'stima': stima, 'cat': cat})

        if mostra_art:
            # Header con pillola grande
            pct_header = mostra_art[0]['p']
            st.markdown(f"""
                <div class="custom-header">
                    <span>📦 <b>{art}</b></span>
                    <div class="pill-bg">
                        <div class="pill-fill" style="width: {pct_header}%;"></div>
                        <div class="pill-text">{pct_header}%</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
            with st.expander(""):
                for m in mostra_art:
                    st.markdown(f"""
                        <div class="status-row {m['css']}">
                            <span>Consegna: {m['r']['Data_Consegna'].strftime('%d/%m/%Y')} | Q.tà: {m['r']['Qta_Effettiva']:,.0f}</span>
                            <span>Stima: {m['stima'].strftime('%d/%m/%Y')} ({m['cat']})</span>
                        </div>
                    """, unsafe_allow_html=True)
