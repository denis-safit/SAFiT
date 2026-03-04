import streamlit as st
import pandas as pd
import os

# --- CONFIGURAZIONE ---
APP_VERSION = "3.5.1-STABLE-BOM"
st.set_page_config(page_title=f"Safit Portal - BETA BOM", layout="wide")

st.markdown("""
    <style>
    .status-card { padding: 15px; border-radius: 10px; margin-bottom: 10px; border: 1px solid #ddd; color: #000 !important; }
    .on-time { background-color: #e8f5e9 !important; border-left: 10px solid #4caf50; } 
    .bom-source { background-color: #f3e5f5 !important; border-left: 10px solid #9c27b0; } 
    .warning { background-color: #fff3e0 !important; border-left: 10px solid #ff9800; }
    .urgent { background-color: #ffebee !important; border-left: 10px solid #f44336; }
    </style>
""", unsafe_allow_html=True)

def clean_num(series):
    return pd.to_numeric(series.replace('[\\.,]', '', regex=True), errors='coerce').fillna(0)

def process_portal_data():
    if not os.path.exists('righe_Ordini_ARCA.xlsx') or not os.path.exists('Avanzamento_access.xlsx'):
        st.error("File dati non trovati!")
        return pd.DataFrame()

    df_arca = pd.read_excel('righe_Ordini_ARCA.xlsx')
    df_acc = pd.read_excel('Avanzamento_access.xlsx')
    
    # DEBUG: Stampiamo le colonne nel terminale per sicurezza
    print("Colonne ARCA rilevate:", df_arca.columns.tolist())
    
    df_acc.columns = [str(c).strip().upper() for c in df_acc.columns]
    df_arca["Articolo C"] = df_arca["Articolo C"].astype(str).str.strip().str.upper()
    df_arca = df_arca[~df_arca["Articolo C"].isin(['NAN', 'NONE', ''])]
    df_arca["Qta Residua"] = clean_num(df_arca["Qta Residua"]).abs()

    stock_map = {}
    for _, r in df_acc.iterrows():
        cod = str(r.get('CODICE', '')).strip().upper()
        if cod and cod != 'NAN':
            stock_map[cod] = {
                'GIA': float(clean_num(pd.Series([r.get('GIA', 0)])).iloc[0]),
                'ACQ': float(clean_num(pd.Series([r.get('ACQ', 0)])).iloc[0]),
                'PROD': float(clean_num(pd.Series([r.get('PROD', 0)])).iloc[0]),
                'FIGLIO': str(r.get('FIGLIO', '')).strip().upper()
            }

    def get_availability(target_art, qta_needed, depth=0):
        if depth > 5 or target_art not in stock_map:
            return None, 0
        info = stock_map[target_art]
        if info['GIA'] >= qta_needed:
            info['GIA'] -= qta_needed
            return target_art, 1
        figlio = info.get('FIGLIO', 'NAN')
        if figlio and figlio != 'NAN' and figlio in stock_map:
            res_art, res_lvl = get_availability(figlio, qta_needed, depth + 1)
            if res_art: return res_art, res_lvl + 1
        if depth == 0:
            if info['ACQ'] + info['PROD'] >= qta_needed:
                return "PIANIFICATO", 0
        return None, -1

    final_data = []
    for _, row in df_arca.iterrows():
        art = row["Articolo C"]
        qta = row["Qta Residua"]
        source_art, level = get_availability(art, qta)
        res = row.to_dict()
        res['ART_KEY'] = art
        if source_art == art:
            res.update({'STATO': 'DISPONIBILE', 'DESC_STATO': 'Giacenza Pronta', 'CSS': 'on-time'})
        elif level > 1:
            res.update({'STATO': 'DISP. DA FIGLIO', 'DESC_STATO': f'Coperto da {source_art}', 'CSS': 'bom-source'})
        elif source_art == "PIANIFICATO":
            res.update({'STATO': 'PIANIFICATO', 'DESC_STATO': 'In arrivo (Acq/Prod)', 'CSS': 'warning'})
        else:
            res.update({'STATO': 'PRODUZIONE', 'DESC_STATO': 'Manca Materiale', 'CSS': 'urgent'})
        final_data.append(res)
    return pd.DataFrame(final_data)

# --- INTERFACCIA ---
st.title("🏭 Safit Portal - Gestione Ordini (BOM Ready)")
df_final = process_portal_data()

if not df_final.empty:
    search = st.text_input("Filtra per Codice Articolo:").strip().upper()
    filtered = df_final[df_final['ART_KEY'].str.contains(search)] if search else df_final.head(30)

    for _, r in filtered.iterrows():
        # Recupero sicuro dei nomi colonne (se non esistono mette "N/D")
        cliente = r.get('Ragione Sociale', r.get('Cliente', r.get('Rag.Soc.', 'N/D')))
        consegna = r.get('Data Consegna', r.get('Data Cons.', 'N/D'))
        
        st.markdown(f"""
        <div class="status-card {r['CSS']}">
            <div style="display: flex; justify-content: space-between; font-weight: bold;">
                <span>📦 {r['ART_KEY']}</span>
                <span>{r['STATO']}</span>
            </div>
            <div style="font-size: 0.9em; color: #555;">
                👤 Cliente: {cliente} | 📉 Qta: {int(r['Qta Residua'])} | 📅 Consegna: {consegna}
            </div>
            <div style="font-size: 0.85em; font-style: italic; margin-top: 5px;">
                ℹ️ {r['DESC_STATO']}
            </div>
        </div>
        """, unsafe_allow_html=True)
        
