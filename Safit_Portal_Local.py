import streamlit as st
import pandas as pd
import os
from datetime import datetime

# --- CONFIGURAZIONE ---
APP_VERSION = "2.7.1-Accuracy-Fix"
st.set_page_config(page_title=f"Safit Portal v{APP_VERSION}", layout="wide")

st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 10px; border-radius: 8px; margin-bottom: 6px; border: 1px solid #ddd; }
    .on-time-row { background-color: #e8f5e9 !important; border-left: 6px solid #4caf50; } 
    .acq-row { background-color: #e3f2fd !important; border-left: 8px solid #2196f3; }    
    .prod-row { background-color: #fffde7 !important; border-left: 8px solid #fbc02d; }   
    .urgent-row { background-color: #ffebee !important; border-left: 8px solid #f44336; } 
    .debug-box { background-color: #f8f9fa !important; color: #333 !important; padding: 10px; border-radius: 8px; border: 1px solid #ccc; margin-bottom: 10px; font-size: 13px; }
    </style>
    """, unsafe_allow_html=True)

def clean_num(val):
    if pd.isna(val) or str(val).strip() == '' or str(val).lower() == 'nan': return 0.0
    # Rimuove spazi e formati strani
    s = str(val).replace(' ', '').replace('\xa0', '')
    # Gestione separatore migliaia (punto) e decimali (virgola)
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except:
        return 0.0

def super_smart_load(filename, target_col):
    """Cerca la riga di intestazione corretta e carica i dati senza slittamenti"""
    if not os.path.exists(filename): return pd.DataFrame()
    
    # Leggiamo le prime 50 righe per trovare l'intestazione
    df_preview = pd.read_excel(filename, header=None, nrows=50)
    header_idx = 0
    for i, row in df_preview.iterrows():
        if target_col in [str(c).strip() for c in row.values]:
            header_idx = i
            break
            
    # Carichiamo il file partendo dalla riga giusta
    df = pd.read_excel(filename, skiprows=header_idx)
    # Puliamo i nomi delle colonne da spazi invisibili
    df.columns = [str(c).strip() for c in df.columns]
    return df

@st.cache_data(ttl=60)
def load_and_process():
    # Caricamento file con la nuova logica ultra-precisa
    df_arca = super_smart_load('righe_Ordini_ARCA.xlsx', "Articolo C")
    df_acc = super_smart_load('Avanzamento_access.xlsx', "CODICE")
    
    if df_arca.empty or df_acc.empty: return pd.DataFrame(), {}

    # 1. MAPPATURA ACCESS (SUI CAMPI CHE MI HAI INDICATO)
    db_access = {}
    for _, r in df_acc.iterrows():
        cod = str(r.get('CODICE', '')).strip().upper()
        if not cod or cod == 'NAN': continue
        
        # Estrazione valori con pulizia numerica
        gia = clean_num(r.get('GIA', 0))
        inacq = clean_num(r.get('INACQ', 0))
        tmp = clean_num(r.get('TMP', 0))
        rwi = clean_num(r.get('RWI', 0))
        trs = clean_num(r.get('TRS', 0))
        
        # DEBUG per Denis: se il codice è quello dell'esempio, stampiamo i valori trovati
        if cod == "PPU060400090000000":
             print(f"DEBUG {cod}: GIA={gia}, SLD_M={r.get('SLD_M')}")

        db_access[cod] = {
            'GIA': gia,
            'INACQ': inacq,
            'PROD': tmp + rwi + trs,
            'FIGLIO': str(r.get('FIGLIO', '')).strip().upper() if pd.notna(r.get('FIGLIO')) else None,
            'RAW_GIA': gia, # Salviamo per il box di riepilogo
            'RAW_ACQ': inacq,
            'RAW_PROD': tmp + rwi + trs
        }

    # 2. LOGICA DI COPERTURA ORDINI
    c_tipo, c_art, c_qta, c_dat, c_cli = "Codice Documento", "Articolo C", "Qta Residua", "Data Consegna", "Cliente Fornitore CD"
    df_arca[c_qta] = df_arca[c_qta].apply(clean_num)
    df_arca[c_dat] = pd.to_datetime(df_arca[c_dat], errors='coerce')
    
    df_ordini = df_arca[df_arca[c_tipo].isin(['OCI', 'OCA'])].sort_values(c_dat)
    
    final_data = []
    # Copia locale per scalare le disponibilità
    working_stock = {k: v.copy() for k, v in db_access.items()}

    for _, row in df_ordini.iterrows():
        art = str(row[c_art]).strip().upper()
        qta_req = row[c_qta]
        s = working_stock.get(art, {'GIA': 0, 'INACQ': 0, 'PROD': 0})
        
        status, color = ("MANCANTE", "urgent-row")
        
        # Sequenza di consumo: Giacenza -> Acquisto -> Produzione
        if s['GIA'] >= qta_req:
            s['GIA'] -= qta_req
            status, color = "DISPONIBILE", "on-time-row"
        elif (s['GIA'] + s['INACQ']) >= qta_req:
            diff = qta_req - s['GIA']
            s['GIA'] = 0
            s['INACQ'] -= diff
            status, color = "IN ACQUISTO", "acq-row"
        elif (s['GIA'] + s['INACQ'] + s['PROD']) >= qta_req:
            diff = qta_req - (s['GIA'] + s['INACQ'])
            s['GIA'] = 0
            s['INACQ'] = 0
            s['PROD'] -= diff
            status, color = "IN PRODUZIONE", "prod-row"
            
        res = row.to_dict()
        res.update({'ST': status, 'CS': color, 'ART_KEY': art})
        final_data.append(res)

    return pd.DataFrame(final_data), db_access

# --- INTERFACCIA ---
df_res, db_snap = load_and_process()

st.title(f"Safit Portal - Verifica Precisione v{APP_VERSION}")

if not df_res.empty:
    search = st.sidebar.text_input("🔍 Verifica Codice (es. PPU060400090000000):").upper()
    df_f = df_res[df_res['ART_KEY'].str.contains(search)] if search else df_res

    for art, g in df_f.groupby('ART_KEY'):
        with st.expander(f"📦 {art} - {g['Articolo D'].iloc[0]}", expanded=True):
            # Recuperiamo i dati ORIGINALI dal file per questo articolo
            s = db_snap.get(art, {})
            st.markdown(f"""
            <div class="debug-box">
                <b>DATI REALI DA FILE ACCESS:</b><br>
                Giacenza (GIA): <b style="color:blue; font-size:16px;">{int(s.get('RAW_GIA',0))}</b> (Deve essere 17061 per il PPU06...900)<br>
                In Acquisto (INACQ): <b>{int(s.get('RAW_ACQ',0))}</b> | 
                Lavorazione (TMP+RWI+TRS): <b>{int(s.get('RAW_PROD',0))}</b>
            </div>
            """, unsafe_allow_html=True)
            
            for _, r in g.iterrows():
                st.markdown(f'<div class="status-row {r["CS"]}"><span>📅 {r["Data Consegna"].strftime("%d/%m/%Y")} | Q: {int(r["Qta Residua"])} | {r["Cliente Fornitore CD"]}</span><span><b>{r["ST"]}</b></span></div>', unsafe_allow_html=True)
