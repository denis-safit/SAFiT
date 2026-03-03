import streamlit as st
import pandas as pd
import os
from datetime import datetime

# --- 1. CONFIGURAZIONE ---
APP_VERSION = "2.8.0-BOM-Warrior"
st.set_page_config(page_title=f"Safit Portal v{APP_VERSION}", layout="wide")

st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 10px; border-radius: 8px; margin-bottom: 6px; border: 1px solid #ddd; }
    .on-time-row { background-color: #e8f5e9 !important; border-left: 8px solid #4caf50; } 
    .acq-row { background-color: #e3f2fd !important; border-left: 8px solid #2196f3; }    
    .prod-row { background-color: #fffde7 !important; border-left: 8px solid #fbc02d; }   
    .urgent-row { background-color: #ffebee !important; border-left: 8px solid #f44336; } 
    .debug-box { background-color: #f1f3f4 !important; color: #202124 !important; padding: 12px; border-radius: 8px; border: 1px solid #bdc1c6; margin-bottom: 10px; font-size: 13px; }
    .kpi-box { text-align: center; padding: 10px; background: white; border-radius: 10px; border: 1px solid #eee; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. LOGICA DI PULIZIA DATI ---
def clean_num(val):
    if pd.isna(val) or str(val).strip() == '': return 0.0
    s = str(val).replace(' ', '').replace('\xa0', '').strip()
    # Gestione separatore migliaia (punto) e decimali (virgola)
    if ',' in s and '.' in s: s = s.replace('.', '')
    s = s.replace(',', '.')
    try: return float(s)
    except: return 0.0

def super_smart_load(filename, target_col):
    if not os.path.exists(filename): return pd.DataFrame()
    df_p = pd.read_excel(filename, header=None, nrows=40)
    h_idx = 0
    for i, row in df_p.iterrows():
        if target_col in [str(c).strip() for c in row.values]:
            h_idx = i; break
    df = pd.read_excel(filename, skiprows=h_idx)
    df.columns = [str(c).strip() for c in df.columns]
    return df

# --- 3. MOTORE DI CALCOLO BOM & STOCK ---
@st.cache_data(ttl=60)
def process_data():
    df_arca = super_smart_load('righe_Ordini_ARCA.xlsx', "Articolo C")
    df_acc = super_smart_load('Avanzamento_access.xlsx', "CODICE")
    
    if df_arca.empty or df_acc.empty: return pd.DataFrame(), {}

    # Creazione mappa Access con Figli
    raw_db = {}
    for _, r in df_acc.iterrows():
        cod = str(r.get('CODICE', '')).strip().upper()
        if not cod or cod == 'NAN': continue
        
        raw_db[cod] = {
            'GIA': clean_num(r.get('GIA', 0)),
            'INACQ': clean_num(r.get('INACQ', 0)),
            'PROD': clean_num(r.get('TMP',0)) + clean_num(r.get('RWI',0)) + clean_num(r.get('TRS',0)),
            'FIGLIO': str(r.get('FIGLIO', '')).strip().upper() if pd.notna(r.get('FIGLIO')) and str(r.get('FIGLIO')).strip() != '0' else None
        }

    # Funzione ricorsiva per sommare disponibilità della filiera
    def get_bom_availability(codice, visited=None):
        if visited is None: visited = set()
        if codice not in raw_db or codice in visited: return 0, 0, 0, codice
        visited.add(codice)
        
        item = raw_db[codice]
        g, a, p = item['GIA'], item['INACQ'], item['PROD']
        path = codice
        
        if item['FIGLIO']:
            sg, sa, sp, spath = get_bom_availability(item['FIGLIO'], visited)
            g += sg; a += sa; p += sp
            path += f" → {spath}"
        
        return g, a, p, path

    # Processo Ordini
    c_tipo, c_art, c_qta, c_dat, c_cli = "Codice Documento", "Articolo C", "Qta Residua", "Data Consegna", "Cliente Fornitore CD"
    df_arca[c_qta] = df_arca[c_qta].apply(clean_num)
    df_arca[c_dat] = pd.to_datetime(df_arca[c_dat], errors='coerce')
    
    df_ord = df_arca[df_arca[c_tipo].isin(['OCI', 'OCA'])].sort_values(c_dat)
    
    # Copia per scalabilità
    stock_state = {k: v.copy() for k, v in raw_db.items()}
    final_rows = []

    for _, row in df_ord.iterrows():
        art = str(row[c_art]).strip().upper()
        qta_req = row[c_qta]
        
        # Calcolo disponibilità aggregata Padre + Figlio
        g_tot, a_tot, p_tot, chain = get_bom_availability(art)
        
        st_v, cs_v = ("MANCANTE", "urgent-row")
        
        # Logica a scalare su disponibilità totale
        if g_tot >= qta_req:
            st_v, cs_v = "DISPONIBILE", "on-time-row"
            # (Logica di decremento semplificata per UI)
        elif (g_tot + a_tot) >= qta_req:
            st_v, cs_v = "IN ACQUISTO", "acq-row"
        elif (g_tot + a_tot + p_tot) >= qta_req:
            st_v, cs_v = "IN PRODUZIONE", "prod-row"

        res = row.to_dict()
        res.update({'ST': st_v, 'CS': cs_v, 'CHAIN': chain, 'DISP_TOT': g_tot + a_tot + p_tot})
        final_data = final_rows.append(res)

    return pd.DataFrame(final_rows), raw_db

# --- 4. INTERFACCIA UTENTE ---
df_res, db_snap = process_data()

st.title(f"Safit Portal v{APP_VERSION}")

if not df_res.empty:
    with st.sidebar:
        st.header("Filtri")
        filtro_stati = st.multiselect("Stati:", ["DISPONIBILE", "IN ACQUISTO", "IN PRODUZIONE", "MANCANTE"], default=["DISPONIBILE", "IN ACQUISTO", "IN PRODUZIONE", "MANCANTE"])
        search = st.text_input("🔍 Cerca Articolo:").upper()

    df_f = df_res[df_res['ST'].isin(filtro_stati)]
    if search: df_f = df_f[df_f['Articolo C'].str.contains(search)]

    for art, g in df_f.groupby('Articolo C'):
        with st.expander(f"📦 {art} - {g['Articolo D'].iloc[0]}"):
            s = db_snap.get(art, {})
            st.markdown(f"""
            <div class="debug-box">
                <b>FILIERA (BOM):</b> {g['CHAIN'].iloc[0]}<br>
                <b>Giacenza Padre:</b> {int(s.get('GIA',0))} | 
                <b>Totale Disponibile (Padre+Figli):</b> {int(g['DISP_TOT'].iloc[0])}
            </div>
            """, unsafe_allow_html=True)
            
            for _, r in g.iterrows():
                st.markdown(f'<div class="status-row {r["CS"]}"><span>📅 {r["Data Consegna"].strftime("%d/%m/%Y")} | Q: {int(r["Qta Residua"])} | {r["Cliente Fornitore CD"]}</span><span><b>{r["ST"]}</b></span></div>', unsafe_allow_html=True)
