import streamlit as st
import pandas as pd
import os
import plotly.express as px
from datetime import datetime

# --- 1. CONFIGURAZIONE & STILE ---
APP_VERSION = "3.0.0-Total-Rebirth"
st.set_page_config(page_title=f"Safit Portal v{APP_VERSION}", layout="wide")

st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 10px; border-radius: 8px; margin-bottom: 6px; border: 1px solid #ddd; color: black; }
    .on-time-row { background-color: #e8f5e9 !important; border-left: 6px solid #4caf50; } 
    .acq-row { background-color: #e3f2fd !important; border-left: 8px solid #2196f3; }    
    .prod-row { background-color: #fffde7 !important; border-left: 8px solid #fbc02d; }   
    .urgent-row { background-color: #ffebee !important; border-left: 8px solid #f44336; } 
    .debug-box { background-color: #f8f9fa !important; color: #333 !important; padding: 12px; border-radius: 8px; border: 1px solid #ccc; margin-bottom: 10px; font-size: 13px; }
    .kpi-card { background: white; padding: 20px; border-radius: 10px; border: 1px solid #eee; text-align: center; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNZIONI CORE ---
def clean_num(val):
    if pd.isna(val) or str(val).strip() == '': return 0.0
    s = str(val).replace(' ', '').replace('\xa0', '').strip()
    if ',' in s and '.' in s: s = s.replace('.', '')
    s = s.replace(',', '.')
    try: return float(s)
    except: return 0.0

def super_smart_load(filename, target_col):
    if not os.path.exists(filename): return pd.DataFrame()
    df_p = pd.read_excel(filename, header=None, nrows=40)
    h_idx = 0
    for i, row in df_p.iterrows():
        if target_col.upper() in [str(c).strip().upper() for c in row.values]:
            h_idx = i; break
    df = pd.read_excel(filename, skiprows=h_idx)
    df.columns = [str(c).strip() for c in df.columns]
    return df.loc[:, ~df.columns.duplicated()]

@st.cache_data(ttl=60)
def load_all_data():
    df_arca = super_smart_load('righe_Ordini_ARCA.xlsx', "Articolo C")
    df_acc = super_smart_load('Avanzamento_access.xlsx', "CODICE")
    
    if df_arca.empty or df_acc.empty: return pd.DataFrame(), {}

    # Mappa Access con logic Padre-Figlio
    db_acc = {}
    for _, r in df_acc.iterrows():
        c = str(r.get('CODICE', '')).strip().upper()
        if not c or c == 'NAN': continue
        db_acc[c] = {
            'GIA': clean_num(r.get('GIA', 0)),
            'INACQ': clean_num(r.get('INACQ', 0)),
            'PROD': clean_num(r.get('TMP',0)) + clean_num(r.get('RWI',0)) + clean_num(r.get('TRS',0)),
            'FIGLIO': str(r.get('FIGLIO', '')).strip().upper() if pd.notna(r.get('FIGLIO')) and str(r.get('FIGLIO')).strip() not in ['0',''] else None
        }

    # Funzione BOM
    def get_bom(cod, visited=None):
        if visited is None: visited = set()
        if cod not in db_acc or cod in visited: return 0, 0, 0, cod
        visited.add(cod)
        item = db_acc[cod]
        g, a, p = item['GIA'], item['INACQ'], item['PROD']
        path = cod
        if item['FIGLIO']:
            sg, sa, sp, spath = get_bom(item['FIGLIO'], visited)
            g+=sg; a+=sa; p+=sp; path += f" → {spath}"
        return g, a, p, path

    # Processo Arca
    c_tipo, c_art, c_qta, c_dat, c_cli = "Codice Documento", "Articolo C", "Qta Residua", "Data Consegna", "Cliente Fornitore CD"
    df_arca[c_qta] = df_arca[c_qta].apply(clean_num)
    df_arca[c_dat] = pd.to_datetime(df_arca[c_dat], errors='coerce')
    df_ord = df_arca[df_arca[c_tipo].isin(['OCI', 'OCA'])].sort_values(c_dat)
    
    final = []
    for _, row in df_ord.iterrows():
        art = str(row[c_art]).strip().upper()
        g_t, a_t, p_t, chain = get_bom(art)
        
        st_v, cs_v = ("MANCANTE", "urgent-row")
        if g_t >= row[c_qta]: st_v, cs_v = "DISPONIBILE", "on-time-row"
        elif (g_t + a_t) >= row[c_qta]: st_v, cs_v = "IN ACQUISTO", "acq-row"
        elif (g_t + a_t + p_t) >= row[c_qta]: st_v, cs_v = "IN PRODUZIONE", "prod-row"
            
        res = row.to_dict()
        res.update({'ST': st_v, 'CS': cs_v, 'CHAIN': chain, 'DISP_TOT': g_t + a_t + p_t})
        final.append(res)
    
    return pd.DataFrame(final), db_acc

# --- 3. INTERFACCIA ---
if "auth" not in st.session_state: st.session_state.auth = False

with st.sidebar:
    st.image("https://www.safit.it/wp-content/uploads/2021/05/logo-safit.png", width=150) # Inserire URL logo reale se disponibile
    if not st.session_state.auth:
        u = st.text_input("Utente")
        p = st.text_input("Password", type="password")
        if st.button("Accedi"):
            if u == "denis" and p == "denis2026": 
                st.session_state.auth = True
                st.rerun()
    else:
        if st.button("Esci"):
            st.session_state.auth = False
            st.rerun()

if st.session_state.auth:
    df, db = load_all_data()
    
    if not df.empty:
        # Sidebar Filtri
        with st.sidebar:
            st.divider()
            cli_list = ["TUTTI"] + sorted(df['Cliente Fornitore CD'].unique().tolist())
            sel_cli = st.selectbox("Filtro Cliente:", cli_list)
            sel_stati = st.multiselect("Stati:", ["DISPONIBILE", "IN ACQUISTO", "IN PRODUZIONE", "MANCANTE"], default=["DISPONIBILE", "IN ACQUISTO", "IN PRODUZIONE", "MANCANTE"])
            search = st.text_input("🔍 Cerca Articolo:").upper()

        # Filtro Dati
        df_f = df[df['ST'].isin(sel_stati)]
        if sel_cli != "TUTTI": df_f = df_f[df_f['Cliente Fornitore CD'] == sel_cli]
        if search: df_f = df_f[df_f['Articolo C'].str.contains(search) | df_f['Articolo D'].str.contains(search)]

        # Dashboard KPI
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f'<div class="kpi-card"><span>Pezzi Totali</span><br><b style="font-size:24px;">{int(df_f["Qta Residua"].sum()):,}</b></div>', unsafe_allow_html=True)
        with c2: 
            perc = (len(df_f[df_f['ST'] == 'DISPONIBILE']) / len(df_f) * 100) if len(df_f)>0 else 0
            st.markdown(f'<div class="kpi-card"><span>Disponibilità %</span><br><b style="font-size:24px;">{perc:.1f}%</b></div>', unsafe_allow_html=True)
        
        # Grafico
        fig = px.pie(df_f, names='ST', color='ST', color_discrete_map={'DISPONIBILE':'#4caf50','IN ACQUISTO':'#2196f3','IN PRODUZIONE':'#fbc02d','MANCANTE':'#f44336'})
        st.plotly_chart(fig, use_container_width=True)

        # Tabella Articoli
        for art, g in df_f.groupby('Articolo C'):
            with st.expander(f"📦 {art} - {g['Articolo D'].iloc[0]}"):
                s = db.get(art, {})
                st.markdown(f"""
                <div class="debug-box">
                    <b>FILIERA:</b> {g['CHAIN'].iloc[0]}<br>
                    <b>Giacenza Reale (GIA):</b> <span style="color:blue; font-weight:bold; font-size:16px;">{int(s.get('GIA',0))}</span> | 
                    In Acquisto: {int(s.get('INACQ',0))} | Produzione: {int(s.get('PROD',0))}
                </div>
                """, unsafe_allow_html=True)
                for _, r in g.iterrows():
                    st.markdown(f'<div class="status-row {r["CS"]}"><span>📅 {r["Data Consegna"].strftime("%d/%m/%Y")} | Q: {int(r["Qta Residua"])} | {r["Cliente Fornitore CD"]}</span><span><b>{r["ST"]}</b></span></div>', unsafe_allow_html=True)
