import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO
import plotly.express as px

# --- 1. CONFIGURAZIONE E STILE (RIPRISTINATO ORIGINALE) ---
APP_VERSION = "3.4.8"
st.set_page_config(page_title=f"Safit Portal v{APP_VERSION}", layout="wide")

st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 10px; border-radius: 8px; margin-bottom: 6px; border: 1px solid #ddd; color: #000 !important; font-size: 14px; }
    .on-time-row { background-color: #e8f5e9 !important; border-left: 6px solid #4caf50; } 
    .acq-row { background-color: #e3f2fd !important; border-left: 8px solid #2196f3; }    
    .prod-row { background-color: #fffde7 !important; border-left: 8px solid #fbc02d; }   
    .urgent-row { background-color: #ffebee !important; border-left: 8px solid #f44336; } 
    .oca-row { background-color: #f5f5f5 !important; border-left: 8px solid #9e9e9e; color: #666 !important; }
    .debug-box { background-color: #f8f9fa !important; color: #333 !important; padding: 12px; border-radius: 8px; border: 1px dotted #bbb; margin-bottom: 10px; display: flex; justify-content: space-around; font-size: 13px; font-weight: bold; }
    .kpi-card { background-color: #ffffff; border: 1px solid #e0e0e0; padding: 15px; border-radius: 10px; text-align: center; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .kpi-val { font-size: 24px; font-weight: bold; color: #1f77b4; }
    .user-info { padding: 10px; background: #f8f9fa; border-radius: 5px; border: 1px solid #eee; margin-bottom: 20px; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. GESTIONE UTENTI DINAMICA (ORIGINALE) ---
@st.cache_data
def get_user_db():
    if os.path.exists('utenti.xlsx'):
        try:
            df_u = pd.read_excel('utenti.xlsx')
            df_u.columns = [str(c).strip() for c in df_u.columns]
            return df_u.set_index('username')[['password', 'cliente_arca']].T.to_dict('list')
        except Exception as e:
            st.error(f"Errore lettura utenti.xlsx: {e}")
    return {"safit_admin": ["admin2026", "TUTTI"], "denis": ["denis2026", "TUTTI"]}

USER_DB = get_user_db()

# --- 3. FUNZIONI TECNICHE (ORIGINALE) ---
def clean_num(serie):
    s = serie.astype(str).str.replace(' ', '').str.replace('\xa0', '')
    def fix_val(val):
        if val.lower() in ['nan', '', 'none']: return '0'
        if ',' in val and '.' in val: return val.replace('.', '').replace(',', '.')
        elif ',' in val: return val.replace(',', '.')
        return val
    return pd.to_numeric(s.apply(fix_val), errors='coerce').fillna(0)

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.drop(columns=['CS', 'DT_EXP', 'ART_KEY', 'ST', 'PRES_DISP_RT'], errors='ignore').to_excel(writer, index=False)
    return output.getvalue()

def smart_load(filename, key_col):
    if not os.path.exists(filename): return pd.DataFrame()
    df_p = pd.read_excel(filename, header=None, nrows=20)
    h_row = 0
    for i, row in df_p.iterrows():
        if key_col in row.astype(str).values:
            h_row = i; break
    df = pd.read_excel(filename, skiprows=h_row)
    df.columns = [str(c).strip() for c in df.columns]
    if "CODICE" not in key_col: df = df.ffill() 
    return df

# --- 4. MOTORE DI CALCOLO (INTEGRATO OFF) ---
@st.cache_data(ttl=300)
def load_and_process():
    try:
        df_arca_raw = smart_load('righe_Ordini_ARCA.xlsx', "Articolo C")
        if df_arca_raw.empty: return pd.DataFrame(), {}
        
        c_tipo, c_art, c_qta, c_dat, c_cli = "Codice Documento", "Articolo C", "Qta Residua", "Data Consegna", "Cliente Fornitore CD"
        
        df_arca_raw[c_qta] = clean_num(df_arca_raw[c_qta]).abs()
        df_arca_raw[c_dat] = pd.to_datetime(df_arca_raw[c_dat], errors='coerce')
        df_arca_raw[c_art] = df_arca_raw[c_art].astype(str).str.strip().str.upper()

        # --- RECUPERO DATE DISPONIBILITÀ (OFF) ---
        df_off = df_arca_raw[df_arca_raw[c_tipo] == 'OFF'].copy()
        mappa_disp_rt = df_off.sort_values(c_dat).groupby(c_art)[c_dat].first().to_dict()

        df_arca = df_arca_raw[df_arca_raw[c_tipo].isin(['OCI', 'OCA'])]
        df_arca = df_arca[df_arca[c_qta] > 0].dropna(subset=[c_dat, c_art])

        df_acc = smart_load('Avanzamento_access.xlsx', "CODICE")
        stock_map = {}
        if not df_acc.empty:
            df_acc.columns = [str(c).strip().upper() for c in df_acc.columns]
            for _, r in df_acc.iterrows():
                art_code = str(r['CODICE']).strip().upper()
                gia = clean_num(pd.Series([r.get('GIA', 0)])).iloc[0]
                acq = clean_num(pd.Series([r.get('INACQ', 0)])).iloc[0]
                prod = sum([clean_num(pd.Series([r.get(f, 0)])).iloc[0] for f in ['LANCIATI', 'GRZ', 'TMP', 'RWI', 'TRS', 'ACQ', 'TRF']])
                stock_map[art_code] = {'GIA': gia, 'ACQ': acq, 'PROD': prod}

        df_orders = df_arca.sort_values(by=[c_art, c_dat])
        final_results = []
        current_stocks = {k: v.copy() for k, v in stock_map.items()}

        for index, row in df_orders.iterrows():
            art_code = row[c_art]
            qta_ordine = float(row[c_qta])
            scorte = current_stocks.get(art_code, {'GIA': 0, 'ACQ': 0, 'PROD': 0})
            
            # Data OFF
            dt_off = mappa_disp_rt.get(art_code, None)
            dt_off_str = dt_off.strftime('%d/%m/%Y') if pd.notnull(dt_off) else "-"
            
            if scorte['GIA'] >= qta_ordine:
                scorte['GIA'] -= qta_ordine; stato, colore = "DISPONIBILE", "on-time-row"
            elif (scorte['GIA'] + scorte['ACQ']) >= qta_ordine:
                scorte['ACQ'] -= (qta_ordine - scorte['GIA']); scorte['GIA'] = 0; stato, colore = "ACQUISTO", "acq-row"
            elif (scorte['GIA'] + scorte['ACQ'] + scorte['PROD']) >= qta_ordine:
                scorte['PROD'] -= (qta_ordine - scorte['GIA'] - scorte['ACQ']); scorte['GIA'] = 0; scorte['ACQ'] = 0; stato, colore = "PRODUZIONE", "prod-row"
            else:
                scorte['GIA'] = 0; scorte['ACQ'] = 0; scorte['PROD'] = 0; stato, colore = "MANCANTE", "urgent-row"
            
            if row[c_tipo] == 'OCA' and stato == "MANCANTE": stato, colore = "DA PIANIFICARE", "oca-row"

            current_stocks[art_code] = scorte
            res = row.to_dict()
            res.update({'ST': stato, 'CS': colore, 'DT_EXP': row[c_dat], 'PRES_DISP_RT': dt_off_str, 'ART_KEY': art_code, 'CLI_NAME': str(row[c_cli])})
            final_results.append(res)
                
        return pd.DataFrame(final_results), stock_map
    except Exception as e:
        st.error(f"Errore: {e}"); return pd.DataFrame(), {}

# --- 5. LOGICA DI ACCESSO ---
if "auth" not in st.session_state: st.session_state.auth = False
if not st.session_state.auth:
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=250)
        u = st.text_input("Username").strip()
        p = st.text_input("Password", type="password").strip()
        if st.button("Accedi", use_container_width=True):
            if u in USER_DB and str(USER_DB[u][0]) == p:
                st.session_state.auth, st.session_state.user, st.session_state.permesso = True, u, USER_DB[u][1]
                st.rerun()
            else: st.error("Credenziali non valide")
else:
    # --- 6. DASHBOARD PRINCIPALE ---
    df_final, stock_map = load_and_process()
    if df_final.empty:
        st.warning("Nessun dato trovato."); st.stop()

    if st.session_state.permesso != "TUTTI":
        df_final = df_final[df_final['CLI_NAME'] == st.session_state.permesso]

    # Info Utente
    st.markdown(f'<div class="user-info">Utente: <b>{st.session_state.user}</b> | Accesso: {st.session_state.permesso}</div>', unsafe_allow_html=True)
    
    # KPI
    c1, c2, c3, c4 = st.columns(4)
    with c1: st.markdown(f'<div class="kpi-card"><div class="kpi-val">{len(df_final)}</div>Ordini Totali</div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="kpi-card"><div class="kpi-val">{len(df_final[df_final["ST"]=="DISPONIBILE"])}</div>Pronti</div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="kpi-card"><div class="kpi-val">{len(df_final[df_final["ST"].isin(["ACQUISTO","PRODUZIONE"])])}</div>In Lavorazione</div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="kpi-card"><div class="kpi-val" style="color:red">{len(df_final[df_final["ST"]=="MANCANTE"])}</div>Mancanti</div>', unsafe_allow_html=True)

    # Filtri
    st.write("---")
    search = st.text_input("Cerca Articolo o Ragione Sociale").upper()
    if search:
        df_final = df_final[df_final['Articolo C'].str.contains(search) | df_final['CLI_NAME'].str.contains(search)]

    # Debug Box (Opzionale)
    with st.expander("Vedi riepilogo tecnico"):
        st.markdown(f'<div class="debug-box"><span>Righe totali: {len(df_final)}</span><span>Articoli unici: {df_final["Articolo C"].nunique()}</span></div>', unsafe_allow_html=True)

    # Elenco Ordini
    for _, r in df_final.iterrows():
        st.markdown(f"""
            <div class="status-row {r['CS']}">
                <div style="flex:2"><b>{r['Articolo C']}</b><br><small>{r['CLI_NAME']}</small></div>
                <div style="flex:1">Qta: {int(r['Qta Residua'])}</div>
                <div style="flex:1.2">Consegna: {r['DT_EXP'].strftime('%d/%m/%Y')}</div>
                <div style="flex:2; color: #1f77b4;"><b>Presunta data disponibilità in tempo reale: {r['PRES_DISP_RT']}</b></div>
                <div style="flex:1; text-align:right;"><b>{r['ST']}</b></div>
            </div>
        """, unsafe_allow_html=True)

    # Export
    st.write("---")
    st.download_button("Scarica Report Excel", data=to_excel(df_final), file_name=f"Report_Safit_{datetime.now().strftime('%Y%m%d')}.xlsx")
    if st.button("Logout"):
        st.session_state.auth = False; st.rerun()
