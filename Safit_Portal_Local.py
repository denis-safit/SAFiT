import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO
import plotly.express as px

# --- 1. CONFIGURAZIONE E STILE (Dalla v3.2.0) ---
APP_VERSION = "3.2.1-Master-AuthFix"
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

# --- 2. GESTIONE UTENTI DINAMICA (Recuperata da v1.4) ---
@st.cache_data
def get_user_db():
    # Prova a leggere il file utenti.xlsx come nella v1.4
    if os.path.exists('utenti.xlsx'):
        try:
            df_u = pd.read_excel('utenti.xlsx')
            df_u.columns = [str(c).strip() for c in df_u.columns]
            # Crea un dizionario {username: [password, cliente_arca]}
            return df_u.set_index('username')[['password', 'cliente_arca']].T.to_dict('list')
        except Exception as e:
            st.error(f"Errore lettura utenti.xlsx: {e}")
    
    # Utenti di emergenza se il file non esiste
    return {
        "safit_admin": ["admin2026", "TUTTI"],
        "denis": ["denis2026", "TUTTI"]
    }

USER_DB = get_user_db()

# --- 3. FUNZIONI TECNICHE (Invariate) ---
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
        df.drop(columns=['CS', 'DT_EXP', 'ART_KEY', 'ST'], errors='ignore').to_excel(writer, index=False)
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

# --- 4. MOTORE DI CALCOLO (LOGICA MASTER 3.2.8: PROTEZIONE DATE ORIGINALI) ---
@st.cache_data(ttl=300)
def load_and_process():
    try:
        # 4.1 Caricamento Ordini ARCA
        df_arca = smart_load('righe_Ordini_ARCA.xlsx', "Articolo C")
        if df_arca.empty: 
            return pd.DataFrame(), {}
        
        c_tipo, c_art, c_qta, c_dat, c_cli = "Codice Documento", "Articolo C", "Qta Residua", "Data Consegna", "Cliente Fornitore CD"
        
        # Pulizia dati: trasformiamo i negativi in positivi e convertiamo le date
        df_arca[c_qta] = clean_num(df_arca[c_qta]).abs()
        df_arca[c_dat] = pd.to_datetime(df_arca[c_dat], errors='coerce')
        
        # Teniamo solo ordini validi OCI/OCA
        df_arca = df_arca[df_arca[c_tipo].isin(['OCI', 'OCA'])]
        df_arca = df_arca[df_arca[c_qta] > 0].dropna(subset=[c_dat, c_art])

        # 4.2 Caricamento Scorte ACCESS
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

        # 4.3 CALCOLO SEQUENZIALE SENZA SOVRASCRIVERE LE DATE
        df_orders = df_arca.sort_values(by=[c_art, c_dat])
        final_results = []
        
        # Copia delle scorte per il calcolo a scalare
        current_stocks = {k: v.copy() for k, v in stock_map.items()}

        for index, row in df_orders.iterrows():
            art_code = str(row[c_art]).upper()
            qta_ordine = float(row[c_qta])
            scorte = current_stocks.get(art_code, {'GIA': 0, 'ACQ': 0, 'PROD': 0})
            
            # Usiamo SEMPRE la data originale di ARCA per la visualizzazione
            data_originale = row[c_dat]
            
            if scorte['GIA'] >= qta_ordine:
                scorte['GIA'] -= qta_ordine
                stato, colore = "DISPONIBILE", "on-time-row"
            elif (scorte['GIA'] + scorte['ACQ']) >= qta_ordine:
                scorte['ACQ'] -= (qta_ordine - scorte['GIA'])
                scorte['GIA'] = 0
                stato, colore = "ACQUISTO", "acq-row"
            elif (scorte['GIA'] + scorte['ACQ'] + scorte['PROD']) >= qta_ordine:
                scorte['PROD'] -= (qta_ordine - scorte['GIA'] - scorte['ACQ'])
                scorte['GIA'] = 0; scorte['ACQ'] = 0
                stato, colore = "PRODUZIONE", "prod-row"
            else:
                scorte['GIA'] = 0; scorte['ACQ'] = 0; scorte['PROD'] = 0
                stato, colore = "MANCANTE", "urgent-row"
            
            if row[c_tipo] == 'OCA' and stato == "MANCANTE":
                stato, colore = "DA PIANIFICARE", "oca-row"

            current_stocks[art_code] = scorte

            # Salvataggio: DT_EXP ora è uguale alla data di ARCA
            res = row.to_dict()
            res.update({
                'ST': stato, 
                'CS': colore, 
                'DT_EXP': data_originale, # <-- CORREZIONE: usiamo la data reale
                'ART_KEY': art_code,
                'CLI_NAME': str(row[c_cli])
            })
            final_results.append(res)
                
        return pd.DataFrame(final_results), stock_map

    except Exception as e:
        st.error(f"Errore: {e}")
        return pd.DataFrame(), {}

# --- 5. LOGICA DI ACCESSO ---
if "auth" not in st.session_state: st.session_state.auth = False
if not st.session_state.auth:
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=250)
        u = st.text_input("Username").strip()
        p = st.text_input("Password", type="password").strip()
        if st.button("Accedi", use_container_width=True):
            # Controllo incrociato con USER_DB (file utenti.xlsx)
            if u in USER_DB and str(USER_DB[u][0]) == p:
                st.session_state.auth = True
                st.session_state.user = u
                st.session_state.permesso = USER_DB[u][1]
                st.rerun()
            else:
                st.error("Credenziali non valide o utente non trovato in utenti.xlsx")
    st.stop()

# --- 6. DASHBOARD (Invariata) ---
df_res, stock_raw = load_and_process()
if not df_res.empty:
    with st.sidebar:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
        st.markdown(f'<div class="user-info">👤 <b>{st.session_state.user}</b></div>', unsafe_allow_html=True)
        if st.button("🚪 Log-out", use_container_width=True): st.session_state.auth = False; st.rerun()
        st.markdown("---")
        
        if st.session_state.permesso == "TUTTI":
            sel_cli = st.selectbox("Seleziona Cliente:", ["TUTTI"] + sorted(df_res['CLI_NAME'].unique().tolist()))
        else:
            sel_cli = st.session_state.permesso
            st.info(f"Filtro Cliente: {sel_cli}")

        sel_stati = [s for s in ["DISPONIBILE", "ACQUISTO", "PRODUZIONE", "MANCANTE", "DA PIANIFICARE"] if st.checkbox(s, value=True, key=f"ch_{s}")]
        search = st.text_input("🔍 Cerca Articolo:").upper()

    df_f = df_res[df_res['CLI_NAME'] == sel_cli] if sel_cli != "TUTTI" else df_res.copy()
    df_f = df_f[df_f['ST'].isin(sel_stati)]
    if search: df_f = df_f[df_f['ART_KEY'].str.contains(search)]

    st.sidebar.download_button("📊 Esporta Report", data=to_excel(df_f), file_name=f"Safit_Report_{datetime.now().strftime('%d%m')}.xlsx", use_container_width=True)

    st.title("Pannello Controllo Consegne Safit")
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f'<div class="kpi-card"><div style="font-size:11px">PEZZI FILTRATI</div><div class="kpi-val">{int(df_f["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div style="font-size:11px; color:#4caf50">PRONTI</div><div class="kpi-val">{int(df_f[df_f["ST"]=="DISPONIBILE"]["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card"><div style="font-size:11px; color:#2196f3">IN ACQUISTO</div><div class="kpi-val">{int(df_f[df_f["ST"]=="ACQUISTO"]["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi-card"><div style="font-size:11px; color:#f44336">MANCANTI</div><div class="kpi-val">{int(df_f[df_f["ST"]=="MANCANTE"]["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(px.pie(df_f, values='Qta Residua', names='ST', color='ST', title="Stato Copertura",
                               color_discrete_map={'DISPONIBILE':'#4caf50','ACQUISTO':'#2196f3','PRODUZIONE':'#fbc02d','MANCANTE':'#f44336','DA PIANIFICARE':'#9e9e9e'}), use_container_width=True)
    with c2:
        df_f['Famiglia'] = df_f['Articolo D'].apply(lambda x: " ".join(str(x).split()[:2]).upper())
        st.plotly_chart(px.pie(df_f.groupby('Famiglia')['Qta Residua'].sum().reset_index().sort_values('Qta Residua', ascending=False).head(10), 
                               values='Qta Residua', names='Famiglia', hole=0.4, title="Top 10 Famiglie"), use_container_width=True)

    st.markdown("---")
    for art, g in df_f.groupby('ART_KEY'):
        with st.expander(f"📦 {art} - {g['Articolo D'].iloc[0]} ({len(g)} ordini)"):
            s_i = stock_raw.get(art, {'GIA': 0, 'ACQ': 0, 'PROD': 0})
            st.markdown(f'<div class="debug-box"><span>📦 GIA: {int(s_i["GIA"])}</span><span>🚚 ACQ: {int(s_i["ACQ"])}</span><span>⚙️ PROD: {int(s_i["PROD"])}</span></div>', unsafe_allow_html=True)
            for _, r in g.iterrows():
                tag = "📋 [PREV]" if r['Codice Documento'] == "OCA" else "🛒 [ORD]"
                st.markdown(f'<div class="status-row {r["CS"]}"><span>{tag} 📅 {r["DT_EXP"].strftime("%d/%m/%Y")} | Q: {int(r["Qta Residua"])} | {r["CLI_NAME"]}</span><span><b>{r["ST"]}</b></span></div>', unsafe_allow_html=True)
