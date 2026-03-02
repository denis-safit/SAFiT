import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO
import plotly.express as px

# --- 1. CONFIGURAZIONE E STILE ---
APP_VERSION = "2.3.0-Titanium"
st.set_page_config(page_title=f"Safit Portal v{APP_VERSION}", layout="wide")

st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 10px; border-radius: 8px; margin-bottom: 6px; border: 1px solid #ddd; color: #000 !important; font-size: 14px; }
    .on-time-row { background-color: #e8f5e9 !important; border-left: 6px solid #4caf50; } 
    .acq-row { background-color: #e3f2fd !important; border-left: 8px solid #2196f3; }    
    .prod-row { background-color: #fffde7 !important; border-left: 8px solid #fbc02d; }   
    .urgent-row { background-color: #ffebee !important; border-left: 8px solid #f44336; } 
    .oca-row { background-color: #f5f5f5 !important; border-left: 8px solid #9e9e9e; color: #666 !important; }
    .debug-box { background-color: #f0f2f6 !important; color: #111 !important; padding: 10px; border-radius: 6px; border: 1px solid #ccc; margin-bottom: 10px; font-weight: 600; display: flex; justify-content: space-around; }
    .kpi-card { background-color: #ffffff; border: 1px solid #e0e0e0; padding: 15px; border-radius: 10px; text-align: center; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .kpi-val { font-size: 24px; font-weight: bold; color: #1f77b4; }
    .user-info { padding: 10px; background: #f8f9fa; border-radius: 5px; border: 1px solid #eee; margin-bottom: 20px; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. DATABASE UTENTI E PERMESSI ---
USER_DB = {
    "safit_admin": ["admin2026", "TUTTI"],
    "denis": ["denis2026", "TUTTI"],
    "cliente_test": ["safit2026", "NOME_CLIENTE_ESEMPIO"] # Qui metti il nome esatto che appare in Arca
}

# --- 3. FUNZIONI TECNICHE ---
def clean_num(serie):
    s = serie.astype(str).str.replace(' ', '').str.replace('\xa0', '')
    def fix_val(val):
        if val.lower() in ['nan', '', 'none']: return '0'
        return val.replace('.', '').replace(',', '.') if ',' in val else val
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
    df = df.ffill() 
    return df

# --- 4. MOTORE DI CALCOLO ATP ---
@st.cache_data(ttl=300)
def load_and_process():
    try:
        # Carico Arca (Ordini Clienti e Fornitori)
        df_full = smart_load('righe_Ordini_ARCA.xlsx', "Articolo C")
        if df_full.empty: return pd.DataFrame(), {}
        
        c_tipo, c_art, c_qta, c_dat, c_cli = "Codice Documento", "Articolo C", "Qta Residua", "Data Consegna", "Cliente Fornitore CD"
        df_full[c_dat] = pd.to_datetime(df_full[c_dat], errors='coerce')
        df_full[c_qta] = clean_num(df_full[c_qta])
        df_full = df_full.dropna(subset=[c_dat, c_art])

        # Carico Access (Stock e Avanzamento)
        df_acc = smart_load('Avanzamento_access.xlsx', "CODICE")
        stock = {}
        if not df_acc.empty:
            for _, r in df_acc.iterrows():
                # Somma ACQ, TMP, RWI, TRF come richiesto
                p_val = sum([clean_num(pd.Series([r.get(f, 0)])).iloc[0] for f in ['ACQ', 'TMP', 'RWI', 'TRF']])
                stock[str(r['CODICE']).strip().upper()] = {
                    'GIA': clean_num(pd.Series([r.get('GIA', 0)])).iloc[0], 
                    'PROD': p_val
                }

        # Gestione Acquisti Fornitori (OFF / OFR)
        df_arr = df_full[df_full[c_tipo].isin(['OFF', 'OFR'])].copy()
        arca_arr = {art: g.sort_values(c_dat)[[c_dat, c_qta]].values.tolist() for art, g in df_arr.groupby(c_art)}

        # Processo OCI (Ordini) e OCA (Previsioni)
        df_orders = df_full[df_full[c_tipo].isin(['OCI', 'OCA'])].sort_values([c_dat, c_tipo])
        
        final = []
        for _, row in df_orders.iterrows():
            art = str(row[c_art]).strip().upper()
            qta = row[c_qta]
            s = stock.get(art, {'GIA': 0, 'PROD': 0})
            arr_list = arca_arr.get(art, [])
            
            st_v, col, dt_e = ("MANCANTE", "urgent-row", row[c_dat] + timedelta(days=45))
            if row[c_tipo] == 'OCA': st_v, col = "DA PIANIFICARE", "oca-row"

            # 1. Copertura con Giacenza
            if s['GIA'] >= qta:
                s['GIA'] -= qta; st_v, col, dt_e = "DISPONIBILE", "on-time-row", row[c_dat]
            else:
                qta -= s['GIA']; s['GIA'] = 0
                # 2. Copertura con Acquisti (OFF/OFR) - Prende la data reale
                for a in arr_list:
                    if a[1] > 0:
                        if a[1] >= qta:
                            a[1] -= qta; st_v, col, dt_e = "ACQUISTO", "acq-row", a[0]; qta = 0; break
                        else:
                            qta -= a[1]; a[1] = 0
                # 3. Copertura con Produzione (Access)
                if qta > 0 and s['PROD'] >= qta:
                    s['PROD'] -= qta; st_v, col, dt_e = "PRODUZIONE", "prod-row", datetime.now() + timedelta(days=21)

            res = row.to_dict()
            res.update({'ST': st_v, 'CS': col, 'DT_EXP': dt_e, 'ART_KEY': art, 'CLI_NAME': str(row[c_cli])})
            final.append(res)
        
        return pd.DataFrame(final), stock
    except Exception as e:
        st.error(f"Errore nel motore: {e}"); return pd.DataFrame(), {}

# --- 5. LOGICA DI ACCESSO ---
if "auth" not in st.session_state: st.session_state.auth = False

if not st.session_state.auth:
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=250)
        u = st.text_input("Utente"); p = st.text_input("Password", type="password")
        if st.button("Accedi", use_container_width=True):
            if u in USER_DB and USER_DB[u][0] == p:
                st.session_state.auth = True
                st.session_state.user = u
                st.session_state.permesso = USER_DB[u][1]
                st.rerun()
            else: st.error("Credenziali errate")
    st.stop()

# --- 6. DASHBOARD E FILTRI ---
df_res, stock_raw = load_and_process()

if not df_res.empty:
    with st.sidebar:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
        st.markdown(f'<div class="user-info">👤 <b>{st.session_state.user}</b></div>', unsafe_allow_html=True)
        if st.button("🚪 Esci"): st.session_state.auth = False; st.rerun()
        st.markdown("---")
        
        # Filtro Cliente (Bloccato se non è admin)
        if st.session_state.permesso == "TUTTI":
            sel_cli = st.selectbox("Cliente:", ["TUTTI"] + sorted(df_res['CLI_NAME'].unique().tolist()))
        else:
            sel_cli = st.session_state.permesso
            st.info(f"Filtro bloccato su: {sel_cli}")

        # Filtro Stato con bottoni Radio (le tue "opzioni")
        sel_st = st.radio("Stato Ordine:", ["TUTTI", "DISPONIBILE", "ACQUISTO", "PRODUZIONE", "MANCANTE", "DA PIANIFICARE"])
        search = st.text_input("🔍 Cerca Articolo:").upper()

    # Applica Filtri
    df_f = df_res[df_res['CLI_NAME'] == sel_cli] if sel_cli != "TUTTI" else df_res.copy()
    if sel_st != "TUTTI": df_f = df_f[df_f['ST'] == sel_st]
    if search: df_f = df_f[df_f['ART_KEY'].str.contains(search)]

    # DOWNLOAD
    st.sidebar.download_button("📊 Scarica Excel", data=to_excel(df_f), file_name="Piano_Safit.xlsx", use_container_width=True)

    # UI PRINCIPALE
    st.title(f"Piano Consegne: {sel_cli}")
    
    # KPI Row
    k1, k2, k3, k4 = st.columns(4)
    tot_q = df_f['Qta Residua'].sum()
    k1.markdown(f'<div class="kpi-card"><div style="font-size:11px">PEZZI TOTALI</div><div class="kpi-val">{int(tot_q):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div style="font-size:11px">PRONTI</div><div class="kpi-val" style="color:#4caf50">{int(df_f[df_f["ST"]=="DISPONIBILE"]["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card"><div style="font-size:11px">IN ARRIVO</div><div class="kpi-val" style="color:#2196f3">{int(df_f[df_f["ST"]=="ACQUISTO"]["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi-card"><div style="font-size:11px">MANCANTI</div><div class="kpi-val" style="color:#f44336">{int(df_f[df_f["ST"]=="MANCANTE"]["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)

    # Grafici
    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(px.pie(df_f, values='Qta Residua', names='ST', color='ST', title="Stato Copertura",
                               color_discrete_map={'DISPONIBILE':'#4caf50','ACQUISTO':'#2196f3','PRODUZIONE':'#fbc02d','MANCANTE':'#f44336','DA PIANIFICARE':'#9e9e9e'}), use_container_width=True)
    with c2:
        df_f['Famiglia'] = df_f['Articolo D'].apply(lambda x: " ".join(str(x).split()[:2]).upper())
        st.plotly_chart(px.pie(df_f.groupby('Famiglia')['Qta Residua'].sum().reset_index().sort_values('Qta Residua', ascending=False).head(10), 
                               values='Qta Residua', names='Famiglia', hole=0.4, title="Top 10 Famiglie"), use_container_width=True)

    st.markdown("---")

    # Lista Ordini
    for art, g in df_f.groupby('ART_KEY'):
        with st.expander(f"📦 {art} - {g['Articolo D'].iloc[0]} ({len(g)} righe)"):
            info = stock_raw.get(art, {'GIA': 0, 'PROD': 0})
            st.markdown(f'<div class="debug-box"><span>STOCCAGGIO: {int(info["GIA"])}</span><span>IN PRODUZIONE (ACQ+TMP+RWI+TRF): {int(info["PROD"])}</span></div>', unsafe_allow_html=True)
            for _, r in g.iterrows():
                tag = "📋 [PREV]" if r['Codice Documento'] == "OCA" else "🛒 [ORD]"
                st.markdown(f'<div class="status-row {r["CS"]}"><span>{tag} 📅 {r["DT_EXP"].strftime("%d/%m/%Y")} | Q: {int(r["Qta Residua"])} | {r["CLI_NAME"]}</span><span><b>{r["ST"]}</b></span></div>', unsafe_allow_html=True)
