import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO
import plotly.express as px

# --- 1. CONFIGURAZIONE E STILE ---
APP_VERSION = "1.7.2-Final-Check"
st.set_page_config(page_title=f"Safit Portal v{APP_VERSION}", layout="wide")

st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 10px; border-radius: 8px; margin-bottom: 6px; border: 1px solid #ddd; color: #000 !important; font-size: 14px; }
    .on-time-row { background-color: #e8f5e9 !important; border-left: 6px solid #4caf50; } 
    .acq-row { background-color: #e3f2fd !important; border-left: 8px solid #2196f3; }    
    .prod-row { background-color: #fffde7 !important; border-left: 8px solid #fbc02d; }   
    .urgent-row { background-color: #ffebee !important; border-left: 8px solid #f44336; } 
    .oca-row { background-color: #f5f5f5 !important; border-left: 8px solid #9e9e9e; color: #666 !important; }
    .debug-box { background-color: #f0f2f6 !important; color: #111 !important; padding: 8px 12px; border-radius: 6px; border: 1px solid #ccc; margin-bottom: 10px; font-family: sans-serif; font-size: 13px; font-weight: 600; display: flex; justify-content: space-between; }
    .kpi-card { background-color: #ffffff; border: 1px solid #e0e0e0; padding: 15px; border-radius: 10px; text-align: center; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .kpi-val { font-size: 24px; font-weight: bold; color: #1f77b4; }
    .kpi-lab { font-size: 12px; color: #666; text-transform: uppercase; }
    .user-info { padding: 10px; background: #f8f9fa; border-radius: 5px; border: 1px solid #eee; margin-bottom: 20px; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNZIONI SUPPORTO ---
def fmt_n(val):
    try: return f"{int(round(float(val), 0)):,}".replace(",", ".")
    except: return "0"

def clean_num(serie):
    s = serie.astype(str).str.replace(' ', '').str.replace('\xa0', '')
    def fix_val(val):
        if val.lower() in ['nan', '', 'none']: return '0'
        return val.replace('.', '').replace(',', '.') if ',' in val else val
    return pd.to_numeric(s.apply(fix_val), errors='coerce').fillna(0)

def to_excel_full(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.drop(columns=['CS'], errors='ignore').to_excel(writer, index=False, sheet_name='Piano_Consegne')
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

# --- 3. MOTORE LOGICO ATP ---
@st.cache_data(ttl=300)
def load_and_process():
    try:
        df_full = smart_load('righe_Ordini_ARCA.xlsx', "Articolo C")
        if df_full.empty: return pd.DataFrame(), {}
        
        c_tipo, c_art, c_des, c_qta, c_dat, c_cli = "Codice Documento", "Articolo C", "Articolo D", "Qta Residua", "Data Consegna", "Cliente Fornitore CD"
        
        # FIX DATE: Trasforma tutto in data vera, se fallisce mette NaT e poi cancelliamo quelle righe
        df_full[c_dat] = pd.to_datetime(df_full[c_dat], errors='coerce')
        df_full = df_full.dropna(subset=[c_dat, c_art])
        df_full[c_qta] = clean_num(df_full[c_qta])
        
        # Caricamento Access
        df_acc = smart_load('Avanzamento_access.xlsx', "CODICE")
        stock = {}
        if not df_acc.empty:
            for _, r in df_acc.iterrows():
                p = sum([clean_num(pd.Series([r.get(f, 0)])).iloc[0] for f in ['LANCIATI', 'GRZ', 'TMP', 'RWI', 'TRS'] if f in df_acc.columns])
                stock[str(r['CODICE']).strip().upper()] = {'GIA': clean_num(pd.Series([r.get('GIA', 0)])).iloc[0], 'PROD': p}

        # Gestione Arrivi (OFF/DCL)
        df_in = df_full[df_full[c_tipo].isin(['OFF', 'DCL'])].copy()
        arrivals = {art: g.sort_values(c_dat)[[c_dat, c_qta]].values.tolist() for art, g in df_in.groupby(c_art)}

        # Split e Ordinamento: OCI e OCA insieme, ordinati per data
        df_orders = df_full[df_full[c_tipo].isin(['OCI', 'OCA'])].sort_values([c_dat, c_tipo], ascending=[True, False])
        
        final = []
        for _, row in df_orders.iterrows():
            art = str(row[c_art]).strip().upper()
            qta = row[c_qta]
            s = stock.get(art, {'GIA': 0, 'PROD': 0})
            arr_list = arrivals.get(art, [])
            
            is_oca = (row[c_tipo] == 'OCA')
            st_v, col, dt_e = ("MANCANTE", "urgent-row", datetime.now() + timedelta(days=45)) if not is_oca else ("DA PIANIFICARE", "oca-row", row[c_dat])

            if s['GIA'] >= qta:
                s['GIA'] -= qta; st_v, col, dt_e = "DISPONIBILE", "on-time-row", row[c_dat]
            else:
                fabb = qta - s['GIA']; s['GIA'] = 0
                for a in arr_list:
                    if a[1] >= fabb:
                        a[1] -= fabb; st_v, col, dt_e = "ACQUISTO", "acq-row", a[0]; fabb = 0; break
                    else: fabb -= a[1]; a[1] = 0
                if fabb > 0 and s['PROD'] >= fabb:
                    s['PROD'] -= fabb; st_v, col, dt_e = "PRODUZIONE", "prod-row", datetime.now() + timedelta(days=21)

            res = row.to_dict()
            res.update({'ST': st_v, 'CS': col, 'DT_E': dt_e, 'ART_KEY': art, 'CLI_NAME': str(row[c_cli])})
            final.append(res)
        
        return pd.DataFrame(final), stock
    except Exception as e:
        st.error(f"Errore: {e}"); return pd.DataFrame(), {}

# --- 4. LOGIN ---
if "authenticated" not in st.session_state: st.session_state["authenticated"] = False
if not st.session_state["authenticated"]:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=300)
        st.title("Safit Portal - Login")
        u = st.text_input("Username").strip(); p = st.text_input("Password", type="password").strip()
        if st.button("Accedi", use_container_width=True):
            if (u == "safit_admin" and p == "admin2026") or u == "denis":
                st.session_state.update({"authenticated": True, "username": u, "user_type": "TUTTI"})
                st.rerun()
            else: st.error("Credenziali errate")
    st.stop()

# --- 5. DASHBOARD ---
df_res, stock_final = load_and_process()
if not df_res.empty:
    with st.sidebar:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
        st.markdown(f'<div class="user-info">👤 Utente: <b>{st.session_state["username"]}</b></div>', unsafe_allow_html=True)
        if st.button("🚪 Esci", type="primary", use_container_width=True):
            st.session_state.authenticated = False; st.rerun()
        st.markdown("---")
        list_cli = ["TUTTI"] + sorted(df_res['CLI_NAME'].unique().tolist())
        sel_cli = st.selectbox("Seleziona Cliente:", list_cli)
        search = st.text_input("🔍 Cerca Articolo:").upper()
        
        df_exp = df_res if sel_cli == "TUTTI" else df_res[df_res['CLI_NAME'] == sel_cli]
        st.download_button("📊 Scarica Excel", data=to_excel_full(df_exp), file_name=f"Report_{sel_cli}.xlsx", use_container_width=True)

    df_f = df_res if sel_cli == "TUTTI" else df_res[df_res['CLI_NAME'] == sel_cli]
    
    # KPI E GRAFICI
    st.title(f"Dashboard Safit: {sel_cli}")
    t_q = df_f['Qta Residua'].sum()
    def get_p(s): return (df_f[df_f['ST'] == s]['Qta Residua'].sum() / t_q * 100) if t_q > 0 else 0
    
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f'<div class="kpi-card"><div class="kpi-lab">Pezzi</div><div class="kpi-val">{fmt_n(t_q)}</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div class="kpi-lab">Pronti %</div><div class="kpi-val" style="color:#4caf50">{get_p("DISPONIBILE"):.1f}%</div></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card"><div class="kpi-lab">In Prod %</div><div class="kpi-val" style="color:#fbc02d">{get_p("PRODUZIONE"):.1f}%</div></div>', unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi-card"><div class="kpi-lab">Mancanti %</div><div class="kpi-val" style="color:#f44336">{get_p("MANCANTE"):.1f}%</div></div>', unsafe_allow_html=True)
    
    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(px.pie(df_f, values='Qta Residua', names='ST', color='ST', 
                               color_discrete_map={'DISPONIBILE':'#4caf50','ACQUISTO':'#2196f3','PRODUZIONE':'#fbc02d','MANCANTE':'#f44336','DA PIANIFICARE':'#9e9e9e'}), use_container_width=True)
    with c2:
        df_f['Famiglia'] = df_f['Articolo D'].apply(lambda x: " ".join(str(x).split()[:2]).upper())
        st.plotly_chart(px.pie(df_f.groupby('Famiglia')['Qta Residua'].sum().reset_index().sort_values('Qta Residua', ascending=False).head(10), values='Qta Residua', names='Famiglia', hole=0.4), use_container_width=True)

    st.markdown("---")
    
    # LISTA
    df_v = df_f[df_f['ART_KEY'].str.contains(search)] if search else df_f
    for art, g in df_v.groupby('ART_KEY'):
        with st.expander(f"📦 {art} - {g['Articolo D'].iloc[0]} ({len(g)} righe)"):
            info = stock_final.get(art, {'GIA': 0, 'PROD': 0})
            st.markdown(f'<div class="debug-box"><span>GIA: <b>{fmt_n(info["GIA"])}</b></span><span>PROD: <b>{fmt_n(info["PROD"])}</b></span></div>', unsafe_allow_html=True)
            for _, r in g.iterrows():
                tag = "📋 [PREV]" if str(r['Codice Documento']) == "OCA" else "🛒 [ORD]"
                st.markdown(f'<div class="status-row {r["CS"]}"><span>{tag} 📅 {r[c_dat].strftime("%d/%m/%Y")} | Q: {fmt_n(r["Qta Residua"])} | {r["CLI_NAME"]}</span><span><b>{r["ST"]}</b></span></div>', unsafe_allow_html=True)
else:
    st.warning("Dati non trovati. Verifica i file Excel.")
