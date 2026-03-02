import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO
import plotly.express as px # Libreria per i grafici

# --- 1. CONFIGURAZIONE E STILE ---
APP_VERSION = "1.5.0-Beta-Admin"
st.set_page_config(page_title=f"Safit Portal v{APP_VERSION}", layout="wide")

st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 10px; border-radius: 8px; margin-bottom: 6px; border: 1px solid #ddd; color: #000 !important; font-size: 14px; }
    .on-time-row { background-color: #e8f5e9 !important; border-left: 6px solid #4caf50; } 
    .acq-row { background-color: #e3f2fd !important; border-left: 6px solid #2196f3; }    
    .prod-row { background-color: #fffde7 !important; border-left: 8px solid #fbc02d; }   
    .urgent-row { background-color: #ffebee !important; border-left: 8px solid #f44336; } 
    .debug-box { background-color: #f0f2f6 !important; color: #111 !important; padding: 8px 12px; border-radius: 6px; border: 1px solid #ccc; margin-bottom: 10px; font-family: sans-serif; font-size: 13px; font-weight: 600; display: flex; justify-content: space-between; white-space: nowrap; overflow-x: auto; }
    .kpi-card { background-color: #ffffff; border: 1px solid #e0e0e0; padding: 15px; border-radius: 10px; text-align: center; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .kpi-val { font-size: 24px; font-weight: bold; color: #1f77b4; }
    .kpi-lab { font-size: 12px; color: #666; text-transform: uppercase; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNZIONI UTILI ---
def to_excel_full(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.drop(columns=['cs'], errors='ignore').to_excel(writer, index=False, sheet_name='Piano_Consegne')
    return output.getvalue()

def clean_num(serie):
    s = serie.astype(str).str.replace(' ', '').str.replace('\xa0', '')
    def fix_val(val):
        if val.lower() in ['nan', '', 'none']: return '0'
        if ',' in val and '.' in val: return val.replace('.', '').replace(',', '.')
        elif ',' in val: return val.replace(',', '.')
        return val
    return pd.to_numeric(s.apply(fix_val), errors='coerce').fillna(0)

def find_col_exact(df, target):
    for c in df.columns:
        if str(c).strip().upper() == target.upper(): return c
    return None

def find_col(df, targets):
    for c in df.columns:
        if any(t.upper() in str(c).upper() for t in targets): return c
    return None

def smart_load(filename):
    df_p = pd.read_excel(filename, header=None, nrows=15)
    h_row = 0
    for i, row in df_p.iterrows():
        row_s = " ".join([str(x) for x in row.values])
        if any(k in row_s for k in ["Articolo", "Cliente", "Data", "CODICE"]):
            h_row = i; break
    df = pd.read_excel(filename, skiprows=h_row)
    df.columns = [str(c).strip() for c in df.columns]
    return df

# --- 3. CARICAMENTO DATI ---
@st.cache_data
def load_all():
    try:
        df_a = smart_load('righe_Ordini_ARCA.xlsx')
        c_cli = find_col(df_a, ['Cliente Fornitore', 'CD'])
        c_art = find_col(df_a, ['Articolo C', 'Cod. Art'])
        c_des = find_col(df_a, ['Articolo D', 'Descriz'])
        c_dat = find_col(df_a, ['Data'])
        c_qta = find_col(df_a, ['Qta Residua', 'Qta Doc'])
        
        for col in [c_cli, c_art, c_des]:
            if col: df_a[col] = df_a[col].ffill()
        
        df_a = df_a.dropna(subset=[c_art, c_dat])
        df_a['Art_Key'] = df_a[c_art].astype(str).str.strip().str.upper()
        df_a['Desc_Full'] = df_a[c_des].astype(str).fillna("N.D.")
        df_a['Data_Dt'] = pd.to_datetime(df_a[c_dat], errors='coerce')
        df_a['Qta_Res'] = clean_num(df_a[c_qta])

        if os.path.exists('Avanzamento_access.xlsx'):
            df_t = smart_load('Avanzamento_access.xlsx')
            c_code = find_col_exact(df_t, 'CODICE')
            c_gia = find_col_exact(df_t, 'GIA')
            c_acq = find_col_exact(df_t, 'INACQ')
            if c_code:
                df_t['Key_Acc'] = df_t[c_code].astype(str).str.strip().str.upper()
                df_t['GIA_V'] = clean_num(df_t[c_gia]) if c_gia else 0
                df_t['ACQ_V'] = clean_num(df_t[c_acq]) if c_acq else 0
                df_t['PROD_V'] = 0
                for f in ['LANCIATI', 'GRZ', 'TMP', 'RWI', 'TRS']:
                    cf = find_col_exact(df_t, f)
                    if cf: df_t['PROD_V'] += clean_num(df_t[cf])
                return pd.merge(df_a, df_t[['Key_Acc', 'GIA_V', 'ACQ_V', 'PROD_V']], left_on='Art_Key', right_on='Key_Acc', how='left').fillna(0)
        return df_a
    except Exception as e:
        st.error(f"Errore: {e}"); return pd.DataFrame()

# --- 4. GESTIONE UTENTI ---
@st.cache_data
def get_user_db():
    if os.path.exists('utenti.xlsx'):
        try:
            df_u = pd.read_excel('utenti.xlsx')
            return df_u.set_index('username')[['password', 'cliente_arca']].T.to_dict('list')
        except: pass
    return {'safit_admin': ['admin2026', 'TUTTI']}

USER_DB = get_user_db()
if "authenticated" not in st.session_state: st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=300)
        st.title("Safit Portal - Login")
        u = st.text_input("Username").strip()
        p = st.text_input("Password", type="password").strip()
        if st.button("Accedi"):
            if u in USER_DB and str(USER_DB[u][0]) == p:
                st.session_state.update({"authenticated": True, "user_type": USER_DB[u][1], "username": u})
                st.rerun()
            else: st.error("Credenziali errate")
    st.stop()

# --- 5. ELABORAZIONE ---
data = load_all()
if not data.empty:
    c_cli_n = find_col(data, ['Cliente Fornitore', 'CD'])
    
    # 1. Filtro Cliente
    if st.session_state["user_type"] == "TUTTI":
        with st.sidebar:
            st.markdown("### 🛠️ Amministrazione")
            sel_cli = st.selectbox("Seleziona Cliente:", ["TUTTI I CLIENTI"] + sorted(data[c_cli_n].unique().astype(str)))
    else:
        sel_cli = st.session_state["user_type"]

    df_filtered = data if sel_cli == "TUTTI I CLIENTI" else data[data[c_cli_n] == sel_cli]

    # 2. Logica ATP Globale
    results = []
    for art in df_filtered['Art_Key'].unique():
        df_art = df_filtered[df_filtered['Art_Key'] == art].copy().sort_values('Data_Dt')
        m, a, p = float(df_art['GIA_V'].iloc[0]), float(df_art['ACQ_V'].iloc[0]), float(df_art['PROD_V'].iloc[0])
        for _, r in df_art.iterrows():
            q, oggi = float(r['Qta_Res']), datetime.now()
            if m >= q: m -= q; s, c, d = "DISPONIBILE", "on-time-row", r['Data_Dt']
            elif (m+a) >= q: a -= (q-m); m=0; s, c, d = "ACQUISTO", "acq-row", oggi+timedelta(12)
            elif (m+a+p) >= q: p -= (q-m-a); m=0; a=0; s, c, d = "PRODUZIONE", "prod-row", oggi+timedelta(22)
            else: s, c, d = "MANCANTE", "urgent-row", oggi+timedelta(40)
            res = r.to_dict(); res.update({'st': s, 'cs': c, 'dt_e': d}); results.append(res)
    
    df_res = pd.DataFrame(results)

    # --- PARTE ADMIN: KPI & GRAFICI ---
    if st.session_state["user_type"] == "TUTTI":
        st.title(f"Dashboard Analitica: {sel_cli}")
        
        # Calcolo KPI
        tot_qta = df_res['Qta_Res'].sum()
        def get_perc(stato): return (df_res[df_res['st'] == stato]['Qta_Res'].sum() / tot_qta * 100) if tot_qta > 0 else 0

        k1, k2, k3, k4 = st.columns(4)
        with k1: st.markdown(f'<div class="kpi-card"><div class="kpi-lab">Pezzi Ordinati</div><div class="kpi-val">{tot_qta:,.0f}</div></div>', unsafe_allow_html=True)
        with k2: st.markdown(f'<div class="kpi-card"><div class="kpi-lab">Pronti %</div><div class="kpi-val" style="color:#4caf50">{get_perc("DISPONIBILE"):.1f}%</div></div>', unsafe_allow_html=True)
        with k3: st.markdown(f'<div class="kpi-card"><div class="kpi-lab">In Prod %</div><div class="kpi-val" style="color:#fbc02d">{get_perc("PRODUZIONE"):.1f}%</div></div>', unsafe_allow_html=True)
        with k4: st.markdown(f'<div class="kpi-card"><div class="kpi-lab">Mancanti %</div><div class="kpi-val" style="color:#f44336">{get_perc("MANCANTE"):.1f}%</div></div>', unsafe_allow_html=True)

        st.markdown("---")
        
        # GRAFICO A TORTA RAGGRUPPATO
        col_g1, col_g2 = st.columns([1, 1])
        with col_g1:
            st.subheader("Stato Ordini (Volume)")
            fig_st = px.pie(df_res, values='Qta_Res', names='st', color='st',
                         color_discrete_map={'DISPONIBILE':'#4caf50','ACQUISTO':'#2196f3','PRODUZIONE':'#fbc02d','MANCANTE':'#f44336'})
            st.plotly_chart(fig_st, use_container_width=True)

        with col_g2:
            st.subheader("Famiglie Articoli (Top 10)")
            # Raggruppamento per le prime 2 parole della descrizione
            df_res['Famiglia'] = df_res['Desc_Full'].apply(lambda x: " ".join(str(x).split()[:2]).upper())
            df_fam = df_res.groupby('Famiglia')['Qta_Res'].sum().reset_index().sort_values('Qta_Res', ascending=False).head(10)
            fig_fam = px.pie(df_fam, values='Qta_Res', names='Famiglia', hole=0.4)
            st.plotly_chart(fig_fam, use_container_width=True)

    # --- PARTE LISTA (Per tutti) ---
    with st.sidebar:
        st.markdown("---")
        lista_art = sorted(df_res['Art_Key'].unique().astype(str))
        search_art = st.selectbox("🔍 Cerca Articolo:", ["TUTTI"] + lista_art)
        f_disp, f_acq, f_prod, f_manc = st.checkbox("🟢 Magazzino", True), st.checkbox("🔵 Acquisti", True), st.checkbox("🟡 Produzione", True), st.checkbox("🔴 Mancante", True)
        if not df_res.empty:
            st.download_button("📊 Scarica Excel", data=to_excel_full(df_res), file_name=f"Report_{sel_cli}.xlsx")

    # Filtri finali
    df_show = df_res.copy()
    if search_art != "TUTTI": df_show = df_show[df_show['Art_Key'] == search_art]
    allowed = [s for s, f in zip(["DISPONIBILE", "ACQUISTO", "PRODUZIONE", "MANCANTE"], [f_disp, f_acq, f_prod, f_manc]) if f]
    df_show = df_show[df_show['st'].isin(allowed)]

    if st.session_state["user_type"] != "TUTTI":
        st.title(f"Piano Consegne: {sel_cli}")

    for art in sorted(df_show['Art_Key'].unique()):
        df_sub = df_show[df_show['Art_Key'] == art]
        with st.expander(f"📦 {art} - {df_sub['Desc_Full'].iloc[0]}"):
            st.markdown(f'''<div class="debug-box">
                <span>GIA: <b>{df_sub["GIA_V"].iloc[0]:,.0f}</b></span>
                <span>ACQ: <b>{df_sub["ACQ_V"].iloc[0]:,.0f}</b></span>
                <span>PROD: <b>{df_sub["PROD_V"].iloc[0]:,.0f}</b></span>
            </div>''', unsafe_allow_html=True)
            for _, r in df_sub.iterrows():
                st.markdown(f"""<div class="status-row {r['cs']}">
                    <span>📅 <b>{r['Data_Dt'].strftime('%d/%m/%Y')}</b> | Q: {r['Qta_Res']:,.0f}</span>
                    <span><b>{r['st']}</b> ({r['dt_e'].strftime('%d/%m/%Y')})</span>
                </div>""", unsafe_allow_html=True)
else:
    st.warning("Nessun dato caricato.")
