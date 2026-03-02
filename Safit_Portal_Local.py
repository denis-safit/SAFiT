import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
from io import BytesIO

# --- 1. CONFIGURAZIONE E STILE ---
APP_VERSION = "1.3.6"
st.set_page_config(page_title=f"Safit Portal v{APP_VERSION}", layout="wide")

st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 12px; border-radius: 10px; margin-bottom: 8px; border: 1px solid #ddd; }
    .on-time-row { background-color: #e8f5e9; border-left: 8px solid #4caf50; color: #1b5e20; } 
    .acq-row { background-color: #e3f2fd; border-left: 8px solid #2196f3; color: #0d47a1; }    
    .prod-row { background-color: #fffde7; border-left: 8px solid #fbc02d; color: #5d4037; }   
    .urgent-row { background-color: #ffebee; border-left: 8px solid #f44336; color: #b71c1c; } 
    .debug-box { background-color: #f8f9fa; padding: 10px; border-radius: 5px; border: 1px dashed #ccc; margin-bottom: 10px; font-family: monospace; font-size: 13px; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNZIONE EXPORT (Ripristinata) ---
def to_excel_full(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export = df.copy()
        # Pulizia per export
        if 'cs' in df_export.columns: df_export = df_export.drop(columns=['cs'])
        df_export.to_excel(writer, index=False, sheet_name='Piano_Consegne')
    return output.getvalue()

# --- 3. GESTIONE UTENTI ---
@st.cache_data
def get_user_db():
    if os.path.exists('utenti.xlsx'):
        try:
            df_u = pd.read_excel('utenti.xlsx')
            df_u.columns = [str(c).strip() for c in df_u.columns]
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

# --- 4. MOTORE DI CARICAMENTO ---
def clean_num(serie):
    s = serie.astype(str).str.replace(' ', '').str.replace('\xa0', '')
    def fix_val(val):
        if val.lower() in ['nan', '', 'none']: return '0'
        if ',' in val and '.' in val: return val.replace('.', '').replace(',', '.')
        elif ',' in val: return val.replace(',', '.')
        return val
    return pd.to_numeric(s.apply(fix_val), errors='coerce').fillna(0)

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

@st.cache_data
def load_all():
    try:
        # Caricamento ARCA
        df_a = smart_load('righe_Ordini_ARCA.xlsx')
        c_cli, c_art, c_des = find_col(df_a, ['Cliente Fornitore', 'CD']), find_col(df_a, ['Articolo C', 'Cod. Art']), find_col(df_a, ['Articolo D', 'Descriz'])
        c_dat, c_qta = find_col(df_a, ['Data']), find_col(df_a, ['Qta Residua', 'Qta Doc'])
        for col in [c_cli, c_art, c_des]:
            if col: df_a[col] = df_a[col].ffill()
        df_a = df_a.dropna(subset=[c_art, c_dat])
        df_a['Art_Key'] = df_a[c_art].astype(str).str.strip().str.upper()
        df_a['Data_Dt'] = pd.to_datetime(df_a[c_dat], errors='coerce')
        df_a['Qta_Res'] = clean_num(df_a[c_qta])

        # Caricamento ACCESS (Nuove Colonne)
        if os.path.exists('Avanzamento_access.xlsx'):
            df_t = smart_load('Avanzamento_access.xlsx')
            c_code = find_col(df_t, ['CODICE', 'Codice'])
            if c_code:
                df_t['Key_Acc'] = df_t[c_code].astype(str).str.strip().str.upper()
                df_t['GIA_V'] = clean_num(df_t[find_col(df_t, ['GIA'])])
                df_t['ACQ_V'] = clean_num(df_t[find_col(df_t, ['INACQ'])])
                df_t['PROD_V'] = 0
                for f in ['LANCIATI', 'GRZ', 'TMP', 'RWI', 'TRS']:
                    cf = find_col(df_t, [f])
                    if cf: df_t['PROD_V'] += clean_num(df_t[cf])
                return pd.merge(df_a, df_t[['Key_Acc', 'GIA_V', 'ACQ_V', 'PROD_V']], left_on='Art_Key', right_on='Key_Acc', how='left').fillna(0)
        return df_a
    except Exception as e:
        st.error(f"Errore caricamento: {e}"); return pd.DataFrame()

# --- 5. INTERFACCIA E LOGICA ---
data = load_all()
if not data.empty:
    c_cli_n = find_col(data, ['Cliente Fornitore'])
    
    with st.sidebar:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
        st.markdown("### 🛠️ Pannello Controllo")
        
        if st.session_state["user_type"] == "TUTTI":
            sel_cli = st.selectbox("Seleziona Cliente:", sorted(data[c_cli_n].unique().astype(str)))
        else:
            sel_cli = st.session_state["user_type"]
        
        df_cli = data[data[c_cli_n] == sel_cli].copy()

        # Ricerca Articolo Protetta
        lista_art = sorted(df_cli['Art_Key'].unique().astype(str))
        search_art = st.selectbox("🔍 Cerca Articolo:", ["TUTTI"] + lista_art)
        if search_art == "TUTTI": search_art = ""

        st.markdown("---")
        st.markdown("#### 📡 Stato Disponibilità")
        f_disp, f_acq, f_prod, f_manc = st.checkbox("🟢 Magazzino", True), st.checkbox("🔵 Acquisti", True), st.checkbox("🟡 Produzione", True), st.checkbox("🔴 Mancante", True)

    # Logica ATP
    results = []
    for art in df_cli['Art_Key'].unique():
        df_art = df_cli[df_cli['Art_Key'] == art].copy().sort_values('Data_Dt')
        m, a, p = float(df_art['GIA_V'].iloc[0]), float(df_art['ACQ_V'].iloc[0]), float(df_art['PROD_V'].iloc[0])
        for _, r in df_art.iterrows():
            q, oggi = float(r['Qta_Res']), datetime.now()
            if m >= q: m -= q; s, c, d = "DISPONIBILE", "on-time-row", r['Data_Dt']
            elif (m+a) >= q: a -= (q-m); m=0; s, c, d = "ACQUISTO", "acq-row", oggi+timedelta(12)
            elif (m+a+p) >= q: p -= (q-m-a); m=0; a=0; s, c, d = "PRODUZIONE", "prod-row", oggi+timedelta(22)
            else: s, c, d = "MANCANTE", "urgent-row", oggi+timedelta(40)
            res = r.to_dict(); res.update({'st': s, 'cs': c, 'dt_e': d}); results.append(res)
    
    df_f = pd.DataFrame(results)
    if search_art: df_f = df_f[df_f['Art_Key'] == search_art]
    
    allowed = []
    if f_disp: allowed.append("DISPONIBILE")
    if f_acq: allowed.append("ACQUISTO")
    if f_prod: allowed.append("PRODUZIONE")
    if f_manc: allowed.append("MANCANTE")
    df_show = df_f[df_f['st'].isin(allowed)]

    # PULSANTE DOWNLOAD (Ripristinato)
    with st.sidebar:
        st.markdown("---")
        if not df_show.empty:
            st.download_button("📊 Scarica Report Excel", data=to_excel_full(df_show), file_name=f"Report_{sel_cli}.xlsx")

    st.title(f"Piano Consegne: {sel_cli}")
    for art in sorted(df_show['Art_Key'].unique()):
        df_sub = df_show[df_show['Art_Key'] == art]
        c_des_f = find_col(df_sub, ['Articolo D', 'Descriz'])
        desc_val = df_sub[c_des_f].iloc[0] if c_des_f else "N.D."
        with st.expander(f"📦 {art} - {desc_val}"):
            st.markdown(f'<div class="debug-box">Giacenza: {df_sub["GIA_V"].iloc[0]:,.0f} | In Arrivo: {df_sub["ACQ_V"].iloc[0]:,.0f} | Produzione: {df_sub["PROD_V"].iloc[0]:,.0f}</div>', unsafe_allow_html=True)
            for _, r in df_sub.iterrows():
                st.markdown(f"""<div class="status-row {r['cs']}">
                    <span>📅 <b>{r['Data_Dt'].strftime('%d/%m/%Y')}</b> | Q.tà: {r['Qta_Res']:,.0f}</span>
                    <span><b>{r['st']}</b> (Est: {r['dt_e'].strftime('%d/%m/%Y')})</span>
                </div>""", unsafe_allow_html=True)
else:
    st.warning("Nessun dato caricato.")
