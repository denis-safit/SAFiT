import streamlit as st
import pandas as pd
import os
from datetime import datetime
from io import BytesIO
import plotly.express as px
import re
from bom_engine import get_coverage  

# --- 1. CONFIGURAZIONE ---
st.set_page_config(page_title="Safit Portal v3.8", layout="wide")

st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 10px; border-radius: 8px; margin-bottom: 6px; border: 1px solid #ddd; color: #000 !important; font-size: 14px; }
    .on-time-row { background-color: #e8f5e9 !important; border-left: 6px solid #4caf50; } 
    .acq-row { background-color: #e3f2fd !important; border-left: 6px solid #2196f3; }    
    .prod-row { background-color: #fffde7 !important; border-left: 6px solid #fbc02d; }   
    .urgent-row { background-color: #ffebee !important; border-left: 8px solid #f44336; } 
    .oca-row { background-color: #f5f5f5 !important; border-left: 8px solid #9e9e9e; color: #666 !important; }
    .bom-row { background-color: #f3e5f5 !important; border-left: 8px solid #9c27b0; }
    .debug-box { background-color: #f8f9fa !important; color: #333 !important; padding: 12px; border-radius: 8px; border: 1px dotted #bbb; margin-bottom: 10px; display: flex; justify-content: space-around; font-size: 13px; font-weight: bold; }
    .kpi-card { background-color: #ffffff; border: 1px solid #e0e0e0; padding: 15px; border-radius: 10px; text-align: center; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .kpi-val { font-size: 24px; font-weight: bold; color: #1f77b4; }
    .user-info { padding: 10px; background: #f8f9fa; border-radius: 5px; border: 1px solid #eee; margin-bottom: 20px; text-align: center; }
    </style>
    """, unsafe_allow_html=True)

# --- 2. FUNZIONI TECNICHE ---
@st.cache_data
def get_user_db():
    if os.path.exists('utenti.xlsx'):
        try:
            df_u = pd.read_excel('utenti.xlsx')
            df_u.columns = [str(c).strip() for c in df_u.columns]
            return df_u.set_index('username')[['password', 'cliente_arca']].T.to_dict('list')
        except: pass
    return {"safit_admin": ["admin2026", "TUTTI"]}

def clean_num(serie):
    s = serie.astype(str).str.replace(' ', '').str.replace('\xa0', '')
    def fix_val(val):
        if val.lower() in ['nan', '', 'none']: return '0'
        return val.replace('.', '').replace(',', '.') if ',' in val and '.' in val else val.replace(',', '.')
    return pd.to_numeric(s.apply(fix_val), errors='coerce').fillna(0)

def normalize_art_code(val):
    """
    Normalizza codici articolo per evitare mismatch tra Excel (spazi invisibili, NBSP, ecc).
    """
    if val is None:
        return ''
    s = str(val)
    # Rimuove NBSP e varianti di whitespace invisibili.
    s = s.replace('\xa0', ' ').replace('\u202f', ' ').replace('\u2009', ' ')
    # Rimuove caratteri invisibili/di controllo comuni (zero-width, BOM, joiners).
    s = (
        s.replace('\u200b', '')  # zero-width space
         .replace('\u200c', '') # zero-width non-joiner
         .replace('\u200d', '') # zero-width joiner
         .replace('\u2060', '') # word joiner
         .replace('\ufeff', '') # BOM
    )
    # Rimuove qualsiasi whitespace rimanente (anche unicode) e fa uppercase.
    s = re.sub(r'\s+', '', s).strip()
    return s.upper()

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

# --- 3. MOTORE DI CALCOLO ---
@st.cache_data(ttl=1800)
def load_and_process():
    try:
        df_arca = smart_load('righe_Ordini_ARCA.xlsx', "Articolo C")
        if df_arca.empty: return pd.DataFrame(), {}
        
        c_tipo, c_art, c_qta, c_dat, c_cli = "Codice Documento", "Articolo C", "Qta Residua", "Data Consegna", "Cliente Fornitore CD"
        df_arca[c_qta] = clean_num(df_arca[c_qta]).abs()
        df_arca[c_dat] = pd.to_datetime(df_arca[c_dat], errors='coerce')
        df_arca = df_arca[df_arca[c_tipo].isin(['OCI', 'OCA']) & (df_arca[c_qta] > 0)].dropna(subset=[c_dat, c_art])

        df_acc = smart_load('Avanzamento_access.xlsx', "CODICE")
        stock_map = {}
        if not df_acc.empty:
            df_acc.columns = [str(c).strip().upper() for c in df_acc.columns]
            # `PROD` deve rappresentare solo la produzione (non doppiare la parte di acquisto).
            prod_cols = ['LANCIATI', 'GRZ', 'TMP', 'RWI', 'TRS', 'TRF']
            for _, r in df_acc.iterrows():
                art_code = normalize_art_code(r.get('CODICE', ''))
                gia = clean_num(pd.Series([r.get('GIA', 0)])).iloc[0]
                # Se nel file coesistono entrambe le colonne, usiamo `INACQ` se valorizzata,
                # altrimenti `ACQ` (fallback).
                inacq_val = clean_num(pd.Series([r.get('INACQ', 0)])).iloc[0]
                acq_val = clean_num(pd.Series([r.get('ACQ', 0)])).iloc[0]
                acq = inacq_val if float(inacq_val) != 0 else acq_val
                prod = sum([clean_num(pd.Series([r.get(f, 0)])).iloc[0] for f in prod_cols])
                stock_map[art_code] = {'GIA': gia, 'ACQ': acq, 'PROD': prod, 'FIGLIO': str(r.get('FIGLIO', 'NAN'))}

        df_orders = df_arca.sort_values(by=[c_art, c_dat])
        final_results = []
        curr_stocks = {k: v.copy() for k, v in stock_map.items()}

        for index, row in df_orders.iterrows():
            art_code = normalize_art_code(row.get(c_art, ''))
            qta_ordine = float(row[c_qta])
            fonte = get_coverage(art_code, qta_ordine, curr_stocks)
            
            if fonte:
                if str(fonte).strip().upper() == art_code:
                    stato, colore = "DISPONIBILE", "on-time-row"
                else:
                    stato, colore = "COPERTO BOM", "bom-row"
            else:
                s = curr_stocks.get(art_code, {'GIA':0, 'ACQ': 0, 'PROD': 0})
                if (s['GIA'] + s['ACQ']) >= qta_ordine: stato, colore = "ACQUISTO", "acq-row"
                elif (s['GIA'] + s['ACQ'] + s['PROD']) >= qta_ordine: stato, colore = "PRODUZIONE", "prod-row"
                else: stato, colore = "MANCANTE", "urgent-row"
            
            if row[c_tipo] == 'OCA' and stato == "MANCANTE": stato, colore = "DA PIANIFICARE", "oca-row"
            
            res = row.to_dict()
            res.update({'ST': stato, 'CS': colore, 'ART_KEY': art_code, 'DT_EXP': row[c_dat], 'CLI_NAME': str(row[c_cli])})
            final_results.append(res)
                
        return pd.DataFrame(final_results), stock_map
    except Exception as e:
        st.error(f"Errore Motore: {e}")
        return pd.DataFrame(), {}

# --- 4. GESTIONE ACCESSO ---
if "auth" not in st.session_state: 
    st.session_state.auth = False

if not st.session_state.auth:
    USER_DB = get_user_db()
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=250)
        u = st.text_input("Username").strip()
        p = st.text_input("Password", type="password").strip()
        if st.button("Accedi", use_container_width=True):
            if u in USER_DB and str(USER_DB[u][0]) == p:
                st.session_state.auth = True
                st.session_state.user = u
                st.session_state.permesso = USER_DB[u][1]
                st.rerun()
            else:
                st.error("Credenziali errate.")
    st.stop()

# ===========================================================
# FUNZIONI VISTA CLIENTE
# ===========================================================
LABEL_CLI = {
    'DISPONIBILE':    ('Pronto per la spedizione', '#4caf50', 'on-time-row'),
    'COPERTO BOM':    ('Pronto (componente)',       '#4caf50', 'on-time-row'),
    'ACQUISTO':       ('In arrivo a magazzino',     '#2196f3', 'acq-row'),
    'PRODUZIONE':     ('In lavorazione',            '#fbc02d', 'prod-row'),
    'MANCANTE':       ('In pianificazione',         '#f44336', 'urgent-row'),
    'DA PIANIFICARE': ('Da confermare',             '#9e9e9e', 'oca-row'),
}

def pbar_html(pct, color):
    p = min(max(float(pct), 0), 100)
    inside  = str(round(p)) + "%" if p > 15 else ""
    outside = str(round(p)) + "%" if p <= 15 else ""
    return (
        '<div style="background:#e9ecef;border-radius:20px;height:16px;width:100%;overflow:hidden;margin:4px 0 8px 0;">'
        '<div style="width:' + str(round(p,1)) + '%;background:' + color + ';height:100%;border-radius:20px;'
        'display:flex;align-items:center;justify-content:flex-end;padding-right:6px;'
        'font-size:10px;font-weight:700;color:#fff;box-sizing:border-box;">' + inside + '</div>'
        '</div>'
        '<div style="font-size:10px;color:#555;text-align:right;margin-top:-6px;">' + outside + '</div>'
    )

def render_vista_cliente(df_cli, stock_raw):
    """Vista pulita per il cliente — nessun dato interno visibile."""
    COLOR_MAP_CLI = {v[0]: v[1] for v in LABEL_CLI.values()}

    st.title("📦 I tuoi Ordini")

    if df_cli.empty:
        st.info("Nessun ordine aperto al momento.")
        return

    # KPI cliente
    tot_qta   = int(df_cli['Qta Residua'].sum())
    n_pronti  = int(df_cli[df_cli['ST'].isin(['DISPONIBILE','COPERTO BOM'])]['Qta Residua'].sum())
    n_lavoro  = int(df_cli[df_cli['ST'].isin(['ACQUISTO','PRODUZIONE'])]['Qta Residua'].sum())
    n_mancanti= int(df_cli[df_cli['ST'].isin(['MANCANTE','DA PIANIFICARE'])]['Qta Residua'].sum())
    pct_pronto = round(n_pronti / tot_qta * 100) if tot_qta > 0 else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f'<div class="kpi-card"><div style="font-size:11px">TOTALE PEZZI</div><div class="kpi-val">{tot_qta:,}</div></div>'.replace(",","."), unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#4caf50">PRONTI</div><div class="kpi-val">{n_pronti:,}</div></div>'.replace(",","."), unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#2196f3">IN LAVORAZIONE</div><div class="kpi-val">{n_lavoro:,}</div></div>'.replace(",","."), unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#f44336">DA PIANIFICARE</div><div class="kpi-val">{n_mancanti:,}</div></div>'.replace(",","."), unsafe_allow_html=True)

    # Grafico torta stato ordini cliente
    df_cli_chart = df_cli.copy()
    df_cli_chart['Stato Cliente'] = df_cli_chart['ST'].map(lambda x: LABEL_CLI.get(x, (x,'#aaa',''))[0])
    df_torta = df_cli_chart.groupby('Stato Cliente')['Qta Residua'].sum().reset_index()
    col_t, col_info = st.columns([1, 1])
    with col_t:
        fig_cli = px.pie(df_torta, values='Qta Residua', names='Stato Cliente',
                         color='Stato Cliente',
                         color_discrete_map={
                             'Pronto per la spedizione':'#4caf50',
                             'Pronto (componente)':     '#4caf50',
                             'In arrivo a magazzino':   '#2196f3',
                             'In lavorazione':          '#fbc02d',
                             'In pianificazione':       '#f44336',
                             'Da confermare':           '#9e9e9e',
                         })
        fig_cli.update_traces(textinfo='label+percent',
                              hovertemplate='<b>%{label}</b><br>Qta: %{value:,.0f}<extra></extra>')
        fig_cli.update_layout(margin=dict(t=10,b=0,l=0,r=0), showlegend=False, height=280)
        st.plotly_chart(fig_cli, use_container_width=True)
    with col_info:
        st.markdown("##### Avanzamento complessivo")
        st.markdown(f"**{pct_pronto}%** degli ordini è pronto per la spedizione")
        st.markdown(pbar_html(pct_pronto, '#4caf50'), unsafe_allow_html=True)
        st.markdown("---")
        st.caption("🟢 Pronto — disponibile in magazzino")
        st.caption("🔵 In arrivo — materiale in fase di acquisto")
        st.caption("🟡 In lavorazione — in produzione")
        st.caption("🔴 In pianificazione — da programmare")

    st.markdown("---")

    # Dettaglio ordini per articolo
    for art, g in df_cli.groupby('ART_KEY'):
        desc     = g['Articolo D'].iloc[0]
        qta_tot  = int(g['Qta Residua'].sum())
        stati_g  = g['ST'].tolist()
        # Colore expander = stato peggiore
        if 'MANCANTE' in stati_g or 'DA PIANIFICARE' in stati_g:
            badge = "🔴"
        elif 'PRODUZIONE' in stati_g or 'ACQUISTO' in stati_g:
            badge = "🟡"
        else:
            badge = "🟢"

        with st.expander(f"{badge} {art} — {desc} | {qta_tot:,} pz".replace(",",".")):
            # Barra avanzamento articolo
            # "Disponibile" per il cliente deve riflettere quanta parte della richiesta
            # può essere coperta dalla sola GIACENZA (GIA), anche se la riga è classificata
            # come "ACQUISTO" perché la GIA non è sufficiente a coprire tutto.
            s_i = stock_raw.get(normalize_art_code(art), {'GIA': 0, 'ACQ': 0, 'PROD': 0})
            gia_left = float(s_i.get('GIA', 0))
            qta_pronta = 0.0
            g_sorted = g.sort_values(by='DT_EXP') if 'DT_EXP' in g.columns else g
            for _, r in g_sorted.iterrows():
                q = float(r.get('Qta Residua', 0))
                take = min(gia_left, q)
                qta_pronta += take
                gia_left -= take
                if gia_left <= 0:
                    break
            qta_pronta = int(round(qta_pronta))
            pct_art    = round(qta_pronta / qta_tot * 100) if qta_tot > 0 else 0
            bar_col    = '#4caf50' if pct_art >= 100 else ('#fbc02d' if pct_art > 0 else '#f44336')
            st.markdown(
                f"**Disponibile: {qta_pronta:,} / {qta_tot:,} pz**".replace(",","."),
            )
            st.markdown(pbar_html(pct_art, bar_col), unsafe_allow_html=True)

            # Debug cliente: usa la stessa logica admin su `stock_raw`
            # Usiamo `st.caption` invece di HTML per renderlo sempre visibile in UI.
            st.caption(f"DEBUG stock -> GIA: {int(s_i['GIA'])} | ACQ: {int(s_i['ACQ'])} | PROD: {int(s_i['PROD'])}")

            # Righe ordine
            for _, r in g.iterrows():
                testo, colore, css = LABEL_CLI.get(r['ST'], (r['ST'], '#aaa', 'oca-row'))
                data_str = r['DT_EXP'].strftime("%d/%m/%Y") if pd.notnull(r['DT_EXP']) else "N.D."
                st.markdown(
                    f'<div class="status-row {css}">'
                    f'<span>📅 Consegna: <b>{data_str}</b> | Q.tà: <b>{int(r["Qta Residua"]):,}</b></span>'.replace(",",".")
                    + f'<span><b>{testo}</b></span>'
                    f'</div>',
                    unsafe_allow_html=True
                )

# ===========================================================
# --- 5. DASHBOARD ---
# ===========================================================
df_res, stock_raw = load_and_process()
if 'last_update' not in st.session_state:
    st.session_state.last_update = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

if not df_res.empty:
    # --- SIDEBAR COMUNE ---
    with st.sidebar:
        if os.path.exists('Logo SAFIT.JPG'):
            st.image('Logo SAFIT.JPG', use_container_width=True)
        st.markdown(f'<div class="user-info">👤 <b>{st.session_state.user}</b></div>', unsafe_allow_html=True)
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("🚪 Esci", use_container_width=True):
                st.session_state.auth = False
                st.rerun()
        with col_b:
            if st.button("🔄 Aggiorna", use_container_width=True):
                st.cache_data.clear()
                st.session_state.last_update = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
                st.rerun()
        st.caption("📅 Dati al: " + st.session_state.get('last_update', '--'))
        st.markdown("---")

        is_admin = st.session_state.permesso == "TUTTI"
        if is_admin:
            sel_cli = st.selectbox("Seleziona Cliente:", ["TUTTI"] + sorted(df_res['CLI_NAME'].unique().tolist()))
        else:
            sel_cli = st.session_state.permesso

        search = st.text_input("🔍 Cerca Articolo:").upper()

    df_f = df_res[df_res['CLI_NAME'] == sel_cli] if sel_cli != "TUTTI" else df_res.copy()
    if search: df_f = df_f[df_f['ART_KEY'].str.contains(search)]

    # ===========================================================
    # VISTA CLIENTE
    # ===========================================================
    if not is_admin:
        render_vista_cliente(df_f, stock_raw)
        st.stop()

    # ===========================================================
    # VISTA ADMIN (tutto il pannello originale)
    # ===========================================================
    st.sidebar.download_button("📊 Esporta Report", data=to_excel(df_f),
                               file_name=f"Safit_Report_{datetime.now().strftime('%d%m')}.xlsx",
                               use_container_width=True)

    st.title("Pannello Controllo Safit")
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f'<div class="kpi-card"><div style="font-size:11px">FILTRATI</div><div class="kpi-val">{int(df_f["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div style="font-size:11px; color:#4caf50">PRONTI (GIA)</div><div class="kpi-val">{int(df_f[df_f["ST"]=="DISPONIBILE"]["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card"><div style="font-size:11px; color:#9c27b0">BOM (FIGLI)</div><div class="kpi-val">{int(df_f[df_f["ST"]=="COPERTO BOM"]["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi-card"><div style="font-size:11px; color:#f44336">MANCANTI</div><div class="kpi-val">{int(df_f[df_f["ST"]=="MANCANTE"]["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)

    df_f = df_f.copy()
    df_f['Famiglia'] = df_f['Articolo D'].apply(lambda x: " ".join(str(x).split()[:2]).upper())

    COLOR_MAP = {'DISPONIBILE':'#4caf50','COPERTO BOM':'#9c27b0','ACQUISTO':'#2196f3',
                 'PRODUZIONE':'#fbc02d','MANCANTE':'#f44336','DA PIANIFICARE':'#9e9e9e'}

    if 'filtro_stati' not in st.session_state:    st.session_state.filtro_stati    = []
    if 'filtro_famiglie' not in st.session_state: st.session_state.filtro_famiglie = []

    df_fam_chart   = df_f.groupby('Famiglia')['Qta Residua'].sum().reset_index().sort_values('Qta Residua', ascending=False).head(10)
    df_stato_chart = df_f.groupby('ST')['Qta Residua'].sum().sort_values(ascending=False).reset_index()
    stati_disponibili    = df_stato_chart['ST'].tolist()
    famiglie_disponibili = df_fam_chart['Famiglia'].tolist()

    st.markdown("""<style>
    div[data-testid="stButton"] button {
        border-radius:6px!important;font-size:12px!important;padding:4px 6px!important;
        width:100%!important;white-space:nowrap!important;overflow:hidden!important;text-overflow:ellipsis!important;
    }</style>""", unsafe_allow_html=True)

    col_fs, col_g1, col_g2, col_ff = st.columns([1, 2, 2, 1])
    with col_fs:
        st.markdown("**Stato**")
        st.caption("✓ = attivo")
        for stato in stati_disponibili:
            attivo = stato in st.session_state.filtro_stati
            label  = f"✓ {stato}" if attivo else stato
            if st.button(label, key=f"btn_stato_{stato}", use_container_width=True,
                         type="primary" if attivo else "secondary"):
                if attivo: st.session_state.filtro_stati.remove(stato)
                else:      st.session_state.filtro_stati.append(stato)
                st.rerun()

    with col_g1:
        st.markdown("##### Stato Copertura")
        fig_stato = px.pie(df_stato_chart, values='Qta Residua', names='ST', color='ST', color_discrete_map=COLOR_MAP)
        fig_stato.update_traces(textinfo='label+percent', hovertemplate='<b>%{label}</b><br>Qta: %{value:,.0f}<br>%{percent}<extra></extra>')
        fig_stato.update_layout(margin=dict(t=10,b=0,l=0,r=0), showlegend=False, height=320)
        st.plotly_chart(fig_stato, use_container_width=True)

    with col_g2:
        st.markdown("##### Top 10 Famiglie")
        fig_fam = px.pie(df_fam_chart, values='Qta Residua', names='Famiglia', hole=0.35)
        fig_fam.update_traces(textinfo='label+percent', hovertemplate='<b>%{label}</b><br>Qta: %{value:,.0f}<br>%{percent}<extra></extra>', sort=True)
        fig_fam.update_layout(margin=dict(t=10,b=0,l=0,r=0), showlegend=False, height=320)
        st.plotly_chart(fig_fam, use_container_width=True)

    with col_ff:
        st.markdown("**Famiglia**")
        st.caption("✓ = attivo")
        for fam in famiglie_disponibili:
            attivo = fam in st.session_state.filtro_famiglie
            label_short = fam[:14] + "…" if len(fam) > 14 else fam
            label = f"✓ {label_short}" if attivo else label_short
            if st.button(label, key=f"btn_fam_{fam}", use_container_width=True,
                         type="primary" if attivo else "secondary"):
                if attivo: st.session_state.filtro_famiglie.remove(fam)
                else:      st.session_state.filtro_famiglie.append(fam)
                st.rerun()

    with st.sidebar:
        st.markdown("---")
        if st.button("🔄 Reset tutti i filtri", use_container_width=True, key="btn_reset_grafici"):
            st.session_state.filtro_stati    = []
            st.session_state.filtro_famiglie = []
            st.rerun()
        if st.session_state.filtro_stati:
            st.info("🔵 Stati: " + ", ".join(st.session_state.filtro_stati))
        if st.session_state.filtro_famiglie:
            st.info("📂 Famiglie: " + ", ".join(st.session_state.filtro_famiglie))

    df_view = df_f.copy()
    if st.session_state.filtro_famiglie:
        df_view = df_view[df_view['Famiglia'].isin(st.session_state.filtro_famiglie)]
    if st.session_state.filtro_stati:
        df_view = df_view[df_view['ST'].isin(st.session_state.filtro_stati)]

    filtri_attivi = []
    if st.session_state.filtro_famiglie: filtri_attivi.append(f"Famiglie: **{', '.join(st.session_state.filtro_famiglie)}**")
    if st.session_state.filtro_stati:    filtri_attivi.append(f"Stati: **{', '.join(st.session_state.filtro_stati)}**")
    if filtri_attivi:
        st.info("🔍 Filtri attivi — " + " | ".join(filtri_attivi) + f" — {len(df_view)} ordini trovati")
    else:
        st.caption(f"Tutti gli ordini: {len(df_view)}")

    st.markdown("---")
    for art, g in df_view.groupby('ART_KEY'):
        with st.expander(f"📦 {art} - {g['Articolo D'].iloc[0]} ({len(g)} ordini)"):
            s_i = stock_raw.get(normalize_art_code(art), {'GIA': 0, 'ACQ': 0, 'PROD': 0})
            st.markdown(f'<div class="debug-box"><span>📦 GIA: {int(s_i["GIA"])}</span><span>🚚 ACQ: {int(s_i["ACQ"])}</span><span>⚙️ PROD: {int(s_i["PROD"])}</span></div>', unsafe_allow_html=True)
            for _, r in g.iterrows():
                tag = "📋 [PREV]" if r['Codice Documento'] == "OCA" else "🛒 [ORD]"
                st.markdown(f'<div class="status-row {r["CS"]}"><span>{tag} 📅 {r["DT_EXP"].strftime("%d/%m/%Y")} | Q: {int(r["Qta Residua"])} | {r["CLI_NAME"]}</span><span><b>{r["ST"]}</b></span></div>', unsafe_allow_html=True)
