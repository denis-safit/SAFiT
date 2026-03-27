import streamlit as st
import pandas as pd
import os
from datetime import datetime
from io import BytesIO
import plotly.express as px
import re
from bom_engine import get_coverage  
from kpi_avanzati import render_kpi_avanzati

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

# --- CONFIGURAZIONE BTL ---
ANTICIPO_BTL_GG      = 3   # giorni anticipo BTL (Barletta → Friola)
ANTICIPO_ATOPLAST_GG = 3   # giorni anticipo Atoplast → Friola
PATH_STORICO_DATE    = "righe_ordini_storico_con_date.xlsx"

# --- 2. FUNZIONI TECNICHE ---
@st.cache_data
def get_user_db():
    if os.path.exists('utenti.xlsx'):
        try:
            df_u = pd.read_excel('utenti.xlsx')
            df_u.columns = [str(c).strip() for c in df_u.columns]
            return df_u.set_index('username')[['password', 'cliente_arca']].T.to_dict('list')
        except: pass
    return {"safit_admin": ["admin2026", "TUTTI"], "btl": ["btl2026", "BTL"], "atoplast": ["atoplast2026", "ATOPLAST"]}

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

def to_excel_full(df):
    """Esporta l'intero DataFrame senza scartare colonne (export 'completo')."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
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
    # Rimuove righe completamente vuote che ffill potrebbe aver riempito
    df = df.dropna(how='all')
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
        # Rimuove righe duplicate esatte per evitare ordini doppi da export ARCA
        df_arca = df_arca.drop_duplicates()

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
                # Componenti produzione: servono anche per export (es. GRZ).
                prod_components = {}
                for f in prod_cols:
                    prod_components[f] = clean_num(pd.Series([r.get(f, 0)])).iloc[0]
                prod = sum(prod_components.values())
                stock_map[art_code] = {
                    'GIA': gia,
                    'ACQ': acq,
                    'PROD': prod,
                    'FIGLIO': str(r.get('FIGLIO', 'NAN')),
                    **prod_components,
                }

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
            res.update({'ST': stato, 'CS': colore, 'ART_KEY': art_code, 'DT_EXP': row[c_dat], 'CLI_NAME': str(row[c_cli]), 'DATA_ORD': pd.to_datetime(row.get('Data', None), errors='coerce')})
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

def tbar_html(data_ordine, data_consegna, oggi=None):
    """
    Barra temporale: mostra il progresso tra data ordine e data consegna.
    Verde = tempo trascorso, grigio = tempo rimanente, rosso = ritardo.
    """
    if oggi is None:
        oggi = datetime.now()

    # Gestione tipi
    try:
        d_ord  = pd.Timestamp(data_ordine)
        d_cons = pd.Timestamp(data_consegna)
        d_oggi = pd.Timestamp(oggi)
    except Exception:
        return ""

    if pd.isnull(d_ord) or pd.isnull(d_cons):
        return ""

    durata_tot = (d_cons - d_ord).days
    if durata_tot <= 0:
        durata_tot = 1

    giorni_passati  = (d_oggi - d_ord).days
    giorni_mancanti = (d_cons - d_oggi).days

    in_ritardo = d_oggi > d_cons

    if in_ritardo:
        # Tutta la barra è rossa + overflow
        ritardo_gg = (d_oggi - d_cons).days
        pct_verde  = 100
        colore_barra = "#f44336"
        label_stato  = f'<span style="color:#f44336;font-weight:700;font-size:11px;">⚠️ In ritardo di {ritardo_gg} giorni</span>'
    else:
        pct_verde    = min(round(giorni_passati / durata_tot * 100), 100)
        colore_barra = "#4caf50" if pct_verde < 85 else "#ff9800"
        label_stato  = f'<span style="color:#4caf50;font-size:11px;">✅ Mancano <b>{giorni_mancanti}</b> giorni</span>'

    str_ord  = d_ord.strftime("%d/%m/%y")
    str_cons = d_cons.strftime("%d/%m/%y")
    str_oggi = d_oggi.strftime("%d/%m/%y")

    # Marcatore "oggi" sulla barra
    marker_pct = min(max(pct_verde, 0), 98)

    return f"""
<div style="margin:6px 0 10px 0;">
  <div style="display:flex;justify-content:space-between;font-size:10px;
              color:#888;margin-bottom:3px;">
    <span>📋 Ordine: <b style="color:#333;">{str_ord}</b></span>
    <span>{label_stato}</span>
    <span>🎯 Consegna: <b style="color:#333;">{str_cons}</b></span>
  </div>
  <div style="position:relative;background:#e9ecef;border-radius:20px;
              height:18px;width:100%;overflow:visible;">
    <div style="width:{pct_verde}%;background:{colore_barra};height:100%;
                border-radius:20px;transition:width .3s;"></div>
    <div style="position:absolute;top:-3px;left:{marker_pct}%;
                transform:translateX(-50%);width:6px;height:24px;
                background:#1a1a2e;border-radius:3px;opacity:.7;"
         title="Oggi: {str_oggi}"></div>
  </div>
  <div style="display:flex;justify-content:space-between;
              font-size:10px;color:#aaa;margin-top:3px;">
    <span>Inizio</span>
    <span style="color:#555;">Oggi: {str_oggi}</span>
    <span>Scadenza</span>
  </div>
</div>
"""

@st.cache_data(ttl=1800)
def carica_btl_da_storico():
    """Carica OFR + OFF di BTL dal file storico con date."""
    path = PATH_STORICO_DATE  # file nella stessa cartella del portale
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        df = pd.read_excel(path, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        df['Codice Documento'] = df['Codice Documento'].ffill()
        df = df[~df['Codice Documento'].astype(str).str.contains('Totale|NaN|nan', na=True)]
        df = df[df['Articolo C'].notna() & (df['Articolo C'].astype(str) != '(vuoto)')]
        df['Data']          = pd.to_datetime(df['Data'],          errors='coerce')
        df['Data Consegna'] = pd.to_datetime(df['Data Consegna'], errors='coerce')
        df['Data Consegna'] = df.groupby('Numero Documento')['Data Consegna'].ffill()
        df['Qta Doc']       = pd.to_numeric(df['Qta Doc'], errors='coerce').fillna(0)
        df_btl = df[
            df['Cliente Fornitore CD'].str.contains('BTL', case=False, na=False) &
            df['Codice Documento'].isin(['OFR', 'OFF']) &
            (df['Qta Doc'] > 0)
        ].copy()
        df_btl['Tipo'] = df_btl['Codice Documento'].map({'OFR': '🔧 Lavorazione', 'OFF': '🛒 Acquisto'})
        return df_btl
    except Exception as e:
        return pd.DataFrame()


def render_vista_btl(df_res=None):
    """Vista BTL — OFR (lavorazioni) + OFF (acquisti). Usa Qta Residua da ARCA per quantità aperte."""
    st.title("🏭 Lavorazioni & Acquisti BTL")
    st.caption(f"Anticipo consegna a Friola: **{ANTICIPO_BTL_GG} giorni** prima della data consegna")

    df_btl_storico = carica_btl_da_storico()
    if df_btl_storico.empty:
        st.info("Nessun ordine BTL trovato. Verifica che il file righe_ordini_storico_con_date.xlsx sia presente.")
        return

    # Usa Qta Residua da df_res (ARCA) dove disponibile, altrimenti Qta Doc dallo storico
    if df_res is not None and not df_res.empty:
        # Articoli BTL con quantità residua aperta da ARCA
        art_btl = set(df_btl_storico['Articolo C'].unique())
        df_arca_btl = df_res[df_res['ART_KEY'].isin(art_btl)][['ART_KEY','Qta Residua']].groupby('ART_KEY')['Qta Residua'].sum().reset_index()
        df_arca_btl.columns = ['Articolo C', 'Qta Residua ARCA']
        df_btl = df_btl_storico.merge(df_arca_btl, on='Articolo C', how='left')
        # Se Qta Residua ARCA = 0 o NaN → ordine evaso, escludilo
        df_btl = df_btl[df_btl['Qta Residua ARCA'].fillna(0) > 0].copy()
        df_btl['Qta Doc'] = df_btl['Qta Residua ARCA']  # usa residua per i conteggi
    else:
        df_btl = df_btl_storico.copy()

    if df_btl.empty:
        st.success("✅ Nessuna lavorazione BTL aperta al momento.")
        return

    n_lav = df_btl[df_btl['Codice Documento']=='OFR']['Articolo C'].nunique()
    n_acq = df_btl[df_btl['Codice Documento']=='OFF']['Articolo C'].nunique()
    urgenti = df_btl[df_btl['Data Consegna'].notna() &
                     (df_btl['Data Consegna'] <= pd.Timestamp(datetime.now()) + pd.Timedelta(days=7))]
    k1, k2, k3 = st.columns(3)
    k1.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#fbc02d">🔧 IN LAVORAZIONE</div><div class="kpi-val">{n_lav}</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#2196f3">🛒 IN ACQUISTO</div><div class="kpi-val">{n_acq}</div></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#f44336">⚠️ URGENTI ≤7gg</div><div class="kpi-val">{urgenti["Articolo C"].nunique()}</div></div>', unsafe_allow_html=True)
    st.markdown("---")

    for art, g in df_btl.sort_values('Data Consegna', na_position='last').groupby('Articolo C', sort=False):
        desc    = g['Articolo D'].iloc[0]
        qta_tot = int(g['Qta Doc'].sum())
        tipi    = ' + '.join(g['Tipo'].unique().tolist())
        d_cons  = g['Data Consegna'].dropna().min() if g['Data Consegna'].notna().any() else None
        d_ord   = g['Data'].dropna().min()           if g['Data'].notna().any()          else None

        if d_cons is not None:
            d_friola      = d_cons - pd.Timedelta(days=ANTICIPO_BTL_GG)
            giorni_friola = (d_friola - pd.Timestamp(datetime.now())).days
            if giorni_friola < 0:
                badge, urgenza, col_u = "🔴", f"IN RITARDO di {abs(giorni_friola)} gg", "#f44336"
            elif giorni_friola <= 3:
                badge, urgenza, col_u = "🟠", f"URGENTE — {giorni_friola} gg", "#ff9800"
            elif giorni_friola <= 7:
                badge, urgenza, col_u = "🟡", f"A BREVE — {giorni_friola} gg", "#fbc02d"
            else:
                badge, urgenza, col_u = "🟢", f"Mancano {giorni_friola} gg", "#4caf50"
            data_label = f"📦 Friola: {d_friola.strftime('%d/%m/%Y')}"
        else:
            badge, urgenza, col_u, d_friola, data_label = "⚪", "Data N/D", "#9e9e9e", None, "📦 Data N/D"

        with st.expander(f"{badge} {art} — {desc} | {qta_tot:,} pz | {tipi} | {data_label}".replace(",",".")):
            c1, c2, c3 = st.columns(3)
            c1.metric("Quantità", f"{qta_tot:,} pz".replace(",","."))
            c2.metric("A Friola entro", d_friola.strftime("%d/%m/%Y") if d_friola else "N/D")
            c3.metric("Data consegna", d_cons.strftime("%d/%m/%Y") if d_cons else "N/D")
            if d_friola:
                st.markdown(f'<div style="background:#fff8e1;border-left:4px solid {col_u};padding:8px 14px;border-radius:6px;margin:6px 0;color:#1a1a1a!important;font-weight:600;">⏱️ {urgenza} alla consegna a Friola</div>', unsafe_allow_html=True)
                d_start = d_ord if d_ord is not None else d_friola - pd.Timedelta(days=30)
                st.markdown(tbar_html(d_start, d_friola), unsafe_allow_html=True)
            st.markdown("**Dettaglio righe:**")
            for _, r in g.sort_values('Data Consegna', na_position='last').iterrows():
                d_c  = r['Data Consegna']
                d_fs = (d_c - pd.Timedelta(days=ANTICIPO_BTL_GG)).strftime('%d/%m/%Y') if pd.notnull(d_c) else "N/D"
                d_cs = d_c.strftime('%d/%m/%Y') if pd.notnull(d_c) else "N/D"
                css  = 'prod-row' if 'Lavorazione' in str(r.get('Tipo','')) else 'acq-row'
                st.markdown(f'<div class="status-row {css}" style="color:#1a1a1a!important;"><span>{r.get("Tipo","")} | Q: <b>{int(r["Qta Doc"]):,}</b> pz | 📦 Friola: <b>{d_fs}</b> | Scad: {d_cs}</span></div>'.replace(",","."), unsafe_allow_html=True)

def render_vista_atoplast(df_res):
    """Vista dedicata ad Atoplast — solo articoli PRODUZIONE con codice PCP*****."""
    st.title("🏭 Lavorazioni Atoplast")
    st.caption(f"Anticipo consegna a Friola: **{ANTICIPO_ATOPLAST_GG} giorni** prima della data cliente")

    df_atp = df_res[
        (df_res['ST'] == 'PRODUZIONE') &
        (df_res['ART_KEY'].str.startswith('PCP', na=False))
    ].copy()

    if df_atp.empty:
        st.success("✅ Nessuna lavorazione Atoplast in corso al momento.")
        return

    k1, k2, k3 = st.columns(3)
    k1.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#9c27b0">ARTICOLI PCP IN LAVORAZIONE</div><div class="kpi-val">{df_atp["ART_KEY"].nunique()}</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#9c27b0">QUANTITÀ TOTALE</div><div class="kpi-val">{int(df_atp["Qta Residua"].sum()):,}</div></div>'.replace(",","."), unsafe_allow_html=True)
    urgenti = df_atp[df_atp['DT_EXP'] <= pd.Timestamp(datetime.now()) + pd.Timedelta(days=7)]
    k3.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#f44336">URGENTI ≤7gg</div><div class="kpi-val">{urgenti["ART_KEY"].nunique()}</div></div>', unsafe_allow_html=True)
    st.markdown("---")

    for art, g in df_atp.sort_values('DT_EXP').groupby('ART_KEY', sort=False):
        desc           = g['Articolo D'].iloc[0]
        qta_tot        = int(g['Qta Residua'].sum())
        d_cons_cliente = g['DT_EXP'].min()
        d_cons_friola  = d_cons_cliente - pd.Timedelta(days=ANTICIPO_ATOPLAST_GG)
        giorni_friola  = (d_cons_friola - pd.Timestamp(datetime.now())).days

        if giorni_friola < 0:
            badge, urgenza, col_u = "🔴", f"IN RITARDO di {abs(giorni_friola)} gg", "#f44336"
        elif giorni_friola <= 3:
            badge, urgenza, col_u = "🟠", f"URGENTE — {giorni_friola} gg", "#ff9800"
        elif giorni_friola <= 7:
            badge, urgenza, col_u = "🟡", f"A BREVE — {giorni_friola} gg", "#fbc02d"
        else:
            badge, urgenza, col_u = "🟢", f"Mancano {giorni_friola} gg", "#4caf50"

        with st.expander(f"{badge} {art} — {desc} | {qta_tot:,} pz | 📦 Friola: {d_cons_friola.strftime('%d/%m/%Y')}".replace(",",".")):
            c1, c2, c3 = st.columns(3)
            c1.metric("Qta da produrre", f"{qta_tot:,} pz".replace(",","."))
            c2.metric("Consegna a Friola", d_cons_friola.strftime("%d/%m/%Y"))
            c3.metric("Scadenza cliente", d_cons_cliente.strftime("%d/%m/%Y"))
            st.markdown(f'<div style="background:#f3e5f5;border-left:4px solid {col_u};padding:8px 14px;border-radius:6px;margin:6px 0;color:#1a1a1a!important;font-weight:600;">⏱️ {urgenza} alla consegna a Friola</div>', unsafe_allow_html=True)

            d_ord_atp = g['DATA_ORD'].min() if 'DATA_ORD' in g.columns and g['DATA_ORD'].notna().any() else d_cons_friola - pd.Timedelta(days=30)
            st.markdown(tbar_html(d_ord_atp, d_cons_friola), unsafe_allow_html=True)

            st.markdown("**Dettaglio ordini:**")
            for _, r in g.sort_values('DT_EXP').iterrows():
                cli   = str(r.get('CLI_NAME', '')).split(' - ')[-1].strip()
                d_fri = r['DT_EXP'] - pd.Timedelta(days=ANTICIPO_ATOPLAST_GG)
                st.markdown(f'<div class="status-row bom-row" style="color:#1a1a1a!important;"><span>🏭 Q: <b>{int(r["Qta Residua"]):,}</b> pz | 📦 A Friola: <b>{d_fri.strftime("%d/%m/%Y")}</b> | 👤 {cli}</span></div>'.replace(",","."), unsafe_allow_html=True)

def render_vista_cliente(df_cli, stock_raw, nome_cliente=''):
    """Vista pulita per il cliente — nessun dato interno visibile."""
    COLOR_MAP_CLI = {v[0]: v[1] for v in LABEL_CLI.values()}

    st.title("📦 I tuoi Ordini")

    if df_cli.empty:
        st.info("Nessun ordine aperto al momento.")
        return

    # KPI cliente
    tot_qta   = int(df_cli['Qta Residua'].sum())
    # Calcolo allocando la disponibilità (GIA/ACQ/PROD) sulle righe di ciascun articolo
    # in ordine di consegna: così le quantità coperte "in parte" da GIA finiscono nei PRONTI.
    pronti_gia = 0.0
    in_acquisto = 0.0
    in_produzione = 0.0
    mancanti = 0.0
    for art, g in df_cli.groupby('ART_KEY'):
        s_i = stock_raw.get(normalize_art_code(art), {'GIA': 0, 'ACQ': 0, 'PROD': 0})
        gia_left = float(s_i.get('GIA', 0))
        acq_left = float(s_i.get('ACQ', 0))
        prod_left = float(s_i.get('PROD', 0))
        g_sorted = g.sort_values(by='DT_EXP') if 'DT_EXP' in g.columns else g
        for _, r in g_sorted.iterrows():
            q = float(r.get('Qta Residua', 0))
            if q <= 0:
                continue
            # Prima copertura con GIA
            take = min(gia_left, q)
            pronti_gia += take
            gia_left -= take
            q -= take
            if q <= 0:
                continue
            # Poi copertura con ACQ
            take = min(acq_left, q)
            in_acquisto += take
            acq_left -= take
            q -= take
            if q <= 0:
                continue
            # Poi copertura con PROD
            take = min(prod_left, q)
            in_produzione += take
            prod_left -= take
            q -= take
            if q > 0:
                mancanti += q

    n_pronti = int(round(pronti_gia))
    n_lavoro = int(round(in_acquisto + in_produzione))
    n_mancanti = int(round(mancanti))
    pct_pronto = round(n_pronti / tot_qta * 100) if tot_qta > 0 else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f'<div class="kpi-card"><div style="font-size:11px">TOTALE PEZZI</div><div class="kpi-val">{tot_qta:,}</div></div>'.replace(",","."), unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#4caf50">PRONTI</div><div class="kpi-val">{n_pronti:,}</div></div>'.replace(",","."), unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#2196f3">IN LAVORAZIONE</div><div class="kpi-val">{n_lavoro:,}</div></div>'.replace(",","."), unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#f44336">DA PIANIFICARE</div><div class="kpi-val">{n_mancanti:,}</div></div>'.replace(",","."), unsafe_allow_html=True)

    # Grafico torta stato ordini cliente (calcolato coerentemente con KPI/Disponibile)
    # Nota: sommare direttamente per `ST` non è più corretto quando una riga è classificata
    # come ACQUISTO/PRODUZIONE ma una parte è coperta da GIA.
    label_key_by_st = {
        'DISPONIBILE': 'Pronto per la spedizione',
        'COPERTO BOM': 'Pronto (componente)',
        'ACQUISTO': 'In arrivo a magazzino',
        'PRODUZIONE': 'In lavorazione',
        'MANCANTE': 'In pianificazione',
        'DA PIANIFICARE': 'Da confermare',
    }
    # Accumulo per etichetta del grafico
    qty_by_label = {k: 0.0 for k in label_key_by_st.values()}

    for art, g in df_cli.groupby('ART_KEY'):
        s_i = stock_raw.get(normalize_art_code(art), {'GIA': 0, 'ACQ': 0, 'PROD': 0})
        gia_left = float(s_i.get('GIA', 0))
        acq_left = float(s_i.get('ACQ', 0))
        prod_left = float(s_i.get('PROD', 0))
        g_sorted = g.sort_values(by='DT_EXP') if 'DT_EXP' in g.columns else g

        for _, r in g_sorted.iterrows():
            q = float(r.get('Qta Residua', 0))
            if q <= 0:
                continue

            # Portion split: GIA -> PRONTI, poi ACQ -> IN ARRIVO, poi PROD -> IN LAVORAZIONE
            take_gia = min(gia_left, q)
            gia_left -= take_gia
            q_rem = q - take_gia

            if take_gia > 0:
                gia_label = 'Pronto (componente)' if str(r.get('ST', '')).strip().upper() == 'COPERTO BOM' else 'Pronto per la spedizione'
                qty_by_label[gia_label] += take_gia

            if q_rem <= 0:
                continue

            take_acq = min(acq_left, q_rem)
            acq_left -= take_acq
            q_rem -= take_acq
            if take_acq > 0:
                qty_by_label['In arrivo a magazzino'] += take_acq

            if q_rem <= 0:
                continue

            take_prod = min(prod_left, q_rem)
            prod_left -= take_prod
            q_rem -= take_prod
            if take_prod > 0:
                qty_by_label['In lavorazione'] += take_prod

            # Resto scoperto: usiamo lo stato originale (MANCANTE/DA PIANIFICARE)
            if q_rem > 0:
                st_line = str(r.get('ST', '')).strip().upper()
                if st_line == 'DA PIANIFICARE':
                    qty_by_label['Da confermare'] += q_rem
                else:
                    qty_by_label['In pianificazione'] += q_rem

    df_torta = (
        pd.DataFrame(
            [{'Stato Cliente': k, 'Qta Residua': int(round(v))} for k, v in qty_by_label.items() if v > 0]
        )
        .sort_values('Qta Residua', ascending=False)
    )
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

    # Tab vista cliente
    tab_ord_cli, tab_kpi_cli = st.tabs(["📦 I miei Ordini", "📊 Statistiche"])

    with tab_ord_cli:
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

            # Barra temporale: per ogni riga mostra avanzamento vs data consegna
            for _, r_t in g_sorted.iterrows():
                d_ord_r  = r_t.get('Data', None) or r_t.get('DT_EXP', None)
                d_cons_r = r_t.get('DT_EXP', None)
                if pd.notnull(d_ord_r) and pd.notnull(d_cons_r):
                    st.markdown(tbar_html(d_ord_r, d_cons_r), unsafe_allow_html=True)
                    break  # una sola barra per articolo (prima riga)

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

    with tab_kpi_cli:
        render_kpi_avanzati(filtro_cliente=nome_cliente)

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
                st.cache_data.clear()  # svuota TUTTA la cache incluso BTL/storico
                st.session_state.last_update = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
                st.rerun()
        st.caption("📅 Dati al: " + st.session_state.get('last_update', '--'))
        st.markdown("---")

        is_admin = st.session_state.permesso == "TUTTI"
        is_btl      = st.session_state.permesso == "BTL"
        is_atoplast = st.session_state.permesso == "ATOPLAST"
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
    if is_btl:
        render_vista_btl(df_res)
        st.stop()

    if is_atoplast:
        render_vista_atoplast(df_res)
        st.stop()

    if not is_admin:
        render_vista_cliente(df_f, stock_raw, nome_cliente=sel_cli if sel_cli != 'TUTTI' else '')
        st.stop()

    # ===========================================================
    # VISTA ADMIN (tutto il pannello originale)
    # ===========================================================
    # Nota: il download/export lo impostiamo più avanti usando `df_view`,
    # così rispetta anche i filtri selezionati per stato/famiglia.

    # ===========================================================
    # VISTA ADMIN: calcoli coerenti con la vista cliente
    # (allocazione PRONTI/ACQUISTO/PRODUZIONE basata su `stock_raw`)
    # ===========================================================
    def allocate_admin_states(df_in, stock_raw_in):
        qty_by_st = {
            'DISPONIBILE': 0.0,
            'COPERTO BOM': 0.0,
            'ACQUISTO': 0.0,
            'PRODUZIONE': 0.0,
            'MANCANTE': 0.0,
            'DA PIANIFICARE': 0.0,
        }

        for art, g in df_in.groupby('ART_KEY'):
            s_i = stock_raw_in.get(normalize_art_code(art), {'GIA': 0, 'ACQ': 0, 'PROD': 0, 'FIGLIO': 'NAN'})
            gia_left = float(s_i.get('GIA', 0))
            acq_left = float(s_i.get('ACQ', 0))
            prod_left = float(s_i.get('PROD', 0))

            figlio_code = str(s_i.get('FIGLIO', 'NAN')).strip().upper()
            figlio_gia_left = 0.0
            if figlio_code and figlio_code != 'NAN':
                figlio_gia_left = float(stock_raw_in.get(figlio_code, {}).get('GIA', 0))

            g_sorted = g.sort_values(by='DT_EXP') if 'DT_EXP' in g.columns else g
            for _, r in g_sorted.iterrows():
                q = float(r.get('Qta Residua', 0))
                if q <= 0:
                    continue

                st_line = str(r.get('ST', '')).strip().upper()

                # Manteniamo la classificazione "full cover" del motore per DISPONIBILE/COPERTO BOM
                if st_line == 'DISPONIBILE':
                    qty_by_st['DISPONIBILE'] += q
                    gia_left = max(gia_left - q, 0.0)
                    continue
                if st_line == 'COPERTO BOM':
                    qty_by_st['COPERTO BOM'] += q
                    if figlio_code and figlio_code != 'NAN':
                        figlio_gia_left = max(figlio_gia_left - q, 0.0)
                    continue

                # Altrimenti splittiamo: GIA -> DISPONIBILE, poi ACQ -> ACQUISTO, poi PROD -> PRODUZIONE
                take_gia = min(gia_left, q)
                qty_by_st['DISPONIBILE'] += take_gia
                gia_left -= take_gia
                q_rem = q - take_gia

                if q_rem <= 0:
                    continue

                take_acq = min(acq_left, q_rem)
                qty_by_st['ACQUISTO'] += take_acq
                acq_left -= take_acq
                q_rem -= take_acq

                if q_rem <= 0:
                    continue

                take_prod = min(prod_left, q_rem)
                qty_by_st['PRODUZIONE'] += take_prod
                prod_left -= take_prod
                q_rem -= take_prod

                if q_rem <= 0:
                    continue

                # Resto scoperto: usiamo lo stato originale (MANCANTE/DA PIANIFICARE)
                if st_line == 'DA PIANIFICARE':
                    qty_by_st['DA PIANIFICARE'] += q_rem
                else:
                    qty_by_st['MANCANTE'] += q_rem

        return qty_by_st

    alloc_admin = allocate_admin_states(df_f, stock_raw)
    n_pronti_admin = int(round(alloc_admin['DISPONIBILE']))
    n_bom_admin = int(round(alloc_admin['COPERTO BOM']))
    n_mancanti_admin = int(round(alloc_admin['MANCANTE']))
    # ===========================================================

    st.title("Pannello Controllo Safit")

    # ── FILTRI GLOBALI — sopra i tab, propagati su tutti ─────────────────────
    df_f = df_f.copy()
    df_f['Famiglia'] = df_f['Articolo D'].apply(lambda x: " ".join(str(x).split()[:2]).upper())

    COLOR_MAP = {'DISPONIBILE':'#4caf50','COPERTO BOM':'#9c27b0','ACQUISTO':'#2196f3',
                 'PRODUZIONE':'#fbc02d','MANCANTE':'#f44336','DA PIANIFICARE':'#9e9e9e'}

    if 'filtro_stati' not in st.session_state:    st.session_state.filtro_stati    = []
    if 'filtro_famiglie' not in st.session_state: st.session_state.filtro_famiglie = []

    with st.expander("🔽 Filtri — Stato e Famiglia", expanded=False):
        df_fam_chart_g  = df_f.groupby('Famiglia')['Qta Residua'].sum().reset_index().sort_values('Qta Residua', ascending=False).head(10)
        famiglie_disp_g = df_fam_chart_g['Famiglia'].tolist()
        stati_disp_g    = [s for s in COLOR_MAP if df_f[df_f['ST']==s]['Qta Residua'].sum() > 0]

        col_fs, col_ff, col_reset = st.columns([2, 2, 1])
        with col_fs:
            st.markdown("**Filtra per Stato**")
            for stato in stati_disp_g:
                attivo = stato in st.session_state.filtro_stati
                label  = f"✓ {stato}" if attivo else stato
                if st.button(label, key=f"btn_stato_{stato}", use_container_width=True,
                             type="primary" if attivo else "secondary"):
                    if attivo: st.session_state.filtro_stati.remove(stato)
                    else:      st.session_state.filtro_stati.append(stato)
                    st.rerun()
        with col_ff:
            st.markdown("**Filtra per Famiglia**")
            for fam in famiglie_disp_g:
                attivo = fam in st.session_state.filtro_famiglie
                label_s = fam[:16] + "…" if len(fam) > 16 else fam
                label   = f"✓ {label_s}" if attivo else label_s
                if st.button(label, key=f"btn_fam_{fam}", use_container_width=True,
                             type="primary" if attivo else "secondary"):
                    if attivo: st.session_state.filtro_famiglie.remove(fam)
                    else:      st.session_state.filtro_famiglie.append(fam)
                    st.rerun()
        with col_reset:
            st.markdown("&nbsp;", unsafe_allow_html=True)
            if st.button("🔄 Reset filtri", use_container_width=True, key="btn_reset_globale"):
                st.session_state.filtro_stati    = []
                st.session_state.filtro_famiglie = []
                st.rerun()
            if st.session_state.filtro_stati:
                st.info("Stati: " + ", ".join(st.session_state.filtro_stati))
            if st.session_state.filtro_famiglie:
                st.info("Famiglie: " + ", ".join(st.session_state.filtro_famiglie))

    # Applica filtri globali — df_view usato da tutti i tab
    df_view = df_f.copy()
    if st.session_state.filtro_famiglie:
        df_view = df_view[df_view['Famiglia'].isin(st.session_state.filtro_famiglie)]
    if st.session_state.filtro_stati:
        df_view = df_view[df_view['ST'].isin(st.session_state.filtro_stati)]

    tab_det, tab_op, tab_kpi, tab_btl, tab_atp = st.tabs(
        ["🔍 Dettaglio Ordini", "📋 KPI Operativi", "📊 KPI Avanzati", "🏭 Lavorazioni BTL", "🔵 Lavorazioni Atoplast"]
    )

    with tab_op:
      k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f'<div class="kpi-card"><div style="font-size:11px">FILTRATI</div><div class="kpi-val">{int(df_f["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div style="font-size:11px; color:#4caf50">PRONTI (GIA)</div><div class="kpi-val">{n_pronti_admin:,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card"><div style="font-size:11px; color:#9c27b0">BOM (FIGLI)</div><div class="kpi-val">{n_bom_admin:,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
    k4.markdown(f'<div class="kpi-card"><div style="font-size:11px; color:#f44336">MANCANTI</div><div class="kpi-val">{n_mancanti_admin:,}</div></div>'.replace(",", "."), unsafe_allow_html=True)

    # Grafici copertura per tab_op (usano df_view già filtrato)
    df_fam_chart   = df_view.groupby('Famiglia')['Qta Residua'].sum().reset_index().sort_values('Qta Residua', ascending=False).head(10)
    df_stato_chart = (
        pd.DataFrame(
            [{'ST': k, 'Qta Residua': int(round(v))} for k, v in alloc_admin.items()]
        )
        .sort_values('Qta Residua', ascending=False)
        .reset_index(drop=True)
    )

    col_g1, col_g2 = st.columns([2, 2])
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

    # Download Excel: include TUTTI i campi del dataset filtrato in base
    # alle selezioni attive (famiglie/stati) che l'utente vede a schermo.
    # Arricchimento export: per ogni riga d'ordine aggiungiamo disponibilita`
    # e quantita` coperte (GIA/ACQ/GRZ + BOM) basandoci su `stock_raw` e sullo stato riga.
    prod_cols_export = ['LANCIATI', 'GRZ', 'TMP', 'RWI', 'TRS', 'TRF']
    df_export = df_view.copy()

    # Pre-crea colonne per evitare KeyError su DataFrame
    for col in ['FIGLIO', 'GIA', 'ACQ', 'PROD', 'DISP_BOM_GIA']:
        if col not in df_export.columns:
            df_export[col] = 0.0
    for f in prod_cols_export:
        if f not in df_export.columns:
            df_export[f] = 0.0

    for col in [
        'COP_GIA', 'COP_ACQ', 'COP_PROD', 'COP_BOM',
        'COP_GRZ', 'MANCANTE_QTA',
    ]:
        if col not in df_export.columns:
            df_export[col] = 0.0

    for i, row in df_export.iterrows():
        art = row.get('ART_KEY', '')
        art_n = normalize_art_code(art)
        q = float(row.get('Qta Residua', 0) or 0)
        s = stock_raw.get(art_n, {'GIA': 0.0, 'ACQ': 0.0, 'PROD': 0.0, 'FIGLIO': 'NAN'})
        figlio_code = normalize_art_code(s.get('FIGLIO', 'NAN'))
        figlio_code = figlio_code if figlio_code and figlio_code != 'NAN' else ''
        s_figlio = stock_raw.get(figlio_code, {}) if figlio_code else {}

        gia_disp = float(s.get('GIA', 0) or 0)
        acq_disp = float(s.get('ACQ', 0) or 0)
        prod_disp = float(s.get('PROD', 0) or 0)
        df_export.at[i, 'GIA'] = gia_disp
        df_export.at[i, 'ACQ'] = acq_disp
        df_export.at[i, 'PROD'] = prod_disp
        df_export.at[i, 'FIGLIO'] = s.get('FIGLIO', 'NAN')
        df_export.at[i, 'DISP_BOM_GIA'] = float(s_figlio.get('GIA', 0) or 0)
        for f in prod_cols_export:
            df_export.at[i, f] = float(s.get(f, 0) or 0)

        st_line = str(row.get('ST', '')).strip().upper()
        cop_gia = 0.0
        cop_acq = 0.0
        cop_prod = 0.0
        cop_bom = 0.0
        cop_grz = 0.0
        manc = 0.0

        if q <= 0:
            manc = 0.0
        elif st_line == 'DISPONIBILE':
            cop_gia = min(gia_disp, q)
            manc = q - cop_gia
        elif st_line == 'COPERTO BOM':
            cop_bom = q
        else:
            # Splitting per ordine: GIA -> ACQ -> PROD (con breakdown su componenti in ordine elenco)
            q_rem = q
            take_gia = min(gia_disp, q_rem)
            cop_gia = take_gia
            q_rem -= take_gia

            take_acq = min(acq_disp, q_rem)
            cop_acq = take_acq
            q_rem -= take_acq

            take_prod = min(prod_disp, q_rem)
            cop_prod = take_prod
            q_rem -= take_prod

            # Breakdown PROD su componenti (utile per vedere GRZ)
            prod_left = cop_prod
            for f in prod_cols_export:
                if prod_left <= 0:
                    break
                avail = float(s.get(f, 0) or 0)
                take_f = min(avail, prod_left)
                if f == 'GRZ':
                    cop_grz += take_f
                prod_left -= take_f

            manc = q_rem

        df_export.at[i, 'COP_GIA'] = cop_gia
        df_export.at[i, 'COP_ACQ'] = cop_acq
        df_export.at[i, 'COP_PROD'] = cop_prod
        df_export.at[i, 'COP_BOM'] = cop_bom
        df_export.at[i, 'COP_GRZ'] = cop_grz
        df_export.at[i, 'MANCANTE_QTA'] = manc

    st.sidebar.download_button(
        "📊 Esporta Report (filtri attivi, completo + stock)",
        data=to_excel_full(df_export),
        file_name=f"Safit_Report_{datetime.now().strftime('%d%m')}.xlsx",
        use_container_width=True,
    )

    with tab_kpi:
        render_kpi_avanzati(filtro_cliente=sel_cli if sel_cli != "TUTTI" else None, filtro_articolo=search if search else None)

    with tab_btl:
        render_vista_btl(df_res)

    with tab_atp:
        render_vista_atoplast(df_res)

    with tab_det:
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
            # Barra temporale per l'articolo (usa prima riga con date valide)
            g_date = g.sort_values('DT_EXP')
            for _, r_t in g_date.iterrows():
                d_ord_admin  = r_t.get('DATA_ORD', None)
                d_cons_admin = r_t.get('DT_EXP', None)
                if not pd.notnull(d_ord_admin):
                    # Fallback: stima data ordine come 30gg prima della consegna
                    d_ord_admin = pd.Timestamp(d_cons_admin) - pd.Timedelta(days=30) if pd.notnull(d_cons_admin) else None
                if d_ord_admin is not None and pd.notnull(d_cons_admin):
                    st.markdown(tbar_html(d_ord_admin, d_cons_admin), unsafe_allow_html=True)
                    break

            for _, r in g.iterrows():
                tag = "📋 [PREV]" if r['Codice Documento'] == "OCA" else "🛒 [ORD]"
                st.markdown(f'<div class="status-row {r["CS"]}"><span>{tag} 📅 {r["DT_EXP"].strftime("%d/%m/%Y")} | Q: {int(r["Qta Residua"])} | {r["CLI_NAME"]}</span><span><b>{r["ST"]}</b></span></div>', unsafe_allow_html=True)
