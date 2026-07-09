import streamlit as st
import pandas as pd
import os
from datetime import datetime
from io import BytesIO
import plotly.express as px
import re
from bom_engine import get_coverage  
from kpi_avanzati import render_kpi_avanzati
try:
    from kpi_qualita import render_kpi_qualita
    _QUALITA_OK = True
except ImportError:
    _QUALITA_OK = False
try:
    import storico_safit as stor
    _STORICO_DISPONIBILE = True
except ImportError:
    _STORICO_DISPONIBILE = False

# --- 1. CONFIGURAZIONE ---
st.set_page_config(page_title="Safit Portal v3.9", layout="wide")

st.markdown("""
    <style>
    .status-row { display: flex; justify-content: space-between; padding: 10px; border-radius: 8px; margin-bottom: 6px; border: 1px solid #555; color: #1a1a1a !important; font-size: 17px; font-weight: 500; }
    .status-row * { color: #1a1a1a !important; }
    .on-time-row { background-color: #c8e6c9 !important; border-left: 6px solid #2e7d32; }
    .acq-row     { background-color: #bbdefb !important; border-left: 6px solid #1565c0; }
    .prod-row    { background-color: #fff9c4 !important; border-left: 6px solid #f57f17; }
    .urgent-row  { background-color: #ffcdd2 !important; border-left: 8px solid #c62828; }
    .oca-row     { background-color: #e0e0e0 !important; border-left: 8px solid #616161; color: #333 !important; }
    .oca-row *   { color: #333 !important; }
    .bom-row     { background-color: #e1bee7 !important; border-left: 8px solid #6a1b9a; }
    .disp-row    { background-color: #c8e6c9 !important; border-left: 6px solid #2e7d32; }
    .acq-row2    { background-color: #bbdefb !important; border-left: 6px solid #1565c0; }
    .prod-row2   { background-color: #fff9c4 !important; border-left: 6px solid #f57f17; }
    .miss-row    { background-color: #ffcdd2 !important; border-left: 6px solid #c62828; }
    .debug-box { background-color: #e8eaf6 !important; color: #1a1a2e !important; padding: 12px; border-radius: 8px; border: 2px solid #9fa8da; margin-bottom: 10px; display: flex; justify-content: space-around; font-size: 14px; font-weight: bold; }
    .debug-box * { color: #1a1a2e !important; }
    .kpi-card { background-color: #ffffff; border: 1px solid #e0e0e0; padding: 15px; border-radius: 10px; text-align: center; box-shadow: 2px 2px 5px rgba(0,0,0,0.05); }
    .kpi-val { font-size: 24px; font-weight: bold; color: #1f77b4; }
    .user-info { padding: 10px; background: #f8f9fa; border-radius: 5px; border: 1px solid #eee; margin-bottom: 20px; text-align: center; }
    /* Dark mode: forza testo leggibile in tutte le div inline con sfondo chiaro */
    [data-theme="dark"] div[style*="background:#"] { color: #1a1a1a !important; }
    [data-theme="dark"] div[style*="background:#"] span { color: #1a1a1a !important; }
    [data-theme="dark"] div[style*="background:#"] b { color: #1a1a1a !important; }
    [data-theme="dark"] .debug-box { background-color: #1e1e3a !important; color: #e0e0e0 !important; border-color: #4a4a8a !important; }
    [data-theme="dark"] .debug-box * { color: #e0e0e0 !important; }
    [data-theme="dark"] .kpi-card { background-color: #1e1e2e !important; border-color: #444 !important; }
    [data-theme="dark"] .kpi-val { color: #64b5f6 !important; }
    </style>
    """, unsafe_allow_html=True)

# --- CONFIGURAZIONE BTL ---
ANTICIPO_BTL_GG      = 3   # giorni anticipo BTL (Barletta → Friola)
ANTICIPO_ATOPLAST_GG = 3   # giorni anticipo Atoplast → Friola
PATH_STORICO_DATE    = "righe_ordini_storico_con_date.xlsx"
PATH_AVANZAMENTO    = "Avanzamento_access.xlsx"

# --- CONFIGURAZIONE KPI STORICI (SAFIT -> SafitIB) ---
# Adattare i path se i file si trovano in una sottocartella (es. "data/")
if _STORICO_DISPONIBILE:
    from pathlib import Path
    stor.PATH_STORICO_SAFIT = Path("righe_ordini_storico_con_date_SAFIT.xlsx")
    stor.PATH_CORRENTE_IB   = Path("righe_ordini_storico_con_date.xlsx")
    stor.PATH_TRANSCODIFICA = Path("transcodifica.xlsx")

# --- 2. FUNZIONI TECNICHE ---

@st.cache_data(ttl=1800)
def carica_giacenze():
    """Carica GIA per articolo da Avanzamento_access.xlsx."""
    if not os.path.exists(PATH_AVANZAMENTO):
        return {}
    try:
        df = pd.read_excel(PATH_AVANZAMENTO)
        df.columns = [str(c).strip() for c in df.columns]
        df['CODICE'] = df['CODICE'].astype(str).str.strip().str.upper()
        df['GIA']    = pd.to_numeric(df['GIA'], errors='coerce').fillna(0)
        return dict(zip(df['CODICE'], df['GIA']))
    except Exception:
        return {}


def calcola_prima_necessita(articolo, gia, df_storico_full):
    """
    Scala la giacenza GIA sugli OCI/OCA aperti in ordine di Data Consegna.
    Restituisce:
      - data del primo OCI che esaurisce la giacenza (= urgenza reale)
      - se GIA copre tutti gli OCI: prima data OCI come riferimento informativo
      - se nessun OCI: None
    """
    cod_col = 'Codice Documento'
    mask = (
        (df_storico_full['Articolo C'].astype(str).str.upper() == articolo.upper()) &
        (df_storico_full[cod_col].isin(['OCI', 'OCA'])) &
        (df_storico_full['Qta Residua'] > 0) &
        df_storico_full['Data Consegna'].notna()
    )
    df_oci = df_storico_full[mask].sort_values('Data Consegna')
    if df_oci.empty:
        return None, False

    scorta = max(gia, 0)
    for _, r in df_oci.iterrows():
        scorta -= r['Qta Residua']
        if scorta < 0:
            return r['Data Consegna'], True   # urgenza reale: scorta esaurita
    # Giacenza copre tutto: data primo OCI come riferimento
    return df_oci.iloc[0]['Data Consegna'], False


@st.cache_data
def get_user_db():
    if os.path.exists('utenti.xlsx'):
        try:
            df_u = pd.read_excel('utenti.xlsx')
            df_u.columns = [str(c).strip() for c in df_u.columns]
            return df_u.set_index('username')[['password', 'cliente_arca']].T.to_dict('list')
        except: pass
    return {"safit_admin": ["admin2026", "TUTTI"], "btl": ["btl2026", "BTL"], "atoplast": ["atoplast2026", "ATOPLAST"], "zak": ["zak799", "ZAK"]}

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

            # OCA = ordine previsionale: NON scala la GIA
            if row[c_tipo] == 'OCA':
                s = curr_stocks.get(art_code, {'GIA':0, 'ACQ': 0, 'PROD': 0})
                if s['GIA'] >= qta_ordine: stato, colore = 'DISPONIBILE', 'on-time-row'
                elif (s['GIA'] + s['ACQ']) >= qta_ordine: stato, colore = 'ACQUISTO', 'acq-row'
                elif (s['GIA'] + s['ACQ'] + s['PROD']) >= qta_ordine: stato, colore = 'PRODUZIONE', 'prod-row'
                else: stato, colore = 'DA PIANIFICARE', 'oca-row'
            else:
                # OCI = impegno reale: scala la GIA
                fonte = get_coverage(art_code, qta_ordine, curr_stocks)
                if fonte:
                    if str(fonte).strip().upper() == art_code:
                        stato, colore = 'DISPONIBILE', 'on-time-row'
                    else:
                        stato, colore = 'COPERTO BOM', 'bom-row'
                else:
                    s = curr_stocks.get(art_code, {'GIA':0, 'ACQ': 0, 'PROD': 0})
                    if (s['GIA'] + s['ACQ']) >= qta_ordine: stato, colore = 'ACQUISTO', 'acq-row'
                    elif (s['GIA'] + s['ACQ'] + s['PROD']) >= qta_ordine: stato, colore = 'PRODUZIONE', 'prod-row'
                    else: stato, colore = 'MANCANTE', 'urgent-row'
                    # Scala comunque la GIA residua: un OCI non coperto consuma
                    # la GIA disponibile cosi' i clienti successivi non la trovano libera
                    if art_code in curr_stocks:
                        gia_consumata = min(curr_stocks[art_code]['GIA'], qta_ordine)
                        curr_stocks[art_code]['GIA'] -= gia_consumata
                        qta_rem = qta_ordine - gia_consumata
                        acq_consumato = min(curr_stocks[art_code].get('ACQ', 0), qta_rem)
                        curr_stocks[art_code]['ACQ'] = curr_stocks[art_code].get('ACQ', 0) - acq_consumato

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
def _carica_storico_base():
    """Carica e normalizza il file storico — usato da BTL, Atoplast e calcolo necessità."""
    path = PATH_STORICO_DATE
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        df = pd.read_excel(path, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        # ffill Data Consegna: nella pivot ARCA appare solo sulla prima riga del gruppo
        df['Data Consegna'] = pd.to_datetime(df['Data Consegna'], errors='coerce').ffill()
        # ffill Codice Documento con reset sui subtotali
        cod_col = 'Codice Documento'
        filled, last_cod = [], None
        for val in df[cod_col]:
            s = str(val).strip()
            if s in ('nan', 'None', ''):
                filled.append(last_cod)
            elif 'otale' in s or 'Totale' in s:
                filled.append(s); last_cod = None
            else:
                last_cod = s; filled.append(s)
        df[cod_col] = filled
        df = df[~df[cod_col].astype(str).str.contains('Totale|otale', na=False)]
        df = df[df['Articolo C'].notna() & (df['Articolo C'].astype(str).str.strip() != '(vuoto)')]
        df = df[df[cod_col].notna()]
        df['Data']        = pd.to_datetime(df['Data'], errors='coerce')
        df['Qta Doc']     = pd.to_numeric(df['Qta Doc'],     errors='coerce').fillna(0)
        df['Qta Residua'] = pd.to_numeric(df['Qta Residua'], errors='coerce').fillna(0)                             if 'Qta Residua' in df.columns else pd.Series(0, index=df.index)
        return df
    except Exception:
        return pd.DataFrame()


@st.cache_data(show_spinner=False)
def carica_btl_da_storico(_mtime=0):
    """Carica OFR + OFF di BTL aperti (Qta Residua > 0). Data Consegna solo se presente in ARCA."""
    df = _carica_storico_base()
    if df.empty:
        return pd.DataFrame()
    cod_col = 'Codice Documento'
    mask_btl = df['Cliente Fornitore CD'].str.contains('BTL', case=False, na=False)
    df_btl = df[
        mask_btl &
        df[cod_col].isin(['OFR', 'OFF']) &
        (df['Qta Residua'] > 0)
    ].copy()
    # Data Consegna: mantieni solo quella reale da ARCA — NaN = BTL non ha ancora confermato
    # (NON fare fillna con Data ordine: sarebbe fuorviante)
    df_btl['Tipo']    = df_btl[cod_col].map({'OFR': '🔧 Lavorazione', 'OFF': '🛒 Acquisto'})
    df_btl['Qta Doc'] = df_btl['Qta Residua']
    return df_btl


def render_vista_btl(df_res=None, filtro_famiglie=None):
    """Vista BTL — OFR (lavorazioni) + OFF (acquisti). Usa Qta Residua da ARCA per quantità aperte."""
    st.title("🏭 Lavorazioni & Acquisti BTL")
    st.caption(f"Anticipo consegna a Friola: **{ANTICIPO_BTL_GG} giorni** prima della data consegna")

    _mtime = os.path.getmtime(PATH_STORICO_DATE) if os.path.exists(PATH_STORICO_DATE) else 0
    df_btl_storico = carica_btl_da_storico(_mtime)
    if df_btl_storico.empty:
        st.info("Nessun ordine BTL trovato. Verifica che il file righe_ordini_storico_con_date.xlsx sia presente.")
        return

    # Carica storico completo (per calcolo prima necessità) e giacenze
    df_storico_full = _carica_storico_base()
    gia_map         = carica_giacenze()

    df_btl = df_btl_storico.copy()

    # Applica filtro famiglie se attivo
    if filtro_famiglie:
        df_btl['_Famiglia'] = df_btl['Articolo D'].apply(lambda x: ' '.join(str(x).split()[:2]).upper())
        df_btl = df_btl[df_btl['_Famiglia'].isin(filtro_famiglie)]

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

        # Se BTL non ha ancora confermato la data, calcola la prima necessità dagli OCI
        data_confermata = d_cons is not None
        data_necessita  = None
        necessita_urgente = False
        if not data_confermata and not df_storico_full.empty:
            gia = gia_map.get(str(art).strip().upper(), 0)
            data_necessita, necessita_urgente = calcola_prima_necessita(art, gia, df_storico_full)

        # Data di riferimento per calcoli: confermata BTL > necessità OCI > None
        d_rif = d_cons if data_confermata else data_necessita

        if d_rif is not None:
            d_friola      = d_rif - pd.Timedelta(days=ANTICIPO_BTL_GG)
            giorni_friola = (d_friola - pd.Timestamp(datetime.now())).days
            if giorni_friola < 0:
                badge, urgenza, col_u = "🔴", f"IN RITARDO di {abs(giorni_friola)} gg", "#f44336"
            elif giorni_friola <= 3:
                badge, urgenza, col_u = "🟠", f"URGENTE — {giorni_friola} gg", "#ff9800"
            elif giorni_friola <= 7:
                badge, urgenza, col_u = "🟡", f"A BREVE — {giorni_friola} gg", "#fbc02d"
            else:
                badge, urgenza, col_u = "🟢", f"Mancano {giorni_friola} gg", "#4caf50"
            if data_confermata:
                data_label = f"📦 Friola: {d_friola.strftime('%d/%m/%Y')}"
            elif necessita_urgente:
                data_label = f"⚠️ Necessità: {d_rif.strftime('%d/%m/%Y')}"
            else:
                data_label = f"📋 Primo OCI: {d_rif.strftime('%d/%m/%Y')}"
        else:
            badge, urgenza, col_u, d_friola, data_label = "⚪", "Data N/D", "#9e9e9e", None, "📦 Data N/D"

        with st.expander(f"{badge} {art} — {desc} | {qta_tot:,} pa. | {tipi} | {data_label}".replace(",",".")):
            # ── Barra temporale (solo se data confermata) ─────────────────
            if data_confermata:
                d_start = d_ord if d_ord is not None else d_friola - pd.Timedelta(days=30)
                st.markdown(tbar_html(d_start, d_friola), unsafe_allow_html=True)
            elif data_necessita:
                gia_val         = int(gia_map.get(str(art).strip().upper(), 0))
                d_necessita_str = d_rif.strftime("%d/%m/%Y")
                _delta_lbl      = "⚠️ Scorta esaurita" if necessita_urgente else "✅ Coperto da giacenza"
                _info_html      = (
                    f'<span style="margin-right:16px;">📋 Prima necessità (OCI): <b>{d_necessita_str}</b></span>'
                    + f'<span style="margin-right:16px;">🗄️ Giacenza: <b>{gia_val:,} pa.</b></span>'.replace(",",".")
                    + f'<span style="background:{col_u};color:#fff;padding:2px 9px;border-radius:4px;font-weight:700;font-size:0.85rem;">⏱️ {urgenza}</span>'
                    + f'&nbsp;&nbsp;<span style="color:#888;font-size:0.83rem;">{_delta_lbl}</span>'
                )
                st.markdown(f'<div style="padding:4px 2px 2px 2px;font-size:0.9rem;">{_info_html}</div>', unsafe_allow_html=True)
            else:
                st.caption("📅 BTL non ha ancora confermato la data. Nessun ordine cliente aperto trovato.")

            st.markdown("<hr style='margin:6px 0;border:none;border-top:1px solid #e0e0e0;'>", unsafe_allow_html=True)

            # ── Due colonne: righe ordine | istruzioni BTL ───────────────────
            col_righe, col_istr = st.columns(2)

            with col_righe:
                st.markdown('<div style="font-weight:700;font-size:13px;color:#37474f;margin-bottom:4px;">📋 Righe ordine</div>', unsafe_allow_html=True)
                for _, r in g.sort_values('Data Consegna', na_position='last').iterrows():
                    d_c  = r['Data Consegna']
                    d_fs = (d_c - pd.Timedelta(days=ANTICIPO_BTL_GG)).strftime('%d/%m/%Y') if pd.notnull(d_c) else "N/D"
                    d_cs = d_c.strftime('%d/%m/%Y') if pd.notnull(d_c) else "N/D"
                    css  = 'prod-row' if 'Lavorazione' in str(r.get('Tipo','')) else 'acq-row'
                    st.markdown(f'<div class="status-row {css}" style="color:#1a1a1a!important;"><span>{r.get("Tipo","")} | Q: <b>{int(r["Qta Doc"]):,}</b> pa. | 📦 Friola: <b>{d_fs}</b> | Scad: {d_cs}</span></div>'.replace(",","."), unsafe_allow_html=True)

            # ── Istruzioni di lavorazione ─────────────────────────────────────
            istr_map  = carica_istruzioni_btl()
            art_up    = str(art).strip().upper()
            istr_list = istr_map.get(art_up, [])

            def _val(d, key):
                v = str(d.get(key, '')).strip()
                if v.lower() in ['nan', 'none', '', 'n.a.', 'no', '0']:
                    return None
                return v

            def _data_str(d):
                data_i = d.get('Data', None)
                try:
                    if pd.isnull(data_i): return 'N/D'
                    return data_i.strftime('%d/%m/%Y') if hasattr(data_i, 'strftime') else str(data_i)[:10]
                except Exception:
                    return 'N/D'

            def _qta(d):
                try: return int(float(d.get('Quantità', 0)))
                except: return 0

            def _pz_sc(d):
                try:
                    v = d.get('Pz x Scatola', '')
                    return int(float(v)) if str(v).lower() not in ['nan','none','','0'] else 0
                except: return 0

            def _score(d):
                campi = ['Piegatura','Tempra','Verniciatura','Tipo Vernice','Scatola','Note']
                return sum(1 for c in campi if _val(d, c) is not None)

            def deduplica(lista):
                from collections import defaultdict
                gruppi = defaultdict(list)
                for i in lista:
                    key = (str(i.get('Tipo Doc','')).strip().upper(),
                           str(i.get('N. Doc', '')).strip())
                    gruppi[key].append(i)
                risultati = []
                for key, rows in gruppi.items():
                    best = max(rows, key=_score)
                    note_set = []
                    for r in rows:
                        n = _val(r, 'Note')
                        if n and n not in note_set:
                            note_set.append(n)
                    if note_set:
                        best = dict(best)
                        best['Note'] = ' | '.join(note_set)
                    risultati.append(best)
                return risultati

            def render_istr_block(istr, bg, border):
                piegatura = _val(istr, 'Piegatura')
                tempra    = _val(istr, 'Tempra')
                vern      = _val(istr, 'Verniciatura')
                finitura  = _val(istr, 'finitura')
                scatola   = _val(istr, 'Scatola')
                note      = _val(istr, 'Note')
                qta_i     = _qta(istr)
                pz_sc     = _pz_sc(istr)
                righe = []
                if piegatura:
                    righe.append('<span style="display:inline-block;min-width:180px;">&#128295; <b>Piegatura</b></span> ' + piegatura)
                if tempra:
                    righe.append('<span style="display:inline-block;min-width:180px;">&#128293; <b>Tempra</b></span> ' + tempra)
                if vern:
                    vl = vern + (' &mdash; finitura: ' + finitura if finitura else '')
                    righe.append('<span style="display:inline-block;min-width:180px;">&#127912; <b>Verniciatura</b></span> ' + vl)
                if scatola:
                    sl = scatola + (' &mdash; ' + str(pz_sc) + ' pa./sc' if pz_sc > 0 else '')
                    righe.append('<span style="display:inline-block;min-width:180px;">&#128230; <b>Scatola</b></span> ' + sl)
                corpo = '<br>'.join(righe) if righe else '<i style="color:#888;">Nessun dettaglio specificato</i>'
                note_html = ''
                if note:
                    note_html = ('<div style="background:#fff9c4;border-left:4px solid #f9a825;padding:8px 12px;border-radius:5px;margin-top:8px;font-size:14px;font-weight:700;color:#3e2723;">&#9888; NOTE: ' + note + '</div>')
                qta_str = str(qta_i) + ' pz' if qta_i > 0 else ''
                html = (
                    '<div style="background:' + bg + ';border-left:5px solid ' + border + ';padding:12px 16px;border-radius:7px;margin:6px 0;font-size:14px;color:#1a1a1a;">'
                    + ('<div style="font-size:12px;color:#78909c;margin-bottom:6px;">Q.t&agrave;: ' + qta_str + ' &mdash; ' + _data_str(istr) + '</div>' if qta_str else '')
                    + corpo + note_html + '</div>'
                )
                st.markdown(html, unsafe_allow_html=True)

            with col_istr:
                st.markdown('<div style="font-weight:700;font-size:13px;color:#37474f;margin-bottom:4px;">📐 Istruzioni di lavorazione</div>', unsafe_allow_html=True)
                if istr_list:
                    tutti = deduplica(istr_list)
                    for istr in tutti:
                        render_istr_block(istr, bg='#f8f9fa', border='#546e7a')
                else:
                    st.caption("ℹ️ Nessuna istruzione disponibile.")

@st.cache_data(ttl=1800)
def carica_istruzioni_btl():
    """
    Legge btl_istruzioni.xlsx (consolidato dalla cartella locale BTL).
    Restituisce un dict  Articolo → lista di dict con istruzioni di lavorazione.
    """
    path = 'btl_istruzioni.xlsx'
    if not os.path.exists(path):
        return {}
    try:
        df = pd.read_excel(path)
        df.columns = [str(c).strip() for c in df.columns]
        if 'Articolo' not in df.columns:
            return {}
        df['Articolo'] = df['Articolo'].astype(str).str.strip().str.upper()
        istruzioni = {}
        for art, g in df.groupby('Articolo'):
            istruzioni[art] = g.to_dict('records')
        return istruzioni
    except Exception as e:
        return {}


@st.cache_data(show_spinner=False)
def carica_atoplast_da_storico(_mtime=0):
    """Carica OFR + OFF di Atoplast aperti (Qta Residua > 0). Data Consegna solo se presente in ARCA."""
    df = _carica_storico_base()
    if df.empty:
        return pd.DataFrame()
    cod_col  = 'Codice Documento'
    mask_atp = df['Cliente Fornitore CD'].str.contains('ATOPLAST', case=False, na=False)
    df_atp   = df[
        mask_atp &
        df[cod_col].isin(['OFR', 'OFF']) &
        (df['Qta Residua'] > 0)
    ].copy()
    df_atp['Tipo']    = df_atp[cod_col].map({'OFR': '🔧 Lavorazione', 'OFF': '🛒 Acquisto'})
    df_atp['Qta Doc'] = df_atp['Qta Residua']
    return df_atp


def render_vista_atoplast(df_res=None, filtro_famiglie=None):
    """Vista Atoplast — OFR (lavorazioni) + OFF (acquisti) dal file storico."""
    st.title("🏭 Lavorazioni & Acquisti Atoplast")
    st.caption(f"Anticipo consegna a Friola: **{ANTICIPO_ATOPLAST_GG} giorni** prima della data consegna")

    _mtime = os.path.getmtime(PATH_STORICO_DATE) if os.path.exists(PATH_STORICO_DATE) else 0
    df_atp_storico = carica_atoplast_da_storico(_mtime)
    if df_atp_storico.empty:
        st.info("Nessun ordine Atoplast trovato. Verifica che il file righe_ordini_storico_con_date.xlsx sia presente.")
        return

    df_atp = df_atp_storico.copy()

    if filtro_famiglie:
        df_atp['_Famiglia'] = df_atp['Articolo D'].apply(lambda x: ' '.join(str(x).split()[:2]).upper())
        df_atp = df_atp[df_atp['_Famiglia'].isin(filtro_famiglie)]

    if df_atp.empty:
        st.success("✅ Nessuna lavorazione Atoplast aperta al momento.")
        return

    n_lav   = df_atp[df_atp['Codice Documento'] == 'OFR']['Articolo C'].nunique()
    n_acq   = df_atp[df_atp['Codice Documento'] == 'OFF']['Articolo C'].nunique()
    urgenti = df_atp[df_atp['Data Consegna'].notna() &
                     (df_atp['Data Consegna'] <= pd.Timestamp(datetime.now()) + pd.Timedelta(days=7))]
    k1, k2, k3 = st.columns(3)
    k1.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#9c27b0">🔧 IN LAVORAZIONE</div><div class="kpi-val">{n_lav}</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#2196f3">🛒 IN ACQUISTO</div><div class="kpi-val">{n_acq}</div></div>', unsafe_allow_html=True)
    k3.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#f44336">⚠️ URGENTI ≤7gg</div><div class="kpi-val">{urgenti["Articolo C"].nunique()}</div></div>', unsafe_allow_html=True)
    st.markdown("---")

    for art, g in df_atp.sort_values('Data Consegna', na_position='last').groupby('Articolo C', sort=False):
        desc    = g['Articolo D'].iloc[0]
        qta_tot = int(g['Qta Doc'].sum())
        tipi    = ' + '.join(g['Tipo'].unique().tolist())
        d_cons  = g['Data Consegna'].dropna().min() if g['Data Consegna'].notna().any() else None
        d_ord   = g['Data'].dropna().min()           if g['Data'].notna().any()          else None

        if d_cons is not None:
            d_friola      = d_cons - pd.Timedelta(days=ANTICIPO_ATOPLAST_GG)
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

        with st.expander(f"{badge} {art} — {desc} | {qta_tot:,} pa. | {tipi} | {data_label}".replace(",",".")):
            c1, c2, c3 = st.columns(3)
            c1.metric("Quantità", f"{qta_tot:,} pa.".replace(",","."))
            c2.metric("A Friola entro", d_friola.strftime("%d/%m/%Y") if d_friola else "N/D")
            c3.metric("Data consegna", d_cons.strftime("%d/%m/%Y") if d_cons else "N/D")
            if d_friola:
                st.markdown(f'<div style="background:{col_u};border-left:6px solid rgba(0,0,0,0.25);padding:8px 14px;border-radius:6px;margin:6px 0;color:#fff!important;font-weight:700;text-shadow:0 1px 2px rgba(0,0,0,0.4);">⏱️ {urgenza} alla consegna a Friola</div>', unsafe_allow_html=True)
                d_start = d_ord if d_ord is not None else d_friola - pd.Timedelta(days=30)
                st.markdown(tbar_html(d_start, d_friola), unsafe_allow_html=True)
            st.markdown("**Dettaglio righe:**")
            for _, r in g.sort_values('Data Consegna', na_position='last').iterrows():
                d_c  = r['Data Consegna']
                d_fs = (d_c - pd.Timedelta(days=ANTICIPO_ATOPLAST_GG)).strftime('%d/%m/%Y') if pd.notnull(d_c) else "N/D"
                d_cs = d_c.strftime('%d/%m/%Y') if pd.notnull(d_c) else "N/D"
                css  = 'prod-row' if 'Lavorazione' in str(r.get('Tipo', '')) else 'acq-row'
                st.markdown(f'<div class="status-row {css}" style="color:#1a1a1a!important;"><span>{r.get("Tipo","")} | Q: <b>{int(r["Qta Doc"]):,}</b> pa. | 📦 Friola: <b>{d_fs}</b> | Scad: {d_cs}</span></div>'.replace(",","."), unsafe_allow_html=True)

def render_vista_cliente(df_cli, stock_raw, nome_cliente=''):
    """Vista pulita per il cliente — nessun dato interno visibile."""
    COLOR_MAP_CLI = {v[0]: v[1] for v in LABEL_CLI.values()}

    st.title("📦 I tuoi Ordini")

    if df_cli.empty:
        st.info("Nessun ordine aperto al momento.")
        return

    # KPI cliente — usa lo stato ST già calcolato dal motore (che alloca GIA su tutti i clienti
    # in ordine di data, quindi riflette la disponibilità reale per questo cliente)
    tot_qta = int(df_cli['Qta Residua'].sum())
    pronti_gia    = float(df_cli[df_cli['ST'].isin(['DISPONIBILE','COPERTO BOM'])]['Qta Residua'].sum())
    in_acquisto   = float(df_cli[df_cli['ST'] == 'ACQUISTO']['Qta Residua'].sum())
    in_produzione = float(df_cli[df_cli['ST'] == 'PRODUZIONE']['Qta Residua'].sum())
    mancanti      = float(df_cli[df_cli['ST'].isin(['MANCANTE','DA PIANIFICARE'])]['Qta Residua'].sum())

    n_pronti   = int(round(pronti_gia))
    n_lavoro   = int(round(in_acquisto + in_produzione))
    n_mancanti = int(round(mancanti))
    pct_pronto = round(n_pronti / tot_qta * 100) if tot_qta > 0 else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f'<div class="kpi-card"><div style="font-size:11px">TOTALE PAIA</div><div class="kpi-val">{tot_qta:,}</div></div>'.replace(",","."), unsafe_allow_html=True)
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

    # Grafico torta — usa ST del motore (già allocato correttamente su tutti i clienti)
    for _, r in df_cli.iterrows():
        q = float(r.get('Qta Residua', 0))
        if q <= 0:
            continue
        st_line = str(r.get('ST', '')).strip().upper()
        if st_line == 'DISPONIBILE':
            qty_by_label['Pronto per la spedizione'] += q
        elif st_line == 'COPERTO BOM':
            qty_by_label['Pronto (componente)'] += q
        elif st_line == 'ACQUISTO':
            qty_by_label['In arrivo a magazzino'] += q
        elif st_line == 'PRODUZIONE':
            qty_by_label['In lavorazione'] += q
        elif st_line == 'DA PIANIFICARE':
            qty_by_label['Da confermare'] += q
        else:
            qty_by_label['In pianificazione'] += q

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

        with st.expander(f"{badge} {art} — {desc} | {qta_tot:,} pa.".replace(",",".")):
            # Barra avanzamento articolo
            # "Disponibile" per il cliente deve riflettere quanta parte della richiesta
            # può essere coperta dalla sola GIACENZA (GIA), anche se la riga è classificata
            # come "ACQUISTO" perché la GIA non è sufficiente a coprire tutto.
            # Barra avanzamento — usa ST del motore per coerenza con allocazione globale GIA
            g_sorted = g.sort_values(by='DT_EXP') if 'DT_EXP' in g.columns else g
            qta_pronta = int(round(float(
                g[g['ST'].isin(['DISPONIBILE','COPERTO BOM'])]['Qta Residua'].sum()
            )))
            pct_art    = round(qta_pronta / qta_tot * 100) if qta_tot > 0 else 0
            bar_col    = '#4caf50' if pct_art >= 100 else ('#fbc02d' if pct_art > 0 else '#f44336')
            st.markdown(
                f"**Disponibile: {qta_pronta:,} / {qta_tot:,} pa.**".replace(",","."),
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
        kpi_sub = st.tabs(["📋 Operativi", "📊 Avanzati"])

        with kpi_sub[0]:
            tot = int(df_cli['Qta Residua'].sum())
            pr  = int(df_cli[df_cli['ST'].isin(['DISPONIBILE','COPERTO BOM'])]['Qta Residua'].sum())
            acq = int(df_cli[df_cli['ST'] == 'ACQUISTO']['Qta Residua'].sum())
            pro = int(df_cli[df_cli['ST'] == 'PRODUZIONE']['Qta Residua'].sum())
            man = int(df_cli[df_cli['ST'].isin(['MANCANTE','DA PIANIFICARE'])]['Qta Residua'].sum())
            c1,c2,c3,c4,c5 = st.columns(5)
            c1.metric("Totale paia",      f"{tot:,}".replace(",","."))
            c2.metric("Pronti",           f"{pr:,}".replace(",","."))
            c3.metric("In lavorazione",   f"{int(acq+pro):,}".replace(",","."))
            c4.metric("In pianificazione",f"{man:,}".replace(",","."))
            c5.metric("% Pronto",         f"{round(pr/tot*100) if tot>0 else 0}%")
            df_ops = df_cli.groupby('ART_KEY').agg(
                Descrizione=('Articolo D','first'),
                Qta=('Qta Residua','sum'),
                Stato=('ST','first')
            ).reset_index().sort_values('Qta', ascending=False)
            df_ops['Qta']   = df_ops['Qta'].astype(int)
            df_ops['Stato'] = df_ops['Stato'].map(lambda x: LABEL_CLI.get(x, (x,))[0])
            st.dataframe(df_ops.rename(columns={'ART_KEY':'Codice','Qta':'Qta (pa.)','Stato':'Stato'}),
                         use_container_width=True, hide_index=True)

        with kpi_sub[1]:
            try:
                render_kpi_avanzati(
                    filtro_cliente=nome_cliente if nome_cliente else None,
                    filtro_famiglie=None,
                    key_prefix="cli"
                )
            except Exception as e:
                st.warning(f"KPI Avanzati non disponibili: {e}")


# ===========================================================

# ===========================================================
# CRONISTORIA ARTICOLO — OCI/OCA + OFR/OFF per data
# ===========================================================
@st.cache_data(ttl=1800)
@st.cache_data(show_spinner=False)
def _carica_storico_cronistoria(_mtime=0):
    """Carica OCI+OCA+OFR+OFF aperti (Qta Residua > 0) per la cronistoria.
    Legge il file storico direttamente senza skiprows (header a riga 0)."""
    path = PATH_STORICO_DATE
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        df = pd.read_excel(path, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        df['Codice Documento'] = df['Codice Documento'].ffill()
        df['Data Consegna']    = pd.to_datetime(df['Data Consegna'], errors='coerce').ffill()
        df['Numero Documento'] = df['Numero Documento'].ffill()
        df['Qta Doc']          = pd.to_numeric(df['Qta Doc'],     errors='coerce').fillna(0)
        df['Qta Residua']      = pd.to_numeric(df['Qta Residua'], errors='coerce').fillna(0)
        df['Data']             = pd.to_datetime(df.get('Data'),   errors='coerce')
        # Rimuovi righe totale e articoli vuoti
        df = df[~df['Codice Documento'].astype(str).str.contains('Totale|otale', na=False)]
        df = df[df['Articolo C'].notna() & (df['Articolo C'].astype(str).str.strip() != '(vuoto)')]
    except Exception:
        return pd.DataFrame()
    if df.empty:
        return pd.DataFrame()
    doc_ok = ['OCI', 'OCA', 'OFR', 'OFF']
    df = df[df['Codice Documento'].isin(doc_ok) & (df['Qta Residua'] > 0)].copy()
    df['Data Consegna'] = pd.to_datetime(df['Data Consegna'], errors='coerce')
    df['_tipo'] = df['Codice Documento'].map({
        'OCI': 'cliente', 'OCA': 'cliente',
        'OFR': 'fornitore', 'OFF': 'fornitore',
    })
    return df


def render_cronistoria_articolo(art_key, df_ordini=None):
    """Cronistoria OCI/OCA + OFR/OFF affiancati. df_ordini: ordini aperti con ST e DT_EXP."""
    STATO_BADGE = {
        'DISPONIBILE':    ('#4caf50', 'DISPONIBILE'),
        'COPERTO BOM':    ('#9c27b0', 'COPERTO BOM'),
        'ACQUISTO':       ('#2196f3', 'ACQUISTO'),
        'PRODUZIONE':     ('#fbc02d', 'PRODUZIONE'),
        'MANCANTE':       ('#f44336', 'MANCANTE'),
        'DA PIANIFICARE': ('#9e9e9e', 'DA PIANIFICARE'),
    }
    stato_map = {}
    if df_ordini is not None and not df_ordini.empty:
        for _, _r in df_ordini.iterrows():
            _dt = _r.get('DT_EXP', None)
            _st = str(_r.get('ST', '')).strip().upper()
            if pd.notnull(_dt) and _st in STATO_BADGE:
                stato_map[pd.Timestamp(_dt).date()] = STATO_BADGE[_st]
    try:
        _mtime = os.path.getmtime(PATH_STORICO_DATE) if os.path.exists(PATH_STORICO_DATE) else 0
        df_st = _carica_storico_cronistoria(_mtime)
    except Exception as e:
        st.warning(f"Errore caricamento cronistoria: {e}")
        return
    if df_st.empty:
        st.caption(f"Dataset cronistoria vuoto — PATH_STORICO_DATE={PATH_STORICO_DATE}")
        return

    df_art = df_st[
        df_st['Articolo C'].astype(str).str.upper() == art_key.upper()
    ].copy()

    if df_art.empty:
        st.caption("Nessuna riga storica trovata per questo articolo.")
        return

    df_cli = df_art[df_art['_tipo'] == 'cliente'].sort_values('Data Consegna')
    df_forn = df_art[df_art['_tipo'] == 'fornitore'].sort_values('Data Consegna')

    oggi = pd.Timestamp(datetime.now().date())

    col_cli, col_forn = st.columns(2)

    with col_cli:
        st.markdown(
            '<div style="font-size:13px;font-weight:700;color:#1f77b4;'
            'border-bottom:2px solid #1f77b4;padding-bottom:3px;margin-bottom:6px;">'
            '🛒 Richieste Clienti (OCI/OCA)</div>',
            unsafe_allow_html=True
        )
        if df_cli.empty:
            st.caption("Nessuna richiesta cliente aperta.")
        else:
            qta_tot_cli = int(df_cli['Qta Residua'].sum())
            st.caption(f"Totale residuo: **{qta_tot_cli:,} pa.**".replace(",","."))
            for _, r in df_cli.iterrows():
                dc = r['Data Consegna']
                dc_str = dc.strftime("%d/%m/%Y") if pd.notnull(dc) else "N/D"
                qta = int(r['Qta Residua'])
                cod_doc = str(r.get('Codice Documento',''))
                cli = str(r.get('Cliente Fornitore CD','')).split(' - ')[-1].strip()
                nr  = str(r.get('Numero Documento','')).strip()
                # Colore urgenza
                if pd.notnull(dc):
                    gg = (dc - oggi).days
                    if gg < 0:   bc = "#f44336"
                    elif gg <= 7: bc = "#ff9800"
                    else:         bc = "#4caf50"
                else:
                    bc = "#9e9e9e"
                # Badge stato se disponibile
                _dc_date = dc.date() if pd.notnull(dc) else None
                _stato_info = stato_map.get(_dc_date, None) if _dc_date else None
                _badge_html = ''
                if _stato_info:
                    _s_col, _s_txt = _stato_info
                    _badge_html = (
                        f' &nbsp;<span style="background:{_s_col};color:#fff;'
                        f'padding:1px 7px;border-radius:10px;font-size:10px;'
                        f'font-weight:700;">{_s_txt}</span>'
                    )
                st.markdown(
                    f'<div style="border-left:4px solid {bc};'
                    f'padding:8px 12px;border-radius:6px;margin-bottom:4px;font-size:14px;">'
                    f'<b>📅 {dc_str}</b> &nbsp;|&nbsp; {cod_doc} N°{nr}{_badge_html}<br>'
                    f'<span style="font-size:13px;">{cli}</span> &nbsp;|&nbsp; '
                    f'<b>{qta:,} pa.</b>'.replace(",",".")
                    + f'</div>',
                    unsafe_allow_html=True
                )

    with col_forn:
        st.markdown(
            '<div style="font-size:12px;font-weight:700;color:#ff9800;'
            'border-bottom:2px solid #ff9800;padding-bottom:3px;margin-bottom:6px;">'
            '🏭 Ordini Fornitore (OFR/OFF)</div>',
            unsafe_allow_html=True
        )
        if df_forn.empty:
            st.caption("Nessun ordine fornitore aperto.")
        else:
            qta_tot_forn = int(df_forn['Qta Residua'].sum())
            delta = qta_tot_forn - int(df_cli['Qta Residua'].sum()) if not df_cli.empty else qta_tot_forn
            delta_ico = "▲" if delta >= 0 else "▼"
            delta_col = "#4caf50" if delta >= 0 else "#f44336"
            delta_abs = f"{abs(delta):,}".replace(",",".")
            qta_forn_fmt = f"{qta_tot_forn:,}".replace(",",".")
            st.markdown(
                f'Totale residuo: **{qta_forn_fmt} pa.** '
                f'<span style="color:{delta_col};font-weight:700;font-size:11px;">'
                f'({delta_ico} {delta_abs} vs clienti)</span>',
                unsafe_allow_html=True
            )
            for _, r in df_forn.iterrows():
                dc = r['Data Consegna']
                dc_str = dc.strftime("%d/%m/%Y") if pd.notnull(dc) else "N/D"
                qta = int(r['Qta Residua'])
                cod_doc = str(r.get('Codice Documento',''))
                forn = str(r.get('Cliente Fornitore CD','')).split(' - ')[-1].strip()
                nr   = str(r.get('Numero Documento','')).strip()
                if pd.notnull(dc):
                    gg = (dc - oggi).days
                    if gg < 0:    bc = "#f44336"
                    elif gg <= 14: bc = "#fbc02d"
                    else:          bc = "#9c27b0"
                else:
                    bc = "#9e9e9e"
                st.markdown(
                    f'<div style="border-left:4px solid {bc};'
                    f'padding:8px 12px;border-radius:6px;margin-bottom:4px;font-size:14px;">'
                    f'<b>📅 {dc_str}</b> &nbsp;|&nbsp; {cod_doc} N°{nr}<br>'
                    f'<span style="font-size:13px;">{forn}</span> &nbsp;|&nbsp; '
                    f'<b>{qta:,} pa.</b>'.replace(",",".")
                    + f'</div>',
                    unsafe_allow_html=True
                )

    # Barra riepilogo copertura
    if not df_cli.empty:
        qta_c = int(df_cli['Qta Residua'].sum())
        qta_f = int(df_forn['Qta Residua'].sum()) if not df_forn.empty else 0
        pct = min(round(qta_f / qta_c * 100), 100) if qta_c > 0 else 0
        col_bar = "#4caf50" if pct >= 100 else ("#ff9800" if pct >= 50 else "#f44336")
        st.markdown(
            f'<div style="margin-top:8px;font-size:12px;color:#1a1a1a;">'
            f'Copertura fornitore: <b>{pct}%</b> '
            f'({qta_f:,} / {qta_c:,} pa.)'.replace(",",".")
            + f'</div>',
            unsafe_allow_html=True
        )
        p = min(max(pct, 0), 100)
        st.markdown(
            f'<div style="background:#e9ecef;border-radius:20px;height:10px;width:100%;overflow:hidden;margin:3px 0 8px 0;">'
            f'<div style="width:{p}%;background:{col_bar};height:100%;border-radius:20px;"></div>'
            f'</div>',
            unsafe_allow_html=True
        )


# ===========================================================
# FUNZIONE VISTA ZAK — Giacenze magazzino Barletta (Zaccagni)
# ===========================================================
@st.cache_data(show_spinner=False)
def carica_giacenze_zak(_mtime=0):
    """Carica articoli con GIA_SUD > 0 da Avanzamento_access.xlsx."""
    if not os.path.exists(PATH_AVANZAMENTO):
        return pd.DataFrame()
    try:
        df_av = pd.read_excel(PATH_AVANZAMENTO)
        df_av.columns = [str(c).strip() for c in df_av.columns]
        df_av['GIA_SUD'] = pd.to_numeric(df_av['GIA_SUD'], errors='coerce').fillna(0)
        df_zak = df_av[df_av['GIA_SUD'] > 0][['CODICE','GIA_SUD']].copy()
        df_zak.columns = ['Codice Articolo', 'Giacenza (pa.)']
        # Prende descrizioni dallo storico
        if os.path.exists(PATH_STORICO_DATE):
            try:
                df_st = pd.read_excel(PATH_STORICO_DATE, skiprows=2)
                df_st.columns = [str(c).strip() for c in df_st.columns]
                df_st['Articolo C'] = df_st['Articolo C'].astype(str).str.strip()
                desc_map = (df_st.dropna(subset=['Articolo C','Articolo D'])
                            .drop_duplicates('Articolo C')
                            .set_index('Articolo C')['Articolo D']
                            .to_dict())
                df_zak['Descrizione'] = df_zak['Codice Articolo'].map(desc_map).fillna('N/D')
            except Exception:
                df_zak['Descrizione'] = 'N/D'
        else:
            df_zak['Descrizione'] = 'N/D'
        return df_zak[['Codice Articolo', 'Descrizione', 'Giacenza (pa.)']].sort_values('Codice Articolo').reset_index(drop=True)
    except Exception:
        return pd.DataFrame()


def render_vista_zak():
    """Vista ZaK — Giacenze magazzino Barletta (Zaccagni)."""
    st.title("🏪 Giacenze Magazzino ZaK — Barletta")
    st.caption("Giacenze presso **Zaccagni** (GIA_SUD) — aggiornate ad ogni refresh portale")

    _mtime = os.path.getmtime(PATH_AVANZAMENTO) if os.path.exists(PATH_AVANZAMENTO) else 0
    df_zak = carica_giacenze_zak(_mtime)

    if df_zak.empty:
        st.info("Nessuna giacenza trovata presso Zaccagni.")
        return

    # KPI riepilogo
    tot_articoli = len(df_zak)
    tot_paia     = int(df_zak['Giacenza (pa.)'].sum())
    k1, k2 = st.columns(2)
    k1.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#1f77b4">ARTICOLI PRESENTI</div><div class="kpi-val">{tot_articoli}</div></div>', unsafe_allow_html=True)
    k2.markdown(f'<div class="kpi-card"><div style="font-size:11px;color:#4caf50">TOTALE PAIA</div><div class="kpi-val">{tot_paia:,}</div></div>'.replace(",","."), unsafe_allow_html=True)

    st.markdown("---")

    # Cerca articolo
    search_zak = st.text_input("🔍 Cerca codice o descrizione:", key="zak_search").upper()
    df_view = df_zak.copy()
    if search_zak:
        mask = (df_view['Codice Articolo'].str.upper().str.contains(search_zak, na=False) |
                df_view['Descrizione'].str.upper().str.contains(search_zak, na=False))
        df_view = df_view[mask]

    st.caption(f"Visualizzati: **{len(df_view)}** articoli su {tot_articoli}")

    # Tabella principale
    st.dataframe(
        df_view.style.format({'Giacenza (pa.)': lambda x: f"{int(x):,}".replace(',','.')})
               .bar(subset=['Giacenza (pa.)'], color='#bbdefb'),
        use_container_width=True,
        hide_index=True,
        height=min(600, 38 + len(df_view) * 35),
    )

    # Download Excel
    from io import BytesIO
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as w:
        df_view.to_excel(w, index=False, sheet_name='Giacenze_ZaK')
        wb = w.book
        ws = w.sheets['Giacenze_ZaK']
        ws.set_column(0, 0, 22)
        ws.set_column(1, 1, 55)
        ws.set_column(2, 2, 18)
    st.download_button(
        "📥 Esporta Excel",
        data=buf.getvalue(),
        file_name=f"ZaK_Giacenze_{datetime.now().strftime('%d%m%Y')}.xlsx",
        use_container_width=True,
    )

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

        is_admin    = st.session_state.permesso == "TUTTI"
        permesso_up = st.session_state.permesso.upper()
        is_btl      = "BTL" in permesso_up and permesso_up != "TUTTI"
        is_atoplast = "ATOPLAST" in permesso_up and permesso_up != "TUTTI"
        is_zak      = "ZAK" in permesso_up and permesso_up != "TUTTI"
        if is_admin:
            sel_cli = st.selectbox("Seleziona Cliente:", ["TUTTI"] + sorted(df_res['CLI_NAME'].unique().tolist()))
        else:
            sel_cli = st.session_state.permesso

        search = st.text_input("🔍 Cerca Articolo:").upper()

    # Ridichiarate fuori dal with sidebar per essere visibili nel resto del codice
    is_admin    = st.session_state.permesso == "TUTTI"
    permesso_up = st.session_state.permesso.upper()
    is_btl      = "BTL" in permesso_up and permesso_up != "TUTTI"
    is_atoplast = "ATOPLAST" in permesso_up and permesso_up != "TUTTI"
    is_zak      = "ZAK" in permesso_up and permesso_up != "TUTTI"

    # ===========================================================
    # VISTE FORNITORI/SPECIALI — gestite prima del filtro cliente
    # (BTL e ATOPLAST non sono clienti, non usano df_f)
    # ===========================================================
    if is_btl:
        render_vista_btl(df_res)
        st.stop()

    if is_atoplast:
        render_vista_atoplast(df_res)
        st.stop()

    if is_zak:
        render_vista_zak()
        st.stop()

    # Filtro cliente — solo per utenti cliente normali e admin
    if sel_cli == "TUTTI":
        df_f = df_res.copy()
    else:
        # Cerca il codice cliente (es. "C001163") dentro CLI_NAME
        cod_match = sel_cli.split(' - ')[0].strip().upper()
        df_f = df_res[df_res['CLI_NAME'].str.upper().str.contains(cod_match, regex=False)]
        if df_f.empty:
            df_f = df_res[df_res['CLI_NAME'] == sel_cli]
    if search:
        df_f = df_f[df_f['ART_KEY'].str.contains(search)]

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

    # ── Prepara colonne e mappa colori ───────────────────────────────────────
    df_f = df_f.copy()
    df_f['Famiglia'] = df_f['Articolo D'].apply(lambda x: " ".join(str(x).split()[:2]).upper())

    COLOR_MAP = {'DISPONIBILE':'#4caf50','COPERTO BOM':'#9c27b0','ACQUISTO':'#2196f3',
                 'PRODUZIONE':'#fbc02d','MANCANTE':'#f44336','DA PIANIFICARE':'#9e9e9e'}

    # ── Filtri globali ───────────────────────────────────────────────────────
    _TUTTI = "— Tutti —"
    _fam_disp = [_TUTTI] + df_f.groupby('Famiglia')['Qta Residua'].sum().sort_values(ascending=False).head(12).index.tolist()
    _sta_disp = [_TUTTI] + [s for s in COLOR_MAP if df_f[df_f['ST']==s]['Qta Residua'].sum() > 0]

    _cf, _cs = st.columns([3, 2])
    with _cf:
        _sel_fam = st.selectbox("Famiglia", _fam_disp,
            index=0, label_visibility="collapsed", key="op_fam_adm")
    with _cs:
        _sel_sta = st.selectbox("Stato", _sta_disp,
            index=0, label_visibility="collapsed", key="op_sta_adm")

    # Applica filtri — se "— Tutti —" non filtra
    df_view = df_f.copy()
    if _sel_fam != _TUTTI:
        df_view = df_view[df_view['Famiglia'] == _sel_fam]
    if _sel_sta != _TUTTI:
        df_view = df_view[df_view['ST'] == _sel_sta]

    # Propaga ai tab
    st.session_state.filtro_famiglie = [] if _sel_fam == _TUTTI else [_sel_fam]
    st.session_state.filtro_stati    = [] if _sel_sta == _TUTTI else [_sel_sta]

    tab_det, tab_kpi_main, tab_btl, tab_atp, tab_zak = st.tabs(
        ["🔍 Dettaglio Ordini", "📊 KPI", "🏭 Lavorazioni BTL", "🔵 Lavorazioni Atoplast", "🏪 ZaK Barletta"]
    )

    with tab_kpi_main:
        # ── Sotto-tab KPI (uno alla volta) ───────────────────────────────────
        kpi_tab_op, kpi_tab_avz, kpi_tab_qual, kpi_tab_stor = st.tabs([
            "📋 Operativi", "📊 Avanzati", "🎯 Qualità", "📈 Storici"
        ])
        with kpi_tab_op:
            k1, k2, k3, k4 = st.columns(4)
            k1.markdown(f'<div class="kpi-card"><div style="font-size:11px">FILTRATI</div><div class="kpi-val">{int(df_view["Qta Residua"].sum()):,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
            k2.markdown(f'<div class="kpi-card"><div style="font-size:11px; color:#4caf50">PRONTI (GIA)</div><div class="kpi-val">{n_pronti_admin:,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
            k3.markdown(f'<div class="kpi-card"><div style="font-size:11px; color:#9c27b0">BOM (FIGLI)</div><div class="kpi-val">{n_bom_admin:,}</div></div>'.replace(",", "."), unsafe_allow_html=True)
            k4.markdown(f'<div class="kpi-card"><div style="font-size:11px; color:#f44336">MANCANTI</div><div class="kpi-val">{n_mancanti_admin:,}</div></div>'.replace(",", "."), unsafe_allow_html=True)

            # Grafici copertura — usano df_view già filtrato per stato e famiglia
            df_fam_chart   = df_view.groupby('Famiglia')['Qta Residua'].sum().reset_index().sort_values('Qta Residua', ascending=False).head(10)
            df_stato_chart = (
                df_view.groupby('ST')['Qta Residua'].sum().reset_index()
                .rename(columns={'ST': 'ST', 'Qta Residua': 'Qta Residua'})
                .sort_values('Qta Residua', ascending=False)
                .reset_index(drop=True)
            )

            col_g1, col_g2 = st.columns(2)
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

            # ── CONTAGIRI PER FAMIGLIA (KPI Operativi) ──────────────────────────
            # Calcola performance per famiglia usando df_view (già filtrato) e storico
            import math as _math_g
            from kpi_avanzati import carica_storico_arca as _carica_sto

            _df_sto = _carica_sto()
            # Filtra solo OCI (ordini clienti interni) per coerenza con i KPI operativi
            if 'Codice Documento' in _df_sto.columns:
                _df_sto = _df_sto[_df_sto['Codice Documento'] == 'OCI'].copy()
            if not _df_sto.empty and not df_view.empty and 'Famiglia' in df_view.columns:
                st.markdown("---")
                st.markdown("**🎯 Performance per Famiglia** — vendite periodo vs media storica")
                st.caption("Rosso: sotto la media storica | Verde: sopra | 100% = media attesa")

                from datetime import datetime as _dt_op, timedelta as _td_op

                # Selettore periodo
                _gg_map_op = {"Mese in corso": -1, "Anno in corso": -2,
                              "Ultimi 90 gg": 90, "Ultimi 6 mesi": 180,
                              "Ultimo anno": 365, "Tutto lo storico": 9999}
                _pc1, _pc2 = st.columns([1, 2])
                with _pc1:
                    _periodo_op = st.selectbox(
                        "Periodo performance",
                        list(_gg_map_op.keys()) + ["Intervallo personalizzato"],
                        index=4, key="kpi_op_periodo_adm"
                    )

                _oggi_op = _dt_op.now().date()

                if _periodo_op == "Mese in corso":
                    _da_op = _oggi_op.replace(day=1)
                    _a_op  = _oggi_op
                    _df_op = _df_sto[(_df_sto["Data"] >= pd.Timestamp(_da_op)) &
                                      (_df_sto["Data"] <= pd.Timestamp(_a_op))]
                    _gg_op = max(1, (_a_op - _da_op).days + 1)

                elif _periodo_op == "Anno in corso":
                    from calendar import monthrange as _mr_op
                    _ms = (_oggi_op.month - 1) or 12
                    _as = _oggi_op.year if _oggi_op.month > 1 else _oggi_op.year - 1
                    _da_op = _oggi_op.replace(month=1, day=1)
                    _a_op  = _oggi_op.replace(year=_as, month=_ms, day=_mr_op(_as, _ms)[1])
                    _df_op = _df_sto[(_df_sto["Data"] >= pd.Timestamp(_da_op)) &
                                      (_df_sto["Data"] <= pd.Timestamp(_a_op))]
                    _gg_op = max(1, (_a_op - _da_op).days)

                elif _periodo_op == "Intervallo personalizzato":
                    _dmin = _df_sto["Data"].min().date() if not _df_sto.empty else _dt_op(2024,1,1).date()
                    _dmax = _df_sto["Data"].max().date() if not _df_sto.empty else _dt_op.now().date()
                    _pd1, _pd2 = st.columns(2)
                    with _pd1:
                        _da_op = st.date_input("📅 Dal", value=_dmin, key="op_da_adm",
                            min_value=_dmin, max_value=_dmax,
                            key="kpi_op_da", format="DD/MM/YYYY")
                    with _pd2:
                        _a_op = st.date_input("📅 Al", value=_dmax, key="op_a_adm",
                            min_value=_dmin, max_value=_dmax,
                            key="kpi_op_a", format="DD/MM/YYYY")
                    _df_op = _df_sto[(_df_sto["Data"] >= pd.Timestamp(_da_op)) &
                                      (_df_sto["Data"] <= pd.Timestamp(_a_op))]
                    _gg_op = max(1, (pd.Timestamp(_a_op) - pd.Timestamp(_da_op)).days)

                else:
                    _gg_op = _gg_map_op[_periodo_op]
                    if _gg_op < 9999:
                        _cutoff_op = pd.Timestamp(_dt_op.now() - _td_op(days=_gg_op))
                        _df_op = _df_sto[_df_sto["Data"] >= _cutoff_op]
                    else:
                        _df_op = _df_sto.copy()
                        _gg_op = max(1, (_df_sto["Data"].max() - _df_sto["Data"].min()).days)

                # Media giornaliera storica (tutto lo storico OCI)
                _gg_sto = max(1, (_df_sto['Data'].max() - _df_sto['Data'].min()).days)
                _df_sto_fam = _df_sto.groupby('Famiglia')['Qta Doc'].sum().reset_index()
                _df_sto_fam['Media_Giorno_Sto'] = _df_sto_fam['Qta Doc'] / _gg_sto

                # Media giornaliera periodo selezionato
                _df_per_fam = _df_op.groupby('Famiglia')['Qta Doc'].sum().reset_index()
                _df_per_fam.columns = ['Famiglia', 'Qta_Periodo']
                _df_per_fam['Media_Giorno_Per'] = _df_per_fam['Qta_Periodo'] / max(1, _gg_op)

                # Unisci e calcola %
                _df_g = _df_per_fam.merge(_df_sto_fam[['Famiglia','Media_Giorno_Sto']], on='Famiglia', how='left')
                _df_g = _df_g[_df_g['Qta_Periodo'] > 0].copy()
                # Scostamento dalla media storica: 0=pari alla media, -100=zero, +100=doppio
                _df_g['Pct'] = ((_df_g['Media_Giorno_Per'] / _df_g['Media_Giorno_Sto'] - 1) * 100).clip(-100, 100)
                # Per le etichette mostra i valori giornalieri x1000 per leggibilità
                _df_g['Media_Periodo'] = (_df_g['Media_Giorno_Sto'] * _gg_op).round(0).astype(int)
                _df_g = _df_g.sort_values('Qta_Periodo', ascending=False).head(10)

                def _gauge_op(pct, qta, media, fam):
                    """Contagiri scala -100/0/+100 con settori colorati a step 25pt."""
                    import plotly.graph_objects as go
                    def _pta(p):
                        # -100->180(sx), 0->90(cima), +100->0(dx)
                        return _math_g.radians(90 - (p / 100) * 90)
                    N = 40
                    RE, RI = 1.0, 0.58

                    def _arco(s, e, re, ri):
                        xs, ys = [], []
                        for i in range(N+1):
                            a = _pta(s + (e-s)*i/N)
                            xs.append(re*_math_g.cos(a)); ys.append(re*_math_g.sin(a))
                        for i in range(N+1):
                            a = _pta(e - (e-s)*i/N)
                            xs.append(ri*_math_g.cos(a)); ys.append(ri*_math_g.sin(a))
                        xs.append(xs[0]); ys.append(ys[0])
                        return xs, ys

                    settori = [
                        (-100, -75, '#B71C1C'),
                        ( -75, -50, '#E53935'),
                        ( -50, -25, '#EF9A9A'),
                        ( -25,   0, '#FFCDD2'),
                        (   0,  25, '#C8E6C9'),
                        (  25,  50, '#66BB6A'),
                        (  50,  75, '#2E7D32'),
                        (  75, 100, '#1B5E20'),
                    ]

                    fig = go.Figure()
                    for s, e, col in settori:
                        sx, sy = _arco(s, e, RE, RI)
                        fig.add_trace(go.Scatter(x=sx, y=sy, fill='toself',
                            fillcolor=col, line=dict(color=col, width=0),
                            showlegend=False, hoverinfo='skip', mode='lines'))

                    # Linea 0% gialla (= media storica)
                    a0 = _pta(0)
                    fig.add_trace(go.Scatter(
                        x=[RI*_math_g.cos(a0), RE*_math_g.cos(a0)],
                        y=[RI*_math_g.sin(a0), RE*_math_g.sin(a0)],
                        mode='lines', line=dict(color='#FFD700', width=3),
                        showlegend=False, hoverinfo='skip'))

                    # Lancetta — bianca con bordo scuro per visibilità su tutti gli sfondi
                    al = _pta(pct)
                    fig.add_trace(go.Scatter(
                        x=[0, 0.88*_math_g.cos(al)], y=[0, 0.88*_math_g.sin(al)],
                        mode='lines', line=dict(color='#FFFFFF', width=4),
                        showlegend=False, hoverinfo='skip'))
                    # Bordo lancetta (leggermente più larga e scura sotto)
                    fig.add_trace(go.Scatter(
                        x=[0, 0.88*_math_g.cos(al)], y=[0, 0.88*_math_g.sin(al)],
                        mode='lines', line=dict(color='#333333', width=7),
                        showlegend=False, hoverinfo='skip'))
                    fig.add_trace(go.Scatter(
                        x=[0, 0.88*_math_g.cos(al)], y=[0, 0.88*_math_g.sin(al)],
                        mode='lines', line=dict(color='#FFFFFF', width=3),
                        showlegend=False, hoverinfo='skip'))
                    tc = [_math_g.radians(i) for i in range(361)]
                    fig.add_trace(go.Scatter(
                        x=[0.07*_math_g.cos(t) for t in tc],
                        y=[0.07*_math_g.sin(t) for t in tc],
                        fill='toself', fillcolor='#FFFFFF',
                        line=dict(color='#333', width=1),
                        showlegend=False, hoverinfo='skip', mode='lines'))

                    # Tick — colore neutro visibile su entrambi i temi
                    for tp, lb in [(-100,'-100%'),(-50,'-50%'),(0,'0'),(50,'+50%'),(100,'+100%')]:
                        at = _pta(tp)
                        fig.add_trace(go.Scatter(
                            x=[RE*_math_g.cos(at), (RE+0.07)*_math_g.cos(at)],
                            y=[RE*_math_g.sin(at), (RE+0.07)*_math_g.sin(at)],
                            mode='lines', line=dict(color='#AAAAAA', width=1),
                            showlegend=False, hoverinfo='skip'))
                        fig.add_annotation(
                            x=(RE+0.26)*_math_g.cos(at), y=(RE+0.26)*_math_g.sin(at),
                            text=lb, showarrow=False,
                            font=dict(size=8, color='#AAAAAA'),
                            xanchor='center', yanchor='middle')

                    # Valori testo — colori ad alto contrasto su entrambi i temi
                    col_p = '#00C853' if pct >= 0 else '#FF5252'  # verde/rosso brillante
                    segno = '+' if pct >= 0 else ''
                    # Scostamento %
                    fig.add_annotation(x=0, y=-0.16,
                        text=f"<b>{segno}{pct:.0f}%</b>",
                        showarrow=False, font=dict(size=24, color=col_p),
                        xanchor='center', yanchor='top')
                    # Nome famiglia
                    fig.add_annotation(x=0, y=-0.40,
                        text=f"<b>{fam}</b>",
                        showarrow=False, font=dict(size=12, color='#E0E0E0'),
                        xanchor='center', yanchor='top')
                    # Valori assoluti: totale, pa/gg periodo, pa/gg storico
                    _gg_per_label = max(1, _gg_op)
                    _pagg_per = qta / _gg_per_label
                    _pagg_sto = media / _gg_per_label  # media è già media_giorno_sto * gg_op
                    _pagg_sto_real = _df_sto_fam[_df_sto_fam['Famiglia']==fam]['Media_Giorno_Sto'].values
                    _pagg_sto_val = float(_pagg_sto_real[0]) if len(_pagg_sto_real) > 0 else 0
                    fig.add_annotation(x=0, y=-0.56,
                        text=f"{qta:,.0f} pa. | {_pagg_per:.1f} pa./gg → sto: {_pagg_sto_val:.1f} pa./gg".replace(",","."),
                        showarrow=False, font=dict(size=9, color='#BBBBBB'),
                        xanchor='center', yanchor='top')

                    fig.update_layout(
                        height=230,
                        margin=dict(t=8, b=8, l=8, r=8),
                        paper_bgcolor='rgba(0,0,0,0)',
                        plot_bgcolor='rgba(0,0,0,0)',
                        xaxis=dict(visible=False, range=[-1.55, 1.55], scaleanchor='y'),
                        yaxis=dict(visible=False, range=[-0.92, 1.35]),
                        showlegend=False,
                    )
                    return fig

                if not _df_g.empty:
                    _N_COL = 4
                    _righe = [_df_g['Famiglia'].tolist()[i:i+_N_COL]
                              for i in range(0, len(_df_g), _N_COL)]
                    for _riga in _righe:
                        _cols = st.columns(_N_COL)
                        for _j, _fam in enumerate(_riga):
                            _row = _df_g[_df_g['Famiglia'] == _fam].iloc[0]
                            with _cols[_j]:
                                st.plotly_chart(
                                    _gauge_op(float(_row['Pct']),
                                              int(_row['Qta_Periodo']),
                                              int(_row['Media_Periodo']), _fam),
                                    use_container_width=True,
                                    config={'displayModeBar': False},
                                    key=f"gop_{_fam}"
                                )
                        for _j in range(len(_riga), _N_COL):
                            _cols[_j].empty()


    # Download Excel: include TUTTI i campi del dataset filtrato in base
    # alle selezioni attive (famiglie/stati) che l'utente vede a schermo.
    # Arricchimento export: per ogni riga d'ordine aggiungiamo disponibilita`
    # e quantita` coperte (GIA/ACQ/GRZ + BOM) basandoci su `stock_raw` e sullo stato riga.
    prod_cols_export = ['LANCIATI', 'GRZ', 'TMP', 'RWI', 'TRS', 'TRF']
    df_export = df_view.copy()

    # Pre-crea colonne per evitare KeyError su DataFrame
    if 'FIGLIO' not in df_export.columns:
        df_export['FIGLIO'] = 'NAN'
    else:
        df_export['FIGLIO'] = df_export['FIGLIO'].astype(str)
    for col in ['GIA', 'ACQ', 'PROD', 'DISP_BOM_GIA']:
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

    with kpi_tab_avz:
        render_kpi_avanzati(
            filtro_cliente=sel_cli if sel_cli != "TUTTI" else None,
            filtro_articolo=search if search else None,
            filtro_famiglie=st.session_state.filtro_famiglie or None,
            key_prefix="adm"
        )

    with kpi_tab_qual:
        if _QUALITA_OK:
            render_kpi_qualita(
                filtro_cliente=sel_cli if sel_cli != "TUTTI" else None,
                filtro_famiglie=st.session_state.filtro_famiglie or None
            )
        else:
            st.error("Modulo kpi_qualita.py non trovato.")

    with kpi_tab_stor:
        if _STORICO_DISPONIBILE:
            stor.render_kpi_storici(
                filtro_cliente=sel_cli if sel_cli != "TUTTI" else None,
                filtro_articolo=search if search else None,
                filtro_famiglie=st.session_state.filtro_famiglie or None
            )
        else:
            st.error("Modulo storico_safit.py non trovato.")

    with tab_btl:
        render_vista_btl(df_res, filtro_famiglie=st.session_state.filtro_famiglie)

    with tab_atp:
        render_vista_atoplast(df_res, filtro_famiglie=st.session_state.filtro_famiglie)


    with tab_zak:
        render_vista_zak()
    with tab_det:
        if df_view.empty:
            st.info("Nessun ordine trovato con i filtri attivi.")
        else:
            for art, g in df_view.groupby('ART_KEY'):
                desc    = g['Articolo D'].iloc[0] if 'Articolo D' in g.columns else ''
                qta_tot = int(g['Qta Residua'].sum())
                s_i     = stock_raw.get(art, {'GIA': 0, 'ACQ': 0, 'PROD': 0, 'FIGLIO': 'NAN'})
                with st.expander(f"📦 {art} — {desc} | Residuo: {qta_tot:,} pa.".replace(",",".")):
                    st.markdown(
                        f'<div class="debug-box">'
                        f'<span>📦 GIA: {int(s_i["GIA"])}</span>'
                        f'<span>🚚 ACQ: {int(s_i["ACQ"])}</span>'
                        f'<span>⚙️ PROD: {int(s_i["PROD"])}</span>'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                    # Cronistoria: OCI/OCA + OFR/OFF affiancati per data
                    render_cronistoria_articolo(art, g)
