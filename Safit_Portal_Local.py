import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE PAGINA E VERSIONE ---
APP_VERSION = "1.1.01"
st.set_page_config(page_title=f"Safit - Portale Avanzamento {APP_VERSION}", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;600;700&display=swap');

    html, body, .stApp { font-family: 'IBM Plex Sans', sans-serif; }
    .main { background-color: #f7f8fa; }
    .stApp { margin-top: -30px; }

    .status-row {
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap;
        padding: 10px 15px;
        border-radius: 8px;
        margin-bottom: 4px;
        gap: 10px;
        font-size: 14px;
    }
    .on-time-row      { background-color: #f1f8e9; border-left: 6px solid #4caf50; color: #1b5e20; }
    .client-delay-row { background-color: #e3f2fd; border-left: 6px solid #2196f3; color: #0d47a1; }
    .delay-row        { background-color: #fff8e1; border-left: 6px solid #ffc107; color: #5d4037; }
    .prod-delay-row   { background-color: #ffebee; border-left: 6px solid #f44336; color: #b71c1c; }

    .pbar-wrap { background:#fff; border-radius:10px; padding:12px 16px; margin-bottom:10px;
                 border:1px solid #e8eaf0; box-shadow:0 1px 4px rgba(0,0,0,.05); }
    .pbar-lbl  { font-size:11px; color:#666; font-weight:700; letter-spacing:.5px;
                 text-transform:uppercase; margin-bottom:6px; }
    .pbar-bg   { background:#e9ecef; border-radius:20px; height:18px; width:100%; overflow:hidden; }
    .pbar-fill { height:100%; border-radius:20px; display:flex; align-items:center;
                 justify-content:flex-end; padding-right:8px; font-size:11px;
                 font-weight:700; color:#fff; min-width:0; box-sizing:border-box; }
    .pbar-ext  { font-size:11px; font-weight:700; color:#444; text-align:right; margin-top:2px; }

    .tl-wrap   { display:flex; align-items:flex-start; gap:0; margin:8px 0 2px 0; width:100%; }
    .tl-step   { display:flex; flex-direction:column; align-items:center; flex:1; }
    .tl-circle { width:32px; height:32px; border-radius:50%; display:flex; align-items:center;
                 justify-content:center; font-size:14px; z-index:2;
                 box-shadow:0 2px 5px rgba(0,0,0,.15); }
    .tl-circle.done   { background:#4caf50; }
    .tl-circle.active { background:#ff9800; animation:tl-pulse 1.5s infinite; }
    .tl-circle.pend   { background:#e0e0e0; }
    .tl-lbl    { font-size:10px; margin-top:4px; text-align:center; font-weight:600;
                 color:#555; max-width:64px; line-height:1.2; }
    .tl-conn   { flex:1; height:4px; margin-top:14px; z-index:1; }
    .tl-conn.done { background:#4caf50; }
    .tl-conn.pend { background:#e0e0e0; }

    .sum-card  { background:#fff; border-radius:12px; padding:16px 20px; margin-bottom:14px;
                 border:1px solid #e8eaf0; box-shadow:0 2px 8px rgba(0,0,0,.06); }
    .sum-title { font-size:13px; font-weight:700; color:#333; margin-bottom:12px; }
    .kpi-row   { display:flex; gap:10px; flex-wrap:wrap; margin-bottom:12px; }
    .kpi-box   { flex:1; min-width:80px; text-align:center; padding:10px 6px; border-radius:8px; }
    .kpi-num   { font-size:22px; font-weight:700; }
    .kpi-lbl   { font-size:10px; font-weight:600; text-transform:uppercase; opacity:.75; margin-top:2px; }
    .kpi-g { background:#f1f8e9; color:#2e7d32; }
    .kpi-b { background:#e3f2fd; color:#1565c0; }
    .kpi-y { background:#fff8e1; color:#e65100; }
    .kpi-r { background:#ffebee; color:#b71c1c; }

    .sbar      { display:flex; height:20px; border-radius:12px; overflow:hidden; width:100%; gap:2px; }
    .sbar-seg  { height:100%; display:flex; align-items:center; justify-content:center;
                 font-size:10px; font-weight:700; color:#fff; min-width:4px; }

    @keyframes tl-pulse {
        0%,100% { transform:scale(1);   box-shadow:0 0 0 0   rgba(255,152,0,.4); }
        50%      { transform:scale(1.1); box-shadow:0 0 0 6px rgba(255,152,0,0);  }
    }

    .stExpander div { height:auto !important; min-height:min-content !important; }
    .ver-tag { font-size:10px; color:#999; text-align:right; }
    hr.slim  { margin:6px 0; border:none; border-top:1px solid #eee; }
    </style>
""", unsafe_allow_html=True)

# --- 2. UTENTI ---
@st.cache_data
def load_users():
    if os.path.exists('utenti.xlsx'):
        try:
            df_u = pd.read_excel('utenti.xlsx')
            df_u.columns = [str(c).strip() for c in df_u.columns]
            return df_u.set_index('username')[['password', 'cliente_arca']].T.to_dict('list')
        except:
            pass
    return {'safit_admin': ['admin2026', 'TUTTI']}

USER_DB = load_users()

def check_password():
    if 'authenticated' not in st.session_state:
        st.session_state['authenticated'] = False
    if not st.session_state['authenticated']:
        c1, c2, c3 = st.columns([1, 2, 1])
        with c2:
            if os.path.exists('Logo SAFIT.JPG'):
                st.image('Logo SAFIT.JPG', width=300)
            st.title("Accesso Area Riservata")
            user = st.text_input("Username")
            pw   = st.text_input("Password", type="password")
            if st.button("Accedi"):
                if user in USER_DB and str(USER_DB[user][0]) == pw:
                    st.session_state['authenticated'] = True
                    st.session_state['user_type']     = USER_DB[user][1]
                    st.session_state['username']      = user
                    st.rerun()
                else:
                    st.error("Username o Password errati")
        return False
    return True

if not check_password():
    st.stop()

# --- 3. FUNZIONI TECNICHE ---
def aggiungi_gg_lav(data_inizio, giorni):
    d = data_inizio
    while giorni > 0:
        d += timedelta(days=1)
        if d.weekday() < 5:
            giorni -= 1
    return d

def pulisci_numero(serie):
    """
    Gestisce tutti i formati numerici italiani:
    - "20.001,00" (stringa con punto migliaia e virgola decimale)
    - "20001.00"  (float già convertito da pandas)
    - 20001       (intero)
    - NaN / None
    """
    # Se la serie è già numerica (pandas ha già fatto il parsing), ritorna direttamente
    if pd.api.types.is_numeric_dtype(serie):
        return serie.fillna(0)
    # Altrimenti tratta come stringa italiana: rimuovi punto migliaia, converti virgola→punto
    return pd.to_numeric(
        serie.astype(str)
             .str.strip()
             .str.replace('.', '', regex=False)   # rimuove separatore migliaia
             .str.replace(',', '.', regex=False),  # virgola decimale → punto
        errors='coerce'
    ).fillna(0)

@st.cache_data
def load_data():
    try:
        # --- Ricerca file case-insensitive (compatibile Linux/GitHub) ---
        files = os.listdir('.')
        file_arca = next((f for f in files if f.lower() == 'righe_ordini_arca.xlsx'), 'righe_Ordini_ARCA.xlsx')
        file_acc  = next((f for f in files if f.lower() == 'avanzamento_access.xlsx'), 'Avanzamento_access.xlsx')

        df = pd.read_excel(file_arca, sheet_name='Foglio1', skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        for col in ['Cliente Fornitore CD', 'Articolo C', 'Articolo D', 'Data']:
            if col in df.columns:
                df[col] = df[col].ffill()
        df['Data_Consegna'] = pd.to_datetime(df['Data'], errors='coerce')
        col_res = 'Qta Residua' if 'Qta Residua' in df.columns else 'Qta Doc'
        df['Qta_Effettiva'] = pd.to_numeric(df[col_res], errors='coerce').fillna(0)
        df = df[df['Qta_Effettiva'] > 0]
        # Normalizza chiave join: strip + uppercase
        df['Articolo C'] = df['Articolo C'].astype(str).str.strip().str.upper()

        if os.path.exists(file_acc):
            df_t = pd.read_excel(file_acc, skiprows=1)
            df_t.columns = [str(c).strip() for c in df_t.columns]

            # Cerca colonna codice con nomi alternativi
            col_codice = None
            for possibile in ['Codice', 'CODICE', 'Cod', 'COD', 'Art_Key', 'Articolo']:
                if possibile in df_t.columns:
                    col_codice = possibile
                    break
            if col_codice is None and len(df_t.columns) > 0:
                col_codice = df_t.columns[0]  # usa la prima colonna come fallback

            df_t = df_t.rename(columns={col_codice: 'Art_Key'})
            # Normalizza chiave join: strip + uppercase (stesso trattamento di Articolo C)
            df_t['Art_Key'] = df_t['Art_Key'].astype(str).str.strip().str.upper()

            # Pulisci tutte le colonne numeriche presenti
            campi_num = ['Gia', 'Acq', 'Lan', 'Grz', 'Tmp', 'Rwi', 'Trs']
            for c in campi_num:
                if c in df_t.columns:
                    df_t[c] = pulisci_numero(df_t[c])
                else:
                    df_t[c] = 0.0

            # Lavorazione = somma di tutti i materiali in attesa / lavorazione
            df_t['Lavorazione_Totale'] = (
                df_t['Acq'] + df_t['Lan'] + df_t['Grz'] +
                df_t['Tmp'] + df_t['Rwi'] + df_t['Trs']
            )

            # Salva snapshot pre-merge per debug
            df._debug_avanzamento = df_t[['Art_Key', 'Gia', 'Lavorazione_Totale']].copy()
            df._debug_col_codice  = col_codice

            df = pd.merge(
                df,
                df_t[['Art_Key', 'Gia', 'Lavorazione_Totale']],
                left_on='Articolo C', right_on='Art_Key', how='left'
            )

            # Dopo il merge i NaN = articoli senza scheda tecnica → 0
            df['Gia']                = df['Gia'].fillna(0)
            df['Lavorazione_Totale'] = df['Lavorazione_Totale'].fillna(0)

        return df
    except Exception as e:
        st.error("Errore caricamento dati: " + str(e))
        import traceback
        st.code(traceback.format_exc())
        return pd.DataFrame()

data    = load_data()
oggi_dt = datetime.now()

# --- 4. CALCOLO STATO RIGHE ---
def calcola_righe(df_art):
    """
    Calcola lo stato di ogni riga d'ordine con logica FIFO sulla giacenza.
    Ritorna (lista_righe, giacenza_iniziale, lavorazione_totale).
    """
    st_gia = float(df_art['Gia'].iloc[0]) \
        if 'Gia' in df_art.columns and pd.notnull(df_art['Gia'].iloc[0]) else 0.0
    st_lav = float(df_art['Lavorazione_Totale'].iloc[0]) \
        if 'Lavorazione_Totale' in df_art.columns and pd.notnull(df_art['Lavorazione_Totale'].iloc[0]) else 0.0

    gia_res = st_gia
    righe   = []

    for _, row in df_art.sort_values('Data_Consegna').iterrows():
        qta      = float(row['Qta_Effettiva'])
        req_date = row['Data_Consegna']

        if gia_res >= qta:
            gia_res -= qta
            eta, nota, cat = oggi_dt, "Pronto", "DISPONIBILE"
        elif (gia_res + st_lav) >= qta:
            gia_res = 0
            eta, nota, cat = aggiungi_gg_lav(oggi_dt, 10), "In Lavorazione", "LAVORAZIONE"
        else:
            gia_res = 0
            eta, nota, cat = aggiungi_gg_lav(oggi_dt, 25), "Nuova Produzione", "LAVORAZIONE"

        is_ritardo = (pd.notnull(req_date) and eta.date() > req_date.date())

        if cat == "DISPONIBILE":
            if is_ritardo:
                css, nota_d, bar_col = "client-delay-row", "Pronto (Ritardo Ritiro)", "#2196f3"
            else:
                css, nota_d, bar_col = "on-time-row",      "Pronto",                  "#4caf50"
        else:
            if is_ritardo:
                css, nota_d, bar_col = "prod-delay-row", f"{nota} (In Ritardo)", "#f44336"
            else:
                css, nota_d, bar_col = "delay-row",      nota,                   "#ffc107"

        step_idx = 2 if cat == "DISPONIBILE" else 1

        righe.append({
            'css': css, 'date': req_date, 'qta': qta,
            'eta': eta, 'nota': nota_d, 'bar_col': bar_col,
            'cat': cat, 'is_ritardo': is_ritardo,
            'step_idx': step_idx,
        })

    return righe, st_gia, st_lav

# --- 5. HTML HELPERS ---
TIMELINE_LABELS = ["Ordinato", "In Lavorazione", "Pronto", "Consegnato"]
TIMELINE_ICONS  = ["📋", "⚙️", "✅", "🚚"]

def html_progress_bar(pct, color, label_text):
    pct = min(max(float(pct), 0), 100)
    pct_s   = "{:.1f}".format(pct)
    inside  = "{:.0f}%".format(pct) if pct > 12 else ""
    outside = "{:.0f}%".format(pct) if pct <= 12 else ""
    return (
        '<div class="pbar-wrap">'
        '<div class="pbar-lbl">' + label_text + '</div>'
        '<div class="pbar-bg">'
        '<div class="pbar-fill" style="width:' + pct_s + '%;background:' + color + ';">'
        + inside +
        '</div>'
        '</div>'
        '<div class="pbar-ext">' + outside + '</div>'
        '</div>'
    )

def html_timeline(step_idx):
    parts = []
    for i, (lbl, icon) in enumerate(zip(TIMELINE_LABELS, TIMELINE_ICONS)):
        cls = "done" if i < step_idx else ("active" if i == step_idx else "pend")
        parts.append(
            '<div class="tl-step">'
            '<div class="tl-circle ' + cls + '">' + icon + '</div>'
            '<div class="tl-lbl">' + lbl + '</div>'
            '</div>'
        )
        if i < len(TIMELINE_LABELS) - 1:
            conn = "done" if i < step_idx else "pend"
            parts.append('<div class="tl-conn ' + conn + '"></div>')
    return '<div class="tl-wrap">' + ''.join(parts) + '</div>'

def html_stacked_bar(cnts):
    total = sum(cnts.values()) or 1
    segs  = [
        (cnts.get('OK',  0), '#4caf50', '✓'),
        (cnts.get('BLU', 0), '#2196f3', '●'),
        (cnts.get('LAV', 0), '#ffc107', '⚙'),
        (cnts.get('RIT', 0), '#f44336', '!'),
    ]
    parts = []
    for cnt, col, sym in segs:
        if cnt == 0:
            continue
        w   = cnt / total * 100
        lbl = sym if w > 8 else ''
        parts.append(
            '<div class="sbar-seg" style="width:' + "{:.1f}".format(w) + '%;background:' + col + ';">' + lbl + '</div>'
        )
    return '<div class="sbar">' + ''.join(parts) + '</div>'

# --- 6. SIDEBAR ---
with st.sidebar:
    if os.path.exists('Logo SAFIT.JPG'):
        st.image('Logo SAFIT.JPG', use_container_width=True)
    st.write("Utente: **" + st.session_state['username'] + "**")
    st.markdown('<p class="ver-tag">Versione: ' + APP_VERSION + '</p>', unsafe_allow_html=True)

    if st.session_state['user_type'] == 'TUTTI':
        clienti_list = sorted([str(x) for x in data['Cliente Fornitore CD'].unique()])
        sel_cli = st.selectbox("👤 Seleziona Cliente:", clienti_list)
    else:
        sel_cli = st.session_state['user_type']

    filtro_label = st.radio("Filtra per stato:", ["Mostra tutto", "Solo Disponibili", "In Lavorazione", "In Ritardo"])
    st.markdown("---")
    debug_mode = st.checkbox("🔧 Debug giacenze", value=False, help="Mostra i valori grezzi letti da Avanzamento_access.xlsx")
    if st.button("Logout"):
        st.session_state['authenticated'] = False
        st.rerun()

# --- 7. MAIN ---
st.title("🚜 Portale Avanzamento Produzione")
df_cli = data[data['Cliente Fornitore CD'] == sel_cli].copy()

# --- DEBUG PANEL ---
if debug_mode:
    st.warning("🔧 **Modalità Debug attiva**")

    tab1, tab2, tab3 = st.tabs(["📋 Dopo merge (risultato)", "🔑 Chiavi Avanzamento (pre-merge)", "🔑 Chiavi Ordini (pre-merge)"])

    with tab1:
        st.caption("Valori Giacenza e Lavorazione che il portale sta usando. Se sono 0, il merge non ha trovato corrispondenza.")
        if 'Gia' in data.columns:
            df_dbg = (
                data[['Articolo C', 'Gia', 'Lavorazione_Totale']]
                .drop_duplicates(subset='Articolo C')
                .sort_values('Articolo C')
            )
            df_dbg_cli = df_dbg[df_dbg['Articolo C'].isin(df_cli['Articolo C'].unique())]
            st.dataframe(df_dbg_cli.rename(columns={
                'Articolo C':'Codice (in ordini)',
                'Gia':'Giacenza', 'Lavorazione_Totale':'Lavorazione'}),
                use_container_width=True, hide_index=True)

    with tab2:
        st.caption("Codici ESATTI letti da Avanzamento_access.xlsx (dopo strip+uppercase). Devono corrispondere alla colonna a sinistra.")
        if hasattr(data, '_debug_avanzamento'):
            st.dataframe(data._debug_avanzamento.head(50), use_container_width=True, hide_index=True)
            st.caption("Colonna codice usata: " + str(getattr(data, '_debug_col_codice', '?')))
        else:
            st.info("Dati pre-merge non disponibili (ricarica la pagina con cache pulita).")
            files = os.listdir('.')
            file_acc = next((f for f in files if f.lower() == 'avanzamento_access.xlsx'), None)
            if file_acc:
                df_raw = pd.read_excel(file_acc, skiprows=1, nrows=5)
                df_raw.columns = [str(c).strip() for c in df_raw.columns]
                st.write("Colonne trovate:", list(df_raw.columns))
                st.write("Prime 5 righe grezze:")
                st.dataframe(df_raw, use_container_width=True)

    with tab3:
        st.caption("Codici ESATTI negli ordini del cliente selezionato (dopo strip+uppercase).")
        art_cli = sorted(df_cli['Articolo C'].unique())
        st.write(art_cli[:30])

if df_cli.empty:
    st.info("Nessun dato disponibile per il cliente selezionato.")
    st.stop()

articoli_dis  = sorted([str(x) for x in df_cli['Articolo C'].unique()])
sel_art       = st.selectbox("🔍 Cerca Codice Prodotto:", ["Tutti i prodotti"] + articoli_dis)
articoli_view = articoli_dis if sel_art == "Tutti i prodotti" else [sel_art]

# --- RIEPILOGO CLIENTE ---
if sel_art == "Tutti i prodotti":
    cnts = {'OK': 0, 'BLU': 0, 'LAV': 0, 'RIT': 0}
    for art in articoli_dis:
        righe, _, _ = calcola_righe(df_cli[df_cli['Articolo C'] == art])
        for r in righe:
            if r['cat'] == 'DISPONIBILE' and not r['is_ritardo']:
                cnts['OK']  += 1
            elif r['cat'] == 'DISPONIBILE' and r['is_ritardo']:
                cnts['BLU'] += 1
            elif r['cat'] == 'LAVORAZIONE' and not r['is_ritardo']:
                cnts['LAV'] += 1
            else:
                cnts['RIT'] += 1

    tot   = sum(cnts.values()) or 1
    pct_p = (cnts['OK'] + cnts['BLU']) / tot * 100
    pct_ps = "{:.0f}".format(pct_p)

    sum_html = (
        '<div class="sum-card">'
        '<div class="sum-title">📊 Riepilogo Stato Ordini — ' + sel_cli + '</div>'
        '<div class="kpi-row">'
        '<div class="kpi-box kpi-g"><div class="kpi-num">' + str(cnts['OK'])  + '</div><div class="kpi-lbl">Pronti</div></div>'
        '<div class="kpi-box kpi-b"><div class="kpi-num">' + str(cnts['BLU']) + '</div><div class="kpi-lbl">Da Ritirare</div></div>'
        '<div class="kpi-box kpi-y"><div class="kpi-num">' + str(cnts['LAV']) + '</div><div class="kpi-lbl">In Lavorazione</div></div>'
        '<div class="kpi-box kpi-r"><div class="kpi-num">' + str(cnts['RIT']) + '</div><div class="kpi-lbl">In Ritardo</div></div>'
        '</div>'
        '<div class="pbar-lbl">Avanzamento complessivo (' + pct_ps + '% pronto)</div>'
        + html_stacked_bar(cnts) +
        '</div>'
    )
    st.markdown(sum_html, unsafe_allow_html=True)

# --- DETTAGLIO ARTICOLI ---
for art in articoli_view:
    df_art  = df_cli[df_cli['Articolo C'] == art]
    desc    = str(df_art['Articolo D'].iloc[0]) if 'Articolo D' in df_art.columns else ""
    qta_tot = df_art['Qta_Effettiva'].sum()

    righe, st_gia, st_lav = calcola_righe(df_art)

    # Applica filtro
    righe_vis = []
    for r in righe:
        if filtro_label == "Mostra tutto":
            righe_vis.append(r)
        elif filtro_label == "Solo Disponibili" and r['cat'] == 'DISPONIBILE':
            righe_vis.append(r)
        elif filtro_label == "In Lavorazione" and r['cat'] == 'LAVORAZIONE':
            righe_vis.append(r)
        elif filtro_label == "In Ritardo" and r['is_ritardo']:
            righe_vis.append(r)

    if not righe_vis:
        continue

    # Colore barra giacenza
    qta_pronta  = min(st_gia, qta_tot)
    pct_art     = (qta_pronta / qta_tot * 100) if qta_tot > 0 else 0
    bar_art_col = "#4caf50" if pct_art >= 100 else ("#ffc107" if pct_art > 0 else "#f44336")

    with st.expander("📦 " + art + " — " + desc + " | Residuo: " + "{:,.0f}".format(qta_tot)):

        pbar = html_progress_bar(
            pct_art,
            bar_art_col,
            "Giacenza disponibile: " + "{:,.0f}".format(qta_pronta) + " / " + "{:,.0f}".format(qta_tot) + " pz"
        )
        st.markdown(pbar, unsafe_allow_html=True)
        st.markdown('<hr class="slim">', unsafe_allow_html=True)

        for r in righe_vis:
            date_str = r['date'].strftime("%d/%m/%Y") if pd.notnull(r['date']) else "N.D."
            eta_str  = r['eta'].strftime("%d/%m/%Y")

            st.markdown(
                '<div class="status-row ' + r['css'] + '">'
                '<span><b>Consegna:</b> ' + date_str + ' | <b>Q.tà:</b> ' + "{:,.0f}".format(r['qta']) + '</span>'
                '<span><b>Stima Disponibilità:</b> ' + eta_str + ' (' + r['nota'] + ')</span>'
                '</div>',
                unsafe_allow_html=True
            )
            st.markdown(html_timeline(r['step_idx']), unsafe_allow_html=True)
            st.markdown('<hr class="slim">', unsafe_allow_html=True)




