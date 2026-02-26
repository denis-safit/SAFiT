import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE PAGINA ---
APP_VERSION = "1.1.05"
st.set_page_config(page_title="Safit - Portale Avanzamento " + APP_VERSION, layout="wide")

# --- 2. GESTIONE UTENTI ---
@st.cache_data
def load_users():
    file_u = 'utenti.xlsx'
    if os.path.exists(file_u):
        try:
            df_u = pd.read_excel(file_u)
            df_u.columns = [str(c).strip() for c in df_u.columns]
            return df_u.set_index('username')[['password', 'cliente_arca']].T.to_dict('list')
        except:
            return {'safit_admin': ['admin2026', 'TUTTI']}
    return {'safit_admin': ['admin2026', 'TUTTI']}

# --- 3. FUNZIONI TECNICHE ---
def aggiungi_giorni_lavorativi(data_inizio, giorni):
    data_corrente = data_inizio
    while giorni > 0:
        data_corrente += timedelta(days=1)
        if data_corrente.weekday() < 5:
            giorni -= 1
    return data_corrente

def pulisci_numero(serie):
    return pd.to_numeric(
        serie.astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False),
        errors='coerce').fillna(0)

@st.cache_data
def load_data():
    try:
        df = pd.read_excel('righe_Ordini_ARCA.xlsx', sheet_name='Foglio1', skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        for col in ['Cliente Fornitore CD', 'Articolo C', 'Articolo D', 'Data']:
            if col in df.columns:
                df[col] = df[col].ffill()
        df['Data_Consegna'] = pd.to_datetime(df['Data'], errors='coerce')
        col_res = 'Qta Residua' if 'Qta Residua' in df.columns else 'Qta Doc'
        df['Qta_Effettiva'] = pd.to_numeric(df[col_res], errors='coerce').fillna(0)
        df = df[df['Qta_Effettiva'] > 0]
        if os.path.exists('Avanzamento_access.xlsx'):
            df_tech = pd.read_excel('Avanzamento_access.xlsx', skiprows=1)
            df_tech.columns = [str(c).strip() for c in df_tech.columns]
            if 'Codice' in df_tech.columns:
                df_tech = df_tech.rename(columns={'Codice': 'Art_Key'})
                campi_num = ['Gia', 'Acq', 'Lan', 'Grz', 'Tmp', 'Rwi', 'Trs']
                for c in campi_num:
                    if c in df_tech.columns:
                        df_tech[c] = pulisci_numero(df_tech[c])
                df_tech['Lavorazione_Totale'] = (
                    df_tech.get('Acq', 0) + df_tech.get('Lan', 0) +
                    df_tech.get('Grz', 0) + df_tech.get('Tmp', 0) +
                    df_tech.get('Rwi', 0) + df_tech.get('Trs', 0))
                df = pd.merge(df, df_tech[['Art_Key', 'Gia', 'Lavorazione_Totale']],
                              left_on='Articolo C', right_on='Art_Key', how='left')
        return df
    except:
        return pd.DataFrame()

# =============================================================================
# CSS + STILE  (dopo le funzioni @cache per evitare tokenize.TokenError)
# =============================================================================
CSS = (
    "<style>"
    ".main{background-color:#fcfcfc}"
    ".stApp{margin-top:-30px}"
    ".status-row{"
        "display:flex;justify-content:space-between;align-items:center;"
        "flex-wrap:wrap;padding:12px 15px;border-radius:8px;"
        "margin-bottom:8px;gap:10px;font-size:14px}"
    ".on-time-row{background-color:#f1f8e9;border-left:6px solid #4caf50;color:#1b5e20}"
    ".client-delay-row{background-color:#e3f2fd;border-left:6px solid #2196f3;color:#0d47a1}"
    ".delay-row{background-color:#fff8e1;border-left:6px solid #ffc107;color:#5d4037}"
    ".prod-delay-row{background-color:#ffebee;border-left:6px solid #f44336;color:#b71c1c}"
    ".stExpander div{height:auto!important;min-height:min-content!important}"
    ".version-tag{font-size:10px;color:#999;text-align:right}"
    # progress bar
    ".pbar-wrap{background:#fff;border-radius:8px;padding:10px 14px;"
        "margin-bottom:10px;border:1px solid #e8eaf0}"
    ".pbar-lbl{font-size:11px;color:#666;font-weight:700;letter-spacing:.4px;"
        "text-transform:uppercase;margin-bottom:5px}"
    ".pbar-bg{background:#e9ecef;border-radius:20px;height:16px;width:100%;overflow:hidden}"
    ".pbar-fill{height:100%;border-radius:20px;display:flex;align-items:center;"
        "justify-content:flex-end;padding-right:6px;font-size:10px;"
        "font-weight:700;color:#fff;box-sizing:border-box}"
    # timeline
    ".tl-wrap{display:flex;align-items:flex-start;margin:6px 0 2px 0;width:100%}"
    ".tl-step{display:flex;flex-direction:column;align-items:center;flex:1}"
    ".tl-circle{width:30px;height:30px;border-radius:50%;display:flex;align-items:center;"
        "justify-content:center;font-size:13px;box-shadow:0 2px 4px rgba(0,0,0,.15)}"
    ".tl-done{background:#4caf50}"
    ".tl-active{background:#ff9800;animation:tlpulse 1.5s infinite}"
    ".tl-pend{background:#e0e0e0}"
    ".tl-lbl{font-size:10px;margin-top:3px;text-align:center;font-weight:600;"
        "color:#555;max-width:62px;line-height:1.2}"
    ".tl-line{flex:1;height:4px;margin-top:13px}"
    ".tl-line-done{background:#4caf50}"
    ".tl-line-pend{background:#e0e0e0}"
    # kpi summary
    ".sum-card{background:#fff;border-radius:10px;padding:14px 18px;margin-bottom:12px;"
        "border:1px solid #e8eaf0;box-shadow:0 2px 6px rgba(0,0,0,.05)}"
    ".sum-title{font-size:13px;font-weight:700;color:#333;margin-bottom:10px}"
    ".kpi-row{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:10px}"
    ".kpi-box{flex:1;min-width:80px;text-align:center;padding:10px 6px;border-radius:8px}"
    ".kpi-num{font-size:22px;font-weight:700}"
    ".kpi-lbl{font-size:10px;font-weight:600;text-transform:uppercase;opacity:.75;margin-top:2px}"
    ".kpi-g{background:#f1f8e9;color:#2e7d32}"
    ".kpi-b{background:#e3f2fd;color:#1565c0}"
    ".kpi-y{background:#fff8e1;color:#e65100}"
    ".kpi-r{background:#ffebee;color:#b71c1c}"
    ".sbar{display:flex;height:18px;border-radius:10px;overflow:hidden;width:100%;gap:2px}"
    ".sbar-s{height:100%;display:flex;align-items:center;justify-content:center;"
        "font-size:10px;font-weight:700;color:#fff;min-width:4px}"
    "@keyframes tlpulse{"
        "0%,100%{transform:scale(1);box-shadow:0 0 0 0 rgba(255,152,0,.4)}"
        "50%{transform:scale(1.1);box-shadow:0 0 0 5px rgba(255,152,0,0)}}"
    "</style>"
)
st.markdown(CSS, unsafe_allow_html=True)

# =============================================================================
# FUNZIONI HTML GRAFICHE  (stringhe pure, zero f-string con {})
# =============================================================================
TL_LABELS = ["Ordinato", "In Lavorazione", "Pronto", "Consegnato"]
TL_ICONS  = ["📋", "⚙️", "✅", "🚚"]

def html_pbar(pct, color, label):
    p  = min(max(float(pct), 0), 100)
    ps = str(round(p, 1))
    inside  = str(round(p)) + "%" if p > 14 else ""
    outside = str(round(p)) + "%" if p <= 14 else ""
    return (
        '<div class="pbar-wrap">'
        '<div class="pbar-lbl">' + label + '</div>'
        '<div class="pbar-bg">'
        '<div class="pbar-fill" style="width:' + ps + '%;background:' + color + ';">' + inside + '</div>'
        '</div>'
        '<div style="font-size:10px;font-weight:700;color:#555;text-align:right;margin-top:2px;">' + outside + '</div>'
        '</div>'
    )

def html_timeline(step_idx):
    parts = []
    for i, (lbl, icon) in enumerate(zip(TL_LABELS, TL_ICONS)):
        if i < step_idx:
            cls = "tl-circle tl-done"
        elif i == step_idx:
            cls = "tl-circle tl-active"
        else:
            cls = "tl-circle tl-pend"
        parts.append(
            '<div class="tl-step">'
            '<div class="' + cls + '">' + icon + '</div>'
            '<div class="tl-lbl">' + lbl + '</div>'
            '</div>'
        )
        if i < len(TL_LABELS) - 1:
            lc = "tl-line tl-line-done" if i < step_idx else "tl-line tl-line-pend"
            parts.append('<div class="' + lc + '"></div>')
    return '<div class="tl-wrap">' + ''.join(parts) + '</div>'

def html_kpi_summary(cnts, cliente):
    tot   = sum(cnts.values()) or 1
    pct_p = str(int((cnts['OK'] + cnts['BLU']) / tot * 100))
    segs  = [
        (cnts['OK'],  '#4caf50', '✓'),
        (cnts['BLU'], '#2196f3', '●'),
        (cnts['LAV'], '#ffc107', '⚙'),
        (cnts['RIT'], '#f44336', '!'),
    ]
    sbar = ''.join(
        '<div class="sbar-s" style="width:' + str(round(c / tot * 100, 1)) + '%;background:' + col + ';">'
        + (sym if c / tot * 100 > 8 else '') + '</div>'
        for c, col, sym in segs if c > 0
    )
    return (
        '<div class="sum-card">'
        '<div class="sum-title">📊 Riepilogo Stato Ordini — ' + cliente + '</div>'
        '<div class="kpi-row">'
        '<div class="kpi-box kpi-g"><div class="kpi-num">' + str(cnts['OK'])  + '</div><div class="kpi-lbl">Pronti</div></div>'
        '<div class="kpi-box kpi-b"><div class="kpi-num">' + str(cnts['BLU']) + '</div><div class="kpi-lbl">Da Ritirare</div></div>'
        '<div class="kpi-box kpi-y"><div class="kpi-num">' + str(cnts['LAV']) + '</div><div class="kpi-lbl">In Lavorazione</div></div>'
        '<div class="kpi-box kpi-r"><div class="kpi-num">' + str(cnts['RIT']) + '</div><div class="kpi-lbl">In Ritardo</div></div>'
        '</div>'
        '<div style="font-size:11px;font-weight:700;color:#555;margin-bottom:4px;">'
        'Avanzamento complessivo (' + pct_p + '% pronto)</div>'
        '<div class="sbar">' + sbar + '</div>'
        '</div>'
    )

# =============================================================================
# AVVIO APP
# =============================================================================
USER_DB = load_users()
data    = load_data()
oggi_dt = datetime.now()

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if not st.session_state["authenticated"]:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if os.path.exists('Logo SAFIT.JPG'):
                st.image('Logo SAFIT.JPG', width=300)
            st.title("Accesso Area Riservata")
            user = st.text_input("Username")
            pw   = st.text_input("Password", type="password")
            if st.button("Accedi"):
                if user in USER_DB and str(USER_DB[user][0]) == pw:
                    st.session_state["authenticated"] = True
                    st.session_state["user_type"]     = USER_DB[user][1]
                    st.session_state["username"]      = user
                    st.rerun()
                else:
                    st.error("Username o Password errati")
        return False
    return True

if not check_password():
    st.stop()

# =============================================================================
# SIDEBAR
# =============================================================================
with st.sidebar:
    if os.path.exists('Logo SAFIT.JPG'):
        st.image('Logo SAFIT.JPG', use_container_width=True)
    st.write("Utente: **" + st.session_state['username'] + "**")
    st.markdown('<p class="version-tag">Versione: ' + APP_VERSION + '</p>', unsafe_allow_html=True)

    if st.session_state["user_type"] == "TUTTI":
        clienti_list = sorted([str(x) for x in data['Cliente Fornitore CD'].unique()])
        sel_cli = st.selectbox("👤 Seleziona Cliente:", clienti_list)
    else:
        sel_cli = st.session_state["user_type"]

    filtro_label = st.radio("Filtra per stato:", ["Mostra tutto", "Solo Disponibili", "In Lavorazione", "In Ritardo"])
    if st.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

# =============================================================================
# LOGICA CENTRALE — identica all'originale + grafiche
# =============================================================================
st.title("🚜 Portale Avanzamento Produzione")
df_cli = data[data['Cliente Fornitore CD'] == sel_cli].copy()

if not df_cli.empty:
    articoli_dis  = sorted([str(x) for x in df_cli['Articolo C'].unique()])
    sel_art       = st.selectbox("🔍 Cerca Codice Prodotto:", ["Tutti i prodotti"] + articoli_dis)
    articoli_view = articoli_dis if sel_art == "Tutti i prodotti" else [sel_art]

    # --- KPI RIEPILOGO ---
    if sel_art == "Tutti i prodotti":
        cnts = {'OK': 0, 'BLU': 0, 'LAV': 0, 'RIT': 0}
        for art in articoli_dis:
            df_a = df_cli[df_cli['Articolo C'] == art].sort_values('Data_Consegna')
            gia  = float(df_a['Gia'].iloc[0]) if 'Gia' in df_a.columns and pd.notnull(df_a['Gia'].iloc[0]) else 0.0
            lav  = float(df_a['Lavorazione_Totale'].iloc[0]) if 'Lavorazione_Totale' in df_a.columns and pd.notnull(df_a['Lavorazione_Totale'].iloc[0]) else 0.0
            for _, row in df_a.iterrows():
                qta = float(row['Qta_Effettiva'])
                req = row['Data_Consegna']
                if gia >= qta:
                    gia -= qta; eta = oggi_dt; cat = "DISP"
                elif (gia + lav) >= qta:
                    gia = 0; eta = aggiungi_giorni_lavorativi(oggi_dt, 10); cat = "LAV"
                else:
                    gia = 0; eta = aggiungi_giorni_lavorativi(oggi_dt, 25); cat = "LAV"
                late = pd.notnull(req) and eta.date() > req.date()
                if   cat == "DISP" and not late: cnts['OK']  += 1
                elif cat == "DISP" and late:     cnts['BLU'] += 1
                elif cat == "LAV"  and not late: cnts['LAV'] += 1
                else:                            cnts['RIT'] += 1
        st.markdown(html_kpi_summary(cnts, sel_cli), unsafe_allow_html=True)

    # --- DETTAGLIO ARTICOLI ---
    for art in articoli_view:
        df_art = df_cli[df_cli['Articolo C'] == art].sort_values('Data_Consegna')
        desc   = df_art['Articolo D'].iloc[0] if 'Articolo D' in df_art.columns else ""

        # Giacenza e lavorazione — IDENTICI all'originale
        st_gia = float(df_art['Gia'].iloc[0]) if 'Gia' in df_art.columns and pd.notnull(df_art['Gia'].iloc[0]) else 0.0
        st_lav = float(df_art['Lavorazione_Totale'].iloc[0]) if 'Lavorazione_Totale' in df_art.columns and pd.notnull(df_art['Lavorazione_Totale'].iloc[0]) else 0.0

        # Barra giacenza — snapshot PRIMA del loop (non consuma st_gia)
        qta_tot    = df_art['Qta_Effettiva'].sum()
        qta_pronta = min(st_gia, qta_tot)
        pct_art    = (qta_pronta / qta_tot * 100) if qta_tot > 0 else 0
        bar_col    = "#4caf50" if pct_art >= 100 else ("#ffc107" if pct_art > 0 else "#f44336")

        righe_mostra = []
        for _, row in df_art.iterrows():
            qta      = float(row['Qta_Effettiva'])
            req_date = row['Data_Consegna']

            # LOGICA ORIGINALE — non modificata
            if st_gia >= qta:
                st_gia -= qta
                eta, nota, cat_core = oggi_dt, "Pronto", "DISPONIBILE"
            elif (st_gia + st_lav) >= qta:
                st_gia = 0
                eta, nota, cat_core = aggiungi_giorni_lavorativi(oggi_dt, 10), "In Lavorazione", "LAVORAZIONE"
            else:
                eta, nota, cat_core = aggiungi_giorni_lavorativi(oggi_dt, 25), "Nuova Produzione", "LAVORAZIONE"

            is_ritardo = (pd.notnull(req_date) and eta.date() > req_date.date())

            if cat_core == "DISPONIBILE":
                if is_ritardo:
                    css, nota_display = "client-delay-row", "Pronto (Ritardo Ritiro)"
                else:
                    css, nota_display = "on-time-row", "Pronto"
            else:
                if is_ritardo:
                    css, nota_display = "prod-delay-row", nota + " (In Ritardo)"
                else:
                    css, nota_display = "delay-row", nota

            passa = False
            if filtro_label == "Mostra tutto": passa = True
            elif filtro_label == "Solo Disponibili" and cat_core == "DISPONIBILE": passa = True
            elif filtro_label == "In Lavorazione"  and cat_core == "LAVORAZIONE":  passa = True
            elif filtro_label == "In Ritardo"       and is_ritardo:                 passa = True

            step_idx = 2 if cat_core == "DISPONIBILE" else 1

            if passa:
                righe_mostra.append({
                    'css': css, 'date': req_date, 'qta': qta,
                    'eta': eta, 'nota': nota_display, 'step_idx': step_idx
                })

        if righe_mostra:
            with st.expander("📦 " + art + " — " + str(desc) + " | Residuo: " + str(int(df_art['Qta_Effettiva'].sum()))):

                # Barra giacenza
                st.markdown(
                    html_pbar(pct_art, bar_col,
                              "Giacenza: " + str(int(qta_pronta)) + " / " + str(int(qta_tot)) + " pz"),
                    unsafe_allow_html=True
                )

                for r in righe_mostra:
                    date_str = r['date'].strftime("%d/%m/%Y") if pd.notnull(r['date']) else "N.D."
                    eta_str  = r['eta'].strftime("%d/%m/%Y")
                    qta_str  = str(int(r['qta']))

                    # Riga stato — identica all'originale
                    st.markdown(
                        '<div class="status-row ' + r['css'] + '">'
                        '<span><b>Consegna:</b> ' + date_str + ' | <b>Q.tà:</b> ' + qta_str + '</span>'
                        '<span><b>Stima:</b> ' + eta_str + ' (' + r['nota'] + ')</span>'
                        '</div>',
                        unsafe_allow_html=True
                    )
                    # Timeline
                    st.markdown(html_timeline(r['step_idx']), unsafe_allow_html=True)
                    st.markdown('<hr style="margin:4px 0;border:none;border-top:1px solid #eee;">', unsafe_allow_html=True)






