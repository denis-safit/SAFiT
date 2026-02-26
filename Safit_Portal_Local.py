import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# --- 1. CONFIGURAZIONE PAGINA E VERSIONE ---
APP_VERSION = "1.1.00"
st.set_page_config(page_title=f"Safit - Portale Avanzamento {APP_VERSION}", layout="wide")

st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;600;700&display=swap');
    
    html, body, .stApp {{ font-family: 'IBM Plex Sans', sans-serif; }}
    .main {{ background-color: #f7f8fa; }}
    .stApp {{ margin-top: -30px; }}
    
    /* RIGHE STATO */
    .status-row {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap; 
        padding: 10px 15px;
        border-radius: 8px;
        margin-bottom: 6px;
        gap: 10px;
        font-size: 14px;
    }}
    .on-time-row {{ background-color: #f1f8e9; border-left: 6px solid #4caf50; color: #1b5e20; }}
    .client-delay-row {{ background-color: #e3f2fd; border-left: 6px solid #2196f3; color: #0d47a1; }}
    .delay-row {{ background-color: #fff8e1; border-left: 6px solid #ffc107; color: #5d4037; }}
    .prod-delay-row {{ background-color: #ffebee; border-left: 6px solid #f44336; color: #b71c1c; }}
    
    /* PROGRESS BAR CONTAINER */
    .progress-section {{
        background: #fff;
        border-radius: 10px;
        padding: 14px 18px;
        margin-bottom: 10px;
        border: 1px solid #e8eaf0;
        box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    }}
    .progress-label {{
        font-size: 12px;
        color: #666;
        margin-bottom: 6px;
        font-weight: 600;
        letter-spacing: 0.5px;
        text-transform: uppercase;
    }}
    .progress-bar-bg {{
        background: #e9ecef;
        border-radius: 20px;
        height: 18px;
        width: 100%;
        overflow: hidden;
        position: relative;
    }}
    .progress-bar-fill {{
        height: 100%;
        border-radius: 20px;
        transition: width 0.6s ease;
        display: flex;
        align-items: center;
        justify-content: flex-end;
        padding-right: 8px;
        font-size: 11px;
        font-weight: 700;
        color: white;
        min-width: 36px;
    }}
    .progress-bar-text-outside {{
        font-size: 12px;
        font-weight: 700;
        color: #444;
        text-align: right;
        margin-top: 3px;
    }}
    
    /* TIMELINE STEPS */
    .timeline-container {{
        display: flex;
        align-items: center;
        gap: 0;
        margin: 8px 0 2px 0;
        width: 100%;
    }}
    .step-block {{
        display: flex;
        flex-direction: column;
        align-items: center;
        flex: 1;
        position: relative;
    }}
    .step-circle {{
        width: 34px;
        height: 34px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 16px;
        font-weight: bold;
        z-index: 2;
        box-shadow: 0 2px 6px rgba(0,0,0,0.15);
        transition: transform 0.2s;
    }}
    .step-circle.done {{ background: #4caf50; color: white; }}
    .step-circle.active {{ background: #ff9800; color: white; animation: pulse 1.5s infinite; }}
    .step-circle.pending {{ background: #e0e0e0; color: #999; }}
    .step-label {{
        font-size: 10px;
        margin-top: 4px;
        text-align: center;
        font-weight: 600;
        color: #555;
        max-width: 70px;
        line-height: 1.2;
    }}
    .step-connector {{
        flex: 1;
        height: 4px;
        margin-top: -18px;
        z-index: 1;
    }}
    .step-connector.done {{ background: #4caf50; }}
    .step-connector.pending {{ background: #e0e0e0; }}
    
    /* RIEPILOGO CLIENTE */
    .summary-card {{
        background: white;
        border-radius: 12px;
        padding: 16px 20px;
        margin-bottom: 14px;
        border: 1px solid #e8eaf0;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }}
    .summary-title {{
        font-size: 13px;
        font-weight: 700;
        color: #333;
        margin-bottom: 12px;
        letter-spacing: 0.3px;
    }}
    .kpi-row {{
        display: flex;
        gap: 12px;
        flex-wrap: wrap;
        margin-bottom: 14px;
    }}
    .kpi-box {{
        flex: 1;
        min-width: 90px;
        text-align: center;
        padding: 10px 8px;
        border-radius: 8px;
    }}
    .kpi-box .kpi-num {{ font-size: 22px; font-weight: 700; }}
    .kpi-box .kpi-lbl {{ font-size: 10px; font-weight: 600; text-transform: uppercase; opacity: 0.75; margin-top: 2px; }}
    .kpi-green {{ background: #f1f8e9; color: #2e7d32; }}
    .kpi-blue {{ background: #e3f2fd; color: #1565c0; }}
    .kpi-yellow {{ background: #fff8e1; color: #e65100; }}
    .kpi-red {{ background: #ffebee; color: #b71c1c; }}
    
    /* STACKED BAR */
    .stacked-bar {{
        display: flex;
        height: 20px;
        border-radius: 12px;
        overflow: hidden;
        width: 100%;
        gap: 2px;
    }}
    .stacked-segment {{
        height: 100%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 10px;
        font-weight: 700;
        color: white;
        min-width: 4px;
        transition: width 0.5s ease;
    }}
    
    @keyframes pulse {{
        0%, 100% {{ transform: scale(1); box-shadow: 0 0 0 0 rgba(255,152,0,0.4); }}
        50% {{ transform: scale(1.08); box-shadow: 0 0 0 6px rgba(255,152,0,0); }}
    }}
    
    .stExpander div {{ height: auto !important; min-height: min-content !important; }}
    .version-tag {{ font-size: 10px; color: #999; text-align: right; }}
    </style>
    """, unsafe_allow_html=True)

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

USER_DB = load_users()

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state["authenticated"] = False
    if not st.session_state["authenticated"]:
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', width=300)
            st.title("Accesso Area Riservata")
            user = st.text_input("Username")
            pw = st.text_input("Password", type="password")
            if st.button("Accedi"):
                if user in USER_DB and str(USER_DB[user][0]) == pw:
                    st.session_state["authenticated"] = True
                    st.session_state["user_type"] = USER_DB[user][1]
                    st.session_state["username"] = user
                    st.rerun()
                else: st.error("Username o Password errati")
        return False
    return True

if not check_password(): st.stop()

# --- 3. FUNZIONI TECNICHE ---
def aggiungi_giorni_lavorativi(data_inizio, giorni):
    data_corrente = data_inizio
    while giorni > 0:
        data_corrente += timedelta(days=1)
        if data_corrente.weekday() < 5: giorni -= 1
    return data_corrente

def pulisci_numero(serie):
    return pd.to_numeric(serie.astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False), errors='coerce').fillna(0)

@st.cache_data
def load_data():
    try:
        df = pd.read_excel('righe_Ordini_ARCA.xlsx', sheet_name='Foglio1', skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        for col in ['Cliente Fornitore CD', 'Articolo C', 'Articolo D', 'Data']:
            if col in df.columns: df[col] = df[col].ffill()
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
                    if c in df_tech.columns: df_tech[c] = pulisci_numero(df_tech[c])
                df_tech['Lavorazione_Totale'] = df_tech.get('Acq', 0) + df_tech.get('Lan', 0) + \
                                               df_tech.get('Grz', 0) + df_tech.get('Tmp', 0) + \
                                               df_tech.get('Rwi', 0) + df_tech.get('Trs', 0)
                df = pd.merge(df, df_tech[['Art_Key', 'Gia', 'Lavorazione_Totale']], 
                              left_on='Articolo C', right_on='Art_Key', how='left')
        return df
    except: return pd.DataFrame()

data = load_data()
oggi_dt = datetime.now()

# --- 4. FUNZIONI GRAFICHE ---

def render_progress_bar(pct, color):
    """Barra di progresso colorata con percentuale."""
    pct_clamp = min(max(pct, 0), 100)
    label_inside = f"{pct_clamp:.0f}%" if pct_clamp > 15 else ""
    label_outside = f"{pct_clamp:.0f}%" if pct_clamp <= 15 else ""
    return f"""
    <div class="progress-bar-bg">
        <div class="progress-bar-fill" style="width:{pct_clamp}%; background:{color};">
            {label_inside}
        </div>
    </div>
    <div class="progress-bar-text-outside">{label_outside}</div>
    """

def render_timeline(step_index, labels, icons):
    """Timeline con step evidenziati. step_index: 0=ordine, 1=lavorazione, 2=pronto, 3=consegnato."""
    html = '<div class="timeline-container">'
    for i, (lbl, icon) in enumerate(zip(labels, icons)):
        if i < step_index:
            circle_cls = "done"
        elif i == step_index:
            circle_cls = "active"
        else:
            circle_cls = "pending"
        html += f'<div class="step-block"><div class="step-circle {circle_cls}">{icon}</div><div class="step-label">{lbl}</div></div>'
        if i < len(labels) - 1:
            conn_cls = "done" if i < step_index else "pending"
            html += f'<div class="step-connector {conn_cls}"></div>'
    html += '</div>'
    return html

TIMELINE_LABELS = ["Ordinato", "In Lavorazione", "Pronto", "Consegnato"]
TIMELINE_ICONS  = ["📋", "⚙️", "✅", "🚚"]

def get_step_and_pct(cat_core, nota):
    """Restituisce (step_index, percentuale_avanzamento)."""
    if "Pronto" in nota:
        return 2, 100
    elif "Lavorazione" in nota:
        return 1, 50
    elif "Nuova Produzione" in nota:
        return 1, 20
    else:
        return 0, 5

def render_article_summary_bar(counts):
    """Barra impilata colorata con i conteggi per stato."""
    total = sum(counts.values()) or 1
    segments = [
        (counts.get('DISPONIBILE_OK', 0), '#4caf50', '✓'),
        (counts.get('DISPONIBILE_RITARDO', 0), '#2196f3', '●'),
        (counts.get('LAV_OK', 0), '#ffc107', '⚙'),
        (counts.get('LAV_RITARDO', 0), '#f44336', '!'),
    ]
    html = '<div class="stacked-bar">'
    for cnt, color, sym in segments:
        if cnt == 0: continue
        w = cnt / total * 100
        label = sym if w > 8 else ''
        html += f'<div class="stacked-segment" style="width:{w:.1f}%; background:{color};">{label}</div>'
    html += '</div>'
    return html

# --- 5. SIDEBAR ---
with st.sidebar:
    if os.path.exists('Logo SAFIT.JPG'): st.image('Logo SAFIT.JPG', use_container_width=True)
    st.write(f"Utente: **{st.session_state['username']}**")
    st.markdown(f'<p class="version-tag">Versione: {APP_VERSION}</p>', unsafe_allow_html=True)
    
    if st.session_state["user_type"] == "TUTTI":
        clienti_list = sorted([str(x) for x in data['Cliente Fornitore CD'].unique()])
        sel_cli = st.selectbox("👤 Seleziona Cliente:", clienti_list)
    else: sel_cli = st.session_state["user_type"]
    
    filtro_label = st.radio("Filtra per stato:", ["Mostra tutto", "Solo Disponibili", "In Lavorazione", "In Ritardo"])
    if st.button("Logout"):
        st.session_state["authenticated"] = False
        st.rerun()

# --- 6. LOGICA CENTRALE ---
st.title("🚜 Portale Avanzamento Produzione")
df_cli = data[data['Cliente Fornitore CD'] == sel_cli].copy()

if not df_cli.empty:
    articoli_dis = sorted([str(x) for x in df_cli['Articolo C'].unique()])
    sel_art = st.selectbox("🔍 Cerca Codice Prodotto:", ["Tutti i prodotti"] + articoli_dis)
    articoli_view = articoli_dis if sel_art == "Tutti i prodotti" else [sel_art]

    # --- RIEPILOGO CLIENTE ---
    if sel_art == "Tutti i prodotti":
        tot_counts = {'DISPONIBILE_OK': 0, 'DISPONIBILE_RITARDO': 0, 'LAV_OK': 0, 'LAV_RITARDO': 0}
        for art in articoli_dis:
            df_art = df_cli[df_cli['Articolo C'] == art]
            st_gia = float(df_art['Gia'].iloc[0]) if 'Gia' in df_art.columns and pd.notnull(df_art['Gia'].iloc[0]) else 0.0
            st_lav = float(df_art['Lavorazione_Totale'].iloc[0]) if 'Lavorazione_Totale' in df_art.columns and pd.notnull(df_art['Lavorazione_Totale'].iloc[0]) else 0.0
            for _, row in df_art.iterrows():
                qta = float(row['Qta_Effettiva'])
                req_date = row['Data_Consegna']
                if st_gia >= qta:
                    st_gia -= qta
                    cat_core = "DISPONIBILE"
                    eta = oggi_dt
                elif (st_gia + st_lav) >= qta:
                    st_gia = 0
                    cat_core = "LAVORAZIONE"
                    eta = aggiungi_giorni_lavorativi(oggi_dt, 10)
                else:
                    cat_core = "LAVORAZIONE"
                    eta = aggiungi_giorni_lavorativi(oggi_dt, 25)
                is_ritardo = (pd.notnull(req_date) and eta.date() > req_date.date())
                
                if cat_core == "DISPONIBILE":
                    tot_counts['DISPONIBILE_RITARDO' if is_ritardo else 'DISPONIBILE_OK'] += 1
                else:
                    tot_counts['LAV_RITARDO' if is_ritardo else 'LAV_OK'] += 1

        totale_righe = sum(tot_counts.values()) or 1
        pct_pronto = (tot_counts['DISPONIBILE_OK'] + tot_counts['DISPONIBILE_RITARDO']) / totale_righe * 100

        st.markdown(f"""
        <div class="summary-card">
            <div class="summary-title">📊 Riepilogo Stato Ordini — {sel_cli}</div>
            <div class="kpi-row">
                <div class="kpi-box kpi-green"><div class="kpi-num">{tot_counts['DISPONIBILE_OK']}</div><div class="kpi-lbl">Pronti</div></div>
                <div class="kpi-box kpi-blue"><div class="kpi-num">{tot_counts['DISPONIBILE_RITARDO']}</div><div class="kpi-lbl">Da Ritirare</div></div>
                <div class="kpi-box kpi-yellow"><div class="kpi-num">{tot_counts['LAV_OK']}</div><div class="kpi-lbl">In Lavorazione</div></div>
                <div class="kpi-box kpi-red"><div class="kpi-num">{tot_counts['LAV_RITARDO']}</div><div class="kpi-lbl">In Ritardo</div></div>
            </div>
            <div class="progress-label">Avanzamento complessivo ({pct_pronto:.0f}% pronto)</div>
            {render_article_summary_bar(tot_counts)}
        </div>
        """, unsafe_allow_html=True)

    # --- DETTAGLIO ARTICOLI ---
    for art in articoli_view:
        df_art = df_cli[df_cli['Articolo C'] == art].sort_values('Data_Consegna')
        desc = df_art['Articolo D'].iloc[0] if 'Articolo D' in df_art.columns else ""
        
        st_gia = float(df_art['Gia'].iloc[0]) if 'Gia' in df_art.columns and pd.notnull(df_art['Gia'].iloc[0]) else 0.0
        st_lav = float(df_art['Lavorazione_Totale'].iloc[0]) if 'Lavorazione_Totale' in df_art.columns and pd.notnull(df_art['Lavorazione_Totale'].iloc[0]) else 0.0
        qta_totale = df_art['Qta_Effettiva'].sum()
        
        # Per la barra articolo: calcolo qty già pronte vs totale
        qta_pronta_art = min(st_gia, qta_totale)
        pct_art = (qta_pronta_art / qta_totale * 100) if qta_totale > 0 else 0

        righe_mostra = []
        st_gia_iter = st_gia  # iterazione separata per display
        for _, row in df_art.iterrows():
            qta = float(row['Qta_Effettiva'])
            req_date = row['Data_Consegna']
            
            if st_gia_iter >= qta:
                st_gia_iter -= qta
                eta, nota, cat_core = oggi_dt, "Pronto", "DISPONIBILE"
            elif (st_gia_iter + st_lav) >= qta:
                st_gia_iter = 0
                eta, nota, cat_core = aggiungi_giorni_lavorativi(oggi_dt, 10), "In Lavorazione", "LAVORAZIONE"
            else:
                eta, nota, cat_core = aggiungi_giorni_lavorativi(oggi_dt, 25), "Nuova Produzione", "LAVORAZIONE"

            is_ritardo = (pd.notnull(req_date) and eta.date() > req_date.date())
            
            if cat_core == "DISPONIBILE":
                if is_ritardo:
                    css, nota_display, bar_color = "client-delay-row", "Pronto (Ritardo Ritiro)", "#2196f3"
                else:
                    css, nota_display, bar_color = "on-time-row", "Pronto", "#4caf50"
            else:
                if is_ritardo:
                    css, nota_display, bar_color = "prod-delay-row", f"{nota} (In Ritardo)", "#f44336"
                else:
                    css, nota_display, bar_color = "delay-row", nota, "#ffc107"

            passa = False
            if filtro_label == "Mostra tutto": passa = True
            elif filtro_label == "Solo Disponibili" and cat_core == "DISPONIBILE": passa = True
            elif filtro_label == "In Lavorazione" and cat_core == "LAVORAZIONE": passa = True
            elif filtro_label == "In Ritardo" and is_ritardo: passa = True

            step_idx, pct_riga = get_step_and_pct(cat_core, nota_display)

            if passa:
                righe_mostra.append({
                    'css': css, 'date': req_date, 'qta': qta, 'eta': eta,
                    'nota': nota_display, 'bar_color': bar_color,
                    'step_idx': step_idx, 'pct_riga': pct_riga
                })

        if righe_mostra:
            # Colore barra articolo in base allo stato prevalente
            art_bar_color = "#4caf50" if pct_art >= 100 else ("#ffc107" if pct_art > 0 else "#f44336")
            
            with st.expander(f"📦 {art} — {desc} | Residuo: {qta_totale:,.0f}"):
                # Mini riepilogo articolo con barra
                st.markdown(f"""
                <div class="progress-section">
                    <div class="progress-label">Giacenza disponibile: {min(st_gia, qta_totale):,.0f} / {qta_totale:,.0f} pz</div>
                    {render_progress_bar(pct_art, art_bar_color)}
                </div>
                """, unsafe_allow_html=True)
                
                st.markdown("---")
                
                for r in righe_mostra:
                    # Riga stato
                    st.markdown(
                        f'<div class="status-row {r["css"]}">'
                        f'<span><b>Consegna:</b> {r["date"].strftime("%d/%m/%Y") if pd.notnull(r["date"]) else "N.D."} | <b>Q.tà:</b> {r["qta"]:,.0f}</span>'
                        f'<span><b>Stima:</b> {r["eta"].strftime("%d/%m/%Y")} ({r["nota"]})</span>'
                        f'</div>',
                        unsafe_allow_html=True
                    )
                    # Timeline per questa riga
                    st.markdown(
                        render_timeline(r['step_idx'], TIMELINE_LABELS, TIMELINE_ICONS),
                        unsafe_allow_html=True
                    )
                    st.markdown("<hr style='margin:6px 0; border:none; border-top:1px solid #eee;'>", unsafe_allow_html=True)

else:
    st.info("Nessun dato disponibile per il cliente selezionato.")
