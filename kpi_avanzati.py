# ==============================================================================
# kpi_avanzati.py — Modulo KPI Avanzati per Safit Portal  v2.0
# ==============================================================================
# Struttura attesa del file ARCA (righe_ordini_storico_con_date.xlsx):
#   Colonne: Codice Documento | Data Consegna | Numero Documento |
#            Articolo C | Articolo D | Cliente Fornitore CD | Data | Qta Doc | Valore
#
# INTEGRAZIONE in Safit_Portal_Local.py:
#   1. Copia kpi_avanzati.py nella stessa cartella del portale
#   2. Aggiungi in cima:  from kpi_avanzati import render_kpi_avanzati
#   3. Alla fine della vista admin (dopo il download button):
#        st.markdown("---")
#        render_kpi_avanzati()
# ==============================================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta, date
from io import BytesIO
import os

# ── Percorsi file ─────────────────────────────────────────────────────────────
import os as _os
_DIR = _os.path.dirname(_os.path.abspath(__file__))
PATH_STORICO   = _os.path.join(_DIR, "righe_ordini_storico_con_date.xlsx")
PATH_CONSEGNE  = _os.path.join(_DIR, "dettagli_consegne.xlsx")

# ── CSS ───────────────────────────────────────────────────────────────────────
KPI_CSS = """
<style>
.kpi-adv {
    background: linear-gradient(135deg,#ffffff,#f8faff);
    border:1px solid #e3e8f0; border-left:4px solid #1f77b4;
    padding:14px 16px; border-radius:10px; margin-bottom:6px;
    box-shadow:0 2px 6px rgba(0,0,0,0.05);
    color:#1a1a2e !important;
}
.kpi-adv * { color:#1a1a2e !important; }
.kpi-adv.g { border-left-color:#4caf50; }
.kpi-adv.o { border-left-color:#ff9800; }
.kpi-adv.r { border-left-color:#f44336; }
.kpi-adv.p { border-left-color:#9c27b0; }
.kpi-adv-t { font-size:10px; color:#888 !important; text-transform:uppercase;
             letter-spacing:.5px; margin-bottom:3px; }
.kpi-adv-v { font-size:26px; font-weight:700; color:#1a1a2e !important; line-height:1; }
.kpi-adv-s { font-size:11px; color:#555 !important; margin-top:3px; }
.sec-h { font-size:16px; font-weight:700; color:#1a1a2e !important;
         border-bottom:2px solid #e3e8f0; padding-bottom:5px;
         margin:20px 0 12px 0; }
.alert-box { background:#fff8e1; border-left:4px solid #ff9800;
             padding:10px 14px; border-radius:6px; margin:6px 0;
             font-size:13px; color:#1a1a1a !important; }
.alert-box b { color:#1a1a1a !important; }
.alert-box * { color:#1a1a1a !important; }
</style>
"""


# ── Caricamento e pulizia dati ────────────────────────────────────────────────

@st.cache_data(ttl=1800, show_spinner=False)
def carica_storico_arca(path=PATH_STORICO):
    """
    Carica e normalizza il file export ARCA con struttura pivot.
    Gestisce: header a riga 2, Codice Documento con ffill,
    righe di totale, campi vuoti '(vuoto)'.
    """
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        df = pd.read_excel(path, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]

        # Propaga Codice Documento (struttura pivot ARCA)
        df['Codice Documento'] = df['Codice Documento'].ffill()

        # Rimuovi righe di totale e righe vuote
        mask_totale = df['Codice Documento'].astype(str).str.contains(
            'Totale|NaN|nan', na=True
        )
        df = df[~mask_totale].copy()
        df = df[
            df['Articolo C'].notna() &
            (df['Articolo C'].astype(str) != '(vuoto)')
        ].copy()

        # Normalizza date
        df['Data']          = pd.to_datetime(df['Data'],          errors='coerce')
        df['Data Consegna'] = pd.to_datetime(df['Data Consegna'], errors='coerce')
        # NOTA: non propago Data Consegna tra righe — ogni riga deve avere la propria data
        # Le righe senza Data Consegna vengono escluse dal calcolo puntualità

        # Normalizza quantità
        df['Qta Doc'] = pd.to_numeric(df['Qta Doc'], errors='coerce').fillna(0)
        df['Valore']  = pd.to_numeric(df['Valore'],  errors='coerce').fillna(0)

        # Estrai cliente (es. "C000289 - ORION CALZATURIFICIO SPA" → "ORION CALZATURIFICIO SPA")
        df['Cliente'] = df['Cliente Fornitore CD'].astype(str).str.split(' - ', n=1).str[-1].str.strip()
        df['Cod_Cliente'] = df['Cliente Fornitore CD'].astype(str).str.split(' - ', n=1).str[0].str.strip()

        # Famiglia da prime 2 parole Articolo D
        df['Famiglia'] = df['Articolo D'].apply(
            lambda x: ' '.join(str(x).split()[:2]).upper()
        )

        return df
    except Exception as e:
        st.warning(f"Errore lettura storico ARCA: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=1800, show_spinner=False)
def carica_dettagli_consegne(path=PATH_CONSEGNE):
    """
    Carica il file dettagli_consegne.xlsx esportato da ARCA.
    Contiene già lo scostamento calcolato (datadifference) per ogni riga DVF.
    Colonne chiave:
      Cd_CF, CF_Descrizione, Cd_AR, DORig_Descrizione,
      DataDoc (data reale consegna), DataConsegnaB (data prevista OCI),
      datadifference (gg scostamento: positivo=ritardo, negativo=anticipo),
      Qta, QtaEvasa
    """
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        df = pd.read_excel(path)
        df.columns = [str(c).strip() for c in df.columns]
        # Tieni solo righe con articolo valorizzato (escludi intestazioni DVF)
        df = df[df['Cd_AR'].notna() & (df['Cd_AR'].astype(str).str.strip() != '')].copy()
        # Solo DVF = documenti di trasporto (consegna reale merce)
        # Escludi FVU (fatture), OCI (ordini), e tutti gli altri
        df = df[df['Cd_Do'] == 'DVF'].copy()
        # Solo clienti (Cd_CF inizia con C)
        df = df[df['Cd_CF'].astype(str).str.strip().str.startswith('C')].copy()
        df['DataDoc']        = pd.to_datetime(df['DataDoc'],        errors='coerce')
        df['DataConsegnaB']  = pd.to_datetime(df['DataConsegnaB'],  errors='coerce')
        df['dataconsegnaO']  = pd.to_datetime(df['dataconsegnaO'],  errors='coerce')
        df['datadifference'] = pd.to_numeric(df['datadifference'],  errors='coerce')
        df['Qta']            = pd.to_numeric(df['Qta'],             errors='coerce').fillna(0)
        df['QtaEvasa']       = pd.to_numeric(df['QtaEvasa'],        errors='coerce').fillna(0)
        # Estrai nome cliente pulito
        df['Cliente'] = df['CF_Descrizione'].astype(str).str.strip()
        df['Articolo C'] = df['Cd_AR'].astype(str).str.strip().str.upper()
        df['Famiglia'] = df['DORig_Descrizione'].apply(
            lambda x: ' '.join(str(x).split()[:2]).upper()
        )
        return df
    except Exception as e:
        return pd.DataFrame()


def get_oci_oca(df):
    """Filtra solo OCI con quantità > 0."""
    return df[
        df['Codice Documento'].isin(['OCI']) &
        (df['Qta Doc'] > 0)
    ].copy()


def get_dvf(df):
    """Restituisce i DVF (documenti di vendita = evasioni reali)."""
    return df[df['Codice Documento'] == 'DVF'].copy()


# ── Calcoli KPI ───────────────────────────────────────────────────────────────

def calcola_storicita_rotazione(df_ordini):
    """Storicità e rotazione per articolo."""
    if df_ordini.empty:
        return pd.DataFrame()

    grp = df_ordini.groupby(['Articolo C', 'Articolo D', 'Famiglia']).agg(
        N_Ordini        =('Data', 'count'),
        Qta_Totale      =('Qta Doc', 'sum'),
        Valore_Totale   =('Valore', 'sum'),
        Prima_Richiesta =('Data', 'min'),
        Ultima_Richiesta=('Data', 'max'),
        N_Clienti       =('Cliente', 'nunique'),
    ).reset_index()

    grp['Giorni_Attivita'] = (
        grp['Ultima_Richiesta'] - grp['Prima_Richiesta']
    ).dt.days.clip(lower=1)
    grp['Freq_Settimanale'] = (
        grp['N_Ordini'] / (grp['Giorni_Attivita'] / 7)
    ).round(2)

    return grp.sort_values('Qta_Totale', ascending=False).reset_index(drop=True)


def calcola_frequenza_riordino(df_ordini):
    """Intervallo medio di riordino per articolo × cliente."""
    if df_ordini.empty:
        return pd.DataFrame()

    risultati = []
    for (art, cli), g in df_ordini.groupby(['Articolo C', 'Cliente']):
        date_ordini = sorted(g['Data'].dropna().tolist())
        if len(date_ordini) < 2:
            intervallo = None
        else:
            delta = [(date_ordini[i+1]-date_ordini[i]).days
                     for i in range(len(date_ordini)-1)]
            intervallo = round(float(np.mean(delta)), 1)

        ultimo = max(date_ordini) if date_ordini else pd.NaT
        qta_media = round(g['Qta Doc'].mean(), 0) if len(g) > 0 else 0
        risultati.append({
            'Articolo C':          art,
            'Articolo D':          g['Articolo D'].iloc[0],
            'Famiglia':            g['Famiglia'].iloc[0],
            'Cliente':             cli,
            'Cod_Cliente':         g['Cod_Cliente'].iloc[0],
            'N_Ordini':            len(date_ordini),
            'Qta_Totale':          g['Qta Doc'].sum(),
            'Qta_Media_Ordine':    qta_media,
            'Intervallo_Medio_gg': intervallo,
            'Ultimo_Ordine':       ultimo,
            'Giorni_Da_Ultimo':    (datetime.now() - ultimo).days if pd.notnull(ultimo) else None,
        })

    return pd.DataFrame(risultati).sort_values('N_Ordini', ascending=False)


def calcola_scostamento_consegna(df_cons, filtro_cliente=None, filtro_articolo=None, filtro_famiglie=None, cutoff=None):
    """
    Calcola puntualità consegne dal file dettagli_consegne.xlsx.
    Usa datadifference già calcolato da ARCA — nessun join complicato.
    Positivo = ritardo, Negativo = anticipo, 0 = puntuale.
    """
    if df_cons.empty:
        return pd.DataFrame(), {}

    df = df_cons.copy()

    # Applica filtri
    if filtro_cliente:
        nome = filtro_cliente.split(' - ', 1)[-1].strip() if ' - ' in filtro_cliente else filtro_cliente.strip()
        df = df[df['Cliente'].str.contains(nome, case=False, na=False, regex=False)]
    if filtro_articolo:
        df = df[df['Articolo C'].str.contains(filtro_articolo.upper(), case=False, na=False)]
    if filtro_famiglie:
        df = df[df['Famiglia'].isin(filtro_famiglie)]
    if cutoff is not None:
        df = df[df['DataDoc'] >= cutoff]

    # Usa dataconsegnaO = data scadenza originale dell'OCI (più affidabile di DataConsegnaB)
    # Escludi righe senza entrambe le date
    df = df[df['dataconsegnaO'].notna() & df['DataDoc'].notna()].copy()

    if df.empty:
        return pd.DataFrame(), {}

    # Scostamento = DataDoc (consegna reale) - dataconsegnaO (scadenza OCI originale)
    # Positivo = ritardo, Negativo = anticipo
    df['Scostamento_gg'] = (df['DataDoc'] - df['dataconsegnaO']).dt.days
    df['Data Consegna']  = df['dataconsegnaO']
    df['Data_Evasione']  = df['DataDoc']

    # Escludi anomalie (ritardo > 365 gg = dato spurio)
    df = df[df['Scostamento_gg'].between(-365, 365)].copy()

    # Anticipo o puntuale = tutto ≤ 2 gg di ritardo (inclusi anticipi)
    df['Stato_Consegna'] = df['Scostamento_gg'].apply(
        lambda x: '⚠️ In ritardo' if x > 2 else '✅ Puntuale'
    )

    ritardi = df[df['Scostamento_gg'] > 2]['Scostamento_gg']
    summary = {
        'n_evasi':      len(df),
        'media_scost':  round(ritardi.mean(), 1) if not ritardi.empty else 0.0,
        'mediana_scost':round(df['Scostamento_gg'].median(), 1),
        'pct_puntuali': round((df['Scostamento_gg'] <= 2).mean() * 100, 1),
        'pct_ritardo':  round((df['Scostamento_gg'] > 2).mean() * 100, 1),
    }
    return df, summary


def calcola_trend_mensile(df_ordini):
    """Volume ordini per mese."""
    if df_ordini.empty or 'Data' not in df_ordini.columns:
        return pd.DataFrame()
    df = df_ordini.copy()
    df['Mese'] = df['Data'].dt.to_period('M').astype(str)
    return df.groupby('Mese').agg(
        N_Righe=('Articolo C', 'count'),
        Qta_Totale=('Qta Doc', 'sum'),
        N_Clienti=('Cliente', 'nunique'),
        N_Articoli=('Articolo C', 'nunique'),
    ).reset_index().sort_values('Mese')


# ── Helper UI ─────────────────────────────────────────────────────────────────

def kpi_card(col, title, value, sub="", color=""):
    col.markdown(
        f'<div class="kpi-adv {color}">'
        f'<div class="kpi-adv-t">{title}</div>'
        f'<div class="kpi-adv-v">{value}</div>'
        f'<div class="kpi-adv-s">{sub}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


# ── Rendering principale ──────────────────────────────────────────────────────

def render_kpi_avanzati(path_storico=PATH_STORICO, filtro_cliente=None, filtro_articolo=None, filtro_famiglie=None, key_prefix="kpi"):
    """
    Punto di ingresso. Da chiamare nella vista admin del portale.
    """
    st.markdown(KPI_CSS, unsafe_allow_html=True)

    with st.expander("📊 KPI Avanzati — Storicità · Rotazione · Riordino · Puntualità",
                     expanded=False):

        # ── Caricamento ───────────────────────────────────────────────────────
        with st.spinner("Caricamento dati ARCA..."):
            df_all = carica_storico_arca(path_storico)

        if df_all.empty:
            st.warning(f"File non trovato: `{path_storico}`. "
                       "Copia il file nella cartella del portale e ricarica.")
            return

        df_oci = get_oci_oca(df_all)
        df_dvf = get_dvf(df_all)

        # Filtro cliente pre-applicato (da vista cliente o selezione admin)
        if filtro_cliente:
            nome_cli = filtro_cliente.split(' - ', 1)[-1].strip() if ' - ' in filtro_cliente else filtro_cliente.strip()
            df_oci = df_oci[df_oci['Cliente'].str.contains(nome_cli, case=False, na=False, regex=False)]
            df_dvf = df_dvf[df_dvf['Cliente'].str.contains(nome_cli, case=False, na=False, regex=False)]

        # Filtro articolo dalla sidebar (Cerca Articolo)
        if filtro_articolo:
            df_oci = df_oci[df_oci['Articolo C'].str.contains(filtro_articolo.upper(), case=False, na=False, regex=False)]
            df_dvf = df_dvf[df_dvf['Articolo C'].str.contains(filtro_articolo.upper(), case=False, na=False, regex=False)]

        # Filtro famiglie dai filtri globali del portale
        if filtro_famiglie:
            df_oci = df_oci[df_oci['Famiglia'].isin(filtro_famiglie)]
            df_dvf = df_dvf[df_dvf['Famiglia'].isin(filtro_famiglie)]

        st.caption(
            f"📁 Storico caricato: **{len(df_all):,}** righe totali | "
            f"**{len(df_oci):,}** OCI | **{len(df_dvf):,}** DVF".replace(",",".")
        )

        # ── Filtri ────────────────────────────────────────────────────────────
        import hashlib as _hl
        _key_suffix = _hl.md5((
            str(key_prefix) +
            str(filtro_cliente or "") +
            str(filtro_articolo or "") +
            str(sorted(filtro_famiglie or []))
        ).encode()).hexdigest()[:12]

        c1, c2 = st.columns([1, 2])
        with c1:
            periodo = st.selectbox(
                "Periodo",
                ["Mese in corso", "Anno in corso",
                 "Ultimi 90 gg", "Ultimi 6 mesi", "Ultimo anno",
                 "Tutto lo storico", "Intervallo personalizzato"],
                index=4, key=f"kpi_periodo_{_key_suffix}"
            )
        with c2:
            famiglie = sorted(df_oci['Famiglia'].dropna().unique().tolist())
            sel_fam  = st.multiselect("Famiglia", famiglie,
                                       key=f"kpi_fam_{_key_suffix}",
                                       placeholder="Tutte le famiglie")

        # Se intervallo personalizzato: mostra i due date_input
        if periodo == "Intervallo personalizzato":
            data_min_d = df_oci["Data"].min().date() if not df_oci.empty else date(2024, 1, 1)
            data_max_d = df_oci["Data"].max().date() if not df_oci.empty else datetime.now().date()
            dc1, dc2 = st.columns(2)
            with dc1:
                data_da = st.date_input(
                    "📅 Dal", value=data_min_d,
                    min_value=data_min_d, max_value=data_max_d,
                    key=f"kpi_da_{_key_suffix}", format="DD/MM/YYYY"
                )
            with dc2:
                data_a = st.date_input(
                    "📅 Al", value=data_max_d,
                    min_value=data_min_d, max_value=data_max_d,
                    key=f"kpi_a_{_key_suffix}", format="DD/MM/YYYY"
                )

        if filtro_cliente:
            st.caption(f"👤 Dati filtrati per cliente: **{filtro_cliente}**")
        if filtro_articolo:
            st.caption(f"🔍 Dati filtrati per articolo: **{filtro_articolo.upper()}**")
        if filtro_famiglie:
            st.caption(f"📂 Famiglie attive: **{', '.join(filtro_famiglie)}**")

        # Applica filtri temporali
        gg_map = {"Ultimi 90 gg": 90, "Ultimi 6 mesi": 180,
                  "Ultimo anno": 365, "Tutto lo storico": 9999}
        df_f     = df_oci.copy()
        df_dvf_f = df_dvf.copy()
        cutoff   = None
        gg       = 9999

        _oggi = datetime.now().date()

        if periodo == "Mese in corso":
            cutoff_da = pd.Timestamp(_oggi.replace(day=1))
            cutoff_a  = pd.Timestamp(_oggi)
            df_f      = df_f[(df_f["Data"] >= cutoff_da) & (df_f["Data"] <= cutoff_a)]
            df_dvf_f  = df_dvf_f[(df_dvf_f["Data"] >= cutoff_da) & (df_dvf_f["Data"] <= cutoff_a)]
            cutoff    = cutoff_da
            gg        = max(1, (_oggi - _oggi.replace(day=1)).days + 1)
        elif periodo == "Anno in corso":
            # Dal 1° gennaio all'ultimo giorno del mese scorso
            from calendar import monthrange as _mr
            _mese_scorso = (_oggi.month - 1) or 12
            _anno_ms = _oggi.year if _oggi.month > 1 else _oggi.year - 1
            _ultimo_ms = _mr(_anno_ms, _mese_scorso)[1]
            cutoff_da = pd.Timestamp(date(_oggi.year, 1, 1))
            cutoff_a  = pd.Timestamp(date(_anno_ms, _mese_scorso, _ultimo_ms))
            df_f      = df_f[(df_f["Data"] >= cutoff_da) & (df_f["Data"] <= cutoff_a)]
            df_dvf_f  = df_dvf_f[(df_dvf_f["Data"] >= cutoff_da) & (df_dvf_f["Data"] <= cutoff_a)]
            cutoff    = cutoff_da
            gg        = max(1, (cutoff_a.date() - cutoff_da.date()).days)
        elif periodo == "Intervallo personalizzato":
            cutoff_da = pd.Timestamp(data_da)
            cutoff_a  = pd.Timestamp(data_a)
            df_f      = df_f[(df_f["Data"] >= cutoff_da) & (df_f["Data"] <= cutoff_a)]
            df_dvf_f  = df_dvf_f[(df_dvf_f["Data"] >= cutoff_da) & (df_dvf_f["Data"] <= cutoff_a)]
            cutoff    = cutoff_da
        else:
            gg = gg_map[periodo]
            if gg < 9999:
                cutoff   = pd.Timestamp(datetime.now() - timedelta(days=gg))
                df_f     = df_f[df_f["Data"] >= cutoff]
                df_dvf_f = df_dvf_f[df_dvf_f["Data"] >= cutoff]

        if sel_fam:
            df_f = df_f[df_f["Famiglia"].isin(sel_fam)]

        if df_f.empty:
            st.warning("Nessun dato per i filtri selezionati.")
            return

        st.caption(
            f"🔍 Filtro attivo: **{len(df_f):,}** righe | "
            f"**{df_f['Articolo C'].nunique()}** articoli | "
            f"**{df_f['Cliente'].nunique()}** clienti".replace(",",".")
        )

        # ════════════════════════════════════════════════════════════════════
        # SEZIONE 1 — TREND VOLUMI
        # ════════════════════════════════════════════════════════════════════
        st.markdown('<div class="sec-h">📈 Trend Volumi Ordini</div>',
                    unsafe_allow_html=True)

        df_trend = calcola_trend_mensile(df_f)
        if not df_trend.empty:
            t1, t2, t3, t4 = st.columns(4)
            kpi_card(t1, "Righe nel periodo",
                     f"{int(df_f['Qta Doc'].count()):,}".replace(",","."),
                     "ordini elaborati")
            kpi_card(t2, "Quantità totale",
                     f"{int(df_f['Qta Doc'].sum()):,}".replace(",","."),
                     "paia ordinate", "g")
            kpi_card(t3, "Articoli distinti",
                     str(df_f['Articolo C'].nunique()),
                     "codici unici", "p")
            kpi_card(t4, "Valore ordini",
                     f"€ {df_f['Valore'].sum():,.0f}".replace(",","."),
                     "totale periodo", "o")

            fig_trend = make_subplots(specs=[[{"secondary_y": True}]])
            fig_trend.add_trace(
                go.Bar(x=df_trend['Mese'], y=df_trend['Qta_Totale'],
                       name='Quantità', marker_color='#bbdefb', opacity=0.8),
                secondary_y=False
            )
            fig_trend.add_trace(
                go.Scatter(x=df_trend['Mese'], y=df_trend['N_Clienti'],
                           name='N° Clienti', mode='lines+markers',
                           line=dict(color='#1f77b4', width=2),
                           marker=dict(size=6)),
                secondary_y=True
            )
            fig_trend.update_layout(
                height=280, margin=dict(t=10, b=40, l=0, r=0),
                plot_bgcolor='rgba(0,0,0,0)',
                legend=dict(orientation='h', y=-0.25)
            )
            fig_trend.update_yaxes(title_text="Quantità", secondary_y=False)
            fig_trend.update_yaxes(title_text="N° Clienti", secondary_y=True)
            st.plotly_chart(fig_trend, use_container_width=True)

        # ════════════════════════════════════════════════════════════════════
        # SEZIONE 2 — STORICITÀ E ROTAZIONE
        # ════════════════════════════════════════════════════════════════════
        st.markdown('<div class="sec-h">🔄 Storicità e Rotazione Articoli</div>',
                    unsafe_allow_html=True)

        df_rot = calcola_storicita_rotazione(df_f)

        if not df_rot.empty:
            tab_qta, tab_freq, tab_fam = st.tabs(
                ["📊 Top 20 per Quantità", "🔁 Top 20 per Frequenza", "📂 Per Famiglia"]
            )

            with tab_qta:
                top20 = df_rot.head(20)
                fig = px.bar(
                    top20, x='Articolo C', y='Qta_Totale', color='Famiglia',
                    text='Qta_Totale',
                    color_discrete_sequence=px.colors.qualitative.Set2,
                    labels={'Qta_Totale': 'Quantità', 'Articolo C': ''},
                )
                fig.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                fig.update_layout(height=340, margin=dict(t=10,b=60,l=0,r=0),
                                  xaxis_tickangle=-40, plot_bgcolor='rgba(0,0,0,0)',
                                  showlegend=True)
                st.plotly_chart(fig, use_container_width=True)

            with tab_freq:
                top20f = df_rot.sort_values('N_Ordini', ascending=False).head(20)
                fig2 = px.bar(
                    top20f, x='Articolo C', y='N_Ordini', color='Famiglia',
                    text='N_Ordini',
                    color_discrete_sequence=px.colors.qualitative.Pastel,
                    labels={'N_Ordini': 'N° ordini', 'Articolo C': ''},
                )
                fig2.update_traces(texttemplate='%{text}', textposition='outside')
                fig2.update_layout(height=340, margin=dict(t=10,b=60,l=0,r=0),
                                   xaxis_tickangle=-40, plot_bgcolor='rgba(0,0,0,0)',
                                   showlegend=True)
                st.plotly_chart(fig2, use_container_width=True)

            with tab_fam:
                df_fam = df_rot.groupby('Famiglia').agg(
                    Qta_Totale=('Qta_Totale','sum'),
                    N_Articoli=('Articolo C','count'),
                    N_Ordini=('N_Ordini','sum'),
                ).reset_index().sort_values('Qta_Totale', ascending=False)

                col_fa, col_fb = st.columns(2)
                with col_fa:
                    fig_fam = px.pie(
                        df_fam.head(10), values='Qta_Totale', names='Famiglia',
                        hole=0.4, title="Distribuzione quantità per famiglia",
                        color_discrete_sequence=px.colors.qualitative.Set3,
                    )
                    fig_fam.update_traces(textinfo='label+percent')
                    fig_fam.update_layout(height=320, margin=dict(t=40,b=0,l=0,r=0),
                                          showlegend=False)
                    st.plotly_chart(fig_fam, use_container_width=True)
                with col_fb:
                    st.dataframe(
                        df_fam.rename(columns={
                            'Qta_Totale':'Qta Totale',
                            'N_Articoli':'N° Articoli',
                            'N_Ordini':'N° Ordini'
                        }),
                        use_container_width=True, hide_index=True
                    )

            with st.expander("📋 Dettaglio completo rotazione"):
                show_rot = df_rot.copy()
                for c in ['Prima_Richiesta','Ultima_Richiesta']:
                    show_rot[c] = pd.to_datetime(show_rot[c],errors='coerce').dt.strftime('%d/%m/%Y')
                st.dataframe(show_rot, use_container_width=True, hide_index=True)

        # ════════════════════════════════════════════════════════════════════
        # SEZIONE 3 — FREQUENZA DI RIORDINO
        # ════════════════════════════════════════════════════════════════════
        st.markdown('<div class="sec-h">⏰ Frequenza di Riordino</div>',
                    unsafe_allow_html=True)

        df_riord = calcola_frequenza_riordino(df_f)

        if not df_riord.empty:
            df_con_int = df_riord.dropna(subset=['Intervallo_Medio_gg'])

            r1, r2, r3 = st.columns(3)
            kpi_card(r1, "Coppie Art.×Cliente con riordino",
                     str(len(df_con_int)), "≥ 2 ordini nel periodo", "g")
            if not df_con_int.empty:
                media_g = round(df_con_int['Intervallo_Medio_gg'].mean(), 1)
                art_top = df_con_int.sort_values('N_Ordini', ascending=False).iloc[0]
                kpi_card(r2, "Intervallo medio globale",
                         f"{media_g} gg", "tra riordini consecutivi")
                qta_m = int(art_top['Qta_Media_Ordine']) if 'Qta_Media_Ordine' in art_top else 0
                kpi_card(r3, "Articolo più riordinato",
                         art_top['Articolo C'],
                         f"{int(art_top['N_Ordini'])} ordini | ~{qta_m:,} pa./ordine — {art_top['Cliente'][:20]}".replace(",","."), "p")

            # Heatmap clienti × famiglie
            pivot = df_riord.pivot_table(
                index='Cliente', columns='Famiglia',
                values='N_Ordini', aggfunc='sum', fill_value=0
            )
            if pivot.shape[0] > 1 and pivot.shape[1] > 1:
                # Ordina clienti (righe) per totale decrescente — top 15
                top_cli = pivot.sum(axis=1).sort_values(ascending=False).head(15).index
                pivot = pivot.loc[top_cli]
                # Ordina famiglie (colonne) per totale decrescente
                top_fam = pivot.sum(axis=0).sort_values(ascending=False).index
                pivot = pivot[top_fam]
                fig_heat = px.imshow(
                    pivot,
                    labels=dict(x="Famiglia", y="Cliente", color="N° Ordini"),
                    color_continuous_scale="Blues", aspect="auto",
                    title="N° Ordini per Cliente × Famiglia (ordinati per volume)"
                )
                fig_heat.update_layout(height=420, margin=dict(t=40,b=0,l=0,r=0))
                fig_heat.update_xaxes(tickangle=-35)
                st.plotly_chart(fig_heat, use_container_width=True)

            # ── Calcola df_prossimi (usato sia per alert che per export) ──
            df_prossimi = pd.DataFrame()
            if not df_con_int.empty:
                df_alert = df_con_int.copy()
                df_alert['Giorni_Al_Riordino'] = (
                    df_alert['Intervallo_Medio_gg'] - df_alert['Giorni_Da_Ultimo']
                ).round(0)
                df_prossimi = df_alert[
                    df_alert['Giorni_Al_Riordino'].between(-7, 14)
                ].sort_values('Giorni_Al_Riordino')

            with st.expander("📋 Dettaglio frequenza riordino"):
                show_r = df_riord.copy()
                if 'Ultimo_Ordine' in show_r.columns:
                    show_r['Ultimo_Ordine'] = pd.to_datetime(
                        show_r['Ultimo_Ordine'], errors='coerce'
                    ).dt.strftime('%d/%m/%Y')
                st.dataframe(show_r, use_container_width=True, hide_index=True)


        # ════════════════════════════════════════════════════════════════════
        # SEZIONE 4 — SCOSTAMENTO DATE CONSEGNA
        # ════════════════════════════════════════════════════════════════════
        st.markdown('<div class="sec-h">📦 Puntualità Consegne (OCI vs DVF)</div>',
                    unsafe_allow_html=True)

        # Carica dettagli consegne dal file dedicato
        df_det_cons = carica_dettagli_consegne()
        if df_det_cons.empty:
            st.info(
                "File `dettagli_consegne.xlsx` non trovato. "
                "Esporta il report da ARCA e caricalo su GitHub."
            )
            df_cons, summary = pd.DataFrame(), {}
        else:
            df_cons, summary = calcola_scostamento_consegna(
                df_det_cons,
                filtro_cliente=filtro_cliente,
                filtro_articolo=filtro_articolo,
                filtro_famiglie=filtro_famiglie if filtro_famiglie else None,
                cutoff=cutoff if gg < 9999 else None,
            )

        if df_cons.empty:
            st.info(
                "Nessuna riga trovata nel file consegne per i filtri selezionati. "
                "Verifica che il file `dettagli_consegne.xlsx` contenga dati "
                "per il cliente/periodo selezionato."
            )
        else:
            p1, p2, p3, p4 = st.columns(4)
            kpi_card(p1, "Confronti OCI↔DVF",
                     f"{summary['n_evasi']:,}".replace(",","."),
                     "righe incrociate")
            col_p = "g" if summary['pct_puntuali'] >= 80 else "o"
            kpi_card(p2, "% Consegne puntuali",
                     f"{summary['pct_puntuali']}%",
                     f"{summary['pct_ritardo']}% in ritardo", col_p)
            col_m = "g" if summary['media_scost'] <= 2 else \
                    "o" if summary['media_scost'] <= 7 else "r"
            kpi_card(p3, "Ritardo medio (solo ritardi)",
                     f"{summary['media_scost']} gg",
                     "calcolato solo sulle consegne in ritardo", col_m)
            kpi_card(p4, "Mediana scostamento",
                     f"{summary['mediana_scost']} gg",
                     "valore centrale", "p")

            # Distribuzione scostamento
            fig_hist = px.histogram(
                df_cons, x='Scostamento_gg',
                color='Stato_Consegna', nbins=40,
                color_discrete_map={
                    '✅ Puntuale':    '#4caf50',
                    '⚠️ In ritardo': '#f44336',
                },
                labels={'Scostamento_gg': 'Scostamento (gg)', 'count': 'N° confronti'},
                title="Distribuzione scostamento data consegna prevista vs evasione DVF"
            )
            fig_hist.add_vline(x=0, line_dash="dash", line_color="#888",
                               annotation_text=" Puntuale")
            fig_hist.update_layout(height=300, margin=dict(t=40,b=0,l=0,r=0),
                                   plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_hist, use_container_width=True)

            # Trend mensile puntualità
            df_cons['Mese'] = df_cons['DataDoc'].dt.to_period('M').astype(str)
            df_tm = df_cons.groupby('Mese').agg(
                Pct_Puntuali=('Scostamento_gg',
                              lambda x: round((x <= 2).mean()*100, 1)),
                N_Evasi=('Scostamento_gg', 'count'),
                Scost_Medio=('Scostamento_gg', 'mean'),
            ).reset_index().sort_values('Mese')

            fig_tm = make_subplots(specs=[[{"secondary_y": True}]])
            fig_tm.add_trace(
                go.Bar(x=df_tm['Mese'], y=df_tm['N_Evasi'],
                       name='N° confronti', marker_color='#e3f2fd', opacity=0.8),
                secondary_y=False
            )
            fig_tm.add_trace(
                go.Scatter(x=df_tm['Mese'], y=df_tm['Pct_Puntuali'],
                           name='% Puntuali', mode='lines+markers',
                           line=dict(color='#4caf50', width=2),
                           marker=dict(size=7)),
                secondary_y=True
            )
            fig_tm.update_layout(
                title="Trend mensile puntualità",
                height=280, margin=dict(t=40,b=30,l=0,r=0),
                plot_bgcolor='rgba(0,0,0,0)',
                legend=dict(orientation='h', y=-0.3)
            )
            fig_tm.update_yaxes(title_text="N° confronti", secondary_y=False)
            fig_tm.update_yaxes(title_text="% Puntuali", secondary_y=True,
                                range=[0, 110], ticksuffix='%')
            st.plotly_chart(fig_tm, use_container_width=True)

            # Peggiori per ritardo
            with st.expander("📋 Top 30 ritardi maggiori"):
                df_rit = df_cons[df_cons['Scostamento_gg'] > 2].sort_values(
                    'Scostamento_gg', ascending=False
                ).head(30)
                cols = [c for c in ['Articolo C','Articolo D','Cliente',
                                    'Data Consegna','Data_Evasione',
                                    'Scostamento_gg','Qta Doc'] if c in df_rit.columns]
                st.dataframe(df_rit[cols], use_container_width=True, hide_index=True)

        # ── Export Excel ──────────────────────────────────────────────────────
        st.markdown("---")
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as w:
            if not df_rot.empty:
                df_rot.to_excel(w, sheet_name='Rotazione', index=False)
            if not df_riord.empty:
                df_riord.to_excel(w, sheet_name='Freq_Riordino', index=False)
            if not df_cons.empty:
                df_cons.to_excel(w, sheet_name='Puntualita_Consegne', index=False)
            if not df_trend.empty:
                df_trend.to_excel(w, sheet_name='Trend_Mensile', index=False)

        st.download_button(
            "📥 Esporta tutti i KPI in Excel",
            data=buf.getvalue(),
            file_name=f"KPI_Safit_{datetime.now().strftime('%d%m%Y')}.xlsx",
            use_container_width=True,
        )

        # ════════════════════════════════════════════════════════════════════
        # SEZIONE FINALE — RIORDINO ATTESO (ultima in assoluto)
        # ════════════════════════════════════════════════════════════════════
        if not df_prossimi.empty:
            st.markdown('<div class="sec-h">🔔 Articoli con riordino atteso entro 14 giorni</div>',
                        unsafe_allow_html=True)

            # Prepara dati export
            df_sol = df_prossimi.copy()
            df_sol['Ultimo_Ordine_fmt'] = pd.to_datetime(
                df_sol['Ultimo_Ordine'], errors='coerce'
            ).dt.strftime('%d/%m/%Y')
            df_sol['Riordino_Atteso'] = df_sol.apply(
                lambda r: f"tra {int(r['Giorni_Al_Riordino'])} gg"
                if r['Giorni_Al_Riordino'] >= 0
                else f"in ritardo di {abs(int(r['Giorni_Al_Riordino']))} gg",
                axis=1
            )
            df_export_sol = df_sol[[
                'Cliente', 'Famiglia', 'Articolo C', 'Articolo D',
                'N_Ordini', 'Qta_Totale', 'Qta_Media_Ordine', 'Intervallo_Medio_gg',
                'Ultimo_Ordine_fmt', 'Giorni_Da_Ultimo', 'Giorni_Al_Riordino', 'Riordino_Atteso'
            ]].sort_values(['Cliente', 'Giorni_Al_Riordino']).rename(columns={
                'Articolo C':           'Codice Articolo',
                'Articolo D':           'Descrizione',
                'N_Ordini':             'N° Ordini Storici',
                'Qta_Totale':           'Qta Totale Storica',
                'Qta_Media_Ordine':     'Qta Stimata Prossimo Ordine',
                'Intervallo_Medio_gg':  'Intervallo Medio (gg)',
                'Ultimo_Ordine_fmt':    'Ultimo Ordine',
                'Giorni_Da_Ultimo':     'Giorni Da Ultimo Ordine',
                'Giorni_Al_Riordino':   'Giorni Al Riordino',
                'Riordino_Atteso':      'Riordino Atteso',
            })

            buf_sol = BytesIO()
            with pd.ExcelWriter(buf_sol, engine='xlsxwriter') as w:
                df_export_sol.to_excel(w, sheet_name='Solleciti_Riordino', index=False)
                wb = w.book
                ws = w.sheets['Solleciti_Riordino']
                ws.set_column(0, 11, 24)
                fmt_h = wb.add_format({'bold': True, 'bg_color': '#1F4E79',
                                       'font_color': 'white', 'border': 1})
                for col_num, col_name in enumerate(df_export_sol.columns):
                    ws.write(0, col_num, col_name, fmt_h)
                fmt_red = wb.add_format({'bg_color': '#FFCCCC'})
                for row_num in range(1, len(df_export_sol) + 1):
                    gg_val = df_export_sol.iloc[row_num-1].get('Giorni Al Riordino', 0)
                    if pd.notnull(gg_val) and float(gg_val) < 0:
                        ws.set_row(row_num, None, fmt_red)

            # Pulsante download IN TESTA
            st.download_button(
                label=f"📥 Esporta Solleciti ({len(df_export_sol)} articoli) — raggruppati per Cliente/Famiglia",
                data=buf_sol.getvalue(),
                file_name=f"Solleciti_Riordino_{datetime.now().strftime('%d%m%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"btn_solleciti_{_key_suffix}"
            )

            # Elenco righe alert
            st.markdown(f"**{len(df_prossimi)} articoli con riordino atteso entro 14 giorni:**")
            for _, r in df_prossimi.iterrows():
                gg_r = int(r['Giorni_Al_Riordino'])
                ico  = "🟢" if gg_r > 0 else "🔴"
                txt  = f"tra **{gg_r} gg**" if gg_r >= 0 else f"**in ritardo di {abs(gg_r)} gg**"
                ult  = r['Ultimo_Ordine'].strftime('%d/%m/%Y') if pd.notnull(r['Ultimo_Ordine']) else 'N/D'
                qta_att = int(r['Qta_Media_Ordine']) if 'Qta_Media_Ordine' in r else 0
                st.markdown(
                    f'<div class="alert-box">{ico} <b>{r["Articolo C"]}</b> — '
                    f'{r["Cliente"]} | Atteso {txt} | '
                    f'Qta stimata: <b>~{qta_att:,} pa.</b> | '.replace(",",".")
                    + f'Ultimo ordine: {ult} | Intervallo medio: {r["Intervallo_Medio_gg"]} gg'
                    f'</div>',
                    unsafe_allow_html=True
                )


        # ════════════════════════════════════════════════════════════════════
        # SEZIONE — INDICATORI A CONTAGIRI PER FAMIGLIA
        # Lancetta = vendite periodo filtrato
        # 100% = media storica totale per quella famiglia
        # Arco rosso: 0–100%, arco verde: 100–200%
        # ════════════════════════════════════════════════════════════════════
        st.markdown('<div class="sec-h">🎯 Performance per Famiglia</div>',
                    unsafe_allow_html=True)
        st.caption(
            "La lancetta indica le vendite del periodo selezionato. "
            "100% = media storica della famiglia. "
            "Rosso: sotto media — Verde: sopra media."
        )

        # Calcola media storica per famiglia su tutto lo storico OCI
        df_storico_fam = df_oci.groupby('Famiglia')['Qta Doc'].sum().reset_index()
        # Numero di anni distinti nello storico per normalizzare
        # Media giornaliera storica (tutto lo storico OCI)
        gg_sto = max(1, (df_oci['Data'].max() - df_oci['Data'].min()).days)
        df_storico_fam['Media_Giorno_Sto'] = df_storico_fam['Qta Doc'] / gg_sto

        # Calcola durata del periodo filtrato
        if periodo == "Intervallo personalizzato":
            gg_periodo = max(1, (pd.Timestamp(data_a) - pd.Timestamp(data_da)).days)
        else:
            gg_periodo = gg if gg < 9999 else gg_sto
        gg_periodo = max(1, gg_periodo)

        # Media giornaliera periodo selezionato
        df_periodo_fam = df_f.groupby('Famiglia')['Qta Doc'].sum().reset_index()
        df_periodo_fam.columns = ['Famiglia', 'Qta_Periodo']
        df_periodo_fam['Media_Giorno_Per'] = df_periodo_fam['Qta_Periodo'] / gg_periodo

        # Unisci e calcola scostamento %
        df_gauge = df_periodo_fam.merge(
            df_storico_fam[['Famiglia', 'Media_Giorno_Sto']], on='Famiglia', how='left'
        )
        df_gauge = df_gauge[df_gauge['Qta_Periodo'] > 0].copy()
        # Scostamento: 0=media, -100=zero, +100=doppio
        df_gauge['Pct'] = ((df_gauge['Media_Giorno_Per'] / df_gauge['Media_Giorno_Sto'] - 1) * 100).clip(-100, 100)
        # Per etichette: media attesa nel periodo
        df_gauge['Media_Periodo'] = (df_gauge['Media_Giorno_Sto'] * gg_periodo).round(0).astype(int)
        df_gauge = df_gauge.sort_values('Qta_Periodo', ascending=False)

        if df_gauge.empty:
            st.info("Nessuna famiglia con vendite nel periodo selezionato.")
        else:
            # Griglia: max 4 per riga
            N_COL = 4
            famiglie_gauge = df_gauge['Famiglia'].tolist()
            righe = [famiglie_gauge[i:i+N_COL] for i in range(0, len(famiglie_gauge), N_COL)]

            import math as _math

            def _disegna_gauge(pct, qta, media, fam):
                """
                Contagiri scala -100/0/+100.
                0 = media storica giornaliera (linea gialla).
                Settori a step 25pt con sfumature rosso->verde.
                """
                import math as _m
                def _pta(p):
                    return _m.radians(90 - (p / 100) * 90)
                N = 40
                RE, RI = 1.0, 0.55

                def arco_xy(s, e, re, ri):
                    xs, ys = [], []
                    for i in range(N+1):
                        a = _pta(s + (e-s)*i/N)
                        xs.append(re*_m.cos(a)); ys.append(re*_m.sin(a))
                    for i in range(N+1):
                        a = _pta(e - (e-s)*i/N)
                        xs.append(ri*_m.cos(a)); ys.append(ri*_m.sin(a))
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
                    sx, sy = arco_xy(s, e, RE, RI)
                    fig.add_trace(go.Scatter(x=sx, y=sy, fill='toself',
                        fillcolor=col, line=dict(color=col, width=0),
                        showlegend=False, hoverinfo='skip', mode='lines'))

                # Linea 0% gialla
                a0 = _pta(0)
                fig.add_trace(go.Scatter(
                    x=[RI*_m.cos(a0), RE*_m.cos(a0)],
                    y=[RI*_m.sin(a0), RE*_m.sin(a0)],
                    mode='lines', line=dict(color='#FFD700', width=3),
                    showlegend=False, hoverinfo='skip'))

                # Lancetta
                al = _pta(pct)
                fig.add_trace(go.Scatter(
                    x=[0, 0.90*_m.cos(al)], y=[0, 0.90*_m.sin(al)],
                    mode='lines', line=dict(color='#111111', width=3),
                    showlegend=False, hoverinfo='skip'))
                tc = [_m.radians(i) for i in range(361)]
                fig.add_trace(go.Scatter(
                    x=[0.07*_m.cos(t) for t in tc],
                    y=[0.07*_m.sin(t) for t in tc],
                    fill='toself', fillcolor='#111111',
                    line=dict(color='#111111', width=0),
                    showlegend=False, hoverinfo='skip', mode='lines'))

                # Tick
                for tp, lb in [(-100,'-100%'),(-50,'-50%'),(0,'0'),(50,'+50%'),(100,'+100%')]:
                    at = _pta(tp)
                    fig.add_trace(go.Scatter(
                        x=[RE*_m.cos(at), (RE+0.08)*_m.cos(at)],
                        y=[RE*_m.sin(at), (RE+0.08)*_m.sin(at)],
                        mode='lines', line=dict(color='#8B949E', width=1),
                        showlegend=False, hoverinfo='skip'))
                    fig.add_annotation(
                        x=(RE+0.26)*_m.cos(at), y=(RE+0.26)*_m.sin(at),
                        text=lb, showarrow=False,
                        font=dict(size=8, color='#8B949E'),
                        xanchor='center', yanchor='middle')

                col_pct = '#1B5E20' if pct >= 0 else '#B71C1C'
                segno = '+' if pct >= 0 else ''
                fig.add_annotation(x=0, y=-0.25,
                    text=f"<b>{segno}{pct:.0f}%</b>",
                    showarrow=False, font=dict(size=16, color=col_pct),
                    xanchor='center', yanchor='top')
                fig.add_annotation(x=0, y=-0.45,
                    text=f"<b>{fam}</b>",
                    showarrow=False, font=dict(size=11, color='#C9D1D9'),
                    xanchor='center', yanchor='top')
                fig.add_annotation(x=0, y=-0.60,
                    text=f"{qta:,} pa. | med: {media:,} pa.".replace(",","."),
                    showarrow=False, font=dict(size=9, color='#8B949E'),
                    xanchor='center', yanchor='top')

                fig.update_layout(
                    height=260,
                    margin=dict(t=10, b=10, l=10, r=10),
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    xaxis=dict(visible=False, range=[-1.55, 1.55], scaleanchor='y'),
                    yaxis=dict(visible=False, range=[-0.85, 1.35]),
                    showlegend=False,
                )
                return fig


            for riga in righe:
                cols = st.columns(N_COL)
                for j, fam in enumerate(riga):
                    row = df_gauge[df_gauge['Famiglia'] == fam].iloc[0]
                    pct   = float(row['Pct'])
                    qta   = int(row['Qta_Periodo'])
                    media = int(row['Media_Periodo'])
                    fig = _disegna_gauge(pct, qta, media, fam)
                    with cols[j]:
                        st.plotly_chart(fig, use_container_width=True,
                                        config={'displayModeBar': False},
                                        key=f"gauge_{fam}_{_key_suffix}")

                for j in range(len(riga), N_COL):
                    cols[j].empty()
