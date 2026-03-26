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
from datetime import datetime, timedelta
from io import BytesIO
import os

# ── Percorso file storico ARCA ────────────────────────────────────────────────
PATH_STORICO = "righe_ordini_storico_con_date.xlsx"

# ── CSS ───────────────────────────────────────────────────────────────────────
KPI_CSS = """
<style>
.kpi-adv {
    background: linear-gradient(135deg,#ffffff,#f8faff);
    border:1px solid #e3e8f0; border-left:4px solid #1f77b4;
    padding:14px 16px; border-radius:10px; margin-bottom:6px;
    box-shadow:0 2px 6px rgba(0,0,0,0.05);
}
.kpi-adv.g { border-left-color:#4caf50; }
.kpi-adv.o { border-left-color:#ff9800; }
.kpi-adv.r { border-left-color:#f44336; }
.kpi-adv.p { border-left-color:#9c27b0; }
.kpi-adv-t { font-size:10px; color:#999; text-transform:uppercase;
             letter-spacing:.5px; margin-bottom:3px; }
.kpi-adv-v { font-size:26px; font-weight:700; color:#1a1a2e; line-height:1; }
.kpi-adv-s { font-size:11px; color:#777; margin-top:3px; }
.sec-h { font-size:16px; font-weight:700; color:#1a1a2e;
         border-bottom:2px solid #e3e8f0; padding-bottom:5px;
         margin:20px 0 12px 0; }
.alert-box { background:#fff8e1; border-left:4px solid #ff9800;
             padding:10px 14px; border-radius:6px; margin:6px 0;
             font-size:13px; }
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

        # Propaga Data Consegna all'interno dello stesso documento
        df['Data Consegna'] = df.groupby('Numero Documento')['Data Consegna'].ffill()

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


def get_oci_oca(df):
    """Filtra solo OCI e OCA con quantità > 0."""
    return df[
        df['Codice Documento'].isin(['OCI', 'OCA']) &
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
        risultati.append({
            'Articolo C':          art,
            'Articolo D':          g['Articolo D'].iloc[0],
            'Famiglia':            g['Famiglia'].iloc[0],
            'Cliente':             cli,
            'Cod_Cliente':         g['Cod_Cliente'].iloc[0],
            'N_Ordini':            len(date_ordini),
            'Qta_Totale':          g['Qta Doc'].sum(),
            'Intervallo_Medio_gg': intervallo,
            'Ultimo_Ordine':       ultimo,
            'Giorni_Da_Ultimo':    (datetime.now() - ultimo).days if pd.notnull(ultimo) else None,
        })

    return pd.DataFrame(risultati).sort_values('N_Ordini', ascending=False)


def calcola_scostamento_consegna(df_ordini, df_dvf):
    """
    Scostamento data consegna prevista (OCI) vs data evasione reale (DVF).
    Incrocia su Articolo C + Cliente.
    """
    if df_ordini.empty or df_dvf.empty:
        return pd.DataFrame(), {}

    # Per ogni articolo+cliente: ultima data DVF come proxy "data evasione"
    df_dvf_grp = df_dvf.groupby(['Articolo C', 'Cliente'])['Data'].max().reset_index()
    df_dvf_grp.columns = ['Articolo C', 'Cliente', 'Data_Evasione']

    df_join = df_ordini.merge(df_dvf_grp, on=['Articolo C', 'Cliente'], how='inner')
    df_join = df_join.dropna(subset=['Data Consegna', 'Data_Evasione'])

    if df_join.empty:
        return pd.DataFrame(), {}

    df_join['Scostamento_gg'] = (
        df_join['Data_Evasione'] - df_join['Data Consegna']
    ).dt.days

    df_join['Stato_Consegna'] = df_join['Scostamento_gg'].apply(
        lambda x: '✅ Puntuale'   if -2 <= x <= 2
        else ('⚡ Anticipato'     if x < -2
        else  '⚠️ In ritardo')
    )

    summary = {
        'n_evasi':           len(df_join),
        'media_scost':       round(df_join['Scostamento_gg'].mean(), 1),
        'mediana_scost':     round(df_join['Scostamento_gg'].median(), 1),
        'pct_puntuali':      round((df_join['Scostamento_gg'].between(-2, 2)).mean() * 100, 1),
        'pct_ritardo':       round((df_join['Scostamento_gg'] > 2).mean() * 100, 1),
        'pct_anticipati':    round((df_join['Scostamento_gg'] < -2).mean() * 100, 1),
    }
    return df_join, summary


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

def render_kpi_avanzati(path_storico=PATH_STORICO, filtro_cliente=None):
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
        # Il valore dalla sidebar ha formato "C000744 - SPIRALE SRL"
        # Il campo Cliente nel file storico ha solo "SPIRALE SRL"
        # → estraiamo solo la parte dopo " - " se presente
        if filtro_cliente:
            nome_cli = filtro_cliente.split(' - ', 1)[-1].strip() if ' - ' in filtro_cliente else filtro_cliente.strip()
            df_oci = df_oci[df_oci['Cliente'].str.contains(nome_cli, case=False, na=False, regex=False)]
            df_dvf = df_dvf[df_dvf['Cliente'].str.contains(nome_cli, case=False, na=False, regex=False)]

        st.caption(
            f"📁 Storico caricato: **{len(df_all):,}** righe totali | "
            f"**{len(df_oci):,}** OCI/OCA | **{len(df_dvf):,}** DVF".replace(",",".")
        )

        # ── Filtri ────────────────────────────────────────────────────────────
        # Il filtro cliente viene dalla sidebar del portale (filtro_cliente).
        # Qui gestiamo solo periodo e famiglia.
        # Key univoca per evitare conflitti quando si cambia cliente dalla sidebar
        _key_suffix = str(filtro_cliente or "tutti").lower().replace(" ", "_")[:20]

        c1, c2 = st.columns([1, 2])
        with c1:
            periodo = st.selectbox(
                "Periodo", ["Ultimi 90 gg", "Ultimi 6 mesi",
                            "Ultimo anno", "Tutto lo storico"],
                index=2, key=f"kpi_periodo_{_key_suffix}"
            )
        with c2:
            famiglie = sorted(df_oci['Famiglia'].dropna().unique().tolist())
            sel_fam  = st.multiselect("Famiglia", famiglie,
                                       key=f"kpi_fam_{_key_suffix}",
                                       placeholder="Tutte le famiglie")

        if filtro_cliente:
            st.caption(f"👤 Dati filtrati per cliente: **{filtro_cliente}**")

        # Applica filtri temporali
        gg_map = {"Ultimi 90 gg": 90, "Ultimi 6 mesi": 180,
                  "Ultimo anno": 365, "Tutto lo storico": 9999}
        gg = gg_map[periodo]
        cutoff = datetime.now() - timedelta(days=gg)

        df_f = df_oci.copy()
        if gg < 9999:
            df_f = df_f[df_f['Data'] >= cutoff]
        if sel_fam:
            df_f = df_f[df_f['Famiglia'].isin(sel_fam)]

        df_dvf_f = df_dvf.copy()
        if gg < 9999:
            df_dvf_f = df_dvf_f[df_dvf_f['Data'] >= cutoff]

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
                     "pezzi ordinati", "g")
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
                kpi_card(r3, "Articolo più riordinato",
                         art_top['Articolo C'],
                         f"{int(art_top['N_Ordini'])} ordini — {art_top['Cliente'][:25]}", "p")

            # Heatmap clienti × famiglie
            pivot = df_riord.pivot_table(
                index='Cliente', columns='Famiglia',
                values='N_Ordini', aggfunc='sum', fill_value=0
            )
            if pivot.shape[0] > 1 and pivot.shape[1] > 1:
                top_cli = pivot.sum(axis=1).sort_values(ascending=False).head(15).index
                pivot = pivot.loc[top_cli]
                fig_heat = px.imshow(
                    pivot,
                    labels=dict(x="Famiglia", y="Cliente", color="N° Ordini"),
                    color_continuous_scale="Blues", aspect="auto",
                    title="N° Ordini per Cliente × Famiglia"
                )
                fig_heat.update_layout(height=420, margin=dict(t=40,b=0,l=0,r=0))
                st.plotly_chart(fig_heat, use_container_width=True)

            # Alert: articoli vicini al prossimo riordino atteso (±14 gg)
            if not df_con_int.empty:
                df_alert = df_con_int.copy()
                df_alert['Giorni_Al_Riordino'] = (
                    df_alert['Intervallo_Medio_gg'] - df_alert['Giorni_Da_Ultimo']
                ).round(0)
                df_prossimi = df_alert[
                    df_alert['Giorni_Al_Riordino'].between(-7, 14)
                ].sort_values('Giorni_Al_Riordino')

                if not df_prossimi.empty:
                    st.markdown(f"**🔔 {len(df_prossimi)} articoli con riordino atteso entro 14 giorni:**")
                    for _, r in df_prossimi.iterrows():
                        gg_r = int(r['Giorni_Al_Riordino'])
                        ico  = "🟢" if gg_r > 0 else "🔴"
                        txt  = f"tra **{gg_r} gg**" if gg_r >= 0 else f"**in ritardo di {abs(gg_r)} gg**"
                        ult  = r['Ultimo_Ordine'].strftime('%d/%m/%Y') if pd.notnull(r['Ultimo_Ordine']) else 'N/D'
                        st.markdown(
                            f'<div class="alert-box">{ico} <b>{r["Articolo C"]}</b> — '
                            f'{r["Cliente"]} | Atteso {txt} | '
                            f'Ultimo ordine: {ult} | Intervallo medio: {r["Intervallo_Medio_gg"]} gg'
                            f'</div>',
                            unsafe_allow_html=True
                        )

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

        df_cons, summary = calcola_scostamento_consegna(df_f, df_dvf_f)

        if df_cons.empty:
            st.info(
                "Nessun incrocio OCI↔DVF trovato nel periodo selezionato. "
                "Allarga il periodo o verifica che il file contenga sia OCI che DVF "
                "per gli stessi articoli e clienti."
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
            kpi_card(p3, "Scostamento medio",
                     f"{summary['media_scost']} gg",
                     "positivo = ritardo, negativo = anticipo", col_m)
            kpi_card(p4, "Mediana scostamento",
                     f"{summary['mediana_scost']} gg",
                     "valore centrale", "p")

            # Distribuzione scostamento
            fig_hist = px.histogram(
                df_cons, x='Scostamento_gg',
                color='Stato_Consegna', nbins=40,
                color_discrete_map={
                    '✅ Puntuale':    '#4caf50',
                    '⚡ Anticipato': '#2196f3',
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
            df_cons['Mese'] = df_cons['Data'].dt.to_period('M').astype(str)
            df_tm = df_cons.groupby('Mese').agg(
                Pct_Puntuali=('Stato_Consegna',
                              lambda x: round((x == '✅ Puntuale').mean()*100, 1)),
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
