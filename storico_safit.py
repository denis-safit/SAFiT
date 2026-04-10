"""
storico_safit.py  v2.0
======================
Modulo KPI Storici per il Safit Portal.
Integra i dati SAFIT (pre-3/9/2025) con i dati SafitIB (correnti)
e li presenta con lo stesso layout di kpi_avanzati.py.

Path configurabili dall'esterno (Safit_Portal_Local.py li sovrascrive):
    PATH_STORICO_SAFIT  -> righe ordini storici SAFIT
    PATH_CORRENTE_IB    -> righe ordini correnti SafitIB
    PATH_TRANSCODIFICA  -> tabella transcodifica codici articolo
"""

import os
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path
from datetime import datetime, timedelta, date

# ─────────────────────────────────────────────────────────────
# CONFIGURAZIONE PATH (sovrascrivibili da Safit_Portal_Local.py)
# ─────────────────────────────────────────────────────────────
PATH_STORICO_SAFIT = Path("righe_ordini_storico_con_date_SAFIT.xlsx")
PATH_CORRENTE_IB   = Path("righe_ordini_storico_con_date.xlsx")
PATH_TRANSCODIFICA = Path("transcodifica.xlsx")

DOC_REALI   = {'OCA', 'OCI', 'OFF', 'OFR', 'OFI', 'OFA'}
DATA_CAMBIO = pd.Timestamp("2025-09-03")

# ─────────────────────────────────────────────────────────────
# CSS identico a kpi_avanzati
# ─────────────────────────────────────────────────────────────
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

COLOR_SAFIT   = "#1f77b4"
COLOR_SAFITIB = "#ff7f0e"
COLOR_IBRIDO  = "#9467bd"
PALETTE       = {"SAFIT": COLOR_SAFIT, "SAFITIB": COLOR_SAFITIB, "IBRIDO": COLOR_IBRIDO}


# ─────────────────────────────────────────────────────────────
# CARICAMENTO E NORMALIZZAZIONE
# ─────────────────────────────────────────────────────────────

def _ffill_doc(df):
    for col in ['Codice Documento', 'Data Consegna', 'Numero Documento']:
        if col in df.columns:
            df[col] = df[col].ffill()
    return df


def _normalizza(df):
    df = _ffill_doc(df.copy())
    df.columns = [str(c).strip() for c in df.columns]
    df['Data']          = pd.to_datetime(df.get('Data'),          errors='coerce')
    df['Data Consegna'] = pd.to_datetime(df.get('Data Consegna'), errors='coerce')
    df['Qta Doc']       = pd.to_numeric(df.get('Qta Doc'),     errors='coerce').fillna(0)
    df['Valore']        = pd.to_numeric(df.get('Valore'),      errors='coerce').fillna(0)
    df['Qta Residua']   = pd.to_numeric(df.get('Qta Residua'), errors='coerce').fillna(0)
    df = df[df['Codice Documento'].isin(DOC_REALI)].copy()
    df = df[df['Articolo C'].notna() & (df['Articolo C'].astype(str) != '(vuoto)')].copy()
    df['Anno']    = df['Data'].dt.year
    df['Mese']    = df['Data'].dt.to_period('M').astype(str)
    df['Cliente'] = df['Cliente Fornitore CD'].astype(str).str.split(' - ', n=1).str[-1].str.strip()
    df['Famiglia'] = df['Articolo D'].apply(lambda x: ' '.join(str(x).split()[:2]).upper())
    return df


def _leggi_excel_autoheader(path):
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]
    if 'Data' not in df.columns:
        df = pd.read_excel(path, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
    return df


@st.cache_data(show_spinner=False)
def _carica_transcodifica():
    df = pd.read_excel(PATH_TRANSCODIFICA)
    return dict(zip(df['Cd_AR_Vecchio'].astype(str), df['Cd_AR_Nuovo'].astype(str)))


@st.cache_data(show_spinner="Caricamento dataset storico unificato...", ttl=1800)
def carica_dataset_unificato():
    trans   = _carica_transcodifica()
    df_safit = _normalizza(pd.read_excel(PATH_STORICO_SAFIT))
    df_safit['_sorgente']     = 'SAFIT'
    df_safit['Articolo_C_IB'] = df_safit['Articolo C'].astype(str).map(trans)
    df_ib = _normalizza(_leggi_excel_autoheader(PATH_CORRENTE_IB))
    df_ib['_sorgente']     = 'SAFITIB'
    df_ib['Articolo_C_IB'] = df_ib['Articolo C'].astype(str)
    return pd.concat([df_safit, df_ib], ignore_index=True)


# ─────────────────────────────────────────────────────────────
# FUNZIONI KPI PUBBLICHE
# ─────────────────────────────────────────────────────────────

def kpi_annuale(df=None):
    if df is None:
        df = carica_dataset_unificato()
    grp = df.groupby('Anno').agg(
        Valore_Tot=('Valore',           'sum'),
        Qta_Tot   =('Qta Doc',          'sum'),
        N_Ordini  =('Numero Documento', 'nunique'),
        N_Righe   =('Articolo C',       'count'),
    ).reset_index()
    grp['_sorgente'] = grp['Anno'].apply(lambda a: 'SAFITIB' if a > DATA_CAMBIO.year else 'SAFIT')
    grp.loc[grp['Anno'] == DATA_CAMBIO.year, '_sorgente'] = 'IBRIDO'
    return grp.sort_values('Anno')


def kpi_per_tipo_doc(df=None):
    if df is None:
        df = carica_dataset_unificato()
    return df.groupby(['Anno', 'Codice Documento']).agg(
        Valore_Tot=('Valore',  'sum'),
        Qta_Tot   =('Qta Doc', 'sum'),
        N_Ordini  =('Numero Documento', 'nunique'),
    ).reset_index().sort_values(['Anno', 'Codice Documento'])


def articoli_non_mappati(df=None):
    if df is None:
        df = carica_dataset_unificato()
    mask = (df['_sorgente'] == 'SAFIT') & (df['Articolo_C_IB'].isna())
    return (
        df[mask][['Articolo C', 'Articolo D', 'Anno', 'Valore', 'Qta Doc']]
        .groupby(['Articolo C', 'Articolo D']).agg(
            Anni      =('Anno',    lambda x: sorted(x.unique().tolist())),
            Valore_Tot=('Valore',  'sum'),
            Qta_Tot   =('Qta Doc', 'sum'),
        ).reset_index().sort_values('Valore_Tot', ascending=False)
    )


# ─────────────────────────────────────────────────────────────
# CALCOLI INTERNI
# ─────────────────────────────────────────────────────────────

def _calcola_trend_mensile(df):
    if df.empty:
        return pd.DataFrame()
    d = df.copy()
    d['Mese'] = d['Data'].dt.to_period('M').astype(str)
    return d.groupby('Mese').agg(
        N_Righe   =('Articolo C', 'count'),
        Qta_Totale=('Qta Doc',    'sum'),
        Valore_Tot=('Valore',     'sum'),
        N_Clienti =('Cliente',    'nunique'),
        N_Articoli=('Articolo C', 'nunique'),
    ).reset_index().sort_values('Mese')


def _calcola_storicita_rotazione(df):
    if df.empty:
        return pd.DataFrame()
    grp = df.groupby(['Articolo C', 'Articolo D', 'Famiglia', 'Articolo_C_IB']).agg(
        N_Ordini        =('Data',    'count'),
        Qta_Totale      =('Qta Doc', 'sum'),
        Valore_Totale   =('Valore',  'sum'),
        Prima_Richiesta =('Data',    'min'),
        Ultima_Richiesta=('Data',    'max'),
        N_Clienti       =('Cliente', 'nunique'),
    ).reset_index()
    grp['Giorni_Attivita']  = (grp['Ultima_Richiesta'] - grp['Prima_Richiesta']).dt.days.clip(lower=1)
    grp['Freq_Settimanale'] = (grp['N_Ordini'] / (grp['Giorni_Attivita'] / 7)).round(2)
    return grp.sort_values('Qta_Totale', ascending=False).reset_index(drop=True)


def _calcola_frequenza_riordino(df):
    if df.empty:
        return pd.DataFrame()
    risultati = []
    for (art, cli), g in df.groupby(['Articolo C', 'Cliente']):
        date_ord = sorted(g['Data'].dropna().tolist())
        intervallo = None
        if len(date_ord) >= 2:
            delta = [(date_ord[i+1] - date_ord[i]).days for i in range(len(date_ord)-1)]
            intervallo = round(float(np.mean(delta)), 1)
        ultimo    = max(date_ord) if date_ord else pd.NaT
        qta_media = round(g['Qta Doc'].mean(), 0) if len(g) > 0 else 0
        risultati.append({
            'Articolo C':          art,
            'Articolo D':          g['Articolo D'].iloc[0],
            'Famiglia':            g['Famiglia'].iloc[0],
            'Cliente':             cli,
            'N_Ordini':            len(date_ord),
            'Qta_Totale':          g['Qta Doc'].sum(),
            'Qta_Media_Ordine':    qta_media,
            'Intervallo_Medio_gg': intervallo,
            'Ultimo_Ordine':       ultimo,
            'Giorni_Da_Ultimo':    (datetime.now() - ultimo).days if pd.notnull(ultimo) else None,
        })
    return pd.DataFrame(risultati).sort_values('N_Ordini', ascending=False)


def _calcola_trend_annuale(df):
    if df.empty:
        return pd.DataFrame()
    return df.groupby(['Anno', '_sorgente']).agg(
        Qta_Totale=('Qta Doc', 'sum'),
        Valore_Tot=('Valore',  'sum'),
        N_Clienti =('Cliente', 'nunique'),
        N_Articoli=('Articolo C', 'nunique'),
    ).reset_index().sort_values('Anno')


# ─────────────────────────────────────────────────────────────
# HELPER UI
# ─────────────────────────────────────────────────────────────

def _kpi_card(col, title, value, sub="", color=""):
    col.markdown(
        f'<div class="kpi-adv {color}">'
        f'<div class="kpi-adv-t">{title}</div>'
        f'<div class="kpi-adv-v">{value}</div>'
        f'<div class="kpi-adv-s">{sub}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────
# RENDER PRINCIPALE
# ─────────────────────────────────────────────────────────────

def render_kpi_storici(filtro_cliente=None, filtro_articolo=None, filtro_famiglie=None):
    """
    Punto di ingresso — chiamato da Safit_Portal_Local.py nel tab KPI Storici.
    Layout identico a render_kpi_avanzati ma su dataset 2011->oggi.
    """
    st.markdown(KPI_CSS, unsafe_allow_html=True)

    with st.expander(
        "📊 KPI Storici — Trend · Storicita · Rotazione · Riordino (2011 -> oggi)",
        expanded=True,
    ):
        # ── Caricamento ──────────────────────────────────────────────────
        try:
            df_all = carica_dataset_unificato()
        except FileNotFoundError as e:
            st.error(f"File non trovato: `{e.filename}`. Verificare i file nella cartella del portale.")
            return

        if df_all.empty:
            st.warning("Dataset vuoto.")
            return

        df_oci = df_all[df_all['Codice Documento'].isin(['OCI', 'OCA']) & (df_all['Qta Doc'] > 0)].copy()

        st.caption(
            f"Dataset unificato: **{len(df_all):,}** righe totali | "
            f"**{len(df_oci):,}** OCI/OCA | "
            "🔵 SAFIT (pre 3/9/2025) + 🟠 SafitIB (post 3/9/2025)".replace(",", ".")
        )

        # ── Filtri ───────────────────────────────────────────────────────
        c1, c2 = st.columns([1, 2])
        with c1:
            periodo = st.selectbox(
                "Periodo",
                ["Ultimi 90 gg", "Ultimi 6 mesi", "Ultimo anno",
                 "Ultimi 5 anni", "Tutto lo storico", "Intervallo personalizzato"],
                index=4, key="ks_periodo_stor"
            )
        with c2:
            famiglie_disp = sorted(df_oci['Famiglia'].dropna().unique().tolist())
            sel_fam = st.multiselect("Famiglia", famiglie_disp,
                                     key="ks_fam_stor",
                                     placeholder="Tutte le famiglie")

        if periodo == "Intervallo personalizzato":
            data_min_d = df_oci["Data"].min().date() if not df_oci.empty else date(2011, 1, 1)
            data_max_d = df_oci["Data"].max().date() if not df_oci.empty else datetime.now().date()
            dc1, dc2 = st.columns(2)
            with dc1:
                data_da = st.date_input("Dal", value=data_min_d,
                                        min_value=data_min_d, max_value=data_max_d,
                                        key="ks_da_stor", format="DD/MM/YYYY")
            with dc2:
                data_a = st.date_input("Al", value=data_max_d,
                                       min_value=data_min_d, max_value=data_max_d,
                                       key="ks_a_stor", format="DD/MM/YYYY")

        # Filtri opzionali passati dalla sidebar
        if filtro_cliente:
            nome_cli = filtro_cliente.split(' - ', 1)[-1].strip() if ' - ' in filtro_cliente else filtro_cliente.strip()
            df_oci = df_oci[df_oci['Cliente'].str.contains(nome_cli, case=False, na=False, regex=False)]
            st.caption(f"👤 Cliente: **{filtro_cliente}**")
        if filtro_articolo:
            df_oci = df_oci[df_oci['Articolo C'].str.contains(filtro_articolo.upper(), case=False, na=False)]
            st.caption(f"🔍 Articolo: **{filtro_articolo.upper()}**")
        if filtro_famiglie:
            df_oci = df_oci[df_oci['Famiglia'].isin(filtro_famiglie)]

        # Filtro temporale
        gg_map = {"Ultimi 90 gg": 90, "Ultimi 6 mesi": 180,
                  "Ultimo anno": 365, "Ultimi 5 anni": 1825, "Tutto lo storico": 99999}
        df_f = df_oci.copy()
        if periodo == "Intervallo personalizzato":
            df_f = df_f[(df_f["Data"] >= pd.Timestamp(data_da)) & (df_f["Data"] <= pd.Timestamp(data_a))]
        elif periodo != "Tutto lo storico":
            cutoff = pd.Timestamp(datetime.now() - timedelta(days=gg_map[periodo]))
            df_f = df_f[df_f["Data"] >= cutoff]
        if sel_fam:
            df_f = df_f[df_f["Famiglia"].isin(sel_fam)]

        if df_f.empty:
            st.warning("Nessun dato per i filtri selezionati.")
            return

        st.caption(
            f"Filtro attivo: **{len(df_f):,}** righe | "
            f"**{df_f['Articolo C'].nunique()}** articoli | "
            f"**{df_f['Cliente'].nunique()}** clienti".replace(",", ".")
        )

        # ════════════════════════════════════════════════════════════
        # SEZIONE 1 — TREND VOLUMI
        # ════════════════════════════════════════════════════════════
        st.markdown('<div class="sec-h">📈 Trend Volumi Ordini</div>', unsafe_allow_html=True)

        df_trend = _calcola_trend_mensile(df_f)
        if not df_trend.empty:
            t1, t2, t3, t4 = st.columns(4)
            _kpi_card(t1, "Righe nel periodo",
                      f"{int(df_f['Qta Doc'].count()):,}".replace(",", "."), "ordini elaborati")
            _kpi_card(t2, "Quantita totale",
                      f"{int(df_f['Qta Doc'].sum()):,}".replace(",", "."), "paia ordinate", "g")
            _kpi_card(t3, "Articoli distinti",
                      str(df_f['Articolo C'].nunique()), "codici unici", "p")
            _kpi_card(t4, "Valore ordini",
                      f"EUR {df_f['Valore'].sum():,.0f}".replace(",", "."), "totale periodo", "o")

            tab_mens, tab_ann = st.tabs(["Mensile", "Annuale"])

            with tab_mens:
                fig_m = make_subplots(specs=[[{"secondary_y": True}]])
                fig_m.add_trace(
                    go.Bar(x=df_trend['Mese'], y=df_trend['Qta_Totale'],
                           name='Quantita', marker_color='#bbdefb', opacity=0.8),
                    secondary_y=False)
                fig_m.add_trace(
                    go.Scatter(x=df_trend['Mese'], y=df_trend['N_Clienti'],
                               name='N Clienti', mode='lines+markers',
                               line=dict(color='#1f77b4', width=2), marker=dict(size=6)),
                    secondary_y=True)
                fig_m.update_layout(height=280, margin=dict(t=10, b=40, l=0, r=0),
                                    plot_bgcolor='rgba(0,0,0,0)',
                                    legend=dict(orientation='h', y=-0.25))
                fig_m.update_yaxes(title_text="Quantita (pa.)", secondary_y=False)
                fig_m.update_yaxes(title_text="N Clienti", secondary_y=True)
                st.plotly_chart(fig_m, use_container_width=True)

            with tab_ann:
                df_ann = _calcola_trend_annuale(df_oci)
                if not df_ann.empty:
                    fig_ann = make_subplots(specs=[[{"secondary_y": True}]])
                    for sorg, color in PALETTE.items():
                        d = df_ann[df_ann['_sorgente'] == sorg]
                        if not d.empty:
                            fig_ann.add_trace(
                                go.Bar(x=d['Anno'], y=d['Qta_Totale'],
                                       name=sorg, marker_color=color, opacity=0.85),
                                secondary_y=False)
                    ann_cli = df_ann.groupby('Anno')['N_Clienti'].sum()
                    fig_ann.add_trace(
                        go.Scatter(x=ann_cli.index, y=ann_cli.values,
                                   name='N Clienti', mode='lines+markers',
                                   line=dict(color='#333', width=2, dash='dot'),
                                   marker=dict(size=6)),
                        secondary_y=True)
                    fig_ann.update_layout(barmode='stack', height=300,
                                          margin=dict(t=10, b=40, l=0, r=0),
                                          plot_bgcolor='rgba(0,0,0,0)',
                                          legend=dict(orientation='h', y=-0.25))
                    fig_ann.update_yaxes(title_text="Quantita (pa.)", secondary_y=False)
                    fig_ann.update_yaxes(title_text="N Clienti", secondary_y=True)
                    st.plotly_chart(fig_ann, use_container_width=True)

                    df_val = df_oci.groupby(['Anno', '_sorgente'])['Valore'].sum().reset_index()
                    fig_val = px.bar(df_val, x='Anno', y='Valore', color='_sorgente',
                                     color_discrete_map=PALETTE, barmode='stack',
                                     labels={'Valore': 'Valore (EUR)', '_sorgente': 'Fonte'},
                                     title="Valore ordinato per anno")
                    fig_val.update_layout(height=260, margin=dict(t=40, b=10, l=0, r=0),
                                          plot_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig_val, use_container_width=True)

        # ════════════════════════════════════════════════════════════
        # SEZIONE 2 — STORICITA E ROTAZIONE
        # ════════════════════════════════════════════════════════════
        st.markdown('<div class="sec-h">🔄 Storicita e Rotazione Articoli</div>', unsafe_allow_html=True)

        df_rot = _calcola_storicita_rotazione(df_f)

        if not df_rot.empty:
            tab_qta, tab_freq, tab_fam_v = st.tabs(
                ["Top 20 per Quantita", "Top 20 per Frequenza", "Per Famiglia"])

            with tab_qta:
                top20 = df_rot.head(20)
                fig = px.bar(top20, x='Articolo C', y='Qta_Totale', color='Famiglia',
                             text='Qta_Totale',
                             color_discrete_sequence=px.colors.qualitative.Set2,
                             labels={'Qta_Totale': 'Quantita (pa.)', 'Articolo C': ''})
                fig.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                fig.update_layout(height=340, margin=dict(t=10, b=60, l=0, r=0),
                                  xaxis_tickangle=-40, plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig, use_container_width=True)

            with tab_freq:
                top20f = df_rot.sort_values('N_Ordini', ascending=False).head(20)
                fig2 = px.bar(top20f, x='Articolo C', y='N_Ordini', color='Famiglia',
                              text='N_Ordini',
                              color_discrete_sequence=px.colors.qualitative.Pastel,
                              labels={'N_Ordini': 'N ordini', 'Articolo C': ''})
                fig2.update_traces(texttemplate='%{text}', textposition='outside')
                fig2.update_layout(height=340, margin=dict(t=10, b=60, l=0, r=0),
                                   xaxis_tickangle=-40, plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig2, use_container_width=True)

            with tab_fam_v:
                df_fam = df_rot.groupby('Famiglia').agg(
                    Qta_Totale  =('Qta_Totale',    'sum'),
                    N_Articoli  =('Articolo C',     'count'),
                    N_Ordini    =('N_Ordini',       'sum'),
                    Valore_Tot  =('Valore_Totale',  'sum'),
                ).reset_index().sort_values('Qta_Totale', ascending=False)
                col_fa, col_fb = st.columns(2)
                with col_fa:
                    fig_fam = px.pie(df_fam.head(10), values='Qta_Totale', names='Famiglia',
                                     hole=0.4, title="Distribuzione quantita per famiglia",
                                     color_discrete_sequence=px.colors.qualitative.Set3)
                    fig_fam.update_traces(textinfo='label+percent')
                    fig_fam.update_layout(height=320, margin=dict(t=40, b=0, l=0, r=0), showlegend=False)
                    st.plotly_chart(fig_fam, use_container_width=True)
                with col_fb:
                    st.dataframe(df_fam.rename(columns={
                        'Qta_Totale': 'Qta Totale (pa.)', 'N_Articoli': 'N Articoli',
                        'N_Ordini': 'N Ordini', 'Valore_Tot': 'Valore (EUR)',
                    }), use_container_width=True, hide_index=True)

            with st.expander("Dettaglio completo rotazione"):
                show_rot = df_rot.copy()
                for c in ['Prima_Richiesta', 'Ultima_Richiesta']:
                    show_rot[c] = pd.to_datetime(show_rot[c], errors='coerce').dt.strftime('%d/%m/%Y')
                st.dataframe(show_rot, use_container_width=True, hide_index=True)

        # ════════════════════════════════════════════════════════════
        # SEZIONE 3 — FREQUENZA DI RIORDINO
        # ════════════════════════════════════════════════════════════
        st.markdown('<div class="sec-h">⏰ Frequenza di Riordino</div>', unsafe_allow_html=True)

        df_riord = _calcola_frequenza_riordino(df_f)

        if not df_riord.empty:
            df_con_int = df_riord.dropna(subset=['Intervallo_Medio_gg'])
            r1, r2, r3 = st.columns(3)
            _kpi_card(r1, "Coppie Art.xCliente con riordino",
                      str(len(df_con_int)), "2+ ordini nel periodo", "g")
            if not df_con_int.empty:
                media_g = round(df_con_int['Intervallo_Medio_gg'].mean(), 1)
                art_top = df_con_int.sort_values('N_Ordini', ascending=False).iloc[0]
                _kpi_card(r2, "Intervallo medio globale", f"{media_g} gg", "tra riordini consecutivi")
                qta_m = int(art_top['Qta_Media_Ordine']) if 'Qta_Media_Ordine' in art_top else 0
                _kpi_card(r3, "Articolo piu riordinato", art_top['Articolo C'],
                          f"{int(art_top['N_Ordini'])} ordini | ~{qta_m:,} pa./ordine".replace(",", "."), "p")

            pivot = df_riord.pivot_table(index='Cliente', columns='Famiglia',
                                         values='N_Ordini', aggfunc='sum', fill_value=0)
            if pivot.shape[0] > 1 and pivot.shape[1] > 1:
                top_cli = pivot.sum(axis=1).sort_values(ascending=False).head(15).index
                top_fam = pivot.sum(axis=0).sort_values(ascending=False).index
                pivot   = pivot.loc[top_cli][top_fam]
                fig_heat = px.imshow(pivot,
                                     labels=dict(x="Famiglia", y="Cliente", color="N Ordini"),
                                     color_continuous_scale="Blues", aspect="auto",
                                     title="N Ordini per Cliente x Famiglia")
                fig_heat.update_layout(height=420, margin=dict(t=40, b=0, l=0, r=0))
                fig_heat.update_xaxes(tickangle=-35)
                st.plotly_chart(fig_heat, use_container_width=True)

            # Alert riordini attesi
            df_prossimi = pd.DataFrame()
            if not df_con_int.empty:
                df_alert = df_con_int.copy()
                df_alert['Giorni_Al_Riordino'] = (
                    df_alert['Intervallo_Medio_gg'] - df_alert['Giorni_Da_Ultimo']
                ).round(0)
                df_prossimi = df_alert[df_alert['Giorni_Al_Riordino'].between(-7, 14)]\
                    .sort_values('Giorni_Al_Riordino')

            with st.expander("Dettaglio frequenza riordino"):
                show_r = df_riord.copy()
                if 'Ultimo_Ordine' in show_r.columns:
                    show_r['Ultimo_Ordine'] = pd.to_datetime(
                        show_r['Ultimo_Ordine'], errors='coerce').dt.strftime('%d/%m/%Y')
                st.dataframe(show_r, use_container_width=True, hide_index=True)

            if not df_prossimi.empty:
                st.markdown('<div class="sec-h">🔔 Articoli con riordino atteso entro 14 giorni</div>',
                            unsafe_allow_html=True)
                st.markdown(f"**{len(df_prossimi)} articoli:**")
                for _, r in df_prossimi.iterrows():
                    gg_r = int(r['Giorni_Al_Riordino'])
                    ico  = "🟢" if gg_r > 0 else "🔴"
                    txt  = f"tra **{gg_r} gg**" if gg_r >= 0 else f"**in ritardo di {abs(gg_r)} gg**"
                    ult  = r['Ultimo_Ordine'].strftime('%d/%m/%Y') if pd.notnull(r['Ultimo_Ordine']) else 'N/D'
                    qta_att = int(r['Qta_Media_Ordine']) if 'Qta_Media_Ordine' in r else 0
                    st.markdown(
                        f'<div class="alert-box">{ico} <b>{r["Articolo C"]}</b> — '
                        f'{r["Cliente"]} | Atteso {txt} | '
                        f'Qta stimata: <b>~{qta_att:,} pa.</b> | '.replace(",", ".")
                        + f'Ultimo ordine: {ult} | Intervallo medio: {r["Intervallo_Medio_gg"]} gg'
                        f'</div>', unsafe_allow_html=True)

        # ════════════════════════════════════════════════════════════
        # SEZIONE 4 — AUDIT TRANSCODIFICA
        # ════════════════════════════════════════════════════════════
        with st.expander("⚠️ Articoli SAFIT senza transcodifica (audit)"):
            df_nm = articoli_non_mappati(df_all)
            if df_nm.empty:
                st.success("Tutti gli articoli SAFIT sono mappati.")
            else:
                st.warning(f"{len(df_nm)} codici SAFIT non mappati nella tabella di transcodifica.")
                st.dataframe(df_nm.style.format({'Valore_Tot': '{:,.0f}', 'Qta_Tot': '{:,.0f}'}),
                             use_container_width=True, hide_index=True)
