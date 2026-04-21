# ==============================================================================
# kpi_qualita.py — KPI Qualità / Commerciali per Safit Portal v3.9
# ==============================================================================
# Fonte dati: righe_ordini_storico_con_date.xlsx
#   DVF  = fatture clienti (fatturato reale)
#   CFF/CFR = carichi da fornitore (per giorni ritardo)
#   OCI  = ordini clienti (quantità vendute)
# ==============================================================================

import os
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta, date
from io import BytesIO

import os as _os
_DIR = _os.path.dirname(_os.path.abspath(__file__))
PATH_STORICO = _os.path.join(_DIR, "righe_ordini_storico_con_date.xlsx")

# Prefissi codice articolo per categoria
PREF_PUNTALI = ('PPU','PPT','PPA','PPF','PCT','PCP','CPU','PAT','PAC')
PREF_SOLETTE = ('PSO','CSO','CSCH','PSC','CSC')

KPI_CSS = """
<style>
.kq-card {
    background:linear-gradient(135deg,#ffffff,#f8faff);
    border:1px solid #e3e8f0; border-left:4px solid #1f77b4;
    padding:14px 16px; border-radius:10px; margin-bottom:6px;
    box-shadow:0 2px 6px rgba(0,0,0,0.05); color:#1a1a2e !important;
}
.kq-card * { color:#1a1a2e !important; }
.kq-card.g { border-left-color:#4caf50; }
.kq-card.o { border-left-color:#ff9800; }
.kq-card.r { border-left-color:#f44336; }
.kq-card.p { border-left-color:#9c27b0; }
.kq-t { font-size:10px;color:#888 !important;text-transform:uppercase;letter-spacing:.5px;margin-bottom:3px; }
.kq-v { font-size:26px;font-weight:700;color:#1a1a2e !important;line-height:1; }
.kq-s { font-size:11px;color:#555 !important;margin-top:3px; }
.kq-sec { font-size:16px;font-weight:700;color:#1a1a2e !important;
          border-bottom:2px solid #e3e8f0;padding-bottom:5px;margin:20px 0 12px 0; }
</style>
"""

def _card(col, title, value, sub="", color=""):
    col.markdown(
        f'<div class="kq-card {color}">'
        f'<div class="kq-t">{title}</div>'
        f'<div class="kq-v">{value}</div>'
        f'<div class="kq-s">{sub}</div>'
        f'</div>', unsafe_allow_html=True)

def _sec(title):
    st.markdown(f'<div class="kq-sec">{title}</div>', unsafe_allow_html=True)


@st.cache_data(ttl=1800, show_spinner=False)
def _carica_dati(path=PATH_STORICO):
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        df = pd.read_excel(path, skiprows=2)
        df.columns = [str(c).strip() for c in df.columns]
        df['Codice Documento'] = df['Codice Documento'].ffill()
        df['Data Consegna']    = pd.to_datetime(df['Data Consegna'], errors='coerce').ffill()
        df['Numero Documento'] = df['Numero Documento'].ffill()
        df['Data']    = pd.to_datetime(df['Data'],    errors='coerce')
        df['Qta Doc'] = pd.to_numeric(df['Qta Doc'],  errors='coerce').fillna(0)
        df['Valore']  = pd.to_numeric(df['Valore'],   errors='coerce').fillna(0)
        # Rimuovi righe totale e vuote
        df = df[~df['Codice Documento'].astype(str).str.contains('Totale|otale', na=False)]
        df = df[df['Articolo C'].notna() & (df['Articolo C'].astype(str).str.strip() != '(vuoto)')]
        # Campi derivati
        df['Cliente']     = df['Cliente Fornitore CD'].astype(str).str.split(' - ', n=1).str[-1].str.strip()
        df['Cod_Cliente'] = df['Cliente Fornitore CD'].astype(str).str.split(' - ', n=1).str[0].str.strip()
        df['Anno']        = df['Data'].dt.year
        df['Mese']        = df['Data'].dt.to_period('M').astype(str)
        df['cod_up']      = df['Articolo C'].fillna('').astype(str).str.strip().str.upper()
        df['Categoria']   = df['cod_up'].apply(
            lambda x: 'Puntale' if x.startswith(PREF_PUNTALI)
            else ('Soletta' if x.startswith(PREF_SOLETTE) else 'Altro')
        )
        df['Famiglia'] = df['Articolo D'].apply(
            lambda x: ' '.join(str(x).split()[:2]).upper() if pd.notna(x) and str(x).strip() not in ('nan','(vuoto)','') else 'ALTRO'
        )
        return df
    except Exception as e:
        st.warning(f"Errore lettura dati: {e}")
        return pd.DataFrame()


def _applica_filtro_periodo(df, periodo, data_da=None, data_a=None):
    gg_map = {"Ultimi 90 gg": 90, "Ultimi 6 mesi": 180,
              "Ultimo anno": 365, "Ultimi 2 anni": 730, "Tutto lo storico": 99999}
    if periodo == "Intervallo personalizzato" and data_da and data_a:
        return df[(df['Data'] >= pd.Timestamp(data_da)) & (df['Data'] <= pd.Timestamp(data_a))]
    gg = gg_map.get(periodo, 99999)
    if gg < 99999:
        return df[df['Data'] >= pd.Timestamp(datetime.now() - timedelta(days=gg))]
    return df


# ==============================================================================
# SEZIONE 1 — FATTURATO (DVF)
# ==============================================================================
def _render_fatturato(df_dvf):
    _sec("💶 Fatturato (Fatture DVF)")

    if df_dvf.empty:
        st.warning("Nessun dato DVF disponibile.")
        return

    c1, c2, c3, c4 = st.columns(4)
    tot = df_dvf['Valore'].sum()
    n_cli = df_dvf['Cod_Cliente'].nunique()
    media_m = df_dvf.groupby('Mese')['Valore'].sum().mean()
    mese_top = df_dvf.groupby('Mese')['Valore'].sum().idxmax()
    _card(c1, "Fatturato totale periodo", f"€ {tot:,.0f}".replace(",","."), "somma DVF", "g")
    _card(c2, "Clienti fatturati", str(n_cli), "distinti nel periodo", "p")
    _card(c3, "Media mensile", f"€ {media_m:,.0f}".replace(",","."), "per mese")
    _card(c4, "Mese migliore", mese_top, f"€ {df_dvf.groupby('Mese')['Valore'].sum()[mese_top]:,.0f}".replace(",","."), "o")

    tab_mens, tab_ann, tab_cli = st.tabs(["📅 Mensile", "📆 Annuale", "👥 Per cliente"])

    with tab_mens:
        df_m = df_dvf.groupby('Mese')['Valore'].sum().reset_index().sort_values('Mese')
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(go.Bar(x=df_m['Mese'], y=df_m['Valore'],
                             name='Fatturato', marker_color='#1f77b4', opacity=0.85), secondary_y=False)
        fig.add_trace(go.Scatter(x=df_m['Mese'], y=df_m['Valore'].cumsum(),
                                 name='Cumulato', mode='lines', line=dict(color='#ff7f0e', width=2)),
                      secondary_y=True)
        fig.update_layout(height=320, margin=dict(t=10,b=40,l=0,r=0),
                          plot_bgcolor='rgba(0,0,0,0)', legend=dict(orientation='h', y=-0.25))
        fig.update_yaxes(title_text="Fatturato (€)", secondary_y=False)
        fig.update_yaxes(title_text="Cumulato (€)", secondary_y=True)
        st.plotly_chart(fig, use_container_width=True)

    with tab_ann:
        df_a = df_dvf.groupby('Anno')['Valore'].sum().reset_index()
        fig_a = px.bar(df_a, x='Anno', y='Valore', text='Valore',
                       labels={'Valore':'Fatturato (€)'},
                       color_discrete_sequence=['#1f77b4'])
        fig_a.update_traces(texttemplate='€ %{text:,.0f}', textposition='outside')
        fig_a.update_layout(height=300, margin=dict(t=10,b=10,l=0,r=0),
                            plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_a, use_container_width=True)

        # YoY confronto mese per mese
        df_yoy = df_dvf.groupby(['Anno', df_dvf['Data'].dt.month.rename('Mese_n')])['Valore'].sum().reset_index()
        df_yoy['Mese_lbl'] = df_yoy['Mese_n'].apply(lambda m: date(2000, m, 1).strftime('%b'))
        fig_yoy = px.line(df_yoy, x='Mese_lbl', y='Valore', color='Anno',
                          labels={'Valore':'Fatturato (€)', 'Mese_lbl':'Mese'},
                          title="Confronto anno su anno (mese per mese)")
        fig_yoy.update_layout(height=300, margin=dict(t=40,b=10,l=0,r=0),
                               plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_yoy, use_container_width=True)

    with tab_cli:
        df_cli = df_dvf.groupby('Cliente')['Valore'].sum().sort_values(ascending=False).reset_index()
        fig_cli = px.bar(df_cli.head(20), x='Valore', y='Cliente', orientation='h',
                         labels={'Valore':'Fatturato (€)', 'Cliente':''},
                         color_discrete_sequence=['#1f77b4'])
        fig_cli.update_layout(height=500, margin=dict(t=10,b=10,l=0,r=0),
                              plot_bgcolor='rgba(0,0,0,0)', yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_cli, use_container_width=True)


# ==============================================================================
# SEZIONE 2 — ANALISI CLIENTI
# ==============================================================================
def _render_clienti(df_dvf, df_oci):
    _sec("👥 Analisi Clienti")

    if df_dvf.empty:
        st.warning("Nessun dato disponibile.")
        return

    anni = sorted(df_dvf['Anno'].dropna().unique().tolist())
    if len(anni) < 2:
        st.info("Servono almeno 2 anni di dati per l'analisi acquisiti/persi.")

    # Clienti per anno
    cli_per_anno = {a: set(df_dvf[df_dvf['Anno']==a]['Cod_Cliente'].unique()) for a in anni}

    c1, c2, c3, c4 = st.columns(4)
    tot_cli = df_dvf['Cod_Cliente'].nunique()
    _card(c1, "Clienti totali periodo", str(tot_cli), "distinti con almeno 1 DVF")

    if len(anni) >= 2:
        anno_curr = max(anni)
        anno_prev = sorted([a for a in anni if a < anno_curr])[-1]
        acquisiti = cli_per_anno[anno_curr] - cli_per_anno[anno_prev]
        persi     = cli_per_anno[anno_prev] - cli_per_anno[anno_curr]
        stabili   = cli_per_anno[anno_curr] & cli_per_anno[anno_prev]
        _card(c2, f"Acquisiti {anno_curr}", str(len(acquisiti)), f"non presenti in {anno_prev}", "g")
        _card(c3, f"Persi {anno_curr}", str(len(persi)), f"presenti in {anno_prev} ma non ora", "r")
        _card(c4, "Stabili", str(len(stabili)), f"attivi in entrambi gli anni")

        st.markdown(f"**Dettaglio clienti acquisiti {anno_curr}:**")
        if acquisiti:
            df_acq = df_dvf[(df_dvf['Anno']==anno_curr) & (df_dvf['Cod_Cliente'].isin(acquisiti))]
            df_acq_sum = df_acq.groupby('Cliente')['Valore'].sum().sort_values(ascending=False).reset_index()
            st.dataframe(df_acq_sum.style.format({'Valore':'{:,.0f}'}), use_container_width=True, hide_index=True)

        st.markdown(f"**Clienti persi (presenti in {anno_prev}, assenti in {anno_curr}):**")
        if persi:
            df_persi = df_dvf[(df_dvf['Anno']==anno_prev) & (df_dvf['Cod_Cliente'].isin(persi))]
            df_persi_sum = df_persi.groupby('Cliente')['Valore'].sum().sort_values(ascending=False).reset_index()
            st.dataframe(df_persi_sum.style.format({'Valore':'{:,.0f}'}), use_container_width=True, hide_index=True)

    # Clienti sporadici (1 solo ordine nell'anno corrente)
    if len(anni) >= 1:
        anno_curr = max(anni)
        df_curr = df_dvf[df_dvf['Anno'] == anno_curr]
        n_ordini = df_curr.groupby('Cod_Cliente')['Numero Documento'].nunique()
        sporadici = n_ordini[n_ordini == 1].index
        _sec(f"⚠️ Clienti sporadici {anno_curr} (1 solo ordine)")
        df_spor = df_curr[df_curr['Cod_Cliente'].isin(sporadici)].groupby('Cliente')['Valore'].sum().sort_values(ascending=False).reset_index()
        st.caption(f"{len(df_spor)} clienti con un solo ordine nell'anno")
        st.dataframe(df_spor.style.format({'Valore':'{:,.0f}'}), use_container_width=True, hide_index=True)

    # Pareto: clienti A+B (80% del fatturato)
    _sec("📊 Legge di Pareto — Clienti A+B (80% fatturato)")
    df_pareto = df_dvf.groupby('Cliente')['Valore'].sum().sort_values(ascending=False).reset_index()
    df_pareto['Cumulato'] = df_pareto['Valore'].cumsum()
    df_pareto['Pct_Cumulato'] = df_pareto['Cumulato'] / df_pareto['Valore'].sum() * 100
    df_pareto['Pct_Clienti'] = (df_pareto.index + 1) / len(df_pareto) * 100
    df_pareto['Classe'] = df_pareto['Pct_Cumulato'].apply(
        lambda x: 'A' if x <= 50 else ('B' if x <= 80 else 'C'))

    n_ab = (df_pareto['Classe'].isin(['A','B'])).sum()
    pct_cli_ab = round(n_ab / len(df_pareto) * 100, 1)
    pct_fatt_ab = round(df_pareto[df_pareto['Classe'].isin(['A','B'])]['Valore'].sum() / df_pareto['Valore'].sum() * 100, 1)

    p1, p2, p3 = st.columns(3)
    _card(p1, "Clienti classe A+B", str(n_ab), f"{pct_cli_ab}% dei clienti totali", "p")
    _card(p2, "% fatturato A+B", f"{pct_fatt_ab}%", "quota sul totale", "g")
    _card(p3, "Clienti classe C", str(len(df_pareto)-n_ab), "coda lunga", "o")

    fig_pareto = make_subplots(specs=[[{"secondary_y": True}]])
    colors = {'A':'#1f77b4','B':'#ff7f0e','C':'#aec7e8'}
    fig_pareto.add_trace(
        go.Bar(x=df_pareto['Cliente'], y=df_pareto['Valore'],
               marker_color=df_pareto['Classe'].map(colors),
               name='Fatturato', opacity=0.85), secondary_y=False)
    fig_pareto.add_trace(
        go.Scatter(x=df_pareto['Cliente'], y=df_pareto['Pct_Cumulato'],
                   name='% Cumulato', mode='lines', line=dict(color='#d62728', width=2)),
        secondary_y=True)
    fig_pareto.add_hline(y=80, line_dash="dash", line_color="#d62728",
                          annotation_text="80%", secondary_y=True)
    fig_pareto.update_layout(height=380, margin=dict(t=10,b=60,l=0,r=0),
                              plot_bgcolor='rgba(0,0,0,0)',
                              xaxis_tickangle=-45, showlegend=True,
                              legend=dict(orientation='h', y=-0.35))
    fig_pareto.update_yaxes(title_text="Fatturato (€)", secondary_y=False)
    fig_pareto.update_yaxes(title_text="% Cumulato", secondary_y=True, range=[0,105], ticksuffix='%')
    st.plotly_chart(fig_pareto, use_container_width=True)

    with st.expander("📋 Dettaglio clienti con classe Pareto"):
        df_show = df_pareto[['Cliente','Valore','Pct_Cumulato','Classe']].rename(
            columns={'Valore':'Fatturato (€)','Pct_Cumulato':'% Cumulato','Classe':'Classe Pareto'})
        st.dataframe(df_show.style.format({'Fatturato (€)':'{:,.0f}','% Cumulato':'{:.1f}%'}),
                     use_container_width=True, hide_index=True)


# ==============================================================================
# SEZIONE 3 — QUANTITÀ VENDUTE (OCI) — per famiglia merceologica
# ==============================================================================
# Palette colori per famiglie
_FAM_COLORS = [
    '#1f77b4','#ff7f0e','#2ca02c','#d62728','#9467bd',
    '#8c564b','#e377c2','#7f7f7f','#bcbd22','#17becf'
]

def _render_quantita(df_oci):
    _sec("📦 Quantità Vendute per Famiglia Merceologica (OCI)")

    if df_oci.empty:
        st.warning("Nessun dato OCI disponibile.")
        return

    # Famiglie significative (escludi voci non merceologiche)
    ESCLUDI = {'RIMBORSO SPESE','SPESE DI','POLVERE MP','TESSUTO KNOX',
               'SBAVATORE STAMPI','ALTRO','NAN'}
    df_oci = df_oci[~df_oci['Famiglia'].isin(ESCLUDI)].copy()

    fam_tot = df_oci.groupby('Famiglia')['Qta Doc'].sum().sort_values(ascending=False)
    fam_list = fam_tot.index.tolist()
    tot_paia = int(df_oci['Qta Doc'].sum())

    # KPI top famiglie
    cols = st.columns(min(4, len(fam_list)))
    for i, fam in enumerate(fam_list[:4]):
        qta = int(fam_tot[fam])
        pct = round(qta / tot_paia * 100, 1) if tot_paia > 0 else 0
        _card(cols[i], fam, f"{qta:,}".replace(",","."), f"pa. — {pct}% del totale")

    tab_fam, tab_mens, tab_ann, tab_trend = st.tabs(
        ["🥧 Per famiglia", "📅 Mensile", "📆 Annuale", "📈 Trend famiglie"])

    with tab_fam:
        # Torta + barre affiancate
        c1, c2 = st.columns([1, 1])
        with c1:
            fig_pie = px.pie(
                values=fam_tot.values, names=fam_tot.index,
                color_discrete_sequence=_FAM_COLORS, hole=0.35,
                title="Ripartizione % per famiglia")
            fig_pie.update_traces(textinfo='label+percent', sort=True)
            fig_pie.update_layout(height=350, margin=dict(t=40,b=0,l=0,r=0), showlegend=False)
            st.plotly_chart(fig_pie, use_container_width=True)
        with c2:
            fig_bar = px.bar(
                x=fam_tot.values, y=fam_tot.index, orientation='h',
                color=fam_tot.index, color_discrete_sequence=_FAM_COLORS,
                labels={'x':'Quantità (pa.)','y':''},
                title="Quantità totale per famiglia")
            fig_bar.update_layout(height=350, margin=dict(t=40,b=0,l=0,r=0),
                                  plot_bgcolor='rgba(0,0,0,0)',
                                  showlegend=False,
                                  yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_bar, use_container_width=True)

    with tab_mens:
        df_m = (df_oci.groupby(['Mese','Famiglia'])['Qta Doc'].sum()
                .reset_index().sort_values('Mese'))
        fig_m = px.bar(df_m, x='Mese', y='Qta Doc', color='Famiglia',
                       color_discrete_sequence=_FAM_COLORS,
                       labels={'Qta Doc':'Quantità (pa.)','Mese':''},
                       barmode='stack')
        fig_m.update_layout(height=340, margin=dict(t=10,b=60,l=0,r=0),
                             plot_bgcolor='rgba(0,0,0,0)',
                             xaxis_tickangle=-45,
                             legend=dict(orientation='h', y=-0.4))
        st.plotly_chart(fig_m, use_container_width=True)

    with tab_ann:
        df_a = (df_oci.groupby(['Anno','Famiglia'])['Qta Doc'].sum()
                .reset_index().sort_values('Anno'))
        fig_a = px.bar(df_a, x='Anno', y='Qta Doc', color='Famiglia',
                       color_discrete_sequence=_FAM_COLORS,
                       labels={'Qta Doc':'Quantità (pa.)','Anno':''},
                       barmode='stack', text_auto=False)
        fig_a.update_layout(height=340, margin=dict(t=10,b=10,l=0,r=0),
                             plot_bgcolor='rgba(0,0,0,0)',
                             legend=dict(orientation='h', y=-0.25))
        st.plotly_chart(fig_a, use_container_width=True)

        # Tabella riepilogo anno x famiglia
        df_pivot = (df_oci.groupby(['Famiglia','Anno'])['Qta Doc'].sum()
                    .unstack(fill_value=0))
        st.dataframe(df_pivot.style.format('{:,.0f}').background_gradient(
            cmap='Blues', axis=None),
                     use_container_width=True)

    with tab_trend:
        # Trend mensile per singola famiglia selezionabile
        fam_sel = st.multiselect("Seleziona famiglie da confrontare:",
                                  fam_list, default=fam_list[:3],
                                  key="kq_fam_trend")
        if fam_sel:
            df_t = (df_oci[df_oci['Famiglia'].isin(fam_sel)]
                    .groupby(['Mese','Famiglia'])['Qta Doc'].sum()
                    .reset_index().sort_values('Mese'))
            fig_t = px.line(df_t, x='Mese', y='Qta Doc', color='Famiglia',
                            color_discrete_sequence=_FAM_COLORS,
                            labels={'Qta Doc':'Quantità (pa.)','Mese':''},
                            markers=True)
            fig_t.update_layout(height=320, margin=dict(t=10,b=60,l=0,r=0),
                                 plot_bgcolor='rgba(0,0,0,0)',
                                 xaxis_tickangle=-45,
                                 legend=dict(orientation='h', y=-0.4))
            st.plotly_chart(fig_t, use_container_width=True)


# ==============================================================================
# SEZIONE 4 — GIORNI RITARDO FORNITORI
# ==============================================================================
def _render_ritardo_fornitori(df_cff):
    _sec("🚚 Puntualità Fornitori (CFF/CFR)")

    if df_cff.empty:
        st.warning("Nessun dato CFF/CFR disponibile.")
        return

    # Calcolo delta corretto: collega CFF/CFR al relativo OFF/OFR
    # tramite Fornitore+Articolo, usando Data Consegna DELL'ORDINE OFF (non del CFF)
    df_all_local = _carica_dati()

    # Estrai OFF/OFR con Data Consegna valorizzata (solo prima riga per numero documento)
    df_off_loc = df_all_local[df_all_local['Codice Documento'].isin(['OFF','OFR'])].copy()
    df_off_loc['DC_raw'] = pd.to_datetime(
        df_all_local.loc[df_off_loc.index, 'Data Consegna'], errors='coerce')
    # Prende la prima Data Consegna valorizzata per gruppo numero documento
    nr_to_dc = (df_off_loc[df_off_loc['DC_raw'].notna()]
                .drop_duplicates(subset=['Numero Documento'])
                .set_index('Numero Documento')['DC_raw'].to_dict())
    df_off_loc['DC_off'] = df_off_loc['Numero Documento'].map(nr_to_dc)
    df_off_key = df_off_loc[df_off_loc['DC_off'].notna()][
        ['Articolo C','Cliente','Data','DC_off']].copy()
    df_off_key.columns = ['Articolo C','Fornitore','Data_OFF','DC_prevista']
    df_off_key = df_off_key.sort_values('Data_OFF')

    # Collega ogni CFF al suo OFF: stesso Fornitore+Articolo, Data_OFF <= Data_CFF
    df = df_cff[df_cff['Data'].notna()].copy()
    df_merge = df[['Articolo C','Cliente','Data']].copy()
    df_merge.columns = ['Articolo C','Fornitore','Data_CFF']
    merged = df_merge.merge(df_off_key, on=['Articolo C','Fornitore'], how='left')
    merged = merged[merged['Data_OFF'] <= merged['Data_CFF']]
    merged = (merged.sort_values('Data_OFF')
              .groupby(['Articolo C','Fornitore','Data_CFF']).last().reset_index())
    merged['Delta_gg'] = (merged['Data_CFF'] - merged['DC_prevista']).dt.days

    # Rimappa il delta sul df originale
    df = df.copy()
    df['_key'] = list(zip(df['Articolo C'], df['Cliente'], df['Data']))
    merged['_key'] = list(zip(merged['Articolo C'], merged['Fornitore'], merged['Data_CFF']))
    delta_map = merged.set_index('_key')['Delta_gg'].to_dict()
    df['Delta_gg'] = df['_key'].map(delta_map)
    df = df[df['Delta_gg'].notna() & df['Delta_gg'].between(-30, 180)]

    if df.empty:
        st.warning("Nessun delta calcolabile.")
        return

    df['Stato'] = df['Delta_gg'].apply(
        lambda x: '✅ Puntuale' if x <= 3 else '⚠️ In ritardo')

    c1, c2, c3, c4 = st.columns(4)
    pct_punt = round((df['Delta_gg'] <= 3).mean() * 100, 1)
    media_rit = round(df[df['Delta_gg'] > 3]['Delta_gg'].mean(), 1) if (df['Delta_gg'] > 3).any() else 0
    _card(c1, "Carichi analizzati", f"{len(df):,}".replace(",","."), "righe CFF/CFR")
    col_p = "g" if pct_punt >= 80 else "o"
    _card(c2, "% Puntuali (≤3gg)", f"{pct_punt}%", f"{100-pct_punt:.1f}% in ritardo", col_p)
    col_m = "g" if media_rit <= 7 else "o" if media_rit <= 14 else "r"
    _card(c3, "Ritardo medio", f"{media_rit} gg", "solo sui ritardatari", col_m)
    _card(c4, "Mediana delta", f"{df['Delta_gg'].median():.0f} gg",
          "negativo = anticipo", "p")

    tab_dist, tab_forn, tab_trend = st.tabs(["📊 Distribuzione", "🏭 Per fornitore", "📅 Trend mensile"])

    with tab_dist:
        fig_hist = px.histogram(df, x='Delta_gg', color='Stato', nbins=40,
                                color_discrete_map={'✅ Puntuale':'#4caf50','⚠️ In ritardo':'#f44336'},
                                labels={'Delta_gg':'Giorni (negativo=anticipo)', 'count':'N° carichi'},
                                title="Distribuzione giorni ritardo/anticipo")
        fig_hist.add_vline(x=0, line_dash="dash", line_color="#888", annotation_text=" Previsto")
        fig_hist.add_vline(x=3, line_dash="dot", line_color="#ff9800", annotation_text=" +3gg")
        fig_hist.update_layout(height=300, margin=dict(t=40,b=0,l=0,r=0),
                               plot_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig_hist, use_container_width=True)

    with tab_forn:
        df_forn = df.groupby('Cliente').agg(
            N_Carichi=('Delta_gg','count'),
            Delta_Medio=('Delta_gg','mean'),
            Pct_Puntuali=('Delta_gg', lambda x: round((x<=3).mean()*100,1)),
        ).reset_index().sort_values('Delta_Medio', ascending=False)
        df_forn['Delta_Medio'] = df_forn['Delta_Medio'].round(1)

        fig_forn = px.bar(df_forn, x='Delta_Medio', y='Cliente', orientation='h',
                          color='Pct_Puntuali',
                          color_continuous_scale=['#f44336','#ff9800','#4caf50'],
                          range_color=[0,100],
                          labels={'Delta_Medio':'Delta medio (gg)', 'Cliente':'Fornitore',
                                  'Pct_Puntuali':'% Puntuali'},
                          title="Delta medio per fornitore (colorato per % puntualità)")
        fig_forn.add_vline(x=0, line_dash="dash", line_color="#888")
        fig_forn.update_layout(height=max(300, len(df_forn)*30),
                               margin=dict(t=40,b=0,l=0,r=0),
                               plot_bgcolor='rgba(0,0,0,0)',
                               yaxis={'categoryorder':'total ascending'})
        st.plotly_chart(fig_forn, use_container_width=True)

        st.dataframe(df_forn.rename(columns={
            'Cliente':'Fornitore','N_Carichi':'N° Carichi',
            'Delta_Medio':'Delta Medio (gg)','Pct_Puntuali':'% Puntuali'
        }).style.format({'Delta Medio (gg)':'{:.1f}','% Puntuali':'{:.1f}%'}),
                     use_container_width=True, hide_index=True)

    with tab_trend:
        df['Mese'] = df['Data'].dt.to_period('M').astype(str)
        df_trend = df.groupby('Mese').agg(
            Delta_Medio=('Delta_gg','mean'),
            N_Carichi=('Delta_gg','count'),
            Pct_Puntuali=('Delta_gg', lambda x: round((x<=3).mean()*100,1)),
        ).reset_index().sort_values('Mese')

        fig_t = make_subplots(specs=[[{"secondary_y": True}]])
        fig_t.add_trace(go.Bar(x=df_trend['Mese'], y=df_trend['N_Carichi'],
                               name='N° carichi', marker_color='#e3f2fd', opacity=0.8),
                        secondary_y=False)
        fig_t.add_trace(go.Scatter(x=df_trend['Mese'], y=df_trend['Pct_Puntuali'],
                                   name='% Puntuali', mode='lines+markers',
                                   line=dict(color='#4caf50', width=2), marker=dict(size=7)),
                        secondary_y=True)
        fig_t.add_hline(y=80, line_dash="dot", line_color="#ff9800",
                        annotation_text=" Target 80%", secondary_y=True)
        fig_t.update_layout(height=280, margin=dict(t=10,b=40,l=0,r=0),
                             plot_bgcolor='rgba(0,0,0,0)',
                             legend=dict(orientation='h', y=-0.25))
        fig_t.update_yaxes(title_text="N° carichi", secondary_y=False)
        fig_t.update_yaxes(title_text="% Puntuali", secondary_y=True,
                           range=[0,110], ticksuffix='%')
        st.plotly_chart(fig_t, use_container_width=True)


# ==============================================================================
# ENTRY POINT
# ==============================================================================
def render_kpi_qualita(filtro_cliente=None, filtro_famiglie=None):
    """Punto di ingresso — chiamato dal portale nel tab KPI Avanzati."""
    st.markdown(KPI_CSS, unsafe_allow_html=True)

    with st.expander("📊 KPI Qualità — Fatturato · Clienti · Quantità · Fornitori",
                     expanded=False):

        df_all = _carica_dati()
        if df_all.empty:
            st.warning(f"File non trovato: `{PATH_STORICO}`")
            return

        # Filtri periodo
        _key = "kq"
        c1, c2 = st.columns([1, 2])
        with c1:
            periodo = st.selectbox("Periodo",
                ["Ultimi 90 gg","Ultimi 6 mesi","Ultimo anno",
                 "Ultimi 2 anni","Tutto lo storico","Intervallo personalizzato"],
                index=2, key=f"{_key}_periodo")
        with c2:
            fornitori_disp = sorted(
                df_all[df_all['Codice Documento'].isin(['CFF','CFR'])]
                ['Cliente'].dropna().unique().tolist()
            )
            sel_forn = st.multiselect("Filtra fornitore (per ritardi)",
                                      fornitori_disp, key=f"{_key}_forn",
                                      placeholder="Tutti i fornitori")

        data_da = data_a = None
        if periodo == "Intervallo personalizzato":
            dc1, dc2 = st.columns(2)
            data_min = df_all['Data'].min().date() if not df_all.empty else date(2025,1,1)
            data_max = df_all['Data'].max().date() if not df_all.empty else datetime.now().date()
            with dc1:
                data_da = st.date_input("Dal", value=data_min, key=f"{_key}_da", format="DD/MM/YYYY")
            with dc2:
                data_a  = st.date_input("Al",  value=data_max, key=f"{_key}_a",  format="DD/MM/YYYY")

        # Filtra per periodo
        df_f = _applica_filtro_periodo(df_all, periodo, data_da, data_a)

        # Estrai dataset per tipologia
        df_dvf = df_f[df_f['Codice Documento'] == 'DVF'].copy()
        df_oci = df_f[df_f['Codice Documento'].isin(['OCI']) & (df_f['Qta Doc'] > 0)].copy()
        df_cff = df_f[df_f['Codice Documento'].isin(['CFF','CFR'])].copy()

        # Filtra fornitore se selezionato
        if sel_forn:
            df_cff = df_cff[df_cff['Cliente'].isin(sel_forn)]

        st.caption(
            f"DVF: **{len(df_dvf):,}** righe | "
            f"OCI: **{len(df_oci):,}** righe | "
            f"CFF/CFR: **{len(df_cff):,}** righe".replace(",",".")
        )

        # Render sezioni
        _render_fatturato(df_dvf)
        _render_clienti(df_dvf, df_oci)
        _render_quantita(df_oci)
        _render_ritardo_fornitori(df_cff)

        # Export Excel
        st.markdown("---")
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as w:
            if not df_dvf.empty:
                df_dvf.groupby(['Anno','Mese','Cliente'])['Valore'].sum().reset_index().to_excel(
                    w, sheet_name='Fatturato', index=False)
            if not df_oci.empty:
                df_oci.groupby(['Anno','Mese','Categoria'])['Qta Doc'].sum().reset_index().to_excel(
                    w, sheet_name='Quantita_Vendute', index=False)
            pareto = df_dvf.groupby('Cliente')['Valore'].sum().sort_values(ascending=False).reset_index()
            pareto['Pct_Cumulato'] = pareto['Valore'].cumsum() / pareto['Valore'].sum() * 100
            pareto['Classe'] = pareto['Pct_Cumulato'].apply(
                lambda x: 'A' if x<=50 else ('B' if x<=80 else 'C'))
            pareto.to_excel(w, sheet_name='Pareto_Clienti', index=False)
            if not df_cff.empty:
                df_cff_exp = df_cff[df_cff['Data Consegna'].notna() & df_cff['Data'].notna()].copy()
                df_cff_exp['Delta_gg'] = (df_cff_exp['Data'] - df_cff_exp['Data Consegna']).dt.days
                df_cff_exp.to_excel(w, sheet_name='Ritardo_Fornitori', index=False)

        st.download_button(
            "📥 Esporta KPI Qualità in Excel",
            data=buf.getvalue(),
            file_name=f"KPI_Qualita_{datetime.now().strftime('%d%m%Y')}.xlsx",
            use_container_width=True,
        )
