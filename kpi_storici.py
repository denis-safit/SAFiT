"""
kpi_storici.py
==============
Pagina Streamlit "KPI Storici (2011→oggi)" per il Safit Portal.

Da aggiungere al portale come nuova voce nel menu laterale, es:
    elif pagina == "KPI Storici":
        import kpi_storici
        kpi_storici.render()

Dipende da: storico_safit.py (nella stessa directory)
"""

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

import storico_safit as stor


# ─────────────────────────────────────────────────────────────
# COSTANTI UI
# ─────────────────────────────────────────────────────────────
COLOR_SAFIT   = "#1f77b4"   # blu
COLOR_SAFITIB = "#ff7f0e"   # arancio
COLOR_IBRIDO  = "#9467bd"   # viola (2025)

PALETTE = {
    'SAFIT':   COLOR_SAFIT,
    'SAFITIB': COLOR_SAFITIB,
    'IBRIDO':  COLOR_IBRIDO,
}

DOC_LABEL = {
    'OCA': 'Ordini clienti (OCA)',
    'OCI': 'Ordini clienti interni (OCI)',
    'OFF': 'Offerte fornitori (OFF)',
    'OFR': 'Ordini fornitori (OFR)',
    'OFI': 'Ordini fornitori interni (OFI)',
    'OFA': 'Ordini fornitori avanzati (OFA)',
}


# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────

def _color_bar(sorgente_series: pd.Series) -> list[str]:
    return [PALETTE.get(s, COLOR_SAFIT) for s in sorgente_series]


def _fmt_paia(n: float) -> str:
    return f"{n:,.0f} pa.".replace(",", ".")


def _fmt_euro(n: float) -> str:
    return f"€ {n:,.0f}".replace(",", ".")


# ─────────────────────────────────────────────────────────────
# RENDER PRINCIPALE
# ─────────────────────────────────────────────────────────────

def render():
    st.title("📊 KPI Storici — SAFIT 2011 → SafitIB oggi")

    st.info(
        "**Legenda colori:**  "
        f"🔵 Dati **SAFIT** (pre 3/9/2025)  |  "
        f"🟠 Dati **SafitIB** (post 3/9/2025)  |  "
        f"🟣 Anno **ibrido** 2025",
        icon="ℹ️",
    )

    # --- Carica dati ---
    try:
        df = stor.carica_dataset_unificato()
    except FileNotFoundError as e:
        st.error(
            f"File non trovato: `{e.filename}`\n\n"
            "Verificare i path in `storico_safit.py` → `PATH_STORICO_SAFIT`, "
            "`PATH_CORRENTE_IB`, `PATH_TRANSCODIFICA`."
        )
        return

    kpi_ann = stor.kpi_annuale(df)

    # ── Filtro anni ──────────────────────────────────────────
    anni_disponibili = sorted(df['Anno'].dropna().unique().tolist())
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        anno_min = st.selectbox("Anno da", anni_disponibili,
                                index=0, key="ks_anno_min")
    with col_f2:
        anno_max = st.selectbox("Anno a", anni_disponibili,
                                index=len(anni_disponibili) - 1,
                                key="ks_anno_max")

    mask_ann = (kpi_ann['Anno'] >= anno_min) & (kpi_ann['Anno'] <= anno_max)
    kpi_filt = kpi_ann[mask_ann].copy()
    df_filt  = df[(df['Anno'] >= anno_min) & (df['Anno'] <= anno_max)].copy()

    st.divider()

    # ── Tab principali ───────────────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs([
        "📈 Serie storica",
        "🔍 Per articolo",
        "👥 Per cliente",
        "📋 Per tipo documento",
    ])

    # ────────────────────────────────────────────────────────
    # TAB 1 — Serie storica annuale
    # ────────────────────────────────────────────────────────
    with tab1:
        st.subheader("Valore ordinato annuale (€)")

        fig_val = go.Figure()
        fig_val.add_trace(go.Bar(
            x=kpi_filt['Anno'],
            y=kpi_filt['Valore_Tot'],
            marker_color=_color_bar(kpi_filt['_sorgente']),
            hovertemplate='%{x}: %{y:,.0f} €<extra></extra>',
            name='Valore',
        ))
        fig_val.update_layout(
            xaxis_title="Anno",
            yaxis_title="Valore (€)",
            bargap=0.25,
            height=380,
        )
        st.plotly_chart(fig_val, use_container_width=True)

        st.subheader("Quantità ordinate annuale (pa.)")

        fig_qta = go.Figure()
        fig_qta.add_trace(go.Bar(
            x=kpi_filt['Anno'],
            y=kpi_filt['Qta_Tot'],
            marker_color=_color_bar(kpi_filt['_sorgente']),
            hovertemplate='%{x}: %{y:,.0f} pa.<extra></extra>',
            name='Quantità',
        ))
        fig_qta.update_layout(
            xaxis_title="Anno",
            yaxis_title="Quantità (pa.)",
            bargap=0.25,
            height=380,
        )
        st.plotly_chart(fig_qta, use_container_width=True)

        # Metriche riassuntive
        st.subheader("Riepilogo periodo selezionato")
        c1, c2, c3 = st.columns(3)
        c1.metric("Valore totale", _fmt_euro(kpi_filt['Valore_Tot'].sum()))
        c2.metric("Quantità totale", _fmt_paia(kpi_filt['Qta_Tot'].sum()))
        c3.metric("N° ordini distinti", f"{kpi_filt['N_Ordini'].sum():,}".replace(",", "."))

        st.subheader("Dettaglio annuale")
        st.dataframe(
            kpi_filt[['Anno', 'Valore_Tot', 'Qta_Tot', 'N_Ordini', '_sorgente']]
            .rename(columns={
                'Valore_Tot': 'Valore (€)',
                'Qta_Tot':    'Qtà (pa.)',
                'N_Ordini':   'N° Ordini',
                '_sorgente':  'Fonte',
            })
            .set_index('Anno')
            .style.format({'Valore (€)': '{:,.0f}', 'Qtà (pa.)': '{:,.0f}'}),
            use_container_width=True,
        )

        # Vista unificata vs separata
        st.subheader("Vista separata SAFIT / SafitIB per anno")
        df_sep = df_filt.groupby(['Anno', '_sorgente']).agg(
            Valore_Tot=('Valore', 'sum'),
            Qta_Tot=('Qta Doc', 'sum'),
        ).reset_index()
        fig_sep = px.bar(
            df_sep, x='Anno', y='Valore_Tot', color='_sorgente',
            color_discrete_map=PALETTE,
            barmode='stack',
            labels={'Valore_Tot': 'Valore (€)', '_sorgente': 'Fonte'},
            height=350,
        )
        st.plotly_chart(fig_sep, use_container_width=True)

    # ────────────────────────────────────────────────────────
    # TAB 2 — Per articolo (codice SafitIB)
    # ────────────────────────────────────────────────────────
    with tab2:
        st.subheader("Serie storica per articolo (codice SafitIB)")

        # Seleziona articolo
        articoli_ib = sorted(
            df_filt['Articolo_C_IB'].dropna().unique().tolist()
        )
        art_sel = st.selectbox(
            "Seleziona codice articolo SafitIB",
            articoli_ib,
            key="ks_art",
        )

        if art_sel:
            df_art = df_filt[df_filt['Articolo_C_IB'] == art_sel]
            descr  = df_art['Articolo D'].dropna().iloc[0] if not df_art.empty else ""
            st.caption(f"**{art_sel}** — {descr}")

            kpi_art = df_art.groupby(['Anno', '_sorgente']).agg(
                Valore_Tot=('Valore', 'sum'),
                Qta_Tot=('Qta Doc', 'sum'),
            ).reset_index()

            fig_art = px.bar(
                kpi_art, x='Anno', y='Qta_Tot',
                color='_sorgente',
                color_discrete_map=PALETTE,
                barmode='stack',
                labels={'Qta_Tot': 'Qtà (pa.)', '_sorgente': 'Fonte'},
                title=f"Quantità ordinate — {art_sel}",
                height=350,
            )
            st.plotly_chart(fig_art, use_container_width=True)

            fig_art_v = px.bar(
                kpi_art, x='Anno', y='Valore_Tot',
                color='_sorgente',
                color_discrete_map=PALETTE,
                barmode='stack',
                labels={'Valore_Tot': 'Valore (€)', '_sorgente': 'Fonte'},
                title=f"Valore ordinato — {art_sel}",
                height=350,
            )
            st.plotly_chart(fig_art_v, use_container_width=True)

        # Audit articoli non mappati
        with st.expander("⚠️ Articoli SAFIT senza transcodifica (audit)"):
            df_nm = stor.articoli_non_mappati(df)
            if df_nm.empty:
                st.success("Tutti gli articoli SAFIT sono mappati.")
            else:
                st.warning(
                    f"{len(df_nm)} codici articolo SAFIT non hanno corrispondenza "
                    "nella tabella di transcodifica."
                )
                st.dataframe(
                    df_nm.style.format({'Valore_Tot': '{:,.0f}', 'Qta_Tot': '{:,.0f}'}),
                    use_container_width=True,
                )

    # ────────────────────────────────────────────────────────
    # TAB 3 — Per cliente
    # ────────────────────────────────────────────────────────
    with tab3:
        st.subheader("Serie storica per cliente")

        clienti = sorted(df_filt['Cliente Fornitore CD'].dropna().unique().tolist())
        cliente_sel = st.selectbox("Seleziona cliente", clienti, key="ks_cli")

        if cliente_sel:
            df_cli = df_filt[df_filt['Cliente Fornitore CD'] == cliente_sel]
            kpi_cli = df_cli.groupby(['Anno', '_sorgente']).agg(
                Valore_Tot=('Valore', 'sum'),
                Qta_Tot=('Qta Doc', 'sum'),
                N_Ordini=('Numero Documento', 'nunique'),
            ).reset_index()

            fig_cli = px.bar(
                kpi_cli, x='Anno', y='Valore_Tot',
                color='_sorgente',
                color_discrete_map=PALETTE,
                barmode='stack',
                labels={'Valore_Tot': 'Valore (€)', '_sorgente': 'Fonte'},
                title=f"Valore ordinato — {cliente_sel}",
                height=380,
            )
            st.plotly_chart(fig_cli, use_container_width=True)

            st.dataframe(
                kpi_cli[['Anno', 'Valore_Tot', 'Qta_Tot', 'N_Ordini', '_sorgente']]
                .rename(columns={
                    'Valore_Tot': 'Valore (€)',
                    'Qta_Tot':    'Qtà (pa.)',
                    'N_Ordini':   'N° Ordini',
                    '_sorgente':  'Fonte',
                })
                .set_index('Anno')
                .style.format({'Valore (€)': '{:,.0f}', 'Qtà (pa.)': '{:,.0f}'}),
                use_container_width=True,
            )

        # Top 10 clienti per valore nel periodo
        st.subheader(f"Top 10 clienti per valore ({anno_min}–{anno_max})")
        top10 = (
            df_filt.groupby('Cliente Fornitore CD')['Valore']
            .sum()
            .nlargest(10)
            .reset_index()
        )
        fig_top = px.bar(
            top10, x='Valore', y='Cliente Fornitore CD',
            orientation='h',
            labels={'Valore': 'Valore (€)', 'Cliente Fornitore CD': 'Cliente'},
            height=400,
        )
        fig_top.update_traces(marker_color=COLOR_SAFIT)
        st.plotly_chart(fig_top, use_container_width=True)

    # ────────────────────────────────────────────────────────
    # TAB 4 — Per tipo documento
    # ────────────────────────────────────────────────────────
    with tab4:
        st.subheader("KPI per tipo documento e anno")

        kpi_doc = stor.kpi_per_tipo_doc(df_filt)
        kpi_doc['Tipo'] = kpi_doc['Codice Documento'].map(DOC_LABEL).fillna(kpi_doc['Codice Documento'])

        tipi_sel = st.multiselect(
            "Filtra tipo documento",
            options=sorted(kpi_doc['Codice Documento'].unique().tolist()),
            default=sorted(kpi_doc['Codice Documento'].unique().tolist()),
            key="ks_tipdoc",
        )
        kpi_doc_f = kpi_doc[kpi_doc['Codice Documento'].isin(tipi_sel)]

        fig_doc = px.bar(
            kpi_doc_f, x='Anno', y='Valore_Tot',
            color='Codice Documento',
            barmode='stack',
            labels={'Valore_Tot': 'Valore (€)', 'Codice Documento': 'Tipo doc'},
            height=380,
        )
        st.plotly_chart(fig_doc, use_container_width=True)

        st.dataframe(
            kpi_doc_f[['Anno', 'Tipo', 'Valore_Tot', 'Qta_Tot', 'N_Ordini']]
            .rename(columns={
                'Tipo':      'Documento',
                'Valore_Tot':'Valore (€)',
                'Qta_Tot':   'Qtà (pa.)',
                'N_Ordini':  'N° Ordini',
            })
            .style.format({'Valore (€)': '{:,.0f}', 'Qtà (pa.)': '{:,.0f}'}),
            use_container_width=True,
        )
