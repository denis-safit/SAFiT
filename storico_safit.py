"""
storico_safit.py
================
Modulo per l'integrazione dei dati storici SAFIT (pre-3/9/2025)
con i dati correnti SAFITIB nel portale Safit.

Configurazione path (adattare all'ambiente):
    PATH_STORICO_SAFIT  → export ARCA righe ordini Safit (xlsx)
    PATH_CORRENTE_IB    → export ARCA righe ordini SafitIB (xlsx)
    PATH_TRANSCODIFICA  → tabella transcodifica codici articolo (xlsx)

Funzioni pubbliche:
    carica_dataset_unificato()  → DataFrame unificato Safit + SafitIB
    kpi_annuale()               → KPI aggregati per anno
    kpi_per_articolo_ib()       → KPI per codice articolo SafitIB (serie storica)
    kpi_per_cliente()           → KPI per cliente (serie storica)
    kpi_per_tipo_doc()          → KPI per tipo documento (OCA/OCI/OFF/OFR)
"""

import pandas as pd
import streamlit as st
from pathlib import Path

# ─────────────────────────────────────────────────────────────
# CONFIGURAZIONE PATH — adattare all'ambiente locale
# ─────────────────────────────────────────────────────────────
PATH_STORICO_SAFIT = Path("data/righe_ordini_storico_con_date_SAFIT.xlsx")
PATH_CORRENTE_IB   = Path("data/righe_ordini_storico_con_date.xlsx")
PATH_TRANSCODIFICA = Path("data/transcodifica.xlsx")

# Tipi documento "reali" (escluso totali e righe aggregate)
DOC_REALI = {'OCA', 'OCI', 'OFF', 'OFR', 'OFI', 'OFA'}

# Data di cambio ditta
DATA_CAMBIO = pd.Timestamp("2025-09-03")


# ─────────────────────────────────────────────────────────────
# CARICAMENTO E NORMALIZZAZIONE
# ─────────────────────────────────────────────────────────────

def _ffill_doc(df: pd.DataFrame) -> pd.DataFrame:
    """Applica ffill sulle colonne di intestazione documento (logica ARCA)."""
    for col in ['Codice Documento', 'Data Consegna', 'Numero Documento']:
        if col in df.columns:
            df[col] = df[col].ffill()
    return df


def _normalizza(df: pd.DataFrame) -> pd.DataFrame:
    """Normalizza tipi e filtra righe aggregate."""
    df = _ffill_doc(df.copy())
    df['Data']         = pd.to_datetime(df['Data'], errors='coerce')
    df['Data Consegna'] = pd.to_datetime(df['Data Consegna'], errors='coerce')
    df['Qta Doc']      = pd.to_numeric(df['Qta Doc'], errors='coerce').fillna(0)
    df['Valore']       = pd.to_numeric(df['Valore'], errors='coerce').fillna(0)
    df['Qta Residua']  = pd.to_numeric(df['Qta Residua'], errors='coerce').fillna(0)
    # Rimuovi righe totale/complessivo
    df = df[df['Codice Documento'].isin(DOC_REALI)].copy()
    df['Anno'] = df['Data'].dt.year
    df['Mese'] = df['Data'].dt.to_period('M').astype(str)
    return df


@st.cache_data(show_spinner="Caricamento dati storici SAFIT…")
def _carica_transcodifica() -> dict:
    """Carica dizionario {codice_vecchio: codice_nuovo}."""
    df = pd.read_excel(PATH_TRANSCODIFICA)
    return dict(zip(df['Cd_AR_Vecchio'].astype(str),
                    df['Cd_AR_Nuovo'].astype(str)))


@st.cache_data(show_spinner="Unificazione dataset storico…")
def carica_dataset_unificato() -> pd.DataFrame:
    """
    Carica e unifica i dati SAFIT (storici) e SAFITIB (correnti).

    Colonne aggiuntive nel DataFrame risultante:
        _sorgente       : 'SAFIT' o 'SAFITIB'
        Articolo_C_IB   : codice articolo normalizzato in formato SafitIB
        Anno            : anno della riga
        Mese            : periodo mese (es. '2024-03')

    Returns
    -------
    pd.DataFrame
    """
    trans = _carica_transcodifica()

    # --- Storico SAFIT ---
    df_safit = _normalizza(pd.read_excel(PATH_STORICO_SAFIT))
    df_safit['_sorgente']    = 'SAFIT'
    df_safit['Articolo_C_IB'] = df_safit['Articolo C'].astype(str).map(trans)

    # --- Corrente SAFITIB ---
    df_ib = _normalizza(pd.read_excel(PATH_CORRENTE_IB))
    df_ib['_sorgente']    = 'SAFITIB'
    df_ib['Articolo_C_IB'] = df_ib['Articolo C'].astype(str)

    df = pd.concat([df_safit, df_ib], ignore_index=True)
    return df


# ─────────────────────────────────────────────────────────────
# KPI AGGREGATI
# ─────────────────────────────────────────────────────────────

def kpi_annuale(df: pd.DataFrame | None = None) -> pd.DataFrame:
    """
    KPI per anno: Valore totale, Quantità totale, N° ordini distinti.
    Include colonna '_sorgente_prevalente' per sapere se l'anno è
    prevalentemente SAFIT o SAFITIB.
    """
    if df is None:
        df = carica_dataset_unificato()

    grp = df.groupby('Anno').agg(
        Valore_Tot  = ('Valore',           'sum'),
        Qta_Tot     = ('Qta Doc',          'sum'),
        N_Ordini    = ('Numero Documento', 'nunique'),
        N_Righe     = ('Articolo C',       'count'),
    ).reset_index()

    # Flag sorgente per colorazione nei grafici
    grp['_sorgente'] = grp['Anno'].apply(
        lambda a: 'SAFITIB' if a >= DATA_CAMBIO.year else 'SAFIT'
    )
    # 2025 è ibrido: annotazione speciale
    grp.loc[grp['Anno'] == DATA_CAMBIO.year, '_sorgente'] = 'IBRIDO'

    return grp.sort_values('Anno')


def kpi_per_articolo_ib(
    df: pd.DataFrame | None = None,
    codice_ib: str | None = None,
    anni: list[int] | None = None,
) -> pd.DataFrame:
    """
    Serie storica annuale per un codice articolo SafitIB.
    Se codice_ib è None, restituisce tutti gli articoli aggregati per anno.
    """
    if df is None:
        df = carica_dataset_unificato()
    if codice_ib:
        df = df[df['Articolo_C_IB'] == codice_ib]
    if anni:
        df = df[df['Anno'].isin(anni)]

    return df.groupby(['Anno', 'Articolo_C_IB']).agg(
        Valore_Tot = ('Valore',  'sum'),
        Qta_Tot    = ('Qta Doc', 'sum'),
        Descrizione = ('Articolo D', 'first'),
    ).reset_index().sort_values(['Articolo_C_IB', 'Anno'])


def kpi_per_cliente(
    df: pd.DataFrame | None = None,
    codice_cliente: str | None = None,
) -> pd.DataFrame:
    """Serie storica annuale per cliente."""
    if df is None:
        df = carica_dataset_unificato()
    if codice_cliente:
        df = df[df['Cliente Fornitore CD'].str.startswith(codice_cliente, na=False)]

    return df.groupby(['Anno', 'Cliente Fornitore CD']).agg(
        Valore_Tot = ('Valore',  'sum'),
        Qta_Tot    = ('Qta Doc', 'sum'),
    ).reset_index().sort_values(['Cliente Fornitore CD', 'Anno'])


def kpi_per_tipo_doc(df: pd.DataFrame | None = None) -> pd.DataFrame:
    """KPI per tipo documento e anno."""
    if df is None:
        df = carica_dataset_unificato()
    return df.groupby(['Anno', 'Codice Documento']).agg(
        Valore_Tot = ('Valore',  'sum'),
        Qta_Tot    = ('Qta Doc', 'sum'),
        N_Ordini   = ('Numero Documento', 'nunique'),
    ).reset_index().sort_values(['Anno', 'Codice Documento'])


def articoli_non_mappati(df: pd.DataFrame | None = None) -> pd.DataFrame:
    """
    Restituisce le righe SAFIT il cui codice articolo NON ha transcodifica.
    Utile per audit / completamento della tabella di transcodifica.
    """
    if df is None:
        df = carica_dataset_unificato()
    mask = (df['_sorgente'] == 'SAFIT') & (df['Articolo_C_IB'].isna())
    return df[mask][['Articolo C', 'Articolo D', 'Anno', 'Valore', 'Qta Doc']]\
        .groupby(['Articolo C', 'Articolo D']).agg(
            Anni        = ('Anno',    lambda x: sorted(x.unique().tolist())),
            Valore_Tot  = ('Valore',  'sum'),
            Qta_Tot     = ('Qta Doc', 'sum'),
        ).reset_index().sort_values('Valore_Tot', ascending=False)
