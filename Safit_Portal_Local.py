import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# Mostra subito se il file è quello nuovo
st.error("⚠️ TEST DI SINCRONIZZAZIONE ATTIVO ⚠️")

# --- 1. CONFIGURAZIONE PAGINA E VERSIONE ---
APP_VERSION = "1.0.02"
st.set_page_config(page_title=f"Safit - Portale Avanzamento {APP_VERSION}", layout="wide")

st.markdown(f"""
    <style>
    .main {{ background-color: #fcfcfc; }}
    .stApp {{ margin-top: -30px; }}
    .status-row {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        flex-wrap: wrap; 
        padding: 12px 15px;
        border-radius: 8px;
        margin-bottom: 8px;
        gap: 10px;
        font-size: 14px;
    }}
    .delay-row {{ background-color: #fff8e1; border-left: 6px solid #ffc107; color: #5d4037; }}
    .on-time-row {{ background-color: #f1f8e9; border-left: 6px solid #4caf50; color: #1b5e20; }}
    .client-delay-row {{ background-color: #e3f2fd; border-left: 6px solid #2196f3; color: #0d47a1; }}
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
            st.title("Accesso Area Ris
