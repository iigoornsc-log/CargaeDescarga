import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import json
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re
import datetime

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Dashboard Executivo | Magalu Logística", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# --- 2. CSS DE ALTO NÍVEL (UI/UX PREMIUM) ---
st.markdown("""
    <style>
    .stApp { background-color: #F4F7F9; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #E5E7EB; }
    
    /* Cartões de KPI Padrão */
    .kpi-card { background-color: #FFFFFF; border-radius: 12px; padding: 15px; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05); text-align: center; border-bottom: 4px solid #0086FF; margin-bottom: 15px; height: 100%; }
    .kpi-card.danger { border-bottom: 4px solid #FF3366; }
    .kpi-card.success { border-bottom: 4px solid #00C853; }
    
    /* Cartões VIP (Oportunidade FULL) */
    .kpi-card.gold { border-bottom: 4px solid #FFD700; background: linear-gradient(145deg, #ffffff, #fffdf0); }
    .kpi-card.gold .kpi-value { color: #B8860B; }
    
    .kpi-title { color: #6B7280; font-size: 13px; font-weight: 700; text-transform: uppercase; margin-bottom: 5px; }
    .kpi-value { color: #111827; font-size: 22px; font-weight: 800; }
    
    /* Títulos e Containers */
    .main-title { color: #111827; font-size: 32px; font-weight: 800; margin-bottom: 5px; }
    .chart-container { background-color: #FFFFFF; border-radius: 16px; padding: 20px; box-shadow: 0 4px 20px rgba(0, 0, 0, 0.03); margin-top: 15px; margin-bottom: 15px; }
    .chart-container.gold-container { border: 1px solid #FFD700; box-shadow: 0 4px 20px rgba(218, 165, 32, 0.15); }
    
    .stDataFrame { border-radius: 10px; overflow: hidden; }
    </style>
""", unsafe_allow_html=True)

# --- 3. CONEXÃO E CACHE ---
def conectar_google():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        cred_dict = json.loads(st.secrets["google_json"])
        creds = Credentials.from_service_account_info(cred_dict, scopes=scopes)
    except Exception:
        caminho_local = r'C:\Users\IIGOORNSC\Documents\CargaDescarga\credential_key.json'
        creds = Credentials.from_service_account_file(caminho_local, scopes=scopes)
    return gspread.authorize(creds)

@st.cache_data(ttl=600)
def carregar_dados():
    client = conectar_google()
    sh = client.open_by_key("1NWH9BHXgUmS-6WCQ8AjAHbt8DUHIvgQLRJ8hwUSDC7U")
    ws_historico = sh.worksheet("HISTÓRICO 2025")
    dados_h = ws_historico.get_all_values()
    df_h = pd.DataFrame(dados_h[1:], columns=dados_h[0])
    return df_h

# --- 4. TRATAMENTO SÊNIOR ---
def formatar_moeda_br(valor):
    if pd.isna(valor) or valor == 0: return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def tratar_dados(df_h):
    # Padronização Inicial e Desduplicação
    df_h['STATUS'] = df_h['STATUS'].astype(str).str.strip().str.upper()
    df_h['AGENDA WMS'] = df_h['AGENDA WMS'].astype(str).str.strip().str.upper()
    
    df_h['DATA AGENDA'] = df_h['DATA AGENDA'].astype(str).str.strip()
    df_h['DATA AGENDADA'] = pd.to_datetime(df_h['DATA AGENDA'], dayfirst=True, errors='coerce')
    df_h = df_h.dropna(subset=['DATA AGENDADA']).copy()
    
    df_h['PRIORIDADE_STATUS'] = df_h['STATUS'].apply(lambda x: 1 if x == 'OK' else (0 if x == 'AUSENTE' else -1))
    df_h = df_h.sort_values(by=['AGENDA WMS', 'PRIORIDADE_STATUS', 'DATA AGENDADA'], ascending=[True, False, False])
    
    mask_wms_valido = ~df_h['AGENDA WMS'].isin(['', '-', 'NAN', 'NONE'])
    df_com_wms = df_h[mask_wms_valido].drop_duplicates(subset=['AGENDA WMS'], keep='first')
    df_sem_wms = df_h[~mask_wms_valido]
    
    df_h = pd.concat([df_com_wms, df_sem_wms], ignore_index=True)
    df_h = df_h.drop(columns=['PRIORIDADE_STATUS'])

    # Datas
    df_h['ANO'] = df_h['DATA AGENDADA'].dt.year.astype('Int64')
    df_h['MES_ORDENACAO'] = df_h['DATA AGENDADA'].dt.to_period('M')
    meses_pt = {1:'Jan', 2:'Fev', 3:'Mar', 4:'Abr', 5:'Mai', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Set', 10:'Out', 11:'Nov', 12:'Dez'}
    df_h['MES_NOME'] = df_h['DATA AGENDADA'].dt.month.map(meses_pt) + "/" + df_h['ANO'].astype(str)

    # Tratamento de Conferentes (Para a Folha de Pagamento)
    if 'CONFERENTE' in df_h.columns:
        df_h['CONFERENTE'] = df_h['CONFERENTE'].astype(str).str.strip().str.title()
        # Apaga linhas com "Tracinho" ou vazias para não contar como pessoa
        df_h.loc[df_h['CONFERENTE'].isin(['-', '', 'Nan', 'None']), 'CONFERENTE'] = None

    # Limpeza Moeda
    def limpar_moeda(valor):
        if pd.isna(valor) or str(valor).strip() == '': return 0.0
        v = str(valor).upper().replace('R$', '').replace(' ', '').strip()
        if v in ['', '-', 'OK', 'NAOÉCOBRADO']: return 0.0
        if '.' in v and ',' in v: v = v.replace('.', '').replace(',', '.')
        elif '.' in v: v = v.replace('.', '')
        elif ',' in v: v = v.replace(',', '.')
        try: return round(float(v), 2)
        except ValueError: return 0.0

    df_h['VALOR_REAL'] = df_h['VALOR'].apply(limpar_moeda) if 'VALOR' in df_h.columns else 0.0
    
    # Textos
    df_h['LINHA'] = df_h['LINHA'].astype(str).str.strip().str.upper()
    df_h['FORNECEDOR/SELLER'] = df_h['FORNECEDOR/SELLER'].astype(str).str.strip().str.upper()
    df_h['CATEGORIA'] = df_h['CATEGORIA'].astype(str).str.strip().str.upper()

    # ==========================================
    # 1º PASSO: FILTRO GERAL DE OPERAÇÕES INTERNAS (CDs, Filiais, etc)
    # ==========================================
    def eh_interno(forn):
        if forn.startswith(('MAGAZINE', 'FILIAL')): return True
        if re.match
