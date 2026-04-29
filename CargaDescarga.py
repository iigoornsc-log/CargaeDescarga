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
import time
from datetime import date
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io

# ==========================================================
# 1. CONFIGURAÇÃO DA PÁGINA E CSS (THEME MAGALOG CORPORATIVO)
# ==========================================================
st.set_page_config(
    page_title="RITMO| MAGALOG", 
    page_icon="https://play-lh.googleusercontent.com/WogWYMVkkivEHGyHAZvKtFZ4F3mklNQI-PQ6vsOMdKTSZWqr7etD9XHuPKIY0NkzZqk=w240-h480-rw", # Pode ser qualquer link de imagem
    layout="wide", 
    initial_sidebar_state="expanded"
)



st.markdown("""
    <style>
    /* 1. FONTE PREMIUM E ÍCONES GOOGLE MATERIAL */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
    @import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Rounded:opsz,wght,FILL,GRAD@24,600,1,0');
    
    * { font-family: 'Inter', sans-serif !important; }

    /* Classe para alinhar os ícones no HTML perfeitamente com o texto */
    .icon-MAGALOG {
        font-family: 'Material Symbols Rounded' !important;
        font-variation-settings: 'FILL' 1, 'wght' 600, 'GRAD' 0, 'opsz' 24;
        font-weight: normal;
        font-style: normal;
        letter-spacing: normal;
        text-transform: none;
        white-space: nowrap;
        direction: ltr;
        -webkit-font-smoothing: antialiased;
        vertical-align: middle;
        display: inline-block;
        line-height: 1;
        font-size: inherit;
    }

    /* --- CORREÇÃO DOS ÍCONES DA SIDEBAR E SISTEMA --- */
    .material-icons, .material-symbols-rounded, [data-testid="stSidebarCollapseButton"] * {
        font-family: 'Material Symbols Rounded', 'Material Icons', sans-serif !important;
    }

    /* 2. ANIMAÇÃO RGB LUIZALABS */
    @keyframes Glow {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    /* Linha Tech Animada no Topo da Tela */
    .stApp::before {
        content: ""; position: fixed; top: 0; left: 0; right: 0; height: 5px;
        background: linear-gradient(90deg, #0086FF, #FF007F, #00C853, #0086FF, #FF007F);
        background-size: 300% 300%;
        animation: MAGALOGGlow 6s linear infinite;
        z-index: 999999;
    }

    /* Fundo da Aplicação (Soft com micro-gradiente) */
    .stApp {
        background-color: #F0F4F8;
        background-image: radial-gradient(circle at 100% 0%, #E2EDF8 0%, transparent 40%);
    }

    /* 3. TÍTULOS COM DEGRADÊ METÁLICO */
    .MAGALOG-page-title { 
        background: linear-gradient(135deg, #0086FF 0%, #001A57 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 32px; font-weight: 900; letter-spacing: -1px; margin-bottom: 2px;
    }
    .MAGALOG-page-subtitle { color: #64748B; font-size: 15px; font-weight: 500; margin-bottom: 25px; }

    /* 4. ABAS (TABS) CORPORATIVAS */
    [data-baseweb="tab-list"] {
        background-color: #FFFFFF;
        border-radius: 12px;
        padding: 6px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.04);
        gap: 8px;
        margin-bottom: 25px;
    }
    button[data-baseweb="tab"] {
        background-color: transparent !important;
        border-radius: 8px !important;
        padding: 12px 24px !important;
        color: #64748B !important;
        font-weight: 600 !important;
        border: none !important;
        transition: all 0.3s ease !important;
    }
    button[data-baseweb="tab"]:hover {
        background-color: #F1F5F9 !important;
        color: #0F172A !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        background-color: #0086FF !important;
        color: #FFFFFF !important;
        box-shadow: 0 4px 15px rgba(0,134,255,0.35) !important;
    }
    [data-baseweb="tab-border"] { display: none; }

    /* 5. RIBBONS ANIMADOS */
    .MAGALOG-ribbon {
        background: linear-gradient(90deg, #0086FF, #005BFF, #FF007F, #0086FF);
        background-size: 300% 300%;
        animation: MAGALOGGlow 8s ease infinite;
        color: #FFFFFF; padding: 8px 24px; font-size: 13px; font-weight: 700;
        border-radius: 0px 8px 8px 0px; margin-bottom: 15px; margin-top: 10px;
        position: relative; left: -1rem; box-shadow: 0 4px 15px rgba(0,134,255,0.3);
        text-transform: uppercase; letter-spacing: 1px;
    }

    /* 6. CARDS GLASSMORPHISM */
    .MAGALOG-card, div[data-testid="stVerticalBlock"] > div > div[data-testid="stVerticalBlockBorderWrapper"] {
        background: rgba(255, 255, 255, 0.95) !important;
        backdrop-filter: blur(10px) !important;
        border: 1px solid rgba(255,255,255,0.6) !important;
        border-radius: 16px !important;
        box-shadow: 0 8px 30px rgba(0, 0, 0, 0.03) !important;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
        padding: 20px !important;
    }
    div[data-testid="stVerticalBlockBorderWrapper"]:hover {
        transform: translateY(-3px);
        box-shadow: 0 12px 40px rgba(0, 134, 255, 0.08) !important;
    }

    /* DATE INPUT SIDEBAR */
    section[data-testid="stSidebar"] div[data-baseweb="input"] > div {
        background: #FFFFFF !important; border-radius: 12px !important;
        border: 1px solid #E2E8F0 !important; min-height: 48px !important; transition: all 0.3s ease !important;
    }
    section[data-testid="stSidebar"] div[data-baseweb="input"] > div:focus-within {
        border-color: #0086FF !important; box-shadow: 0 0 0 3px rgba(0,134,255,0.15) !important;
    }
    section[data-testid="stSidebar"] input { color: #0F172A !important; font-weight: 600 !important; }
    section[data-testid="stSidebar"] svg { color: #0086FF !important; }

    /* CHIPS MULTISELECT */
    span[data-baseweb="tag"] {
        background-color: #E6F2FF !important; color: #0086FF !important;
        border-radius: 8px !important; border: 1px solid #BAE6FD !important;
        font-weight: 700 !important; padding: 6px 12px !important; margin: 4px 4px 4px 0px !important;
    }
    span[data-baseweb="tag"] svg { fill: #0086FF !important; } 

    /* BOTÃO SIDEBAR */
    section[data-testid="stSidebar"] .stButton>button {
        background: linear-gradient(135deg, #0086FF 0%, #005BFF 100%); color: #FFFFFF;
        border: none; border-radius: 12px; font-weight: 700; font-size: 14px;
        padding: 0.8rem 1rem; box-shadow: 0 6px 20px rgba(0,134,255,0.25); transition: all 0.3s ease;
    }
    section[data-testid="stSidebar"] .stButton>button:hover { transform: translateY(-2px); box-shadow: 0 10px 25px rgba(0,134,255,0.4); }
    section[data-testid="stSidebar"] .stButton>button:active { transform: scale(0.97); }
    
    /* BOTÃO PRIMARY (Verde) */
    button[kind="primary"] {
        background: linear-gradient(135deg, #00C853 0%, #009624 100%) !important; border: none !important; color: white !important; font-weight: 800 !important;
        border-radius: 10px !important; box-shadow: 0 6px 20px rgba(0,200,83,0.3) !important; transition: all 0.3s ease !important;
    }
    button[kind="primary"]:hover { box-shadow: 0 8px 25px rgba(0,200,83,0.5) !important; transform: translateY(-2px) !important; }

    /* SCROLLBAR */
    ::-webkit-scrollbar { width: 8px; height: 8px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 10px; }
    ::-webkit-scrollbar-thumb:hover { background: #94A3B8; }

    /* KPI CARDS */
    .kpi-card { 
        background: #FFFFFF; border-radius: 16px; padding: 20px 15px; 
        border: 1px solid #F1F5F9; box-shadow: 0 4px 20px rgba(0, 0, 0, 0.03);
        margin-bottom: 12px; display: flex; flex-direction: column; 
        align-items: center; justify-content: center; transition: all 0.3s ease;
    }
    .kpi-card:hover { transform: translateY(-4px); box-shadow: 0 12px 30px rgba(0, 134, 255, 0.08); }
    .kpi-title { color: #64748B; font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px;}
    .kpi-value { color: #0F172A; font-size: 24px; font-weight: 900; letter-spacing: -0.5px; }

    /* SIDEBAR E DASHBOARD PREMIUM */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #082a63 0%, #00153d 100%) !important;
        border-right: 1px solid rgba(255,255,255,0.08);
    }
    section[data-testid="stSidebar"] * { color: #EAF2FF; }
    section[data-testid="stSidebar"] .stRadio > div { gap: 8px; }
    section[data-testid="stSidebar"] label[data-baseweb="radio"] {
        background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.08);
        border-radius: 14px; padding: 10px 12px; margin-bottom: 8px; transition: all 0.25s ease;
    }
    section[data-testid="stSidebar"] label[data-baseweb="radio"]:hover {
        background: rgba(255,255,255,0.12); border-color: rgba(85,170,255,0.55);
    }
    .MAGALOG-shell {
        background: linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(247,250,255,0.98) 100%);
        border: 1px solid rgba(255,255,255,0.7); border-radius: 24px; box-shadow: 0 20px 60px rgba(15, 23, 42, 0.08);
        padding: 28px; margin-bottom: 22px;
    }
    .MAGALOG-hero {
        position: relative; overflow: hidden; background: linear-gradient(135deg, #0A4FB3 0%, #062B76 45%, #0D1836 100%);
        border-radius: 28px; padding: 34px 34px 26px 34px; color: #FFFFFF;
        box-shadow: 0 24px 70px rgba(0, 74, 173, 0.25); margin-bottom: 24px;
    }
    .MAGALOG-hero::after {
        content: ""; position: absolute; inset: auto -80px -90px auto; width: 320px; height: 320px;
        background: radial-gradient(circle, rgba(0,255,255,0.22) 0%, rgba(0,255,255,0) 68%);
    }
    .MAGALOG-hero::before {
        content: ""; position: absolute; inset: 0; background: linear-gradient(90deg, rgba(255,255,255,0.05), rgba(255,255,255,0)); pointer-events: none;
    }
    .MAGALOG-hero-badge {
        display: inline-block; background: rgba(255,255,255,0.12); border: 1px solid rgba(255,255,255,0.15);
        border-radius: 999px; padding: 8px 14px; font-size: 12px; font-weight: 800; letter-spacing: .08em;
        margin-bottom: 16px; text-transform: uppercase; backdrop-filter: blur(8px);
    }
    .MAGALOG-hero-title {
        font-size: 40px; font-weight: 900; line-height: 1.02; letter-spacing: -1.2px;
        margin-bottom: 10px; position: relative; z-index: 2;
    }
    .MAGALOG-hero-subtitle {
        color: rgba(255,255,255,0.82); font-size: 16px; max-width: 860px; position: relative; z-index: 2;
    }
    .MAGALOG-grid {
        display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin: 22px 0;
    }
    .MAGALOG-feature-card {
        background: linear-gradient(180deg, #ffffff 0%, #f6f9ff 100%); border: 1px solid #E5EEF9;
        border-radius: 22px; padding: 22px; box-shadow: 0 12px 35px rgba(15,23,42,0.06); min-height: 150px;
    }
    .MAGALOG-feature-icon {
        width: 52px; height: 52px; border-radius: 16px; display:flex; align-items:center;
        justify-content:center; font-size: 24px; margin-bottom: 14px;
        background: linear-gradient(135deg, #E9F3FF 0%, #DDEBFF 100%); color: #0A4FB3;
    }
    .MAGALOG-feature-title { color: #0F172A; font-size: 18px; font-weight: 800; margin-bottom: 6px; }
    .MAGALOG-feature-text { color: #64748B; font-size: 14px; line-height: 1.5; }
    .MAGALOG-info-strip {
        background: linear-gradient(90deg, #0A57C9 0%, #0094FF 55%, #FF6B3D 100%); border-radius: 18px;
        padding: 16px 20px; color: #fff; font-weight: 800; letter-spacing: .04em; text-transform: uppercase;
        box-shadow: 0 14px 40px rgba(0,134,255,0.18); margin-top: 10px;
    }
    .MAGALOG-mini-card {
        background: rgba(255,255,255,0.94); border: 1px solid #E6EDF7; border-radius: 18px;
        padding: 18px; box-shadow: 0 10px 30px rgba(15,23,42,0.05); height: 100%;
    }
    .MAGALOG-mini-label { color: #64748B; font-size: 11px; text-transform: uppercase; font-weight: 800; letter-spacing: .06em; margin-bottom: 6px; }
    .MAGALOG-mini-value { color: #0F172A; font-size: 28px; font-weight: 900; letter-spacing: -0.8px; margin-bottom: 4px; }
    .MAGALOG-mini-desc { color: #64748B; font-size: 13px; }
    </style>
""", unsafe_allow_html=True)

# ==========================================================
# 2. CONEXÃO GOOGLE SHEETS & CACHES (BLINDADO CONTRA API ERROR)
# ==========================================================
def conectar_google():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        cred_dict = json.loads(st.secrets["google_json"])
        creds = Credentials.from_service_account_info(cred_dict, scopes=scopes)
    except:
        creds = Credentials.from_service_account_file(r'C:\Users\IIGOORNSC\Documents\CargaDescarga\credential_key.json', scopes=scopes)
    return gspread.authorize(creds)

@st.cache_data(ttl=600)
def carregar_dados_financeiros():
    for tentativa in range(3):
        try:
            client = conectar_google()
            sh = client.open_by_key("1NWH9BHXgUmS-6WCQ8AjAHbt8DUHIvgQLRJ8hwUSDC7U")
            ws = sh.worksheet("HISTÓRICO 2025")
            return pd.DataFrame(ws.get_all_values()[1:], columns=ws.get_all_values()[0])
        except Exception as e:
            if tentativa == 2: return pd.DataFrame()
            time.sleep(1.5)

@st.cache_data(ttl=60)
def carregar_equipe():
    for tentativa in range(3):
        try:
            client = conectar_google()
            sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
            ws = sh.worksheet("QUADRO CARGA e DESCARGA")
            return pd.DataFrame(ws.get_all_records())
        except Exception as e:
            if tentativa == 2: return pd.DataFrame()
            time.sleep(1.5)

@st.cache_data(ttl=10)
def carregar_log_produtividade():
    for tentativa in range(3):
        try:
            client = conectar_google()
            sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
            ws = sh.worksheet("LOG_PRODUTIVIDADE")
            return pd.DataFrame(ws.get_all_records())
        except Exception as e:
            if tentativa == 2: return pd.DataFrame()
            time.sleep(1.5)

@st.cache_data(ttl=60)
def carregar_aux():
    for tentativa in range(3):
        try:
            client = conectar_google()
            sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
            ws = sh.worksheet("aux")
            return pd.DataFrame(ws.get_all_records())
        except Exception as e:
            if tentativa == 2: return pd.DataFrame()
            time.sleep(1.5)

@st.cache_data(ttl=60)
def carregar_matriz():
    for tentativa in range(3):
        try:
            client = conectar_google()
            sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
            ws = sh.worksheet("MATRIZ_COMPETÊNCIA")
            return pd.DataFrame(ws.get_all_records())
        except Exception as e:
            if tentativa == 2: return pd.DataFrame()
            time.sleep(1.5)

@st.cache_data(ttl=60)
def carregar_docas_finalizadas():
    # A SOLUÇÃO DEFINITIVA DO M³ AQUI: Lê como string pura sem o Google sumir com a vírgula!
    for tentativa in range(3):
        try:
            client = conectar_google()
            sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
            ws = sh.worksheet("DOCAS_FINALIZADAS") 
            dados = ws.get_all_values()
            if len(dados) > 1:
                df = pd.DataFrame(dados[1:], columns=dados[0])
            else:
                df = pd.DataFrame(columns=dados[0] if dados else [])
            return df
        except Exception as e:
            if tentativa == 2: return pd.DataFrame()
            time.sleep(1.5)

@st.cache_data(ttl=60)
def carregar_auxexp():
    for tentativa in range(3):
        try:
            client = conectar_google()
            sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
            ws = sh.worksheet("auxexp")
            return pd.DataFrame(ws.get_all_records())
        except Exception as e:
            if tentativa == 2: return pd.DataFrame()
            time.sleep(1.5)

@st.cache_data(ttl=60)
def carregar_log_absenteismo():
    for tentativa in range(3):
        try:
            client = conectar_google()
            sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
            ws = sh.worksheet("LOG_ABSENTEISMO")
            dados = ws.get_all_values()
            if len(dados) > 1:
                df = pd.DataFrame(dados[1:], columns=dados[0])
                df.columns = df.columns.str.strip().str.upper()
                return df
            else:
                return pd.DataFrame(columns=['DATA', 'ID', 'NOME', 'OCORRÊNCIA', 'TURNO'])
        except Exception as e:
            if tentativa == 2: return pd.DataFrame()
            time.sleep(1.5)

@st.cache_data(ttl=60)
def carregar_historico_escalas():
    for tentativa in range(3):
        try:
            client = conectar_google()
            sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
            ws = sh.worksheet("HISTORICO_ESCALAS")
            dados = ws.get_all_values()
            if len(dados) > 1:
                df = pd.DataFrame(dados[1:], columns=dados[0])
                df.columns = df.columns.str.strip().str.upper()
                return df
            return pd.DataFrame(columns=['DATA', 'TURNO', 'EQUIPE', 'MEMBROS'])
        except Exception as e:
            if tentativa == 2: return pd.DataFrame()
            time.sleep(1.5)


from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io

import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import io

def salvar_recibo_no_drive(html_content, agenda_numero):
    try:
        escopos = ['https://www.googleapis.com/auth/drive']
        google_json = st.secrets["google_json"]
        chaves = json.loads(google_json)

        credenciais = Credentials.from_service_account_info(chaves, scopes=escopos)
        drive_service = build('drive', 'v3', credentials=credenciais)

        # Cria como Google Docs (não HTML puro)
        file_metadata = {
            'name': f'Recibo_OS_{agenda_numero}',
            'parents': ['0AI6HL7wIJQwZUk9PVA'],
            'mimeType': 'application/vnd.google-apps.document'
        }

        media = MediaIoBaseUpload(
            io.BytesIO(html_content.encode('utf-8')),
            mimetype='text/html',
            resumable=True
        )

        arquivo = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id',
            supportsAllDrives=True
        ).execute()

        file_id = arquivo.get('id')

        # 🔥 Agora exporta como PDF
        pdf = drive_service.files().export(
            fileId=file_id,
            mimeType='application/pdf'
        ).execute()

        # 🔥 Salva o PDF final no Drive
        file_metadata_pdf = {
            'name': f'Recibo_OS_{agenda_numero}.pdf',
            'parents': ['0AI6HL7wIJQwZUk9PVA']
        }

        media_pdf = MediaIoBaseUpload(
            io.BytesIO(pdf),
            mimetype='application/pdf',
            resumable=True
        )

        arquivo_pdf = drive_service.files().create(
            body=file_metadata_pdf,
            media_body=media_pdf,
            fields='id',
            supportsAllDrives=True
        ).execute()

        return arquivo_pdf.get('id')

    except Exception as e:
        st.error(f"Erro ao salvar no Drive: {e}")
        return None


def salvar_escala_ia(linhas_para_salvar):
    try:
        client = conectar_google()
        sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
        try:
            ws = sh.worksheet("HISTORICO_ESCALAS")
        except:
            ws = sh.add_worksheet(title="HISTORICO_ESCALAS", rows="100", cols="4")
            ws.append_row(["DATA", "TURNO", "EQUIPE", "MEMBROS"])
        ws.append_rows(linhas_para_salvar)
        return True
    except Exception as e:
        return False

@st.cache_data(ttl=60)
def carregar_dados_cobranca():
    for tentativa in range(3):
        try:
            client = conectar_google()
            sh = client.open_by_key("1NWH9BHXgUmS-6WCQ8AjAHbt8DUHIvgQLRJ8hwUSDC7U")
            
            # Base Principal
            dados_painel = sh.worksheet("Painel de Controle").get_all_values()
            df_painel = pd.DataFrame(dados_painel[1:], columns=dados_painel[0]) if len(dados_painel) > 1 else pd.DataFrame()
            
            # Base LG (Motoristas e Veículos)
            dados_lg = sh.worksheet("Base LG").get_all_values()
            df_lg = pd.DataFrame(dados_lg[1:], columns=dados_lg[0]) if len(dados_lg) > 1 else pd.DataFrame()
            
            return df_painel, df_lg
        except Exception as e:
            if tentativa == 2: return pd.DataFrame(), pd.DataFrame()
            time.sleep(1.5)

def atualizar_status_pagamento(agenda, novo_status):
    try:
        client = conectar_google()
        sh = client.open_by_key("1NWH9BHXgUmS-6WCQ8AjAHbt8DUHIvgQLRJ8hwUSDC7U")
        ws = sh.worksheet("Painel de Controle")
        
        # Encontra a linha da agenda exata
        cell = ws.find(str(agenda).strip())
        if cell:
            header = ws.row_values(1)
            # Cria a coluna STATUS PAGAMENTO caso não exista
            if 'STATUS PAGAMENTO' in header:
                col_idx = header.index('STATUS PAGAMENTO') + 1
            else:
                col_idx = len(header) + 1
                ws.update_cell(1, col_idx, 'STATUS PAGAMENTO')
            
            # Atualiza o status
            ws.update_cell(cell.row, col_idx, novo_status)
            return True
        return False
    except Exception as e:
        return False

@st.cache_data(ttl=60)
def carregar_transportadoras_local():
    for tentativa in range(3):
        try:
            client = conectar_google()
            sh = client.open_by_key("1NWH9BHXgUmS-6WCQ8AjAHbt8DUHIvgQLRJ8hwUSDC7U")
            dados = sh.worksheet("Transportadoras").get_all_values()
            if len(dados) > 1:
                df = pd.DataFrame(dados[1:], columns=dados[0])
                df.columns = df.columns.astype(str).str.strip().str.upper()
                return df
            return pd.DataFrame()
        except:
            if tentativa == 2: return pd.DataFrame()
            time.sleep(1.5)



# ==========================================================
# 3. FUNÇÕES DE GRAVAÇÃO (BACK-END)
# ==========================================================
def gravar_absenteismo(dados_para_gravar):
    client = conectar_google()
    sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
    try:
        ws_log = sh.worksheet("LOG_ABSENTEISMO")
        ws_log.append_rows(dados_para_gravar)
        return True
    except:
        st.error("Erro: Aba 'LOG_ABSENTEISMO' não encontrada.")
        return False

def gravar_conclusao_doca(linhas_conclusao, linha_encerramento_log):
    client = conectar_google()
    sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")

    def normalizar_valor(v):
        if pd.isna(v):
            return ""
        if isinstance(v, (pd.Timestamp, datetime.datetime)):
            return v.strftime("%d/%m/%Y %H:%M:%S")
        if isinstance(v, (datetime.date,)):
            return v.strftime("%d/%m/%Y")
        if hasattr(v, "item"):  # numpy.int64, numpy.float64, etc.
            v = v.item()
        return v

    def normalizar_linha(linha):
        return [normalizar_valor(x) for x in linha]

    try:
        linhas_conclusao_ok = [normalizar_linha(linha) for linha in linhas_conclusao]
        linha_encerramento_ok = normalizar_linha(linha_encerramento_log)

        ws_final = sh.worksheet("DOCAS_FINALIZADAS")
        ws_final.append_rows(linhas_conclusao_ok)

        ws_log = sh.worksheet("LOG_PRODUTIVIDADE")
        ws_log.append_rows([linha_encerramento_ok])

        return True
    except Exception as e:
        st.error(f"Erro ao finalizar doca: {e}")
        return False

def processar_gravacao_doca(doca_sel, agenda_sel, conferente_sel, equipe_sel, conflitos, info_docas, encerrar):
    agora_dt = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
    agora_str = agora_dt.strftime("%d/%m/%Y %H:%M:%S")
    linhas_para_gravar = []
    
    agenda_final = agenda_sel if agenda_sel else "-"
    conferente_final = conferente_sel if conferente_sel else "-"

    if encerrar:
        linhas_para_gravar.append([agora_str, str(doca_sel).strip(), agenda_final, conferente_final, "ENCERRADO"])
    else:
        for pessoa in equipe_sel:
            linhas_para_gravar.append([agora_str, str(doca_sel).strip(), agenda_final, conferente_final, pessoa])
    
    if conflitos and not encerrar:
        docas_afetadas = set(conflitos.values())
        agora_dt = agora_dt + datetime.timedelta(seconds=1)
        agora_str_2 = agora_dt.strftime("%d/%m/%Y %H:%M:%S")
        
        for d_antiga in docas_afetadas:
            eq_antiga = info_docas[d_antiga]['equipe'].copy()
            for p_movida in [p for p, d in conflitos.items() if d == d_antiga]:
                if p_movida in eq_antiga:
                    eq_antiga.remove(p_movida)
            
            if not eq_antiga:
                linhas_para_gravar.append([agora_str_2, d_antiga, info_docas[d_antiga]['agenda'], info_docas[d_antiga]['conferente'], "ENCERRADO"])
            else:
                for p_restante in eq_antiga:
                    linhas_para_gravar.append([agora_str_2, d_antiga, info_docas[d_antiga]['agenda'], info_docas[d_antiga]['conferente'], p_restante])

    client = conectar_google()
    sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
    try:
        ws_log = sh.worksheet("LOG_PRODUTIVIDADE")
        ws_log.append_rows(linhas_para_gravar)
        return True
    except:
        st.error("Erro ao gravar Log.")
        return False

@st.dialog("Confirmação de Transferência")
def exibir_popup_transferencia(doca_sel, agenda_sel, conferente_sel, equipe_sel, conflitos, info_docas):
    st.write("Colaboradores já alocados em outras docas:")
    for p, d in conflitos.items():
        st.markdown(f"- **{p}** (Sairá da Doca **{d}**)")
        
    st.write(f"Confirma a transferência para a **Doca {doca_sel}**?")
    st.markdown("<br>", unsafe_allow_html=True)
    
    c1, c2 = st.columns(2)
    if c1.button("Confirmar Transferência", use_container_width=True):
        with st.spinner("Atualizando docas..."):
            if processar_gravacao_doca(doca_sel, agenda_sel, conferente_sel, equipe_sel, conflitos, info_docas, False):
                carregar_log_produtividade.clear()
                st.rerun() 
                
    if c2.button("Cancelar", use_container_width=True):
        st.rerun()

# --- POP-UP MAGALOG: JUSTIFICATIVA DE ATRASO ---
@st.dialog("Justificativa de Atraso")
def exibir_popup_justificativa(dados_multiplos, linha_log_fecha, categoria_carga, pecas_val, m3_val):
    st.warning("Esta carga ultrapassou o tempo de meta. Por favor, informe o motivo do atraso para finalizar:")
    
    opcoes_atraso = [
        "Veiculo Demorou subir para doca",
        "Falta de suprimentos na doca",
        "Carga mal acomodada",
        "Produto Misturado",
        "Equipe Disponível menor do que a necessária",
        "Troca de turno/refeição/rito",
        "Outro (Descrever abaixo)"
    ]
    
    motivo = st.selectbox("Selecione o motivo principal:", opcoes_atraso)
    detalhe = st.text_area("Detalhes adicionais (opcional):", placeholder="Ex: O caminhão chegou com as caixas tombadas...")
    
    if st.button("Confirmar Finalização", use_container_width=True):
        justificativa_final = f"{motivo} - {detalhe}".strip(" - ")
        
        for linha in dados_multiplos:
            linha.append(justificativa_final)
            linha.append(categoria_carga)
            linha.append(pecas_val) 
            linha.append(m3_val)    
            
        with st.spinner("Gravando justificativa e finalizando..."):
            if gravar_conclusao_doca(dados_multiplos, linha_log_fecha):
                st.cache_data.clear()
                st.rerun()

def gravar_alinhamento(dados_para_gravar):
    client = conectar_google()
    sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
    try:
        ws_alinhamento = sh.worksheet("ALINHAMENTO")
        ws_alinhamento.append_row(dados_para_gravar)
        return True
    except Exception as e:
        st.error(f"Erro ao gravar na aba 'ALINHAMENTO'. Verifique se ela existe. Detalhe: {e}")
        return False

# ==========================================================
# 4. TRATAMENTO FINANCEIRO
# ==========================================================
def formatar_moeda_br(valor):
    if pd.isna(valor) or valor == 0: return "R$ 0,00"
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def tratar_dados(df_h):
    # --------------------------------------------------
    # BLINDAGEM: remove colunas duplicadas da planilha
    # --------------------------------------------------
    df_h = df_h.loc[:, ~pd.Index(df_h.columns).duplicated()].copy()

    df_h['STATUS'] = df_h['STATUS'].astype(str).str.strip().str.upper()
    df_h['AGENDA WMS'] = df_h['AGENDA WMS'].astype(str).str.strip().str.upper()
    df_h['DATA AGENDA'] = df_h['DATA AGENDA'].astype(str).str.strip()
    df_h['DATA AGENDADA'] = pd.to_datetime(df_h['DATA AGENDA'], dayfirst=True, errors='coerce')
    df_h = df_h.dropna(subset=['DATA AGENDADA']).copy()

    df_h['PRIORIDADE_STATUS'] = df_h['STATUS'].apply(
        lambda x: 1 if x == 'OK' else (0 if x == 'AUSENTE' else -1)
    )
    df_h = df_h.sort_values(
        by=['AGENDA WMS', 'PRIORIDADE_STATUS', 'DATA AGENDADA'],
        ascending=[True, False, False]
    )

    mask_wms_valido = ~df_h['AGENDA WMS'].isin(['', '-', 'NAN', 'NONE'])
    df_com_wms = df_h[mask_wms_valido].drop_duplicates(subset=['AGENDA WMS'], keep='first')
    df_sem_wms = df_h[~mask_wms_valido]
    df_h = pd.concat([df_com_wms, df_sem_wms], ignore_index=True).drop(columns=['PRIORIDADE_STATUS'])

    df_h['ANO'] = df_h['DATA AGENDADA'].dt.year.astype('Int64')
    df_h['MES_ORDENACAO'] = df_h['DATA AGENDADA'].dt.to_period('M')
    meses_pt = {
        1:'Jan', 2:'Fev', 3:'Mar', 4:'Abr', 5:'Mai', 6:'Jun',
        7:'Jul', 8:'Ago', 9:'Set', 10:'Out', 11:'Nov', 12:'Dez'
    }
    df_h['MES_NOME'] = df_h['DATA AGENDADA'].dt.month.map(meses_pt) + "/" + df_h['ANO'].astype(str)

    def limpar_moeda(valor):
        if pd.isna(valor) or str(valor).strip() == '':
            return 0.0
        v = str(valor).upper().replace('R$', '').replace(' ', '').strip()
        if v in ['', '-', 'OK', 'NAOÉCOBRADO']:
            return 0.0
        if '.' in v and ',' in v:
            v = v.replace('.', '').replace(',', '.')
        elif '.' in v:
            v = v.replace('.', '')
        elif ',' in v:
            v = v.replace(',', '.')
        try:
            return round(float(v), 2)
        except ValueError:
            return 0.0

    df_h['VALOR_REAL'] = df_h['VALOR'].apply(limpar_moeda) if 'VALOR' in df_h.columns else 0.0
    df_h['LINHA'] = df_h['LINHA'].astype(str).str.strip().str.upper()
    df_h['FORNECEDOR/SELLER'] = df_h['FORNECEDOR/SELLER'].astype(str).str.strip().str.upper()
    df_h['CATEGORIA'] = df_h['CATEGORIA'].astype(str).str.strip().str.upper()

    def eh_interno(forn):
        if forn.startswith(('MAGAZINE', 'FILIAL')):
            return True
        if re.match(r'^C[D\d]', forn):
            return True
        return False

    df_h = df_h[~df_h['FORNECEDOR/SELLER'].apply(eh_interno)]

    # --------------------------------------------------
    # Separa 1P x FULL
    # --------------------------------------------------
    mask_full = df_h['LINHA'].str.contains('FULL', na=False) | df_h['CATEGORIA'].str.contains('FULL', na=False)
    df_full = df_h[mask_full].copy()
    df_main = df_h[~mask_full].copy()

    # Mantém somente linhas/categorias relevantes no 1P
    pag_l = df_main[df_main['VALOR_REAL'] > 0]['LINHA'].unique()
    pag_c = df_main[df_main['VALOR_REAL'] > 0]['CATEGORIA'].unique()
    df_main = df_main[
        (df_main['LINHA'].isin([l for l in pag_l if l != ''])) |
        (df_main['CATEGORIA'].isin([c for c in pag_c if c != '']))
    ].copy()

    # Médias monetárias baseadas em cargas OK do 1P
    df_ok = df_main[(df_main['STATUS'] == 'OK') & (df_main['VALOR_REAL'] > 0)]
    m_linha = df_ok.groupby('LINHA')['VALOR_REAL'].mean().to_dict()
    m_cat = df_ok.groupby('CATEGORIA')['VALOR_REAL'].mean().to_dict()

    # --------------------------------------------------
    # 1P
    # --------------------------------------------------
    df_main['TIPO_RECEITA'] = '1P'
    df_main['VALOR_CONSIDERADO'] = df_main['VALOR_REAL']

    df_main['VALOR_PERDIDO'] = 0.0
    mask_aus = df_main['STATUS'] == 'AUSENTE'
    df_main.loc[mask_aus, 'VALOR_PERDIDO'] = df_main.loc[mask_aus, 'LINHA'].map(m_linha)

    mask_zero = mask_aus & (df_main['VALOR_PERDIDO'].isna() | (df_main['VALOR_PERDIDO'] == 0))
    df_main.loc[mask_zero, 'VALOR_PERDIDO'] = df_main.loc[mask_zero, 'CATEGORIA'].map(m_cat)
    df_main['VALOR_PERDIDO'] = df_main['VALOR_PERDIDO'].fillna(0).round(2)

    # --------------------------------------------------
    # FULL
    # --------------------------------------------------
    df_full['TIPO_RECEITA'] = 'FULL'
    df_full['VALOR_PERDIDO'] = 0.0

    df_full['VALOR_CONSIDERADO'] = df_full['VALOR_REAL']

    mask_full_zero = df_full['VALOR_CONSIDERADO'].isna() | (df_full['VALOR_CONSIDERADO'] == 0)
    df_full.loc[mask_full_zero, 'VALOR_CONSIDERADO'] = df_full.loc[mask_full_zero, 'LINHA'].map(m_linha)

    mask_full_zero = df_full['VALOR_CONSIDERADO'].isna() | (df_full['VALOR_CONSIDERADO'] == 0)
    df_full.loc[mask_full_zero, 'VALOR_CONSIDERADO'] = df_full.loc[mask_full_zero, 'CATEGORIA'].map(m_cat)

    df_full.loc[df_full['CATEGORIA'] == 'DIVERSOS', 'VALOR_CONSIDERADO'] = 350.00
    df_full['VALOR_CONSIDERADO'] = df_full['VALOR_CONSIDERADO'].fillna(500.00).round(2)

    # Compatibilidade com qualquer uso anterior
    df_full['VALOR_ESTIMADO'] = df_full['VALOR_CONSIDERADO']

    # Blindagem final
    df_main = df_main.loc[:, ~pd.Index(df_main.columns).duplicated()].copy()
    df_full = df_full.loc[:, ~pd.Index(df_full.columns).duplicated()].copy()

    return df_main, df_full


def render_hero(titulo, subtitulo, badge="Plataforma Operacional"):
    st.markdown(f"""
        <div class="MAGALOG-hero">
            <div class="MAGALOG-hero-badge">{badge}</div>
            <div class="MAGALOG-hero-title">{titulo}</div>
            <div class="MAGALOG-hero-subtitle">{subtitulo}</div>
        </div>
    """, unsafe_allow_html=True)


def render_home_dashboard():
    render_hero(
        "Sistema Carga e Descarga",
        "Painel central para absenteísmo, gestão de docas, alinhamentos, produtividade e visão financeira com identidade mais executiva.",
        "MAGALOG • Controle em tempo real"
    )

    st.markdown("""
        <div class="MAGALOG-shell">
            <div class="MAGALOG-grid">
                <div class="MAGALOG-feature-card">
                    <div class="MAGALOG-feature-icon"><span class="icon-MAGALOG">assignment_ind</span></div>
                    <div class="MAGALOG-feature-title">Registro de Absenteísmo</div>
                    <div class="MAGALOG-feature-text">Lance faltas, BH, DSR, atestados e acompanhe o status da equipe de forma prática.</div>
                </div>
                <div class="MAGALOG-feature-card">
                    <div class="MAGALOG-feature-icon"><span class="icon-MAGALOG">local_shipping</span></div>
                    <div class="MAGALOG-feature-title">Gestão de Docas</div>
                    <div class="MAGALOG-feature-text">Visão operacional, ajuste equipes demandas e planejamentos de maneira simples</div>
                </div>
                <div class="MAGALOG-feature-card">
                    <div class="MAGALOG-feature-icon"><span class="icon-MAGALOG">calendar_month</span></div>
                    <div class="MAGALOG-feature-title">Registro de Alinhamento</div>
                    <div class="MAGALOG-feature-text">Programe folgas e movimentos da equipe em uma experiência mais organizada e intuitiva.</div>
                </div>
                <div class="MAGALOG-feature-card">
                    <div class="MAGALOG-feature-icon"><span class="icon-MAGALOG">monitoring</span></div>
                    <div class="MAGALOG-feature-title">Produtividade</div>
                    <div class="MAGALOG-feature-text">KPIs mais corporativos para tempo médio, SLA e leitura de performance por agenda e colaborador.</div>
                </div>
                <div class="MAGALOG-feature-card">
                    <div class="MAGALOG-feature-icon"><span class="icon-MAGALOG">attach_money</span></div>
                    <div class="MAGALOG-feature-title">Financeiro</div>
                    <div class="MAGALOG-feature-text">Indicadores com linguagem visual de diretoria para faturamento, perdas e ticket médio.</div>
                </div>
            </div>
            <div class="MAGALOG-info-strip">Eficiência e produtividade em tempo real</div>
        </div>
    """, unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("""
            <div class="MAGALOG-mini-card">
                <div class="MAGALOG-mini-label">Visual novo</div>
                <div class="MAGALOG-mini-value">Dashboard</div>
                <div class="MAGALOG-mini-desc">Hero section, cards premium e sidebar mais elegante.</div>
            </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown("""
            <div class="MAGALOG-mini-card">
                <div class="MAGALOG-mini-label">Impacto</div>
                <div class="MAGALOG-mini-value">Mais UX</div>
                <div class="MAGALOG-mini-desc">Leitura rápida, sensação de produto interno de alto nível.</div>
            </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown("""
            <div class="MAGALOG-mini-card">
                <div class="MAGALOG-mini-label">Lógica preservada</div>
                <div class="MAGALOG-mini-value">100%</div>
                <div class="MAGALOG-mini-desc">Mudança concentrada no front, mantendo sua operação atual.</div>
            </div>
        """, unsafe_allow_html=True)


# ==========================================================
# 5. ROTEADOR DE MÓDULOS (SIDEBAR)
# ==========================================================

st.sidebar.markdown("""
    <div style="padding: 10px 8px 4px 8px; margin-bottom: 14px;">
        <div style="background: linear-gradient(135deg, rgba(255,255,255,0.12), rgba(255,255,255,0.05)); border:1px solid rgba(255,255,255,0.10); border-radius: 22px; padding: 18px 16px; box-shadow: inset 0 1px 0 rgba(255,255,255,0.08);">
            <div style="font-size: 13px; font-weight: 800; letter-spacing:.08em; text-transform: uppercase; color:#9CC8FF; margin-bottom:8px;">MAGALOG</div>
            <div style="font-size: 26px; font-weight: 900; line-height:1.0; color:#FFFFFF; margin-bottom:8px;">R.I.T.M.O</div>
            <div style="font-size: 13px; color:rgba(255,255,255,0.72);">Registro Inteligente de Tempo, Monitoramento e Operação</div>
            <div style="height: 8px; margin-top:14px; border-radius: 999px; background: linear-gradient(90deg, #0086FF, #00D2FF, #FF8A3D, #FF4D6D); background-size:300% 300%; animation: MAGALOGGlow 7s linear infinite;"></div>
        </div>
    </div>
""", unsafe_allow_html=True)

pagina_selecionada = st.sidebar.radio(
    "Navegação",
    [
        "Gestão de Docas", 
        "Gerador de Equipes (I.A.)",
        "Registro Absenteísmo",
        "Registro de Alinhamento", 
        "Produtividade (NS & Equipe)", 
        "Financeiro (Diretoria)",
        "Absenteísmo (RH)",
        "Controle de Cobrança"
    ]
)

st.sidebar.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
if st.sidebar.button("Sincronizar Agora", type="secondary", use_container_width=True):
    with st.spinner("Puxando dados em tempo real da base..."):
        st.cache_data.clear()
        st.rerun()

# ==========================================================
# HOME
# ==========================================================
if pagina_selecionada == "Visão Geral":
    render_home_dashboard()

# ==========================================================
# MÓDULO 1: ABSENTEÍSMO
# ==========================================================
elif pagina_selecionada == "Registro Absenteísmo":
    render_hero('Lançamento de Absenteísmos', 'Controle diário da presença da equipe com busca rápida, status padronizado e gravação direta na base.', 'Módulo operacional')
    
    try:
        df_equipe = carregar_equipe()
        st.markdown('<div class="MAGALOG-card">', unsafe_allow_html=True)
        data_chamada = st.date_input("Data", date.today())
        busca = st.text_input("Buscar Colaborador", placeholder="ID ou Nome...")
        st.markdown('</div>', unsafe_allow_html=True)

        if busca:
            df_filtrado = df_equipe[df_equipe['NOME'].str.contains(busca, case=False, na=False) | df_equipe['ID'].astype(str).str.contains(busca, na=False)].copy()
        else:
            df_filtrado = df_equipe.copy()

        df_filtrado['OCORRÊNCIA'] = "PRESENTE" 
        opcoes_ocorrencia = ["PRESENTE", "FALTA", "DSR", "BH", "LICENÇA", "ATESTADO", "FÉRIAS"]

        st.markdown('<div class="MAGALOG-ribbon">Registro da Equipe</div>', unsafe_allow_html=True)
        
        df_filtrado['ID'] = df_filtrado['ID'].astype(str).str.replace('\.0$', '', regex=True)

        df_filtrado = df_filtrado.copy()

        df_filtrado['ID'] = df_filtrado['ID'].fillna('')
        df_filtrado['ID'] = df_filtrado['ID'].astype(str).str.strip()
        df_filtrado['ID'] = df_filtrado['ID'].replace(['nan', 'None', '<NA>'], '')
        df_filtrado['ID'] = df_filtrado['ID'].str.replace(r'\.0$', '', regex=True)

        df_editado = st.data_editor(
            df_filtrado[['ID', 'NOME', 'CARGO', 'TURNO', 'OCORRÊNCIA']],
            column_config={
                "OCORRÊNCIA": st.column_config.SelectboxColumn("Status", options=opcoes_ocorrencia, required=True, width="medium"),
                "ID": st.column_config.TextColumn("Matrícula", disabled=True, width="small"),
                "NOME": st.column_config.TextColumn("Nome", disabled=True, width="large"),
                "CARGO": None,
                "TURNO": None,
            },
            hide_index=True, use_container_width=True, key="editor_chamada"
        )
        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Gravar no Sistema", use_container_width=True):
            ocorrencias = df_editado[df_editado['OCORRÊNCIA'] != "PRESENTE"]
            if not ocorrencias.empty:
                lista_final = []
                data_str = data_chamada.strftime("%d/%m/%Y")
                for index, row in ocorrencias.iterrows():
                    lista_final.append([data_str, row['ID'], row['NOME'], row['OCORRÊNCIA'], row['TURNO']])
                with st.spinner("Gravando..."):
                    sucesso = gravar_absenteismo(lista_final)
                    if sucesso:
                        st.success(f"{len(lista_final)} registros salvos!")
                        carregar_equipe.clear()
            else:
                st.warning("Nenhuma falta marcada.")

    except Exception as e:
        st.error(f"Erro na conexão com RH: {e}")

# ==========================================================
# MÓDULO 2: GESTÃO DE DOCAS E PRODUTIVIDADE
# ==========================================================
elif pagina_selecionada == "Gestão de Docas":
    render_hero('Gestão de Docas', 'Controle unificado de recebimento e expedição com leitura premium, foco em prioridade e ações rápidas.', 'Plataforma Operacional')
    
    df_log = carregar_log_produtividade()
    df_matriz = carregar_matriz()
    df_aux_rec = carregar_aux()
    df_aux_exp = carregar_auxexp()
    
        # --- 1. PROCESSAMENTO DA EXPEDIÇÃO (CONSOLIDAÇÃO DE DOCAS) ---
    if not df_aux_exp.empty:
        if 'Status Chegada' not in df_aux_exp.columns:
            df_aux_exp['Status Chegada'] = ''
            
        # CAÇADOR INTELIGENTE DAS NOVAS COLUNAS
        col_tot_pecas = next((c for c in df_aux_exp.columns if 'TOTAL PEÇA' in str(c).upper() or 'TOTAL PECA' in str(c).upper()), 'Ped Venda')
        col_tot_m3 = next((c for c in df_aux_exp.columns if 'TOTAL M³' in str(c).upper() or 'TOTAL M3' in str(c).upper() or 'M³ TOTAL' in str(c).upper()), 'M³ Total')
            
        df_aux_exp = df_aux_exp.rename(columns={
            'ID Carga': 'AGENDA WMS',
            'Plano de Transporte': 'CATEGORIA',
            col_tot_m3: 'VOLUME_M3',
            col_tot_pecas: 'QTD_PECAS',
            'Limit Carreg': 'LIMITE_RAW',
            'Doca': 'DOCA_ORIGINAL'
        })
        
        # Garante que as novas colunas sejam numéricas (limpando sujeiras) para a soma ser correta
        df_aux_exp['VOLUME_M3'] = pd.to_numeric(df_aux_exp['VOLUME_M3'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        df_aux_exp['QTD_PECAS'] = pd.to_numeric(df_aux_exp['QTD_PECAS'].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)
        
        def consolidar_doca(d):
            d_str = str(d).strip()
            return d_str[:2] if len(d_str) >= 4 else d_str
        
        df_aux_exp['DOCA'] = df_aux_exp['DOCA_ORIGINAL'].apply(consolidar_doca)
        
        def checar_se_foi_embora(row):
            saida = str(row.get('Saida Veículo', '')).strip()
            lib_mot = str(row.get('Lib Mot', '')).strip()
            status_carreg = str(row.get('Status Carreg', '')).strip().upper()
            if (saida and saida.lower() != 'nan') or (lib_mot and lib_mot.lower() != 'nan') or ('FINALIZADO' in status_carreg):
                return "LIBERADO"
            return "AGUARDANDO"
            
        df_aux_exp['STATUS_CALC'] = df_aux_exp.apply(checar_se_foi_embora, axis=1)
        df_aux_exp = df_aux_exp[df_aux_exp['STATUS_CALC'] != 'LIBERADO'].copy()
        
        if not df_aux_exp.empty:
            df_exp_grouped = df_aux_exp.groupby('DOCA').agg({
                'AGENDA WMS': lambda x: ' | '.join(x.astype(str)),
                'CATEGORIA': lambda x: ' | '.join(x.astype(str)),
                'VOLUME_M3': 'sum',
                'QTD_PECAS': 'sum',
                'LIMITE_RAW': 'min',
                'STATUS_CALC': lambda x: 'AGUARDANDO' if 'AGUARDANDO' in x.values else 'LIBERADO',
                'Status Chegada': lambda x: 'AGUARD CHEGADA' if any('AGUARD' in str(v).upper() for v in x.dropna()) else 'NO PÁTIO'
            }).reset_index()
            
            df_exp_grouped['TIPO_OPERACAO'] = "EXPEDIÇÃO"
            df_exp_grouped['STATUS'] = df_exp_grouped['STATUS_CALC']
            df_exp_grouped['STATUS_CHEGADA_RAW'] = df_exp_grouped['Status Chegada'] 
            df_exp_grouped['LINHA'] = df_exp_grouped['CATEGORIA'] 
            df_exp_grouped['PEÇAS_VAL'] = df_exp_grouped['QTD_PECAS'] # Alimentando o banco final
            df_exp_grouped['M3_VAL'] = df_exp_grouped['VOLUME_M3']    # Alimentando o banco final
            df_exp_grouped['PEÇAS'] = df_exp_grouped['QTD_PECAS'] 
            df_exp_grouped['SKU'] = "-" # Ocultando campo obsoleto
            df_exp_grouped['PAGAMENTO'] = "N/A"
            df_exp_grouped['R$ DESCARGA'] = "-"
            df_exp_grouped['CONFERENTE'] = "Expedição"
            df_exp_grouped['META'] = 120 
            
            def extrair_hora(valor):
                if pd.isna(valor) or str(valor).strip() == "": return "00:00"
                try:
                    dt = pd.to_datetime(valor, dayfirst=True, errors='coerce')
                    if not pd.isna(dt): return dt.strftime("%H:%M") 
                except: pass
                import re
                v_str = str(valor).strip().upper()
                match = re.search(r'(\d{1,2}):(\d{2})', v_str)
                if match:
                    h, m = int(match.group(1)), match.group(2)
                    if 'PM' in v_str and h < 12: h += 12
                    if 'AM' in v_str and h == 12: h = 0
                    return f"{h:02d}:{m}"
                return "00:00"
            
            df_exp_grouped['LIMITE'] = df_exp_grouped['LIMITE_RAW'].apply(extrair_hora)
            df_aux_exp_final = df_exp_grouped
        else:
            df_aux_exp_final = pd.DataFrame()
    else:
        df_aux_exp_final = pd.DataFrame()

    # --- 2. PROCESSAMENTO DO RECEBIMENTO ---
    if not df_aux_rec.empty:
        df_aux_rec['AGENDA WMS'] = df_aux_rec['AGENDA WMS'].astype(str).str.strip()
        df_aux_rec['TIPO_OPERACAO'] = "RECEBIMENTO"
        
        if 'CUB' not in df_aux_rec.columns:
            df_aux_rec['CUB'] = 0
            
        df_aux_rec['PEÇAS_VAL'] = df_aux_rec['PEÇAS'] 
        df_aux_rec['M3_VAL'] = df_aux_rec['CUB']     
        df_aux_rec_final = df_aux_rec
    else:
        df_aux_rec_final = pd.DataFrame()

    # --- 3. FUSÃO E IDENTIFICAÇÃO ---
    df_aux = pd.concat([df_aux_rec_final, df_aux_exp_final], ignore_index=True)
    
    agendas_logadas = [] 
    
    if not df_aux.empty:
        df_aux['AGENDA WMS'] = df_aux['AGENDA WMS'].astype(str)
        if not df_log.empty:
            agendas_logadas = df_log['AGENDA'].astype(str).str.strip().unique().tolist()

        colunas_limpas = [str(c).upper().strip() for c in df_aux.columns]
        col_doca = next((c for c, cu in zip(df_aux.columns, colunas_limpas) if 'DOCA' in cu), None)
        col_conf = next((c for c, cu in zip(df_aux.columns, colunas_limpas) if 'CONFERENTE' in cu or 'LIDER' in cu or 'LÍDER' in cu), None)

    # --- MAPEAMENTO DA MATRIZ DE COMPETÊNCIAS ---
    dict_skills_text = {}
    dict_skills_html = {}
    if not df_matriz.empty:
        for _, row in df_matriz.iterrows():
            nome_matriz = str(row.get('NOME', '')).strip()
            if not nome_matriz: continue
            t_badges, h_badges = [], []
            if str(row.get('ECOM', '')).upper() in ['TRUE', '1', 'SIM']:
                t_badges.append("ECOM")
                h_badges.append("<span style='background:#E0F2FE; color:#0284C7; padding:2px 6px; border-radius:4px; font-size:9.5px; font-weight:800; margin-left:6px; border:1px solid #BAE6FD;'><span class='icon-MAGALOG' style='font-size:11px;'>shopping_cart</span> ECOM</span>")
            if str(row.get('FULL', '')).upper() in ['TRUE', '1', 'SIM']:
                t_badges.append("FULL")
                h_badges.append("<span style='background:#FAE8FF; color:#C026D3; padding:2px 6px; border-radius:4px; font-size:9.5px; font-weight:800; margin-left:6px; border:1px solid #F0ABFC;'><span class='icon-MAGALOG' style='font-size:11px;'>inventory_2</span> FULL</span>")
            if str(row.get('CARREGAMENTO', '')).upper() in ['TRUE', '1', 'SIM']:
                t_badges.append("CARREG")
                h_badges.append("<span style='background:#FFEDD5; color:#EA580C; padding:2px 6px; border-radius:4px; font-size:9.5px; font-weight:800; margin-left:6px; border:1px solid #FDBA74;'><span class='icon-MAGALOG' style='font-size:11px;'>upload</span> CARREG</span>")
            if str(row.get('DESCARGA', '')).upper() in ['TRUE', '1', 'SIM']:
                t_badges.append("DESC")
                h_badges.append("<span style='background:#DCFCE7; color:#16A34A; padding:2px 6px; border-radius:4px; font-size:9.5px; font-weight:800; margin-left:6px; border:1px solid #86EFAC;'><span class='icon-MAGALOG' style='font-size:11px;'>download</span> DESC</span>")
            if str(row.get('ENSINAR', '')).upper() in ['TRUE', '1', 'SIM']:
                t_badges.append("MESTRE")
                h_badges.append("<span style='background:#FEF9C3; color:#CA8A04; padding:2px 6px; border-radius:4px; font-size:9.5px; font-weight:800; margin-left:6px; border:1px solid #FDE047;'><span class='icon-MAGALOG' style='font-size:11px;'>star</span> MESTRE</span>")
            if t_badges:
                dict_skills_text[nome_matriz] = " | ".join(t_badges)
                dict_skills_html[nome_matriz] = "".join(h_badges)

    def renderizar_equipe_html(string_pessoas):
        if pd.isna(string_pessoas) or str(string_pessoas).strip() == '': return ""
        lista_p = [x.strip() for x in str(string_pessoas).split(',')]
        html_final = ""
        for p in lista_p:
            if "PAUSADO" in p: continue
            habilidades = dict_skills_html.get(p, "")
            html_final += f"<div style='background:#FFFFFF; border:1px solid #CBD5E1; padding:6px 10px; border-radius:8px; display:inline-block; margin:4px 6px 4px 0px; font-size:12px; color:#1E293B; box-shadow:0 2px 4px rgba(0,0,0,0.02);'><b>{p}</b>{habilidades}</div>"
        return html_final

    # --- FUNÇÃO INTELIGENTE DE CALCULAR TEMPO ATIVO (LENDO O LOG) ---
    def calcular_tempo_real_ativo(agenda, df_log_completo, hora_ref):
        df_ag = df_log_completo[df_log_completo['AGENDA'].astype(str).str.strip() == str(agenda).strip()].copy()
        if df_ag.empty: return 0, hora_ref
        
        df_ag['DATA_HORA_DT'] = pd.to_datetime(df_ag['DATA_HORA'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
        df_ag = df_ag.sort_values('DATA_HORA_DT')
        
        inicio_real = df_ag['DATA_HORA_DT'].iloc[0]
        eventos = df_ag[['DATA_HORA_DT', 'AUXILIARES']].drop_duplicates(subset=['DATA_HORA_DT'])
        
        tempo_pausa_segundos = 0
        pausa_start = None
        
        for _, row_ev in eventos.iterrows():
            aux_str = str(row_ev['AUXILIARES']).upper()
            is_pausa = "PAUSADO" in aux_str
            is_fim = "ENCERRADO" in aux_str
            
            if is_pausa and pausa_start is None:
                pausa_start = row_ev['DATA_HORA_DT']
            elif not is_pausa and not is_fim and pausa_start is not None:
                tempo_pausa_segundos += (row_ev['DATA_HORA_DT'] - pausa_start).total_seconds()
                pausa_start = None
                
        if pausa_start is not None:
            tempo_pausa_segundos += (hora_ref - pausa_start).total_seconds()
            
        tempo_total = (hora_ref - inicio_real).total_seconds()
        tempo_ativo = max(0, tempo_total - tempo_pausa_segundos)
        return tempo_ativo, inicio_real

    lista_auxiliares = []
    mapa_pessoas = {}
    info_docas = {}
    try:
        df_equipe = carregar_equipe()
        lista_auxiliares = df_equipe[df_equipe['NOME'].notna()]['NOME'].unique().tolist()
        lista_auxiliares = [str(nome).strip() for nome in lista_auxiliares if str(nome).strip() != '']
    except: pass

    if not df_log.empty:
        df_log['DATA_HORA_DT'] = pd.to_datetime(df_log['DATA_HORA'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
        ultimos_h = df_log.groupby('DOCA')['DATA_HORA_DT'].transform('max')
        df_ativos_bruto = df_log[df_log['DATA_HORA_DT'] == ultimos_h]
        df_ativos_agrupado = df_ativos_bruto.groupby(['DOCA', 'AGENDA', 'CONFERENTE'])['AUXILIARES'].apply(lambda x: ', '.join(x.dropna())).reset_index()
        df_ativos_agrupado = df_ativos_agrupado[df_ativos_agrupado['AUXILIARES'] != 'ENCERRADO']
        for _, row in df_ativos_agrupado.iterrows():
            doca_atual = str(row['DOCA']).strip()
            eq_atual = [x.strip() for x in str(row['AUXILIARES']).split(',')]
            info_docas[doca_atual] = {'agenda': row['AGENDA'], 'conferente': row['CONFERENTE'], 'equipe': eq_atual}
            for p in eq_atual: mapa_pessoas[p] = doca_atual

    def checar_fadiga(equipe, agenda, df_log_full, df_auxiliar):
        if not equipe or df_log_full.empty: return []
        palavras_pesadas = ['MADEIRA', 'PESADO', 'ELETRO PESADO']
        
        categoria_atual = ""
        if not df_auxiliar.empty:
            match_agenda = df_auxiliar[df_auxiliar['AGENDA WMS'] == str(agenda).strip()]
            if not match_agenda.empty:
                categoria_atual = str(match_agenda.iloc[0].get('LINHA', match_agenda.iloc[0].get('CATEGORIA', ''))).upper()
                
        is_carga_pesada = any(p in categoria_atual for p in palavras_pesadas)
        fadigados = []
        
        if is_carga_pesada and 'CATEGORIA' in df_log_full.columns:
            agora = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
            limite_24h = agora - datetime.timedelta(hours=24)
            df_ultimas_24h = df_log_full[(df_log_full['DATA_HORA_DT'] >= limite_24h) & (df_log_full['AUXILIARES'] != 'ENCERRADO')]
            
            for pessoa in equipe:
                log_pessoa = df_ultimas_24h[df_ultimas_24h['AUXILIARES'].str.contains(pessoa, na=False, regex=False)]
                if not log_pessoa.empty:
                    categorias_pessoa = log_pessoa['CATEGORIA'].fillna('').str.upper()
                    fez_pesado = categorias_pessoa.apply(lambda x: any(p in x for p in palavras_pesadas)).any()
                    if fez_pesado: fadigados.append(pessoa)
        return fadigados

    def processar_gravacao_doca_v2(doca_n, agenda_n, conferente_n, equipe_n, conflitos_n, info_docas_n, is_pausa=False, motivo_pausa=""):
        agora_dt = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
        agora_str = agora_dt.strftime("%d/%m/%Y %H:%M:%S")
        linhas_para_gravar = []

        cat_atual = "NÃO INFORMADA"
        if not df_aux.empty:
            match_atual = df_aux[df_aux['AGENDA WMS'] == str(agenda_n).strip()]
            if not match_atual.empty: cat_atual = str(match_atual.iloc[0].get('LINHA', match_atual.iloc[0].get('CATEGORIA', ''))).upper()

        # --- LÓGICA CORRIGIDA: ATUALIZA A DOCA ANTIGA SEM MATAR QUEM FICOU ---
        if conflitos_n and not is_pausa:
            docas_afetadas = set(conflitos_n.values())
            
            for doca_antiga in docas_afetadas:
                if doca_antiga in info_docas_n:
                    agenda_antiga = info_docas_n[doca_antiga]['agenda']
                    conf_antiga = info_docas_n[doca_antiga]['conferente']
                    equipe_antiga = info_docas_n[doca_antiga]['equipe'].copy()
                    
                    # Identifica quem está saindo DESTA doca específica e remove da equipe
                    pessoas_saindo = [p for p, d in conflitos_n.items() if d == doca_antiga]
                    for p in pessoas_saindo:
                        if p in equipe_antiga:
                            equipe_antiga.remove(p)
                            
                    # Busca a categoria da doca antiga
                    cat_antiga = "NÃO INFORMADA"
                    if not df_aux.empty:
                        match_ant = df_aux[df_aux['AGENDA WMS'] == str(agenda_antiga).strip()]
                        if not match_ant.empty: cat_antiga = str(match_ant.iloc[0].get('LINHA', match_ant.iloc[0].get('CATEGORIA', ''))).upper()
                    
                    if not equipe_antiga:
                        # Se não sobrou ninguém (todos foram transferidos), a doca encerra.
                        linhas_para_gravar.append([agora_str, doca_antiga, agenda_antiga, conf_antiga, "ENCERRADO", cat_antiga])
                    else:
                        # Se sobrou gente, regrava apenas os que ficaram com o horário atualizado
                        for p_restante in equipe_antiga:
                            linhas_para_gravar.append([agora_str, doca_antiga, agenda_antiga, conf_antiga, p_restante, cat_antiga])
        # ---------------------------------------------------------------------

        if is_pausa:
            linhas_para_gravar.append([agora_str, str(doca_n).strip(), str(agenda_n).strip(), str(conferente_n).strip(), f"PAUSADO - {motivo_pausa}", cat_atual])
        else:
            for pessoa in equipe_n:
                linhas_para_gravar.append([agora_str, str(doca_n).strip(), str(agenda_n).strip(), str(conferente_n).strip(), pessoa, cat_atual])

        try:
            client = conectar_google()
            sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
            ws_log = sh.worksheet("LOG_PRODUTIVIDADE")
            ws_log.append_rows(linhas_para_gravar)
            return True
        except Exception as e:
            st.error(f"Erro ao gravar: {e}")
            return False

    @st.dialog("START na Carga (Alocar Equipe)")
    def popup_start_carga(doca_sel, agenda_sel, conferente_sel, is_retomada=False, equipe_padrao=None):
        if equipe_padrao is None: equipe_padrao = []
        equipe_padrao_valida = [x for x in equipe_padrao if x in lista_auxiliares]
        
        titulo = "Retomar" if is_retomada else "Iniciar"
        st.markdown(f"<div style='font-size:14px; margin-bottom:15px;'><span class='icon-MAGALOG'>pin_drop</span> Doca: <b>{doca_sel}</b> &nbsp;|&nbsp; <span class='icon-MAGALOG'>tag</span> Agenda: <b>{agenda_sel}</b> &nbsp;|&nbsp; <span class='icon-MAGALOG'>badge</span> Líder: <b>{conferente_sel}</b></div>", unsafe_allow_html=True)
        equipe_sel = st.multiselect(f"Selecione os colaboradores para {titulo}:", options=lista_auxiliares, default=equipe_padrao_valida, format_func=lambda x: f"{x}  [{dict_skills_text[x]}]" if x in dict_skills_text else x)
        
        conflitos = {}
        for pessoa in equipe_sel:
            if pessoa in mapa_pessoas:
                if mapa_pessoas[pessoa] != str(doca_sel).strip(): conflitos[pessoa] = mapa_pessoas[pessoa]
        if conflitos:
            st.warning("Os colaboradores abaixo já estão ocupados e serão transferidos:")
            for p, d in conflitos.items(): st.markdown(f"- **{p}** (Sairá da Doca {d})")

        fadigados = checar_fadiga(equipe_sel, agenda_sel, df_log, df_aux)
        bloqueio_ergonomico = False
        
        if fadigados:
            st.markdown(f"<div style='background-color: #FEF2F2; border: 1px solid #DC2626; border-radius: 8px; padding: 15px; margin-top: 15px; margin-bottom: 15px;'><b style='color: #DC2626;'><span class='icon-MAGALOG'>warning</span> ALERTA ERGONÔMICO (SST)</b><br><span style='color: #7F1D1D; font-size: 13px;'>Os colaboradores <b>{', '.join(fadigados)}</b> já atuaram em carga pesada nas últimas 24h.</span></div>", unsafe_allow_html=True)
            ciente = st.checkbox("Declaro ciência do risco e autorizo a alocação.", key="chk_fadiga_popup")
            if not ciente: bloqueio_ergonomico = True
                
        st.markdown('<br>', unsafe_allow_html=True)
        if st.button(f"Confirmar {titulo}", type="primary", use_container_width=True):
            if not doca_sel or doca_sel == "A Definir": st.error("Esta carga precisa ter uma Doca informada antes de iniciar!")
            elif not equipe_sel: st.error("Selecione a equipe!")
            elif bloqueio_ergonomico: st.error("Você precisa assumir o risco ergonômico marcando a caixa de seleção!")
            else:
                with st.spinner(f"{titulo} operação..."):
                    sucesso = processar_gravacao_doca_v2(doca_sel, agenda_sel, conferente_sel, equipe_sel, conflitos, info_docas)
                    if sucesso:
                        st.success(f"Carga {titulo.lower()}!")
                        carregar_log_produtividade.clear()
                        st.rerun()

    # --- POP-UP NOVO: PAUSAR OPERAÇÃO ---
    @st.dialog("Pausar Operação")
    def popup_pausar_operacao(doca_sel, agenda_sel, conferente_sel):
        st.warning("Ao pausar a operação, o cronômetro desta doca deixará de contar. Apenas o tempo ativo real comporá o SLA.")
        motivos_pausa = [
            "Horário de Refeição (Almoço/Janta)",
            "Troca de Turno",
            "Falta de Energia / Sistema",
            "Problema na Carga / Avaria",
            "Falta de Paletes / Insumos",
            "Reunião / Rito",
            "Outros"
        ]
        motivo_sel = st.selectbox("Motivo da Pausa:", motivos_pausa)
        
        if st.button("Confirmar Pausa", use_container_width=True):
            with st.spinner("Pausando operação..."):
                sucesso = processar_gravacao_doca_v2(doca_sel, agenda_sel, conferente_sel, [], {}, info_docas, is_pausa=True, motivo_pausa=motivo_sel)
                if sucesso:
                    st.success("Operação Pausada com sucesso!")
                    carregar_log_produtividade.clear()
                    st.rerun()

    # --- POP-UP NOVO: ADICIONAR OPERADOR ---
    @st.dialog("Adicionar Operador à Doca")
    def popup_adicionar_operador(doca_origem, agenda_origem, conf_origem, equipe_atual, info_docas_global):
        doca_origem_str = str(doca_origem).strip()
        st.markdown(f"<div style='color:#64748B; margin-bottom:15px;'>Adicionando reforço na <b>Doca {doca_origem_str}</b></div>", unsafe_allow_html=True)
        
        disponiveis = [p for p in lista_auxiliares if p not in equipe_atual]
        novos_operadores = st.multiselect("Selecione os colaboradores para adicionar:", options=disponiveis, format_func=lambda x: f"{x}  [{dict_skills_text.get(x, '')}]" if x in dict_skills_text else x)
        
        conflitos = {}
        for pessoa in novos_operadores:
            if pessoa in mapa_pessoas:
                if mapa_pessoas[pessoa] != doca_origem_str: conflitos[pessoa] = mapa_pessoas[pessoa]
                
        if conflitos:
            st.warning("Os colaboradores abaixo já estão ocupados e serão transferidos de suas docas atuais:")
            for p, d in conflitos.items(): st.markdown(f"- **{p}** (Sairá da Doca {d})")

        fadigados = checar_fadiga(novos_operadores, agenda_origem, df_log, df_aux)
        bloqueio_ergonomico = False
        
        if fadigados:
            st.markdown(f"<div style='background-color: #FEF2F2; border: 1px solid #DC2626; border-radius: 8px; padding: 15px; margin-top: 15px; margin-bottom: 15px;'><b style='color: #DC2626;'><span class='icon-MAGALOG'>warning</span> ALERTA DE SAÚDE E SEGURANÇA (SST)</b><br><span style='color: #7F1D1D; font-size: 13px;'>Os colaboradores <b>{', '.join(fadigados)}</b> já atuaram em carga pesada nas últimas 24h. Risco ergonômico!</span></div>", unsafe_allow_html=True)
            ciente = st.checkbox("Declaro ciência do risco e autorizo a alocação.", key="chk_fadiga_add")
            if not ciente: bloqueio_ergonomico = True
            
        st.markdown('<br>', unsafe_allow_html=True)
        if st.button("Confirmar Adição", type="primary", use_container_width=True):
            if not novos_operadores: st.warning("Selecione pelo menos um colaborador para adicionar.")
            elif bloqueio_ergonomico: st.error("Você precisa assumir o risco ergonômico marcando a caixa de seleção!")
            else:
                equipe_combinada = equipe_atual + novos_operadores
                with st.spinner("Atualizando registros no sistema..."):
                    sucesso = processar_gravacao_doca_v2(doca_origem, agenda_origem, conf_origem, equipe_combinada, conflitos, info_docas_global)
                    if sucesso:
                        st.success("Equipe atualizada com sucesso!")
                        carregar_log_produtividade.clear()
                        st.rerun()

    @st.dialog("Gerenciar Operador da Doca")
    def popup_gerenciar_operador(doca_origem, equipe_atual, info_docas_global):
        doca_origem_str = str(doca_origem).strip()
        st.markdown(f"<div style='color:#64748B; margin-bottom:15px;'>Modificando a equipe da <b>Doca {doca_origem_str}</b></div>", unsafe_allow_html=True)
        operador_sel = st.selectbox("Selecione o Operador que deseja movimentar:", equipe_atual)
        acao = st.radio("O que deseja fazer com este colaborador?", ["Retirar da Operação (Ficará Livre)", "Transferir para outra Doca ativa"])
        docas_ativas = [d for d in info_docas_global.keys() if str(d).strip() != doca_origem_str]
        doca_destino = None
        if "Transferir" in acao:
            if not docas_ativas: st.warning("Não há outras docas em processo no momento para transferir.")
            else:
                opcoes_formatadas = [f"Doca {d} (Líder: {info_docas_global[d]['conferente']})" for d in docas_ativas]
                doca_destino = st.selectbox("Transferir para qual Doca?", opcoes_formatadas).split(" ")[1] 
                
        st.markdown('<br>', unsafe_allow_html=True)
        if st.button("Confirmar Alteração", type="primary", use_container_width=True):
            agora_dt = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
            agora_str = agora_dt.strftime("%d/%m/%Y %H:%M:%S")
            linhas_para_gravar = []
            
            agenda_orig = info_docas_global[doca_origem_str]['agenda']
            conf_orig = info_docas_global[doca_origem_str]['conferente']
            cat_orig = "NÃO INFORMADA"
            match_o = df_aux[df_aux['AGENDA WMS'] == str(agenda_orig).strip()]
            if not match_o.empty: cat_orig = str(match_o.iloc[0].get('LINHA', match_o.iloc[0].get('CATEGORIA', ''))).upper()

            nova_equipe_origem = [p for p in equipe_atual if p != operador_sel]
            if not nova_equipe_origem: linhas_para_gravar.append([agora_str, doca_origem_str, agenda_orig, conf_orig, "ENCERRADO", cat_orig])
            else:
                for p in nova_equipe_origem: linhas_para_gravar.append([agora_str, doca_origem_str, agenda_orig, conf_orig, p, cat_orig])
                
            if "Transferir" in acao and doca_destino:
                doca_dest_str = str(doca_destino).strip()
                agenda_dest = info_docas_global[doca_dest_str]['agenda']
                conf_dest = info_docas_global[doca_dest_str]['conferente']
                cat_dest = "NÃO INFORMADA"
                match_d = df_aux[df_aux['AGENDA WMS'] == str(agenda_dest).strip()]
                if not match_d.empty: cat_dest = str(match_d.iloc[0].get('LINHA', match_d.iloc[0].get('CATEGORIA', ''))).upper()
                
                nova_equipe_dest = info_docas_global[doca_dest_str]['equipe'] + [operador_sel]
                for p in nova_equipe_dest: linhas_para_gravar.append([agora_str, doca_dest_str, agenda_dest, conf_dest, p, cat_dest])
                
            with st.spinner("Atualizando registros no sistema..."):
                client = conectar_google()
                sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
                try:
                    ws_log = sh.worksheet("LOG_PRODUTIVIDADE")
                    ws_log.append_rows(linhas_para_gravar)
                    st.success("Equipe atualizada com sucesso!")
                    carregar_log_produtividade.clear()
                    st.rerun()
                except Exception as e: st.error(f"Erro ao gravar: {e}")

    st.markdown("<br>", unsafe_allow_html=True)
    with st.container(border=True):
        c_f1, c_f2 = st.columns(2)
        with c_f1:
            filtro_op = st.selectbox("Filtrar por Operação:", ["Todas", "RECEBIMENTO", "EXPEDIÇÃO"])
        with c_f2:
            filtro_doca = st.text_input("Buscar Doca ou Agenda:", placeholder="Ex: 14 ou 99999")
    st.markdown("<br>", unsafe_allow_html=True)

    aba1, aba2, aba3 = st.tabs(["Visão das Docas (EM PROCESSO)", "Fila de Docas (PENDENTE)", "Atividades Fixas (Suporte)"])



    with aba1:
        if not df_log.empty:
            agendas_logadas = df_log['AGENDA'].astype(str).str.strip().unique().tolist()
            df_ativos = df_ativos_bruto.copy()
            df_ativos = df_ativos.groupby(['DOCA', 'AGENDA', 'CONFERENTE', 'DATA_HORA', 'DATA_HORA_DT'])['AUXILIARES'].apply(lambda x: ', '.join(x.dropna())).reset_index()
            df_ativos = df_ativos[df_ativos['AUXILIARES'] != 'ENCERRADO']
            
            agora_dt = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
            
            def calc_urgencia_processo(row):
                try:
                    agenda_str = str(row['AGENDA']).strip()
                    meta_minutos = 60
                    if not df_aux.empty and agenda_str in df_aux['AGENDA WMS'].values:
                        aux_row = df_aux[df_aux['AGENDA WMS'] == agenda_str].iloc[0]
                        col_meta = next((c for c in df_aux.columns if 'META' in str(c).upper()), None)
                        if col_meta and pd.notna(aux_row[col_meta]): meta_minutos = int(float(str(aux_row[col_meta]).replace(',', '.')))
                    
                    inicio_dt = row['DATA_HORA_DT']
                    if pd.isna(inicio_dt): inicio_dt = agora_dt
                    return meta_minutos - ((agora_dt - inicio_dt).total_seconds() / 60)
                except: return 99999 

            df_ativos['URGENCIA'] = df_ativos.apply(calc_urgencia_processo, axis=1)
            df_ativos = df_ativos.sort_values('URGENCIA', ascending=True)

            if df_ativos.empty: st.info("Nenhuma doca ativa no momento. Pátio limpo!")
            else:
                cards_exibidos_aba1 = 0
                for index, row in df_ativos.iterrows():
                    agenda_str = str(row['AGENDA']).strip()
                    
                    aux_row = pd.Series()
                    if not df_aux.empty and agenda_str in df_aux['AGENDA WMS'].values:
                        aux_row = df_aux[df_aux['AGENDA WMS'] == agenda_str].iloc[0]
                    tipo_op = str(aux_row.get('TIPO_OPERACAO', 'RECEBIMENTO'))
                    
                    if filtro_op != "Todas" and filtro_op not in tipo_op: continue
                    if filtro_doca and filtro_doca not in str(row['DOCA']) and filtro_doca not in agenda_str: continue
                    
                    cards_exibidos_aba1 += 1
                    
                    info = {'LINHA': '-', 'SKU': '-', 'PEÇAS': '-', 'VALOR': '-', 'PAGTO': '-', 'STATUS': '-'}
                    meta_minutos = 60
                    if not aux_row.empty:
                        pagto_str = "<span class='icon-MAGALOG' style='font-size:12px; color:#16A34A;'>check_circle</span> Sim" if str(aux_row.get('PAGAMENTO', '')).upper() == 'TRUE' else "<span class='icon-MAGALOG' style='font-size:12px; color:#EA580C;'>schedule</span> Pendente"
                        valor_desc = aux_row.get('R$ DESCARGA', '-')
                        if str(valor_desc).replace('.','',1).isdigit(): valor_desc = f"R$ {float(valor_desc):,.2f}".replace(',','X').replace('.',',').replace('X','.')
                        linha_val = aux_row.get('LINHA', aux_row.get('CATEGORIA', '-'))
                        
                        info = {'LINHA': linha_val, 'SKU': aux_row.get('SKU', '-'), 'PEÇAS': aux_row.get('PEÇAS', '-'), 'VALOR': valor_desc, 'PAGTO': pagto_str, 'STATUS': aux_row.get('STATUS', '-'), 'PEÇAS_BRUTO': aux_row.get('PEÇAS_VAL', 0), 'M3_BRUTO': aux_row.get('M3_VAL', 0)}
                        
                        try:
                            col_meta = next((c for c in df_aux.columns if 'META' in str(c).upper()), None)
                            if col_meta and pd.notna(aux_row[col_meta]): meta_minutos = int(float(str(aux_row[col_meta]).replace(',', '.')))
                        except: pass

                    # Lógica de Identificação de Pausa e Cálculo Real
                    auxiliares_lista = [x.strip() for x in str(row['AUXILIARES']).split(',')]
                    is_pausado = any("PAUSADO" in p.upper() for p in auxiliares_lista)
                    
                    tempo_ativo_seg, inicio_real_dt = calcular_tempo_real_ativo(agenda_str, df_log, agora_dt)
                    decorrido_min = tempo_ativo_seg / 60
                    restante_min = meta_minutos - decorrido_min

                    if is_pausado:
                        h_p, m_p = int(abs(restante_min) // 60), int(abs(restante_min) % 60)
                        if restante_min >= 0:
                            cor_timer, bg_timer, txt_timer = "#B45309", "#FEF3C7", f"<span class='icon-MAGALOG' style='font-size:12px;'>pause_circle</span> Pausado (Resta {h_p:02d}h{m_p:02d}m)"
                        else:
                            cor_timer, bg_timer, txt_timer = "#DC2626", "#FEF2F2", f"<span class='icon-MAGALOG' style='font-size:12px;'>pause_circle</span> Pausado (Atraso {h_p:02d}h{m_p:02d}m)"
                    else:
                        if restante_min >= 0:
                            h, m = int(restante_min // 60), int(restante_min % 60)
                            cor_timer, bg_timer, txt_timer = "#00C853", "#E6F9EC", f"<span class='icon-MAGALOG' style='font-size:12px;'>hourglass_bottom</span> Resta {h:02d}h{m:02d}m"
                        else:
                            atraso = abs(restante_min)
                            h, m = int(atraso // 60), int(atraso % 60)
                            cor_timer, bg_timer, txt_timer = "#DC2626", "#FEF2F2", f"<span class='icon-MAGALOG' style='font-size:12px;'>warning</span> Atraso -{h:02d}h{m:02d}m"

                    html_equipe_cards = renderizar_equipe_html(row['AUXILIARES'])

                    with st.container(border=True):
                        if is_pausado:
                            cor_tema = "#F59E0B"
                            css_hack = f"<style>div[data-testid='stVerticalBlockBorderWrapper']:has(.card-proc-{index}) {{ border: 2px solid {cor_tema} !important; box-shadow: 0 4px 15px rgba(245,158,11,0.15) !important; background-color: #FFFBEB !important; }}</style><div class='card-proc-{index}'></div>"
                        elif "EXPEDIÇÃO" in tipo_op:
                            cor_tema = "#0086FF"
                            css_hack = f"<style>div[data-testid='stVerticalBlockBorderWrapper']:has(.card-pend-{index}) {{ border: 2px solid {cor_tema} !important; box-shadow: 0 4px 15px rgba(0,134,255,0.15) !important; }}</style><div class='card-pend-{index}'></div>"
                            html_detalhes = f"<div style='font-size: 11.5px; color: #475569; background-color: #F0F9FF; padding: 10px; border-radius: 8px; margin-bottom: 8px; border: 1px solid #BAE6FD;'><b>Planos:</b> <span style='color:#0369A1; font-weight:bold;'>{info['LINHA']}</span><br><div style='margin-top: 4px;'><b>Total Peças:</b> {info['PEÇAS_BRUTO']:,.0f} &nbsp;|&nbsp; <b>M³ Total:</b> {info['M3_BRUTO']:,.2f} &nbsp;|&nbsp; <b>Status:</b> <span style='color:#0284C7; font-weight:bold;'>{info['STATUS']}</span></div></div>".replace(',', 'X').replace('.', ',').replace('X', '.')
                        else:
                            cor_tema = "#EA580C"
                            css_hack = f"<style>div[data-testid='stVerticalBlockBorderWrapper']:has(.card-proc-{index}) {{ border: 2px solid {cor_tema} !important; box-shadow: 0 4px 15px rgba(234,88,12,0.15) !important; }}</style><div class='card-proc-{index}'></div>"
                            
                        html_detalhes = f"<div style='font-size: 11.5px; color: #475569; background-color: #FFFFFF; padding: 10px; border-radius: 8px; margin-bottom: 8px; border: 1px solid #E2E8F0;'><b>Linha/Plano:</b> {info['LINHA']}  |  <b>SKU/Peds:</b> {info['SKU']}  |  <b>Volume:</b> {info['PEÇAS']}<br><div style='margin-top: 4px;'><b>Status WMS:</b> <span style='color:{cor_tema}; font-weight:bold;'>{info['STATUS']}</span></div></div>"

                        st.markdown(css_hack, unsafe_allow_html=True)
                        c_title, c_time = st.columns([5, 5])
                        c_title.markdown(f"<h4 style='margin:0; color:{cor_tema};'>Doca {row['DOCA']}</h4>", unsafe_allow_html=True)
                        
                        if is_pausado:
                            c_time.markdown(f"<div style='text-align:right;'><div style='display:inline-block; font-size:12.5px; font-weight:800; color:{cor_timer}; background-color:{bg_timer}; padding:3px 6px; border-radius:4px; border: 1px solid {cor_timer};'>{txt_timer}</div></div>", unsafe_allow_html=True)
                        else:
                            c_time.markdown(f"<div style='text-align:right;'><div style='font-size:11px; color:#64748B; margin-bottom: 2px;'><span class='icon-MAGALOG' style='font-size:11px;'>watch_later</span> Início Real: {inicio_real_dt.strftime('%H:%M')}</div><div style='display:inline-block; font-size:12.5px; font-weight:800; color:{cor_timer}; background-color:{bg_timer}; padding:3px 6px; border-radius:4px; border: 1px solid {cor_timer};'>{txt_timer} <span style='font-size:10px; font-weight:normal;'>(Meta: {meta_minutos}m)</span></div></div>", unsafe_allow_html=True)
                        
                        st.markdown(f"<div style='font-size: 13px; margin: 4px 0px 4px 0px;'><b>Agenda:</b> {row['AGENDA']} | <b>Líder:</b> {row['CONFERENTE']}</div>", unsafe_allow_html=True)
                        st.markdown(html_detalhes, unsafe_allow_html=True)
                        
                        if is_pausado:
                            motivo = next((p for p in auxiliares_lista if "PAUSADO" in p.upper()), "PAUSADO")
                            st.markdown(f"<div style='background-color: #FEF3C7; padding: 12px; border-radius: 8px; border: 1px dashed #F59E0B; margin-bottom: 15px;'><div style='font-size: 13px; font-weight: 800; color: #B45309;'>Status: {motivo}</div></div>", unsafe_allow_html=True)
                            
                            # Recupera equipe antes da pausa para sugerir no pop-up
                            log_agenda = df_log[df_log['AGENDA'].astype(str).str.strip() == agenda_str]
                            pessoas_antes_pausa = log_agenda[~log_agenda['AUXILIARES'].astype(str).str.upper().str.contains('PAUSADO|ENCERRADO', regex=True)]['AUXILIARES'].unique().tolist()
                            
                            if st.button("Retomar Operação", key=f"btn_ret_{row['DOCA']}_{index}", use_container_width=True, type="primary"):
                                popup_start_carga(row['DOCA'], row['AGENDA'], row['CONFERENTE'], is_retomada=True, equipe_padrao=pessoas_antes_pausa)
                        else:
                            st.markdown(f"<div style='background-color: #F1F5F9; padding: 12px; border-radius: 8px; border: 1px dashed #CBD5E1; margin-bottom: 15px;'><div style='font-size: 11px; font-weight: 800; color: #64748B; margin-bottom: 6px; text-transform: uppercase;'>Operadores Alocados:</div>{html_equipe_cards}</div>", unsafe_allow_html=True)
                            
                            c_eq, c_btn1, c_btn2 = st.columns([4, 3, 3])
                            with c_btn1:
                                if st.button("Pausar", key=f"btn_pse_{row['DOCA']}_{index}", use_container_width=True):
                                    popup_pausar_operacao(row['DOCA'], row['AGENDA'], row['CONFERENTE'])
                            with c_btn2:
                                if st.button("Finalizar", key=f"btn_fin_{row['DOCA']}_{index}", type="primary", use_container_width=True):
                                    clique_dt = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
                                    
                                    # CÁLCULO DE TEMPO REAL DE DESCARGA PARA GRAVAR NA PLANILHA
                                    tempo_ativo_segundos, inicio_real = calcular_tempo_real_ativo(agenda_str, df_log, clique_dt)
                                    total_minutos_final = int(tempo_ativo_segundos // 60)
                                    horas, mins = total_minutos_final // 60, total_minutos_final % 60
                                    tempo_str = f"{horas:02d}:{mins:02d}"
                                    
                                    cat_final = str(info['LINHA']).upper()
                                    pecas_final = info.get('PEÇAS_BRUTO', 0)
                                    m3_final = info.get('M3_BRUTO', 0)
                                    
                                    linhas_conclusao_multiplas = []
                                    for pessoa in auxiliares_lista:
                                        linhas_conclusao_multiplas.append([clique_dt.strftime("%d/%m/%Y"), str(row['DOCA']), str(row['AGENDA']), str(row['CONFERENTE']), len(auxiliares_lista), pessoa, inicio_real.strftime("%d/%m/%Y %H:%M:%S"), clique_dt.strftime("%H:%M:%S"), tempo_str])
                                    
                                    linha_log_fecha = [clique_dt.strftime("%d/%m/%Y %H:%M:%S"), str(row['DOCA']), row['AGENDA'], row['CONFERENTE'], "ENCERRADO", cat_final]
                                    
                                    if restante_min < 0: 
                                        exibir_popup_justificativa(linhas_conclusao_multiplas, linha_log_fecha, cat_final, pecas_final, m3_final)
                                    else:
                                        for linha in linhas_conclusao_multiplas: 
                                            linha.append("No Prazo")
                                            linha.append(cat_final)
                                            linha.append(pecas_final)
                                            linha.append(m3_final)    
                                            
                                        with st.spinner("Finalizando..."):
                                            if gravar_conclusao_doca(linhas_conclusao_multiplas, linha_log_fecha):
                                                st.success("Doca finalizada com sucesso!")
                                                carregar_log_produtividade.clear()
                                                st.rerun()
                            
                            st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True) 
                            
                            c_action1, c_action2 = st.columns(2)
                            with c_action1:
                                if st.button("Adicionar Reforço", key=f"btn_add_op_{row['DOCA']}_{index}", use_container_width=True):
                                    popup_adicionar_operador(row['DOCA'], row['AGENDA'], row['CONFERENTE'], auxiliares_lista, info_docas)
                            with c_action2:
                                if st.button("Mover/Retirar", key=f"btn_mgr_{row['DOCA']}_{index}", use_container_width=True):
                                    popup_gerenciar_operador(row['DOCA'], auxiliares_lista, info_docas)
                                
                if cards_exibidos_aba1 == 0: st.info("Nenhuma doca encontrada com esses filtros.")
        else:
            st.info("O Log de Operações ainda está vazio.")

    with aba2:
        if not df_aux.empty:
            df_pendentes = df_aux[~df_aux['AGENDA WMS'].isin(agendas_logadas)].copy()
            df_pendentes = df_pendentes[df_pendentes['AGENDA WMS'] != '']
            status_ignorados = ['AUSENTE', 'DEVOLVIDA', 'OK','LIBERADO']
            if 'STATUS' in df_pendentes.columns: df_pendentes = df_pendentes[~df_pendentes['STATUS'].astype(str).str.upper().isin(status_ignorados)]
            
            agora_dt = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
            
            def calc_urgencia_pendente(row):
                try:
                    meta_minutos = 60
                    col_meta = next((c for c in df_aux.columns if 'META' in str(c).upper()), None)
                    if col_meta and pd.notna(row[col_meta]): meta_minutos = int(float(str(row[col_meta]).replace(',', '.')))
                    
                    col_limite = next((c for c in df_aux.columns if 'LIMITE' in str(c).upper()), None)
                    if col_limite and pd.notna(row[col_limite]) and str(row[col_limite]).strip() != '':
                        limite_str = str(row[col_limite]).strip()
                        partes = limite_str.split(':')
                        h_lim, m_lim = int(partes[0]), int(partes[1][:2])
                        limite_dt = agora_dt.replace(hour=h_lim, minute=m_lim, second=0, microsecond=0)
                        
                        if h_lim < 12 and agora_dt.hour >= 12:
                            limite_dt += datetime.timedelta(days=1)
                        elif h_lim >= 18 and agora_dt.hour < 12:
                            limite_dt -= datetime.timedelta(days=1)
                            
                        hora_max_inicio = limite_dt - datetime.timedelta(minutes=meta_minutos)
                        return (hora_max_inicio - agora_dt).total_seconds() / 60
                except: pass
                return 99999 

            df_pendentes['URGENCIA'] = df_pendentes.apply(calc_urgencia_pendente, axis=1)
            df_pendentes = df_pendentes.sort_values('URGENCIA', ascending=True)

            if df_pendentes.empty: st.info("Nenhuma agenda aguardando equipe. Pátio zerado!")
            else:
                cards_exibidos_aba2 = 0
                for index, row in df_pendentes.iterrows():
                    tipo_op = str(row.get('TIPO_OPERACAO', 'RECEBIMENTO'))
                    agenda_str = str(row['AGENDA WMS'])
                    doca_str = str(row[col_doca]).strip() if col_doca and pd.notna(row[col_doca]) else "A Definir"
                    
                    if filtro_op != "Todas" and filtro_op not in tipo_op: continue
                    if filtro_doca and filtro_doca not in str(doca_str) and filtro_doca not in agenda_str: continue
                    
                    cards_exibidos_aba2 += 1
                    
                    if doca_str.lower() in ['nan', 'none', '']: doca_str = "A Definir"
                    conf_str = str(row[col_conf]).strip() if col_conf and pd.notna(row[col_conf]) else "A Definir"
                    if conf_str.lower() in ['nan', 'none', '']: conf_str = "A Definir"
                    pagto_str = "<span class='icon-MAGALOG' style='font-size:12px; color:#16A34A;'>check_circle</span> Sim" if str(row.get('PAGAMENTO', '')).upper() == 'TRUE' else "<span class='icon-MAGALOG' style='font-size:12px; color:#EA580C;'>schedule</span> Pendente"
                    valor_desc = row.get('R$ DESCARGA', '-')
                    if str(valor_desc).replace('.','',1).isdigit(): valor_desc = f"R$ {float(valor_desc):,.2f}".replace(',','X').replace('.',',').replace('X','.')
                    linha_val = row.get('LINHA', row.get('CATEGORIA', '-'))
                    
                    info = {'LINHA': linha_val, 'SKU': row.get('SKU', '-'), 'PEÇAS': row.get('PEÇAS', '-'), 'VALOR': valor_desc, 'PAGTO': pagto_str, 'STATUS': row.get('STATUS', '-'), 'PEÇAS_BRUTO': row.get('PEÇAS_VAL', 0), 'M3_BRUTO': row.get('M3_VAL', 0)}
                    
                    meta_minutos = 60
                    limite_str, hora_max_str = "", "-"
                    txt_timer_pend, cor_timer_pend, bg_timer_pend = "<span class='icon-MAGALOG' style='font-size:12px;'>schedule</span> Aguardando...", "#F59E0B", "#FEF3C7"
                    try:
                        col_meta = next((c for c in df_aux.columns if 'META' in str(c).upper()), None)
                        if col_meta and pd.notna(row[col_meta]): meta_minutos = int(float(str(row[col_meta]).replace(',', '.')))
                        col_limite = next((c for c in df_aux.columns if 'LIMITE' in str(c).upper()), None)
                        if col_limite and pd.notna(row[col_limite]) and str(row[col_limite]).strip() != '':
                            limite_str = str(row[col_limite]).strip()
                            partes = limite_str.split(':')
                            h_lim, m_lim = int(partes[0]), int(partes[1][:2])
                            limite_dt = agora_dt.replace(hour=h_lim, minute=m_lim, second=0, microsecond=0)
                            
                            if h_lim < 12 and agora_dt.hour >= 12:
                                limite_dt += datetime.timedelta(days=1)
                            elif h_lim >= 18 and agora_dt.hour < 12:
                                limite_dt -= datetime.timedelta(days=1)
                                
                            hora_max_inicio = limite_dt - datetime.timedelta(minutes=meta_minutos)
                            hora_max_str = hora_max_inicio.strftime("%H:%M")
                            diff_min = (hora_max_inicio - agora_dt).total_seconds() / 60
                            if diff_min >= 0:
                                h, m = int(diff_min // 60), int(diff_min % 60)
                                cor_timer_pend, bg_timer_pend, txt_timer_pend = "#00C853", "#E6F9EC", f"<span class='icon-MAGALOG' style='font-size:12px;'>check_circle</span> Sobra {h:02d}h{m:02d}m p/ Iniciar"
                            else:
                                atraso = abs(diff_min)
                                h, m = int(atraso // 60), int(atraso % 60)
                                cor_timer_pend, bg_timer_pend, txt_timer_pend = "#DC2626", "#FEF2F2", f"<span class='icon-MAGALOG' style='font-size:12px;'>error</span> ATRASADO HÁ {h:02d}h{m:02d}m"
                    except: pass
                    
                    status_chegada = str(row.get('STATUS_CHEGADA_RAW', '')).upper()
                    
                    if "EXPEDIÇÃO" in tipo_op and "AGUARD" in status_chegada:
                        txt_timer_pend = "<span class='icon-MAGALOG' style='font-size:12px;'>local_shipping</span> AGUARDANDO VEÍCULO"
                        cor_timer_pend = "#64748B" 
                        bg_timer_pend = "#F1F5F9"  
                    
                    with st.container(border=True):
                        if "EXPEDIÇÃO" in tipo_op:
                            cor_tema = "#0086FF"
                            css_hack = f"<style>div[data-testid='stVerticalBlockBorderWrapper']:has(.card-proc-{index}) {{ border: 2px solid {cor_tema} !important; box-shadow: 0 4px 15px rgba(0,134,255,0.15) !important; }}</style><div class='card-proc-{index}'></div>"
                            html_detalhes = f"<div style='font-size: 11.5px; color: #475569; background-color: #F0F9FF; padding: 10px; border-radius: 8px; margin-bottom: 8px; border: 1px solid #BAE6FD;'><b>Planos:</b> <span style='color:#0369A1; font-weight:bold;'>{info['LINHA']}</span><br><div style='margin-top: 4px;'><b>Total Peças:</b> {info['PEÇAS_BRUTO']:,.0f} &nbsp;|&nbsp; <b>M³ Total:</b> {info['M3_BRUTO']:,.2f} &nbsp;|&nbsp; <b>Status:</b> <span style='color:#0284C7; font-weight:bold;'>{info['STATUS']}</span></div></div>".replace(',', 'X').replace('.', ',').replace('X', '.')
                        else:
                            cor_tema = "#EA580C"
                            css_hack = f"<style>div[data-testid='stVerticalBlockBorderWrapper']:has(.card-pend-{index}) {{ border: 2px solid {cor_tema} !important; box-shadow: 0 4px 15px rgba(234,88,12,0.15) !important; }}</style><div class='card-pend-{index}'></div>"
                            html_detalhes = f"<div style='font-size: 11.5px; color: #475569; background-color: #FFF7ED; padding: 10px; border-radius: 8px; margin-bottom: 8px; border: 1px solid #FFEDD5;'><b>Linha:</b> {info['LINHA']}  |  <b>SKU:</b> {info['SKU']}  |  <b>Peças:</b> {info['PEÇAS']}<br><div style='margin-top: 4px;'><b>Valor Carga:</b> {info['VALOR']}  |  <b>Pagto:</b> {info['PAGTO']}  |  <b>Status:</b> <span style='color:#EA580C; font-weight:bold;'>{info['STATUS']}</span></div></div>"

                        st.markdown(css_hack, unsafe_allow_html=True)
                        st.markdown(f"""
<div style='display: flex; justify-content: space-between; align-items: center;'>
    <h4 style='margin:0; color:{cor_tema};'>Doca {doca_str}</h4>
    <div style='display:inline-block; font-size:12px; font-weight:800; color:{cor_timer_pend}; background-color:{bg_timer_pend}; padding:3px 6px; border-radius:4px; border: 1px solid {cor_timer_pend};'>{txt_timer_pend}</div>
</div>
<div style='font-size: 13px; margin: 8px 0px 4px 0px; display: flex; justify-content: space-between;'>
    <span><b>Agenda:</b> {agenda_str} | <b>Líder:</b> {conf_str}</span>
</div>
<div style='font-size: 12px; display: flex; gap: 15px; margin-bottom: 8px;'>
    <span style='color:#64748B;'>Meta Operação: <b>{meta_minutos}m</b></span>
    <span style='color:#0086FF;'>Iniciar até: <b>{hora_max_str}</b></span>
    <span style='color:#DC2626;'>Fim Máximo: <b>{limite_str if limite_str else '-'}</b></span>
</div>
{html_detalhes}
""", unsafe_allow_html=True)
                        
                        c_eq_pend, c_btn_pend = st.columns([7, 3])
                        c_eq_pend.markdown(f"""<div style='font-size: 12px; color: #DC2626; background-color: #FEF2F2; padding: 8px; border-radius: 8px; border: 1px solid #FECACA;'><span class="icon-MAGALOG" style="font-size:14px; vertical-align:text-bottom;">person_off</span> <b>Equipe:</b> <span style="font-weight:900;">PENDENTE ALOCAÇÃO</span></div>""", unsafe_allow_html=True)
                        
                        with c_btn_pend:
                            if st.button("Adicionar Equipe", key=f"btn_add_{index}", use_container_width=True): 
                                popup_start_carga(doca_str, agenda_str, conf_str)
                                
                if cards_exibidos_aba2 == 0: st.info("Nenhuma agenda encontrada com esses filtros.")

    with aba3:
        st.markdown("<br><h4 style='color: #0086FF; margin-bottom: 5px;'><span class='icon-MAGALOG'>engineering</span> Atividades Fixas e Suporte</h4>", unsafe_allow_html=True)
        st.markdown("<div style='color: #64748B; font-size: 14px; margin-bottom: 25px;'>Aloque a equipe em funções de suporte (sem volume de peças ou m³) para registrar o tempo trabalhado e evitar que sejam lidos como ociosos no Radar da Liderança. O tempo será gravado normalmente na base de finalizados.</div>", unsafe_allow_html=True)
        
        # O Cardápio de Atividades
        atividades = [
            {"id": "5S", "nome": "5S - Organização e Limpeza", "icon": "cleaning_services", "cor": "#0086FF"},
            {"id": "COURRIER", "nome": "Suporte Interno (Courrier)", "icon": "inventory_2", "cor": "#8B5CF6"},
            {"id": "OUTROS", "nome": "Atividades Externas / Outros", "icon": "construction", "cor": "#F59E0B"}
        ]
        
        cols_atv = st.columns(3)
        
        for i, atv in enumerate(atividades):
            with cols_atv[i]:
                # Criando as assinaturas únicas para o banco de dados
                doca_atv = "ATV-" + atv['id']
                agenda_atv = atv['id']
                conferente_atv = "Líder Suporte"
                
                # Checa se essa atividade já está rodando agora
                is_active = doca_atv in info_docas
                
                st.markdown(f"""
                <div class="MAGALOG-card" style="border-top: 4px solid {atv['cor']} !important; height: 100%;">
                    <div style="display:flex; align-items:center; gap:10px; margin-bottom:15px; border-bottom: 1px solid #F1F5F9; padding-bottom: 10px;">
                        <span class="icon-MAGALOG" style="color:{atv['cor']}; font-size:26px; background: {atv['cor']}15; padding: 8px; border-radius: 8px;">{atv['icon']}</span>
                        <div style="font-weight:900; color:#0F172A; font-size:15px;">{atv['nome']}</div>
                    </div>
                """, unsafe_allow_html=True)
                
                if is_active:
                    equipe_atv = info_docas[doca_atv]['equipe']
                    is_paused = any("PAUSADO" in p.upper() for p in equipe_atv)
                    
                    # Cálculo cronológico real da atividade
                    agora_dt = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
                    tempo_ativo_seg, inicio_real_dt = calcular_tempo_real_ativo(agenda_atv, df_log, agora_dt)
                    h_atv, m_atv = int(tempo_ativo_seg // 3600), int((tempo_ativo_seg % 3600) // 60)
                    tempo_decorrido = f"{h_atv:02d}h{m_atv:02d}m"
                    
                    if is_paused:
                        st.markdown(f"<div style='background:#FEF3C7; color:#B45309; padding:6px; border-radius:6px; font-size:12px; font-weight:bold; text-align:center; margin-bottom:15px;'><span class='icon-MAGALOG' style='font-size:14px; vertical-align: middle;'>pause_circle</span> PAUSADO</div>", unsafe_allow_html=True)
                    else:
                        st.markdown(f"<div style='background:#E6F9EC; color:#00C853; padding:6px; border-radius:6px; font-size:12px; font-weight:bold; text-align:center; margin-bottom:15px;'><span class='icon-MAGALOG' style='font-size:14px; vertical-align: middle;'>timer</span> EM ANDAMENTO ({tempo_decorrido})</div>", unsafe_allow_html=True)
                        
                    html_eq = renderizar_equipe_html(", ".join(equipe_atv))
                    st.markdown(f"<div style='font-size:11px; font-weight:800; color:#64748B; margin-bottom:8px;'>EQUIPE ALOCADA:</div>{html_eq}<br><br>", unsafe_allow_html=True)
                    
                    # BOTOES DE AÇÃO
                    c_b1, c_b2 = st.columns(2)
                    with c_b1:
                        if is_paused:
                            log_agenda = df_log[df_log['AGENDA'].astype(str).str.strip() == agenda_atv]
                            pessoas_antes_pausa = log_agenda[~log_agenda['AUXILIARES'].astype(str).str.upper().str.contains('PAUSADO|ENCERRADO', regex=True)]['AUXILIARES'].unique().tolist()
                            if st.button("Retomar", key=f"ret_{atv['id']}", use_container_width=True, type="primary"):
                                popup_start_carga(doca_atv, agenda_atv, conferente_atv, is_retomada=True, equipe_padrao=pessoas_antes_pausa)
                        else:
                            if st.button("Pausar", key=f"pse_{atv['id']}", use_container_width=True):
                                popup_pausar_operacao(doca_atv, agenda_atv, conferente_atv)
                    with c_b2:
                        if st.button("Encerrar", key=f"fim_{atv['id']}", type="primary", use_container_width=True):
                            clique_dt = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
                            tempo_ativo_segundos, inicio_real = calcular_tempo_real_ativo(agenda_atv, df_log, clique_dt)
                            total_minutos_final = int(tempo_ativo_segundos // 60)
                            horas, mins = total_minutos_final // 60, total_minutos_final % 60
                            tempo_str = f"{horas:02d}:{mins:02d}"
                            
                            cat_final = "ATIVIDADE EXTERNA"
                            equipe_limpa = [p for p in equipe_atv if "PAUSADO" not in p.upper()]
                            
                            # Gravação na planilha final sem peças e m3 (evita distorcer)
                            linhas_conclusao = []
                            for pessoa in equipe_limpa:
                                linhas_conclusao.append([
                                    clique_dt.strftime("%d/%m/%Y"), doca_atv, agenda_atv, conferente_atv, len(equipe_limpa), 
                                    pessoa, inicio_real.strftime("%d/%m/%Y %H:%M:%S"), clique_dt.strftime("%H:%M:%S"), 
                                    tempo_str, "No Prazo", cat_final, 0, 0
                                ])
                                
                            linha_log_fecha = [clique_dt.strftime("%d/%m/%Y %H:%M:%S"), doca_atv, agenda_atv, conferente_atv, "ENCERRADO", cat_final]
                            
                            with st.spinner("Encerrando atividade..."):
                                if gravar_conclusao_doca(linhas_conclusao, linha_log_fecha):
                                    st.success("Atividade encerrada e horas computadas!")
                                    carregar_log_produtividade.clear()
                                    st.rerun()
                                    
                    st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True) 
                    if st.button("Mover/Retirar Alguém", key=f"mgr_{atv['id']}", use_container_width=True):
                        popup_gerenciar_operador(doca_atv, equipe_atv, info_docas)
                        
                else:
                    st.markdown("<div style='color:#94A3B8; font-size:13px; margin-bottom:25px; text-align:center; padding:20px 0;'>Nenhuma equipe alocada no momento.<br>Mão de obra livre.</div>", unsafe_allow_html=True)
                    if st.button("Iniciar Atividade", key=f"start_{atv['id']}", type="primary", use_container_width=True):
                        popup_start_carga(doca_atv, agenda_atv, conferente_atv)
                        
                st.markdown("</div>", unsafe_allow_html=True)


# ==========================================================
# MÓDULO 5: FINANCEIRO (DIRETORIA)
# ==========================================================
elif pagina_selecionada == "Financeiro (Diretoria)":
    try:
        with st.spinner('Sincronizando com Base de Dados Financeira...'):
            df_raw = carregar_dados_financeiros()
            df, df_full = tratar_dados(df_raw)

        # --------------------------------------------
        # Filtros
        # --------------------------------------------
        st.sidebar.markdown('<div class="MAGALOG-ribbon" style="left: 0; font-size: 12px;">Parâmetros Financeiros</div>', unsafe_allow_html=True)

        hoje = datetime.date.today()

        datas_base = []
        if not df.empty and 'DATA AGENDADA' in df.columns:
            datas_base.append(df['DATA AGENDADA'].min().date())
        if not df_full.empty and 'DATA AGENDADA' in df_full.columns:
            datas_base.append(df_full['DATA AGENDADA'].min().date())

        d_min = min(datas_base) if datas_base else datetime.date(2025, 1, 1)
        d_max_limite = max(hoje, datetime.date(2026, 12, 31))

        selecao = st.sidebar.date_input(
            "Selecione o Período",
            value=(d_min, hoje),
            min_value=d_min,
            max_value=d_max_limite
        )

        if isinstance(selecao, tuple) and len(selecao) == 2:
            data_ini, data_fim = selecao
        else:
            data_ini = selecao[0] if isinstance(selecao, (tuple, list)) else selecao
            data_fim = data_ini

        filtro_visao = st.sidebar.radio(
            "Tipo de Receita",
            ["Ambos", "Apenas 1P", "Apenas FULL"]
        )

        # --------------------------------------------
        # Filtro por data
        # --------------------------------------------
        df_main_f = df.copy()
        df_full_f = df_full.copy()

        # Blindagem contra colunas duplicadas
        df_main_f = df_main_f.loc[:, ~pd.Index(df_main_f.columns).duplicated()].copy()
        df_full_f = df_full_f.loc[:, ~pd.Index(df_full_f.columns).duplicated()].copy()

        if not df_main_f.empty:
            mask_data_main = (
                (df_main_f['DATA AGENDADA'].dt.date >= data_ini) &
                (df_main_f['DATA AGENDADA'].dt.date <= data_fim)
            )
            df_main_f = df_main_f[mask_data_main].copy()

        if not df_full_f.empty:
            mask_data_full = (
                (df_full_f['DATA AGENDADA'].dt.date >= data_ini) &
                (df_full_f['DATA AGENDADA'].dt.date <= data_fim)
            )
            df_full_f = df_full_f[mask_data_full].copy()

        # --------------------------------------------
        # Aplicação do filtro de visão
        # --------------------------------------------
        if filtro_visao == "Apenas 1P":
            df_visao = df_main_f.copy()
        elif filtro_visao == "Apenas FULL":
            df_visao = df_full_f.copy()
        else:
            df_visao = pd.concat([df_main_f, df_full_f], ignore_index=True, sort=False)

        df_visao = df_visao.loc[:, ~pd.Index(df_visao.columns).duplicated()].copy()

        render_hero(
            'Visão Oficial de Faturamento',
            f'Período analisado: {data_ini.strftime("%d/%m/%Y")} até {data_fim.strftime("%d/%m/%Y")}. '
            f'Leitura executiva de receita, perdas, ticket médio e participação FULL.',
            f'Financeiro • {filtro_visao}'
        )

        if df_visao.empty:
            st.warning("Não há dados financeiros para o período/filtro selecionado.")
        else:
            rec = df_visao[df_visao['STATUS'] == 'OK'].copy()
            aus = df_visao[df_visao['STATUS'] == 'AUSENTE'].copy()

            total_r = rec['VALOR_CONSIDERADO'].sum() if not rec.empty else 0
            total_p = aus['VALOR_PERDIDO'].sum() if ('VALOR_PERDIDO' in aus.columns and not aus.empty) else 0
            tkt_carga = rec['VALOR_CONSIDERADO'].mean() if not rec.empty else 0

            dias_unicos = rec['DATA AGENDADA'].dt.date.nunique() if not rec.empty else 0
            tkt_dia = total_r / dias_unicos if dias_unicos > 0 else 0

            meses_unicos = rec['MES_ORDENACAO'].nunique() if not rec.empty else 0
            tkt_mes = total_r / meses_unicos if meses_unicos > 0 else 0

            qtd_agendas = rec['AGENDA WMS'].nunique() if not rec.empty else 0

            # --------------------------------------------
            # KPIs
            # --------------------------------------------
            col1, col2, col3, col4, col5 = st.columns(5)

            def render_kpi(titulo, valor, subtitulo, cor_hex):
                return f"""
                <div style="background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px; padding: 20px 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.02); position: relative; overflow: hidden; margin-bottom: 20px;">
                    <div style="position: absolute; top: 0; left: 0; width: 4px; height: 100%; background: {cor_hex};"></div>
                    <div style="color: #64748B; font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: .05em; margin-bottom: 8px;">{titulo}</div>
                    <div style="color: #0F172A; font-size: 26px; font-weight: 900; line-height: 1.1; margin-bottom: 6px;">{valor}</div>
                    <div style="color: #94A3B8; font-size: 12px;">{subtitulo}</div>
                </div>
                """

            with col1:
                st.markdown(
                    render_kpi("Receita Total", formatar_moeda_br(total_r), "Cargas com status OK", "#0086FF"),
                    unsafe_allow_html=True
                )
            with col2:
                st.markdown(
                    render_kpi("Perda por No-show", formatar_moeda_br(total_p), "Estimativa financeira das ausências", "#FF3366"),
                    unsafe_allow_html=True
                )
            with col3:
                st.markdown(
                    render_kpi("Ticket por Carga", formatar_moeda_br(tkt_carga), f"{qtd_agendas} agendas faturadas", "#00C853"),
                    unsafe_allow_html=True
                )
            with col4:
                st.markdown(
                    render_kpi("Ticket por Dia", formatar_moeda_br(tkt_dia), f"{dias_unicos} dias com faturamento", "#F59E0B"),
                    unsafe_allow_html=True
                )
            with col5:
                st.markdown(
                    render_kpi("Ticket por Mês", formatar_moeda_br(tkt_mes), f"{meses_unicos} meses no período", "#8B5CF6"),
                    unsafe_allow_html=True
                )

            # --------------------------------------------
            # Mix 1P x FULL
            # --------------------------------------------
            df_mix = pd.concat([df_main_f, df_full_f], ignore_index=True, sort=False)
            df_mix = df_mix.loc[:, ~pd.Index(df_mix.columns).duplicated()].copy()
            df_mix = df_mix[df_mix['STATUS'] == 'OK'].copy()

            if not df_mix.empty:
                mix_receita = (
                    df_mix.groupby('TIPO_RECEITA', as_index=False)['VALOR_CONSIDERADO']
                    .sum()
                )
            else:
                mix_receita = pd.DataFrame(columns=['TIPO_RECEITA', 'VALOR_CONSIDERADO'])

                        # --------------------------------------------
            # Pizza + Receita por Categoria
            # --------------------------------------------
            col_g1, col_g2 = st.columns(2)

            with col_g1:
                st.markdown('<div class="MAGALOG-card">', unsafe_allow_html=True)
                st.markdown(
                    "<h4 style='color: #334155; margin-bottom: 15px;'><span class='icon-MAGALOG'>pie_chart</span> Participação da Receita</h4>",
                    unsafe_allow_html=True
                )

                if not mix_receita.empty and mix_receita['VALOR_CONSIDERADO'].sum() > 0:
                    mix_receita = mix_receita.copy()
                    total_mix = mix_receita['VALOR_CONSIDERADO'].sum()
                    mix_receita['PERCENTUAL'] = (mix_receita['VALOR_CONSIDERADO'] / total_mix * 100).round(1)
                    mix_receita['VALOR_FORMATADO'] = mix_receita['VALOR_CONSIDERADO'].apply(formatar_moeda_br)

                    c_pizza, c_legenda = st.columns([1.4, 1])

                    with c_pizza:
                        fig_pizza = px.pie(
                            mix_receita,
                            values='VALOR_CONSIDERADO',
                            names='TIPO_RECEITA',
                            hole=0.55,
                            color='TIPO_RECEITA',
                            color_discrete_map={
                                '1P': '#0086FF',
                                'FULL': '#A855F7'
                            }
                        )

                        fig_pizza.update_traces(
                            textposition='inside',
                            textinfo='percent+label',
                            customdata=mix_receita[['VALOR_FORMATADO', 'PERCENTUAL']],
                            hovertemplate="<b>%{label}</b><br>Participação: %{customdata[1]:.1f}%<br>Receita: %{customdata[0]}<extra></extra>"
                        )

                        fig_pizza.update_layout(
                            margin=dict(l=0, r=0, t=20, b=0),
                            height=360,
                            showlegend=False,
                            paper_bgcolor='rgba(0,0,0,0)',
                            plot_bgcolor='rgba(0,0,0,0)'
                        )

                        st.plotly_chart(fig_pizza, use_container_width=True, config={'displayModeBar': False})

                    with c_legenda:
                        st.markdown("<div style='padding-top: 18px;'></div>", unsafe_allow_html=True)

                        cores = {
                            '1P': '#0086FF',
                            'FULL': '#A855F7'
                        }

                        for _, row in mix_receita.sort_values('VALOR_CONSIDERADO', ascending=False).iterrows():
                            cor = cores.get(row['TIPO_RECEITA'], '#64748B')
                            st.markdown(
                                f"""
                                <div style="
                                    border: 1px solid #E2E8F0;
                                    border-left: 6px solid {cor};
                                    border-radius: 12px;
                                    padding: 14px;
                                    margin-bottom: 12px;
                                    background: #FFFFFF;
                                ">
                                    <div style="font-size: 12px; color: #64748B; font-weight: 800;">
                                        {row['TIPO_RECEITA']}
                                    </div>
                                    <div style="font-size: 22px; font-weight: 900;">
                                        {row['PERCENTUAL']:.1f}%
                                    </div>
                                    <div style="font-size: 14px; font-weight: 700;">
                                        {row['VALOR_FORMATADO']}
                                    </div>
                                </div>
                                """,
                                unsafe_allow_html=True
                            )
                else:
                    st.info("Sem receita suficiente para montar a participação entre 1P e FULL.")

                st.markdown('</div>', unsafe_allow_html=True)


            with col_g2:
                st.markdown('<div class="MAGALOG-card">', unsafe_allow_html=True)
                st.markdown(
                    "<h4 style='color: #334155; margin-bottom: 15px;'><span class='icon-MAGALOG'>table_chart</span> Detalhamento de Faturamento</h4>",
                    unsafe_allow_html=True
                )
                
                if not rec.empty:
                    # Agora temos TRÊS abas elegantes e dinâmicas
                    aba_cat, aba_veic, aba_forn = st.tabs(["Por Categoria", "Por Veículo", "Por Fornecedor"])
                    
                    # ==========================================
                    # ABA 1: VISÃO ORIGINAL POR CATEGORIA
                    # ==========================================
                    with aba_cat:
                        df_cat_agendas = (
                            rec.groupby('CATEGORIA', as_index=False)
                            .agg(
                                AGENDAS=('AGENDA WMS', 'nunique'),
                                VALOR=('VALOR_CONSIDERADO', 'sum')
                            )
                            .sort_values(['VALOR', 'AGENDAS'], ascending=[False, False])
                        )
                        total_agendas_cat = df_cat_agendas['AGENDAS'].sum()
                        total_valor_cat = df_cat_agendas['VALOR'].sum()
                        
                        c_top1, c_top2 = st.columns(2)
                        with c_top1:
                            st.markdown(f"""<div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:12px; padding:12px 14px; margin-bottom:12px;"><div style="font-size:11px; color:#64748B; font-weight:800; text-transform:uppercase;">Total de Agendas</div><div style="font-size:24px; color:#0F172A; font-weight:900;">{total_agendas_cat}</div></div>""", unsafe_allow_html=True)
                        with c_top2:
                            st.markdown(f"""<div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:12px; padding:12px 14px; margin-bottom:12px;"><div style="font-size:11px; color:#64748B; font-weight:800; text-transform:uppercase;">Valor Cobrado</div><div style="font-size:24px; color:#0F172A; font-weight:900;">{formatar_moeda_br(total_valor_cat)}</div></div>""", unsafe_allow_html=True)
                            
                        st.dataframe(
                            df_cat_agendas,
                            column_config={
                                'CATEGORIA': st.column_config.TextColumn("Categoria", width="large"),
                                'AGENDAS': st.column_config.NumberColumn("Agendas", format="%d"),
                                'VALOR': st.column_config.NumberColumn("Valor", format="R$ %.2f"),
                            },
                            hide_index=True, use_container_width=True, height=280
                        )
                        
                    # ==========================================
                    # ABA 2: VISÃO POR TIPO DE VEÍCULO CONSOLIDADO
                    # ==========================================
                    with aba_veic:
                        col_cat_carro = next((c for c in rec.columns if 'CAT' in c.upper() and 'CARRO' in c.upper()), None)
                        if col_cat_carro:
                            def extrair_veiculo(val):
                                s = str(val).strip().upper()
                                if s in ['', 'NAN', 'NONE', 'NULL']: return 'NÃO INFORMADO'
                                
                                if '-' in s: s = s.split('-', 1)[1].strip()
                                s = s.replace('_PALETIZADO', '').replace(' PALETIZADO', '').strip()
                                
                                if 'TRUCK' in s or 'TOCO' in s: return 'CAMINHÃO (TRUCK/TOCO)'
                                if 'VAN' in s or 'KOMBI' in s or 'SPRINTER' in s: return 'VAN / KOMBI'
                                if 'FIORINO' in s or 'CARRO' in s: return 'FIORINO / CARRO'
                                if 'VUC' in s or 'HR' in s: return 'VUC / HR'
                                if 'TRÊS' in s or '3/4' in s or 'TRES' in s: return 'TRÊS POR QUATRO'
                                if 'SIDER' in s: return 'SIDER'
                                if 'GRANELEIRO' in s: return 'GRANELEIRO'
                                if 'BITREM' in s: return 'BITREM'
                                if 'BAÚ' in s or 'BAU' in s: return 'CARRETA BAÚ'
                                return s 
                                
                            rec['TIPO_VEICULO'] = rec[col_cat_carro].apply(extrair_veiculo)
                        else:
                            rec['TIPO_VEICULO'] = 'COLUNA NÃO ENCONTRADA'
                            
                        rec['TIPO_VEICULO'] = rec['TIPO_VEICULO'].replace('', 'NÃO INFORMADO')

                        black_list = ['NAO É COBRADO', 'NÃO É COBRADO', 'NÃO INFORMADO', 'NAO INFORMADO', '2 AGENDA']
                        rec_veiculos_limpos = rec[~rec['TIPO_VEICULO'].isin(black_list)].copy()

                        if not rec_veiculos_limpos.empty:
                            df_veiculo_agendas = (
                                rec_veiculos_limpos.groupby('TIPO_VEICULO', as_index=False)
                                .agg(
                                    AGENDAS=('AGENDA WMS', 'nunique'),
                                    VALOR=('VALOR_CONSIDERADO', 'sum')
                                )
                                .sort_values(['VALOR', 'AGENDAS'], ascending=[False, False])
                            )
                            
                            total_agendas_veic = df_veiculo_agendas['AGENDAS'].sum()
                            total_valor_veic = df_veiculo_agendas['VALOR'].sum()
                            
                            c_v1, c_v2 = st.columns(2)
                            with c_v1:
                                st.markdown(f"""<div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:12px; padding:12px 14px; margin-bottom:12px;"><div style="font-size:11px; color:#64748B; font-weight:800; text-transform:uppercase;">Agendas Válidas</div><div style="font-size:24px; color:#0F172A; font-weight:900;">{total_agendas_veic}</div></div>""", unsafe_allow_html=True)
                            with c_v2:
                                st.markdown(f"""<div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:12px; padding:12px 14px; margin-bottom:12px;"><div style="font-size:11px; color:#64748B; font-weight:800; text-transform:uppercase;">Valor (Veículos)</div><div style="font-size:24px; color:#0F172A; font-weight:900;">{formatar_moeda_br(total_valor_veic)}</div></div>""", unsafe_allow_html=True)
                                
                            st.dataframe(
                                df_veiculo_agendas,
                                column_config={
                                    'TIPO_VEICULO': st.column_config.TextColumn("Veículo Consolidado", width="large"),
                                    'AGENDAS': st.column_config.NumberColumn("Agendas", format="%d"),
                                    'VALOR': st.column_config.NumberColumn("Valor", format="R$ %.2f"),
                                },
                                hide_index=True, use_container_width=True, height=280
                            )
                        else:
                            st.info("Nenhum tipo de veículo válido no período filtrado.")

                    # ==========================================
                    # ABA 3: NOVA VISÃO POR FORNECEDOR/SELLER
                    # ==========================================
                    with aba_forn:
                        col_forn = 'FORNECEDOR/SELLER'
                        if col_forn in rec.columns:
                            df_forn_agendas = (
                                rec.groupby(col_forn, as_index=False)
                                .agg(
                                    AGENDAS=('AGENDA WMS', 'nunique'),
                                    VALOR=('VALOR_CONSIDERADO', 'sum')
                                )
                                .sort_values(['VALOR', 'AGENDAS'], ascending=[False, False])
                            )
                            
                            total_agendas_forn = df_forn_agendas['AGENDAS'].sum()
                            total_valor_forn = df_forn_agendas['VALOR'].sum()
                            
                            c_f1, c_f2 = st.columns(2)
                            with c_f1:
                                st.markdown(f"""<div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:12px; padding:12px 14px; margin-bottom:12px;"><div style="font-size:11px; color:#64748B; font-weight:800; text-transform:uppercase;">Agendas por Fornecedor</div><div style="font-size:24px; color:#0F172A; font-weight:900;">{total_agendas_forn}</div></div>""", unsafe_allow_html=True)
                            with c_f2:
                                st.markdown(f"""<div style="background:#F8FAFC; border:1px solid #E2E8F0; border-radius:12px; padding:12px 14px; margin-bottom:12px;"><div style="font-size:11px; color:#64748B; font-weight:800; text-transform:uppercase;">Valor (Fornecedores)</div><div style="font-size:24px; color:#0F172A; font-weight:900;">{formatar_moeda_br(total_valor_forn)}</div></div>""", unsafe_allow_html=True)
                                
                            st.dataframe(
                                df_forn_agendas,
                                column_config={
                                    col_forn: st.column_config.TextColumn("Fornecedor / Seller", width="large"),
                                    'AGENDAS': st.column_config.NumberColumn("Agendas", format="%d"),
                                    'VALOR': st.column_config.NumberColumn("Valor", format="R$ %.2f"),
                                },
                                hide_index=True, use_container_width=True, height=280
                            )
                        else:
                            st.warning("Coluna 'FORNECEDOR/SELLER' não encontrada na base.")
                            
                else:
                    st.info("Sem agendas cobradas no período para exibir faturamento.")
                st.markdown('</div>', unsafe_allow_html=True)


            # --------------------------------------------
            # Evolução mensal
            # --------------------------------------------
            st.markdown('<div class="MAGALOG-card">', unsafe_allow_html=True)
            st.markdown(
                "<h4 style='color: #334155; margin-bottom: 15px;'><span class='icon-MAGALOG'>monitoring</span> Evolução Mensal de Receita e Perdas</h4>",
                unsafe_allow_html=True
            )

            ev_arrec = (
                rec.groupby(['MES_ORDENACAO', 'MES_NOME'], as_index=False)['VALOR_CONSIDERADO']
                .sum()
                .rename(columns={'VALOR_CONSIDERADO': 'ARRECADADO'})
            ) if not rec.empty else pd.DataFrame(columns=['MES_ORDENACAO', 'MES_NOME', 'ARRECADADO'])

            ev_perda = (
                aus.groupby(['MES_ORDENACAO', 'MES_NOME'], as_index=False)['VALOR_PERDIDO']
                .sum()
                .rename(columns={'VALOR_PERDIDO': 'PERDIDO'})
            ) if not aus.empty else pd.DataFrame(columns=['MES_ORDENACAO', 'MES_NOME', 'PERDIDO'])

            ev_mes = pd.merge(
                ev_arrec,
                ev_perda,
                on=['MES_ORDENACAO', 'MES_NOME'],
                how='outer'
            ).fillna(0)

            if not ev_mes.empty:
                ev_mes = ev_mes.sort_values('MES_ORDENACAO')

                text_arrecadado = ev_mes['ARRECADADO'].apply(formatar_moeda_br)
                text_perdido = ev_mes['PERDIDO'].apply(formatar_moeda_br)

                fig = make_subplots(specs=[[{"secondary_y": True}]])

                fig.add_trace(
                    go.Bar(
                        x=ev_mes['MES_NOME'],
                        y=ev_mes['ARRECADADO'],
                        name="Faturado (R$)",
                        marker_color='#0086FF',
                        text=text_arrecadado,
                        textposition='auto',
                        hovertemplate="%{x}<br>Faturado: %{text}<extra></extra>"
                    ),
                    secondary_y=False
                )

                fig.add_trace(
                    go.Scatter(
                        x=ev_mes['MES_NOME'],
                        y=ev_mes['PERDIDO'],
                        name="Perda por No-Show (R$)",
                        mode='lines+markers+text',
                        line=dict(color='#FF3366', width=4),
                        marker=dict(size=8),
                        text=text_perdido,
                        textposition='top center',
                        hovertemplate="%{x}<br>Perdido: %{text}<extra></extra>"
                    ),
                    secondary_y=True
                )

                fig.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    margin=dict(t=30, b=10, l=0, r=0),
                    legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="center", x=0.5)
                )
                fig.update_yaxes(showgrid=False, secondary_y=False)
                fig.update_yaxes(showgrid=False, secondary_y=True)

                st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
            else:
                st.info("Sem dados suficientes para evolução mensal no período.")

            st.markdown('</div>', unsafe_allow_html=True)

            # --------------------------------------------
            # Tabela detalhada
            # --------------------------------------------
            st.markdown('<div class="MAGALOG-card">', unsafe_allow_html=True)
            st.markdown(
                "<h4 style='color: #334155; margin-bottom: 15px;'><span class='icon-MAGALOG'>inventory_2</span> Detalhamento Financeiro por Tipo</h4>",
                unsafe_allow_html=True
            )

            df_visao = df_visao.loc[:, ~pd.Index(df_visao.columns).duplicated()].copy()

            colunas_detalhe = [
                'DATA AGENDADA',
                'AGENDA WMS',
                'FORNECEDOR/SELLER',
                'CATEGORIA',
                'LINHA',
                'TIPO_RECEITA',
                'STATUS',
                'VALOR_CONSIDERADO'
            ]
            colunas_detalhe = [c for c in colunas_detalhe if c in df_visao.columns]

            df_detalhe = df_visao[colunas_detalhe].copy()
            if 'DATA AGENDADA' in df_detalhe.columns:
                df_detalhe = df_detalhe.sort_values('DATA AGENDADA', ascending=False)

            st.dataframe(
                df_detalhe,
                column_config={
                    'DATA AGENDADA': st.column_config.DateColumn("Data"),
                    'AGENDA WMS': "Agenda",
                    'FORNECEDOR/SELLER': "Fornecedor / Seller",
                    'CATEGORIA': "Categoria",
                    'LINHA': "Linha",
                    'TIPO_RECEITA': "Tipo",
                    'STATUS': "Status",
                    'VALOR_CONSIDERADO': st.column_config.NumberColumn("Valor Considerado", format="R$ %.2f")
                },
                hide_index=True,
                use_container_width=True,
                height=420
            )

            st.markdown('</div>', unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Erro no módulo financeiro: {e}")
        
# ==========================================================
# MÓDULO 4: PRODUTIVIDADE, NS E DESEMPENHO
# ==========================================================
elif pagina_selecionada == "Produtividade (NS & Equipe)":
    render_hero(
        'Produtividade & Nível de Serviço',
        'Acompanhe SLA, produção por hora, fluxo de agendas e a evolução diária da operação.',
        'Analytics operacional'
    )
    try:
        with st.spinner("Calculando métricas de performance..."):
            df_fin = carregar_docas_finalizadas()
            
            if df_fin.empty:
                st.warning("Ainda não há dados de cargas finalizadas para gerar os indicadores.")
            else:
                # -----------------------------
                # Padronização das colunas
                # -----------------------------
                df_fin = df_fin.copy()
                df_fin.columns = [str(c).upper().strip() for c in df_fin.columns]
                
                col_data = 'DATA'
                col_agenda = 'AGENDA'
                col_tempo = 'TEMPO'
                col_just = 'JUSTIFICATIVA ATRASO'
                col_cat = 'CATEGORIA'
                col_aux = 'NOME'
                col_pecas = 'PEÇAS' if 'PEÇAS' in df_fin.columns else 'PECAS'
                col_doca_final = next((c for c in df_fin.columns if 'DOCA' in c), None)
                
                if 'M³' in df_fin.columns: col_m3 = 'M³'
                elif 'M3' in df_fin.columns: col_m3 = 'M3'
                else: col_m3 = None
                
                colunas_obrigatorias = [col_data, col_agenda, col_tempo, col_just, col_cat, col_aux, col_pecas]
                colunas_faltantes = [c for c in colunas_obrigatorias if c not in df_fin.columns]
                if col_m3 is None: colunas_faltantes.append('M³')
                
                if colunas_faltantes:
                    st.error(f"Não foi possível gerar os indicadores. Colunas ausentes na DOCAS_FINALIZADAS: {', '.join(colunas_faltantes)}")
                else:
                    def tempo_para_minutos(t_str):
                        try:
                            partes = str(t_str).strip().split(':')
                            if len(partes) == 2:
                                h, m = map(int, partes)
                                return h * 60 + m
                            return 0
                        except: return 0
                        
                    def minutos_para_texto(mins):
                        if pd.isna(mins) or mins == 0: return "00h00m"
                        h = int(mins // 60)
                        m = int(mins % 60)
                        return f"{h:02d}h{m:02d}m"
                        
                    def normalizar_numero_br(valor):
                        if pd.isna(valor): return 0.0
                        texto = str(valor).strip()
                        if texto == "" or texto.upper() in ["NAN", "NONE", "-"]: return 0.0
                        texto = texto.replace("R$", "").replace(" ", "")
                        if "." in texto and "," in texto: texto = texto.replace(".", "").replace(",", ".")
                        elif "," in texto: texto = texto.replace(",", ".")
                        try: return float(texto)
                        except: return 0.0
                        
                    df_fin[col_data] = pd.to_datetime(df_fin[col_data], dayfirst=True, errors='coerce')
                    df_fin = df_fin.dropna(subset=[col_data]).copy()
                    df_fin['MINUTOS'] = df_fin[col_tempo].apply(tempo_para_minutos)
                    df_fin = df_fin[df_fin['MINUTOS'] > 0].copy()
                    df_fin['HORAS'] = df_fin['MINUTOS'] / 60
                    df_fin['VAL_PECAS'] = df_fin[col_pecas].apply(normalizar_numero_br)
                    df_fin['VAL_M3'] = df_fin[col_m3].apply(normalizar_numero_br)
                    
                    c_filtro1, c_filtro2 = st.columns(2)
                    with c_filtro1: data_ini = st.date_input("Data inicial", value=df_fin[col_data].min().date())
                    with c_filtro2: data_fim = st.date_input("Data final", value=df_fin[col_data].max().date())
                    
                    mask_periodo = ((df_fin[col_data].dt.date >= data_ini) & (df_fin[col_data].dt.date <= data_fim))
                    df_fin = df_fin[mask_periodo].copy()
                    
                    if df_fin.empty:
                        st.warning("Não há dados no período selecionado.")
                    else:
                        # --- BASE MACRO CONSOLIDADA (1 agenda = 1 carga) ---
                        df_agendas_unicas = (
                            df_fin.groupby(col_agenda, as_index=False)
                            .agg({
                                col_data: 'first', col_cat: 'first', col_just: 'first',
                                'VAL_PECAS': 'first', 'VAL_M3': 'first', 'MINUTOS': 'first', 'HORAS': 'first',
                                col_doca_final: 'first'
                            }).copy()
                        )
                        total_cargas = df_agendas_unicas.shape[0]
                        
                        # --- LÓGICA DE CLASSIFICAÇÃO POR DOCA ---
                        def classificar_operacao(doca_val):
                            try:
                                num_doca = int(re.sub(r'\D', '', str(doca_val)))
                                if 58 <= num_doca <= 90:
                                    return "RECEBIMENTO"
                                else:
                                    return "EXPEDIÇÃO"
                            except:
                                return "EXPEDIÇÃO"
                                
                        df_agendas_unicas['OPERACAO_CALC'] = df_agendas_unicas[col_doca_final].apply(classificar_operacao)
                        
                        df_rec = df_agendas_unicas[df_agendas_unicas['OPERACAO_CALC'] == "RECEBIMENTO"]
                        df_exp = df_agendas_unicas[df_agendas_unicas['OPERACAO_CALC'] == "EXPEDIÇÃO"]
                        qtd_rec = df_rec.shape[0]
                        pecas_rec = df_rec['VAL_PECAS'].sum()
                        m3_rec = df_rec['VAL_M3'].sum()
                        qtd_exp = df_exp.shape[0]
                        pecas_exp = df_exp['VAL_PECAS'].sum()
                        m3_exp = df_exp['VAL_M3'].sum()
                        
                        total_horas_geral = df_agendas_unicas['HORAS'].sum()
                        tempo_medio_geral = df_agendas_unicas['MINUTOS'].mean()
                        total_pecas_geral = df_agendas_unicas['VAL_PECAS'].sum()
                        total_m3_geral = df_agendas_unicas['VAL_M3'].sum()
                        
                        media_pecas_hora = total_pecas_geral / total_horas_geral if total_horas_geral > 0 else 0
                        media_m3_hora = total_m3_geral / total_horas_geral if total_horas_geral > 0 else 0
                        
                                                # CÁLCULO DE FORA DO PRAZO E NO PRAZO
                        qtd_no_prazo = df_agendas_unicas[df_agendas_unicas[col_just].astype(str).str.upper().str.contains("NO PRAZO", na=False)].shape[0]
                        qtd_fora_prazo = total_cargas - qtd_no_prazo
                        
                        sla_percent = (qtd_no_prazo / total_cargas * 100) if total_cargas > 0 else 0
                        cor_sla = "#00C853" if sla_percent >= 90 else "#F59E0B" if sla_percent >= 75 else "#DC2626"
                        
                        # CORREÇÃO: Busca a coluna oficial de AUXILIARES da planilha para evitar rateio falso
                        col_qtd_aux = next((c for c in df_fin.columns if 'AUXILIAR' in c), None)
                        
                        if col_qtd_aux:
                            # Se a coluna existir, usa o número real gravado no banco
                            df_fin['QTD_AUXILIARES_AGENDA'] = pd.to_numeric(df_fin[col_qtd_aux], errors='coerce').fillna(0)
                            
                            # Fallback: Se alguma linha da coluna vier zerada, tenta contar os nomes
                            mask_zero = df_fin['QTD_AUXILIARES_AGENDA'] <= 0
                            if mask_zero.any():
                                df_fin.loc[mask_zero, 'QTD_AUXILIARES_AGENDA'] = df_fin[mask_zero].groupby(col_agenda)[col_aux].transform('nunique')
                        else:
                            # Se a coluna não existir, faz a contagem de nomes únicos
                            df_fin['QTD_AUXILIARES_AGENDA'] = df_fin.groupby(col_agenda)[col_aux].transform('nunique')
                        
                        # Trava final de segurança para nunca dividir por zero
                        df_fin['QTD_AUXILIARES_AGENDA'] = df_fin['QTD_AUXILIARES_AGENDA'].replace(0, 1)
                        
                        # Rateio justo e oficial
                        df_fin['PECAS_PART'] = df_fin['VAL_PECAS'] / df_fin['QTD_AUXILIARES_AGENDA']
                        df_fin['M3_PART'] = df_fin['VAL_M3'] / df_fin['QTD_AUXILIARES_AGENDA']
                        
                        # =========================================================
                        # REINSERINDO A DECLARAÇÃO DAS ABAS QUE HAVIA SUMIDO
                        # =========================================================
                        aba_macro, aba_equipe = st.tabs(["Visão Macro & NS", "Placar de Líderes (Equipe)"])
                        with aba_macro:
                            c1, c2, c3, c4, c5 = st.columns(5)
                            
                            # Formatação limpa dos números BR para evitar quebra no HTML
                            str_pcs_rec = f"{pecas_rec:,.0f}".replace(",", ".")
                            str_m3_rec = f"{m3_rec:,.1f}".replace(".", ",")
                            str_pcs_exp = f"{pecas_exp:,.0f}".replace(",", ".")
                            str_m3_exp = f"{m3_exp:,.1f}".replace(".", ",")
                            
                            with c1: 
                                st.markdown(f'<div class="kpi-card" style="border-top: 4px solid #0086FF; height: 130px; padding: 10px; text-align: center;"><div class="kpi-title">Total Agendas</div><div class="kpi-value" style="font-size:28px;">{total_cargas}</div></div>', unsafe_allow_html=True)
                                
                            with c2: 
                                st.markdown(f'<div class="kpi-card" style="border-top: 4px solid {cor_sla}; height: 130px; padding: 10px; text-align: center;"><div class="kpi-title">SLA Geral</div><div class="kpi-value" style="font-size:28px; color:{cor_sla};">{sla_percent:.1f}%</div><div style="font-size:11px; color:#64748B; font-weight:700; margin-top:4px; line-height:1.4;">{qtd_no_prazo} no prazo<br>{qtd_fora_prazo} atrasos</div></div>', unsafe_allow_html=True)
                                
                            with c3: 
                                st.markdown(f'<div class="kpi-card" style="border-top: 4px solid #8B5CF6; height: 130px; padding: 10px; text-align: center;"><div class="kpi-title">Média m³/H</div><div class="kpi-value" style="font-size:28px;">{media_m3_hora:.2f}</div></div>', unsafe_allow_html=True)
                                
                            with c4: 
                                st.markdown(f'<div class="kpi-card" style="border-top: 4px solid #0EA5E9; height: 130px; padding: 10px; text-align: center;"><div class="kpi-title">Recebimento</div><div class="kpi-value" style="font-size:28px;">{qtd_rec}</div><div style="font-size:11px; color:#64748B; font-weight:700; margin-top:4px; line-height:1.4;">{str_pcs_rec} pçs<br>{str_m3_rec} m³</div></div>', unsafe_allow_html=True)
                                
                            with c5: 
                                st.markdown(f'<div class="kpi-card" style="border-top: 4px solid #14B8A6; height: 130px; padding: 10px; text-align: center;"><div class="kpi-title">Expedição</div><div class="kpi-value" style="font-size:28px;">{qtd_exp}</div><div style="font-size:11px; color:#64748B; font-weight:700; margin-top:4px; line-height:1.4;">{str_pcs_exp} pçs<br>{str_m3_exp} m³</div></div>', unsafe_allow_html=True)
                            
                            # GRÁFICO EVOLUÇÃO DIÁRIA ATUALIZADO

                            
                            # GRÁFICO EVOLUÇÃO DIÁRIA ATUALIZADO
                            st.markdown('<div class="MAGALOG-card">', unsafe_allow_html=True)
                            st.markdown("<h4 style='color: #334155; margin-bottom: 20px;'><span class='icon-MAGALOG'>trending_up</span> Evolução Diária: Peças vs m³</h4>", unsafe_allow_html=True)
                            
                            df_daily = df_agendas_unicas.groupby(col_data).agg({'VAL_PECAS': 'sum', 'VAL_M3': 'sum'}).reset_index().sort_values(col_data)
                            fig_daily = make_subplots(specs=[[{"secondary_y": True}]])
                            
                            text_pecas = df_daily['VAL_PECAS'].apply(lambda x: f"<b>{int(x):,}</b>".replace(',', '.'))
                            fig_daily.add_trace(
                                go.Bar(
                                    x=df_daily[col_data], 
                                    y=df_daily['VAL_PECAS'], 
                                    name="Total Peças", 
                                    marker_color='#0086FF', 
                                    opacity=1.0, 
                                    text=text_pecas,
                                    textposition='outside', 
                                    textfont=dict(color='#0086FF', size=12),
                                    cliponaxis=False, 
                                    hovertemplate="Data: %{x}<br>Peças: %{y:,.0f}<extra></extra>"
                                ), 
                                secondary_y=False
                            )
                            
                            text_m3 = df_daily['VAL_M3'].apply(lambda x: f"<b>{x:.1f}</b>".replace('.', ','))
                            fig_daily.add_trace(
                                go.Scatter(
                                    x=df_daily[col_data], 
                                    y=df_daily['VAL_M3'], 
                                    name="Total m³", 
                                    line=dict(color='#ffcc00', width=4), 
                                    mode='lines+markers+text', 
                                    marker=dict(size=9, color='#FFFFFF', line=dict(width=3, color='#ffcc00')), 
                                    text=text_m3, 
                                    textposition='top center', 
                                    textfont=dict(color='#ffcc00', size=13),
                                    cliponaxis=False,
                                    hovertemplate="Data: %{x}<br>m³: %{y:,.2f}<extra></extra>"
                                ), 
                                secondary_y=True
                            )
                            
                            fig_daily.update_layout(
                                plot_bgcolor='rgba(0,0,0,0)', 
                                paper_bgcolor='rgba(0,0,0,0)', 
                                margin=dict(l=0, r=0, t=40, b=0), 
                                height=420, 
                                legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1), 
                                hovermode="x unified",
                                font_family="Inter"
                            )
                            
                            fig_daily.update_xaxes(showgrid=False, tickformat="%d/%m", tickfont=dict(color='#64748B', size=11, weight='bold'))
                            fig_daily.update_yaxes(title_text="Qtd Peças", secondary_y=False, showgrid=True, gridcolor='#F1F5F9', showticklabels=False, zeroline=False)
                            fig_daily.update_yaxes(title_text="Volume m³", secondary_y=True, showgrid=False, showticklabels=False, zeroline=False)
                            
                            st.plotly_chart(fig_daily, use_container_width=True, config={'displayModeBar': False})
                            st.markdown('</div>', unsafe_allow_html=True)
                            
                            col_g1, col_g2 = st.columns(2)
                            with col_g1:
                                st.markdown("""<div class="MAGALOG-card"><h4 style="color: #334155; margin-bottom: 15px;"><span class="icon-MAGALOG">format_list_bulleted</span> Detalhado por Categoria</h4>""", unsafe_allow_html=True)
                                
                                # Agrupamento Inteligente (reset_index garante alinhamento perfeito)
                                df_cat_table = df_agendas_unicas.groupby(col_cat).agg(
                                    CARGAS=(col_agenda, 'nunique'),
                                    MEDIA_PECAS=('VAL_PECAS', 'mean'),
                                    MINUTOS=('MINUTOS', 'mean')
                                ).reset_index().sort_values('MINUTOS', ascending=False).reset_index(drop=True)
                                
                                # Tratamento de Textos
                                df_cat_table['MÉDIA TEMPO'] = df_cat_table['MINUTOS'].apply(minutos_para_texto)
                                df_cat_table['MÉDIA PEÇAS'] = df_cat_table['MEDIA_PECAS'].apply(lambda x: f"{x:,.0f}".replace(',', '.'))
                                df_cat_table = df_cat_table.rename(columns={col_cat: 'CATEGORIA', 'CARGAS': 'QTD CARGAS'})
                                
                                # DataFrame apenas com as colunas visíveis (Adeus erro do .hide!)
                                df_exibir_cat = df_cat_table[['CATEGORIA', 'QTD CARGAS', 'MÉDIA PEÇAS', 'MÉDIA TEMPO']]
                                
                                # CSS Dinâmico (Tabela Bonitinha)
                                def estilizar_cat(df_visible):
                                    estilos = pd.DataFrame('', index=df_visible.index, columns=df_visible.columns)
                                    for idx, row in df_visible.iterrows():
                                        # Puxa o valor matemático dos minutos direto da base original, por trás dos panos
                                        m = df_cat_table.loc[idx, 'MINUTOS']
                                        
                                        # Lógica Semáforo para o Tempo
                                        if m > 120:
                                            estilos.loc[idx, 'MÉDIA TEMPO'] = 'color: #991B1B; background-color: #FEE2E2; font-weight: 800;'
                                        elif m > 60:
                                            estilos.loc[idx, 'MÉDIA TEMPO'] = 'color: #92400E; background-color: #FEF3C7; font-weight: 800;'
                                        else:
                                            estilos.loc[idx, 'MÉDIA TEMPO'] = 'color: #065F46; background-color: #D1FAE5; font-weight: 800;'
                                        
                                        estilos.loc[idx, 'CATEGORIA'] = 'font-weight: 800; color: #0F172A;'
                                        estilos.loc[idx, 'QTD CARGAS'] = 'color: #64748B; font-weight: 600; text-align: center;'
                                        estilos.loc[idx, 'MÉDIA PEÇAS'] = 'color: #0086FF; font-weight: 700;'
                                    return estilos
                                
                                # Renderizando a Tabela limpa e segura
                                st.dataframe(
                                    df_exibir_cat.style.apply(estilizar_cat, axis=None), 
                                    use_container_width=True, 
                                    hide_index=True,
                                    height=350
                                )
                                st.markdown('</div>', unsafe_allow_html=True)


                                
                            with col_g2:
                                st.markdown("""<div class="MAGALOG-card"><h4 style="color: #334155; margin-bottom: 15px;"><span class="icon-MAGALOG">pie_chart</span> Motivos de Atraso</h4>""", unsafe_allow_html=True)
                                
                                df_atrasos = df_agendas_unicas[~df_agendas_unicas[col_just].astype(str).str.upper().str.contains("NO PRAZO", na=False)]
                                
                                if not df_atrasos.empty:
                                    df_motivos = df_atrasos[col_just].value_counts().reset_index()
                                    df_motivos.columns = ['Motivo', 'Qtd']
                                    
                                    # Gráfico de Rosca Premium
                                    fig2 = px.pie(
                                        df_motivos, 
                                        values='Qtd', 
                                        names='Motivo', 
                                        hole=0.45, # Rosca mais gordinha
                                        color_discrete_sequence=px.colors.sequential.Reds_r
                                    )
                                    
                                    # Posicionamento inteligente da legenda (Lateral) e Margens (Respiro)
                                    fig2.update_layout(
                                        margin=dict(l=10, r=10, t=20, b=20), 
                                        height=350, 
                                        showlegend=True, 
                                        legend=dict(
                                            orientation="v", # Vertical
                                            yanchor="middle", 
                                            y=0.5, 
                                            xanchor="left", 
                                            x=1.05,          # Legenda à direita da pizza
                                            font=dict(size=10, color='#475569') # Fonte clean
                                        ),
                                        plot_bgcolor='rgba(0,0,0,0)', 
                                        paper_bgcolor='rgba(0,0,0,0)',
                                        font_family="Inter"
                                    )
                                    
                                    # Textos dentro da pizza mais limpos (Só a %)
                                    fig2.update_traces(
                                        textposition='inside', 
                                        textinfo='percent', 
                                        insidetextfont=dict(color='#FFFFFF', size=12, weight='bold'),
                                        hovertemplate="<b>%{label}</b><br>Quantidade: %{value}<br>Participação: %{percent}<extra></extra>"
                                    )
                                    
                                    st.plotly_chart(fig2, use_container_width=True, config={'displayModeBar': False})
                                else: 
                                    st.info("Nenhum atraso registrado no período.")
                                    
                                st.markdown('</div>', unsafe_allow_html=True)

                             # =========================================================
                            # NOVA TABELA: HISTÓRICO DE NÍVEL DE SERVIÇO E DESVIOS
                            # =========================================================
                            st.markdown("<br>", unsafe_allow_html=True)
                            st.markdown("""<div class="MAGALOG-card"><h4 style="color: #334155; margin-bottom: 15px;"><span class="icon-MAGALOG">assignment_late</span> Histórico de Nível de Serviço (SLA)</h4>""", unsafe_allow_html=True)
                            
                            # 1. Puxa as metas atuais da planilha auxiliar para cruzar os dados
                            df_aux_prod = carregar_aux()
                            dict_meta = {}
                            if not df_aux_prod.empty:
                                col_meta_aux = next((c for c in df_aux_prod.columns if 'META' in str(c).upper()), None)
                                col_ag_aux = next((c for c in df_aux_prod.columns if 'AGENDA' in str(c).upper()), None)
                                if col_meta_aux and col_ag_aux:
                                    for _, r in df_aux_prod.iterrows():
                                        try: dict_meta[str(r[col_ag_aux]).strip()] = int(float(str(r[col_meta_aux]).replace(',','.')))
                                        except: pass
                                        
                            def get_meta(ag): return dict_meta.get(str(ag).strip(), 60) # 60 min é a meta padrão do sistema
                            
                            df_ns = df_agendas_unicas.copy()
                            df_ns['META_MIN'] = df_ns[col_agenda].apply(get_meta)
                            
                            # A mágica matemática: Meta - Realizado
                            df_ns['DESVIO'] = df_ns['META_MIN'] - df_ns['MINUTOS']
                            
                            # Resgata o nome do Líder/Conferente da base original
                            col_conf = next((c for c in df_fin.columns if 'CONFERENTE' in c), None)
                            if col_conf:
                                map_conf = dict(zip(df_fin[col_agenda], df_fin[col_conf]))
                                df_ns['LÍDER'] = df_ns[col_agenda].map(map_conf)
                            
                            # Classificação de Cores e Textos para a Tabela
                            df_ns['STATUS'] = df_ns[col_just].apply(lambda x: 'NO PRAZO' if 'NO PRAZO' in str(x).upper() else 'ATRASADO')
                            df_ns['BALANÇO'] = df_ns['DESVIO'].apply(lambda x: f"{int(x)} min (Atraso)" if x < 0 else f"+{int(x)} min (Sobra)")
                            df_ns['TEMPO ESTIMADO'] = df_ns['META_MIN'].apply(minutos_para_texto)
                            df_ns['TEMPO REAL'] = df_ns['MINUTOS'].apply(minutos_para_texto)
                            df_ns['PEÇAS'] = df_ns['VAL_PECAS'].apply(lambda x: f"{int(x):,}".replace(',', '.'))
                            
                            # Seleção final das colunas visíveis
                            cols_ns = [col_agenda]
                            if col_conf: cols_ns.append('LÍDER')
                            cols_ns.extend([col_cat, 'PEÇAS', 'TEMPO ESTIMADO', 'TEMPO REAL', 'STATUS', 'BALANÇO', col_just])
                            
                            # Primeiro Ordena pelo DESVIO, DEPOIS corta as colunas visíveis
                            df_exibir_ns = df_ns.sort_values('DESVIO')[cols_ns]
                            
                            df_exibir_ns = df_exibir_ns.rename(columns={col_agenda: 'AGENDA', col_cat: 'CATEGORIA', col_just: 'JUSTIFICATIVA'})
                            
                            # Inteligência CSS (Cores do Painel)
                            def estilizar_ns(df_style):
                                estilos = pd.DataFrame('', index=df_style.index, columns=df_style.columns)
                                for idx, row in df_style.iterrows():
                                    if row['STATUS'] == 'NO PRAZO':
                                        estilos.loc[idx, 'STATUS'] = 'color: #065F46; background-color: #D1FAE5; font-weight: 800;'
                                        estilos.loc[idx, 'BALANÇO'] = 'color: #10B981; font-weight: 700;'
                                    else:
                                        estilos.loc[idx, 'STATUS'] = 'color: #991B1B; background-color: #FEE2E2; font-weight: 800;'
                                        estilos.loc[idx, 'BALANÇO'] = 'color: #EF4444; font-weight: 700;'
                                    
                                    if 'JUSTIFICATIVA' in df_style.columns:
                                        estilos.loc[idx, 'JUSTIFICATIVA'] = 'color: #64748B; font-size: 11px;'
                                    if 'AGENDA' in df_style.columns:
                                        estilos.loc[idx, 'AGENDA'] = 'font-weight: 800; color: #0F172A;'
                                        
                                    estilos.loc[idx, 'TEMPO ESTIMADO'] = 'color: #64748B; font-weight: 600;' # Estilo neutro para a meta
                                    estilos.loc[idx, 'TEMPO REAL'] = 'color: #0F172A; font-weight: 900;' # Destaque forte para o que realmente aconteceu
                                    
                                return estilos
                                
                            st.dataframe(df_exibir_ns.style.apply(estilizar_ns, axis=None), use_container_width=True, hide_index=True, height=400)
                            st.markdown('</div>', unsafe_allow_html=True)


                        with aba_equipe:
                            st.markdown("""
                            <style>
                            .lb-wrapper { background: rgba(255, 255, 255, 0.85); border: 1px solid rgba(255, 255, 255, 0.4); box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.07); backdrop-filter: blur(8px); border-radius: 16px; padding: 25px; margin-top: 15px;}
                            .lb-header { display: grid; grid-template-columns: 0.5fr 2.5fr 1fr 1fr 1.2fr 1.2fr 1.2fr 1.2fr; color: #64748B; font-weight: 800; font-size: 11px; text-transform: uppercase; padding: 0 15px 12px 15px; border-bottom: 2px solid #F1F5F9; margin-bottom: 12px; align-items: center;}
                            .lb-row { display: grid; grid-template-columns: 0.5fr 2.5fr 1fr 1fr 1.2fr 1.2fr 1.2fr 1.2fr; align-items: center; background: #FFFFFF; margin-bottom: 8px; padding: 12px 15px; border-radius: 12px; border: 1px solid #E2E8F0; transition: all 0.2s; border-left: 6px solid transparent;}
                            .lb-row:hover { transform: translateX(4px); box-shadow: 0 5px 15px rgba(0,0,0,0.05); }
                            .lb-gold { background: linear-gradient(90deg, #FFFBEB 0%, #FFFFFF 30%); border-left-color: #F59E0B; border-color: #FEF3C7;}
                            .lb-silver { background: linear-gradient(90deg, #F8FAFC 0%, #FFFFFF 30%); border-left-color: #94A3B8; border-color: #E2E8F0;}
                            .lb-bronze { background: linear-gradient(90deg, #FFF7ED 0%, #FFFFFF 30%); border-left-color: #B45309; border-color: #FFEDD5;}
                            .lb-danger { background: linear-gradient(90deg, #FEF2F2 0%, #FFFFFF 30%); border-left-color: #EF4444; border-color: #FEE2E2;}
                            .lb-rank { font-size: 18px; font-weight: 900; color: #0F172A; }
                            .lb-gold .lb-rank { color: #D97706; }
                            .lb-silver .lb-rank { color: #64748B; }
                            .lb-bronze .lb-rank { color: #92400E; }
                            .lb-danger .lb-rank { color: #DC2626; }
                            .lb-name { font-size: 14px; font-weight: 800; color: #1E293B; display: flex; align-items: center; gap: 8px;}
                            .lb-stat { font-size: 13px; font-weight: 600; color: #475569; }
                            .lb-highlight { background: #F0F9FF; color: #0284C7; padding: 4px 8px; border-radius: 6px; font-weight: 800; border: 1px solid #BAE6FD; display: inline-block; font-size: 12px;}
                            .lb-danger .lb-highlight { background: #FEF2F2; color: #DC2626; border-color: #FECACA; }
                            .lb-gold .lb-highlight { background: #FFFBEB; color: #D97706; border-color: #FDE68A; }
                            .lb-tooltip { position: relative; display: inline-block; cursor: help; color: #94A3B8; margin-left: 4px; vertical-align: middle;}
                            .lb-tooltip .lb-tooltiptext { visibility: hidden; width: 240px; background-color: #0F172A; color: #F8FAFC; text-align: left; border-radius: 8px; padding: 12px; position: absolute; z-index: 999; bottom: 130%; left: 50%; transform: translateX(-50%); opacity: 0; transition: opacity 0.3s, bottom 0.3s; font-size: 11px; font-weight: 500; text-transform: none; letter-spacing: normal; line-height: 1.4; box-shadow: 0 10px 25px rgba(0,0,0,0.2); border: 1px solid #334155; }
                            .lb-tooltip .lb-tooltiptext::after { content: ""; position: absolute; top: 100%; left: 50%; margin-left: -6px; border-width: 6px; border-style: solid; border-color: #334155 transparent transparent transparent; }
                            .lb-tooltip:hover .lb-tooltiptext { visibility: visible; opacity: 1; bottom: 150%; }
                            .lb-tooltip:hover .icon-MAGALOG { color: #0086FF; }
                            </style>
                            """, unsafe_allow_html=True)
                            
                            st.markdown("<h4 style='color: #0086FF; margin-bottom: 15px;'><span class='icon-MAGALOG'>social_leaderboard</span> Placar de Líderes Operacionais</h4>", unsafe_allow_html=True)
                            
                            c_rank_f1, c_rank_f2 = st.columns([3, 7])
                            with c_rank_f1:
                                cat_list_rank = sorted(df_fin[col_cat].dropna().astype(str).unique().tolist())
                                cat_sel_rank = st.selectbox("Filtrar Equipe por Categoria:", ["Todas as Categorias"] + cat_list_rank)
                            with c_rank_f2:
                                metric_sel = st.radio("Classificar placar por:", ["Total de Cargas", "Peças / Hora", "m³ / Hora", "Menor Tempo Médio"], horizontal=True)
                                
                            df_ops_rank = df_fin.copy()
                            if cat_sel_rank != "Todas as Categorias":
                                df_ops_rank = df_ops_rank[df_ops_rank[col_cat] == cat_sel_rank].copy()
                                
                            if not df_ops_rank.empty:
                                df_rank_data = (
                                    df_ops_rank.groupby(col_aux)
                                    .agg({
                                        col_agenda: 'nunique',
                                        'MINUTOS': 'mean',
                                        'HORAS': 'sum',
                                        'PECAS_PART': 'sum',
                                        'M3_PART': 'sum'
                                    }).reset_index()
                                )
                                df_rank_data['PECAS_H'] = (df_rank_data['PECAS_PART'] / df_rank_data['HORAS']).fillna(0)
                                df_rank_data['M3_H'] = (df_rank_data['M3_PART'] / df_rank_data['HORAS']).fillna(0)
                                
                                if metric_sel == "Total de Cargas": 
                                    df_rank_data = df_rank_data.sort_values(col_agenda, ascending=False).reset_index(drop=True)
                                elif metric_sel == "Peças / Hora": 
                                    df_rank_data = df_rank_data.sort_values('PECAS_H', ascending=False).reset_index(drop=True)
                                elif metric_sel == "m³ / Hora": 
                                    df_rank_data = df_rank_data.sort_values('M3_H', ascending=False).reset_index(drop=True)
                                else: 
                                    df_rank_data = df_rank_data.sort_values('MINUTOS', ascending=True).reset_index(drop=True)
                                
                                df_rank_data['Tempo_Medio'] = df_rank_data['MINUTOS'].apply(minutos_para_texto)
                                
                                html_lb = "<div class='lb-wrapper'>"
                                html_lb += "<div class='lb-header'>"
                                html_lb += "<div>POS</div><div>OPERADOR</div>"
                                html_lb += "<div>CARGAS <div class='lb-tooltip'><span class='icon-MAGALOG' style='font-size:14px;'>help</span><div class='lb-tooltiptext'>Total de agendas participadas.</div></div></div>"
                                html_lb += "<div>T. MÉDIO <div class='lb-tooltip'><span class='icon-MAGALOG' style='font-size:14px;'>help</span><div class='lb-tooltiptext'>Média de tempo gasto por carga.</div></div></div>"
                                html_lb += "<div>PÇS TOTAIS <div class='lb-tooltip'><span class='icon-MAGALOG' style='font-size:14px;'>help</span><div class='lb-tooltiptext'>Volume total rateado pela equipe.</div></div></div>"
                                html_lb += "<div>M³ TOTAIS <div class='lb-tooltip'><span class='icon-MAGALOG' style='font-size:14px;'>help</span><div class='lb-tooltiptext'>m³ total rateado pela equipe.</div></div></div>"
                                html_lb += "<div>PEÇAS / H <div class='lb-tooltip'><span class='icon-MAGALOG' style='font-size:14px;'>help</span><div class='lb-tooltiptext'>Eficiência: Peças ÷ Horas.</div></div></div>"
                                html_lb += "<div>M³ / H <div class='lb-tooltip'><span class='icon-MAGALOG' style='font-size:14px;'>help</span><div class='lb-tooltiptext'>Eficiência: m³ ÷ Horas.</div></div></div>"
                                html_lb += "</div>"
                                
                                total_ops = len(df_rank_data)
                                
                                for idx, row_r in df_rank_data.iterrows():
                                    pos_r = idx + 1
                                    css_c = "lb-gold" if pos_r == 1 else "lb-silver" if pos_r == 2 else "lb-bronze" if pos_r == 3 else "lb-danger" if (pos_r >= total_ops - 2 and total_ops >= 6) else ""
                                    ic_r = "workspace_premium" if pos_r == 1 else "military_tech" if pos_r <= 3 else "warning" if "danger" in css_c else "person"
                                    
                                    html_lb += f"<div class='lb-row {css_c}'>"
                                    html_lb += f"<div class='lb-rank'>{pos_r}º</div>"
                                    html_lb += f"<div class='lb-name'><span class='icon-MAGALOG' style='font-size:20px;'>{ic_r}</span> {row_r[col_aux]}</div>"
                                    html_lb += f"<div class='lb-stat'>{int(row_r[col_agenda])}</div>"
                                    html_lb += f"<div class='lb-stat'>{row_r['Tempo_Medio']}</div>"
                                    html_lb += f"<div class='lb-stat'>{row_r['PECAS_PART']:.1f}</div>"
                                    html_lb += f"<div class='lb-stat'>{row_r['M3_PART']:.2f}</div>"
                                    html_lb += f"<div class='lb-stat'><span class='lb-highlight'>{row_r['PECAS_H']:.1f} pç/h</span></div>"
                                    html_lb += f"<div class='lb-stat'><span class='lb-highlight'>{row_r['M3_H']:.2f} m³/h</span></div></div>"
                                

                                html_lb += "</div>"
                                st.markdown(html_lb, unsafe_allow_html=True)
                                
                                # =========================================================
                                # NOVO BLOCO: RADAR DA LIDERANÇA (INSIGHTS COMPORTAMENTAIS)
                                # =========================================================
                                if len(df_rank_data) >= 3:
                                    st.markdown("<br><h4 style='color: #0F172A; margin-bottom: 12px;'><span class='icon-MAGALOG' style='color: #0086FF;'>radar</span> Radar da Liderança (Análise de Perfil)</h4>", unsafe_allow_html=True)
                                    
                                    # Médias dinâmicas baseadas na equipe atual da tela
                                    avg_pecas_h = df_rank_data['PECAS_H'].mean()
                                    avg_m3_h = df_rank_data['M3_H'].mean()
                                    avg_cargas = df_rank_data[col_agenda].mean()
                                    max_cargas = df_rank_data[col_agenda].max()
                                    
                                    insights_html = ""
                                    alertas_gerados = 0
                                    
                                    for idx, row in df_rank_data.iterrows():
                                        if alertas_gerados >= 5: break # Limita a 5 insights para não lotar a tela
                                        
                                        nome = row[col_aux]
                                        pecas_h = row['PECAS_H']
                                        m3_h = row['M3_H']
                                        cargas = row[col_agenda]
                                        
                                        # REGRA 1: Alta Velocidade, Baixa Participação (Ex: O caso do Robson no seu print)
                                        if (pecas_h > avg_pecas_h * 1.15 or m3_h > avg_m3_h * 1.15) and cargas <= (avg_cargas * 0.8):
                                            insights_html += f"""<div style="background: #FFFBEB; border: 1px solid #FDE68A; border-left: 4px solid #F59E0B; border-radius: 8px; padding: 12px; margin-bottom: 8px; display: flex; align-items: flex-start; gap: 10px;">
                                                <span class='icon-MAGALOG' style='color: #F59E0B; font-size: 20px; margin-top: 2px;'>bolt</span>
                                                <div style="font-size: 13px; color: #475569; line-height: 1.4;"><b>{nome}</b> possui eficiência altíssima ({pecas_h:.1f} pçs/h e {m3_h:.1f} m³/h), mas participou de poucas agendas ({int(cargas)} cargas). <span style="color:#B45309; font-weight:700;">Atenção:</span> Pode indicar "escolha" de cargas muito fáceis ("cherry-picking") ou atuação de meio turno.</div>
                                            </div>"""
                                            alertas_gerados += 1
                                            
                                        # REGRA 2: Baixa Eficiência Crônica (Ex: O caso do Jose no seu print)
                                        elif pecas_h < (avg_pecas_h * 0.85) and m3_h < (avg_m3_h * 0.85) and cargas >= (avg_cargas * 0.7):
                                            insights_html += f"""<div style="background: #FEF2F2; border: 1px solid #FECACA; border-left: 4px solid #DC2626; border-radius: 8px; padding: 12px; margin-bottom: 8px; display: flex; align-items: flex-start; gap: 10px;">
                                                <span class='icon-MAGALOG' style='color: #DC2626; font-size: 20px; margin-top: 2px;'>trending_down</span>
                                                <div style="font-size: 13px; color: #475569; line-height: 1.4;"><b>{nome}</b> assumiu o giro de cargas ({int(cargas)} agendas), mas sua velocidade de entrega ({pecas_h:.1f} pçs/h) está bem abaixo da média do grupo ({avg_pecas_h:.0f} pçs/h). Pode indicar dificuldade no manuseio, fadiga ou excesso de alocação em piores cargas.</div>
                                            </div>"""
                                            alertas_gerados += 1

                                        # REGRA 3: Ociosidade Crítica (Ex: O caso do Leonardo no seu print)
                                        elif cargas < (avg_cargas * 0.5):
                                            insights_html += f"""<div style="background: #F1F5F9; border: 1px solid #E2E8F0; border-left: 4px solid #64748B; border-radius: 8px; padding: 12px; margin-bottom: 8px; display: flex; align-items: flex-start; gap: 10px;">
                                                <span class='icon-MAGALOG' style='color: #64748B; font-size: 20px; margin-top: 2px;'>person_off</span>
                                                <div style="font-size: 13px; color: #475569; line-height: 1.4;"><b>{nome}</b> atuou em um volume criticamente baixo de cargas ({int(cargas)} agendas vs média de {int(avg_cargas)}). Validar urgentemente se ocorreu falta, atestado, desvio de função ou ociosidade oculta.</div>
                                            </div>"""
                                            alertas_gerados += 1
                                            
                                        # REGRA 4: Foco Desbalanceado no Pesado
                                        elif m3_h > (avg_m3_h * 1.25) and pecas_h < (avg_pecas_h * 0.75):
                                            insights_html += f"""<div style="background: #F5F3FF; border: 1px solid #DDD6FE; border-left: 4px solid #8B5CF6; border-radius: 8px; padding: 12px; margin-bottom: 8px; display: flex; align-items: flex-start; gap: 10px;">
                                                <span class='icon-MAGALOG' style='color: #8B5CF6; font-size: 20px; margin-top: 2px;'>fitness_center</span>
                                                <div style="font-size: 13px; color: #475569; line-height: 1.4;"><b>{nome}</b> tem alto volume em m³, mas baixo em peças. Sinal claro de alocação excessiva em cargas de grande porte/pesadas. Checar se o rodízio de esforço na doca está justo.</div>
                                            </div>"""
                                            alertas_gerados += 1
                                            
                                        # REGRA 5: O Carregador de Piano / Motor da Operação (Ex: Sandro e Adilson no print)
                                        elif cargas >= (max_cargas * 0.9) and pecas_h >= avg_pecas_h and m3_h >= (avg_m3_h * 0.9):
                                            insights_html += f"""<div style="background: #ECFDF5; border: 1px solid #A7F3D0; border-left: 4px solid #10B981; border-radius: 8px; padding: 12px; margin-bottom: 8px; display: flex; align-items: flex-start; gap: 10px;">
                                                <span class='icon-MAGALOG' style='color: #10B981; font-size: 20px; margin-top: 2px;'>local_fire_department</span>
                                                <div style="font-size: 13px; color: #475569; line-height: 1.4;"><b>{nome}</b> é o motor da operação hoje. Absorveu o maior volume de cargas da doca ({int(cargas)} agendas) sem perder a velocidade, com indicadores excelentes. Perfil de alta tração.</div>
                                            </div>"""
                                            alertas_gerados += 1
                                    
                                    if insights_html == "":
                                        insights_html = "<div style='font-size: 13px; color: #64748B; background: #F8FAFC; padding: 12px; border-radius: 8px; border: 1px dashed #CBD5E1;'><span class='icon-MAGALOG' style='vertical-align: middle;'>check_circle</span> A equipe apresenta uma distribuição de cargas bem equalizada.</div>"
                                        
                                    st.markdown(insights_html, unsafe_allow_html=True)
                                # =========================================================

                                
                            else: 
                                st.info("Sem dados para esta categoria.")
    except Exception as e:
        st.error(f"Erro no módulo de Produtividade: {e}")

# ==========================================================
# MÓDULO 6: GESTÃO DE PESSOAS E IMPACTO (RH)
# ==========================================================
elif pagina_selecionada == "Absenteísmo (RH)":
    render_hero(
        'Análise de Absenteísmo e Impacto',
        'Cruze os dados de faltas com a produtividade do CD para descobrir o real prejuízo operacional gerado pelas ausências.',
        'Recursos Humanos • Analytics'
    )
    
    try:
        with st.spinner("Analisando impacto do absenteísmo..."):
            df_abs = carregar_log_absenteismo()
            
            if df_abs.empty:
                st.warning("O Log de Absenteísmo está vazio ou não foi encontrado.")
            else:
                # 1. Tratamento de Datas e Filtros
                df_abs['DATA'] = pd.to_datetime(df_abs['DATA'], dayfirst=True, errors='coerce')
                df_abs = df_abs.dropna(subset=['DATA']).copy()
                
                st.sidebar.markdown("### <span class='icon-MAGALOG'>filter_alt</span> Filtros de RH", unsafe_allow_html=True)
                
                c_f1, c_f2 = st.columns(2)
                with c_f1: dt_ini = st.date_input("Data Inicial", value=df_abs['DATA'].min().date())
                with c_f2: dt_fim = st.date_input("Data Final", value=df_abs['DATA'].max().date())
                
                # Filtro inteligente de Motivos (Por padrão, foca no que causa prejuízo real)
                col_ocorrencia = next((c for c in df_abs.columns if 'OCORR' in c), 'OCORRÊNCIA')
                motivos_unicos = df_abs[col_ocorrencia].dropna().unique().tolist()
                motivos_default = [m for m in motivos_unicos if str(m).upper() in ['FALTA', 'ATESTADO']]
                if not motivos_default: motivos_default = motivos_unicos # Se não tiver falta/atestado, marca tudo
                
                motivos_sel = st.sidebar.multiselect("Motivos Considerados (Absenteísmo):", options=motivos_unicos, default=motivos_default)
                
                # Aplicação dos Filtros
                mask = (df_abs['DATA'].dt.date >= dt_ini) & (df_abs['DATA'].dt.date <= dt_fim) & (df_abs[col_ocorrencia].isin(motivos_sel))
                df_filtrado = df_abs[mask].copy()
                
                if df_filtrado.empty:
                    st.info("Nenhuma ocorrência encontrada para os filtros selecionados.")
                else:
                    # Funções de conversão embutidas localmente para evitar o erro de escopo
                    def limpa_numero_br(valor):
                        if pd.isna(valor): return 0.0
                        texto = str(valor).strip()
                        if texto == "" or texto.upper() in ["NAN", "NONE", "-"]: return 0.0
                        texto = texto.replace("R$", "").replace(" ", "")
                        if "." in texto and "," in texto: texto = texto.replace(".", "").replace(",", ".")
                        elif "," in texto: texto = texto.replace(",", ".")
                        try: return float(texto)
                        except: return 0.0
                        
                    def time_to_mins(t_str):
                        try:
                            partes = str(t_str).strip().split(':')
                            if len(partes) == 2:
                                h, m = map(int, partes)
                                return h * 60 + m
                            elif len(partes) == 3:
                                return int(partes[0]) * 60 + int(partes[1]) + float(partes[2]) / 60.0
                            return 0
                        except: return 0

                    # 2. Matemática do Impacto (Calculando a velocidade real do CD em background)
                    df_prod = carregar_docas_finalizadas()
                    avg_pecas_h = 0
                    avg_m3_h = 0
                    
                    if not df_prod.empty:
                        df_prod.columns = [str(c).upper().strip() for c in df_prod.columns]
                        col_ag = next((c for c in df_prod.columns if 'AGENDA' in c), None)
                        col_tmp = next((c for c in df_prod.columns if 'TEMPO' in c), None)
                        col_pc = next((c for c in df_prod.columns if 'PEÇAS' in c or 'PECA' in c), None)
                        col_m3 = next((c for c in df_prod.columns if 'M³' in c or 'M3' in c), None)
                        
                        if col_ag and col_tmp and col_pc:
                            df_p = df_prod.copy()
                            df_p['MINS'] = df_p[col_tmp].apply(time_to_mins)
                            df_p['HORAS'] = df_p['MINS'] / 60
                            df_p['PÇS'] = df_p[col_pc].apply(limpa_numero_br)
                            df_p['M3'] = df_p[col_m3].apply(limpa_numero_br) if col_m3 else 0
                            
                            # Pega 1 linha por agenda para não somar peças duplicadas pelo número de ajudantes
                            df_agendas = df_p.groupby(col_ag).first()
                            tot_h = df_agendas['HORAS'].sum()
                            if tot_h > 0:
                                avg_pecas_h = df_agendas['PÇS'].sum() / tot_h
                                avg_m3_h = df_agendas['M3'].sum() / tot_h
                                
                    # 3. KPIs de Absenteísmo
                    dias_com_registro = max(1, df_filtrado['DATA'].nunique()) # Evita divisão por zero
                    total_ausencias = len(df_filtrado)
                    media_ausencias_dia = total_ausencias / dias_com_registro
                    
                    # Cada pessoa = 427 minutos = 7.11 horas
                    horas_perdidas_dia = (media_ausencias_dia * 427) / 60
                    pecas_perdidas_dia = horas_perdidas_dia * avg_pecas_h
                    m3_perdidos_dia = horas_perdidas_dia * avg_m3_h
                    
                    st.markdown("<h4 style='color: #334155; margin-bottom: 15px;'><span class='icon-MAGALOG'>warning</span> Impacto Médio Diário na Operação</h4>", unsafe_allow_html=True)
                    
                    # Função construtora de KPI com o padrão visual simétrico (130px) do Carga e Descarga
                    def renderizar_kpi_rh(titulo, valor, subtitulo, cor):
                        st.markdown(f'''
                        <div class="kpi-card" style="border-top: 4px solid {cor}; height: 130px; padding: 10px; text-align: center;">
                            <div class="kpi-title">{titulo}</div>
                            <div class="kpi-value" style="font-size:28px; color:{cor};">{valor}</div>
                            <div style="font-size:11px; color:#64748B; font-weight:700; margin-top:4px; line-height:1.4;">{subtitulo}</div>
                        </div>
                        ''', unsafe_allow_html=True)
                    
                    c1, c2, c3, c4 = st.columns(4)
                    with c1: renderizar_kpi_rh("Faltas/Atestados por Dia", f"{media_ausencias_dia:.1f}", f"Total de {total_ausencias} no período", "#E74C3C")
                    with c2: renderizar_kpi_rh("Horas Produtivas Perdidas", f"{horas_perdidas_dia:.1f}h", "Base: 427 min/colaborador", "#F39C12")
                    with c3: renderizar_kpi_rh("Perda produtiva em Peças", f"{pecas_perdidas_dia:,.0f}".replace(',','.'), f"Veloc. CD: {avg_pecas_h:.0f} pçs/h", "#8B5CF6")
                    with c4: renderizar_kpi_rh("Perda produtivda em M³", f"{m3_perdidos_dia:.1f}".replace('.',','), f"Veloc. CD: {avg_m3_h:.1f} m³/h", "#0086FF")
                    
                    st.markdown("<br>", unsafe_allow_html=True)

                    
                    # 4. Gráficos Analíticos
                    col_g1, col_g2 = st.columns(2)
                    with col_g1:
                        st.markdown("""<div class="MAGALOG-card"><h4 style="color: #334155; margin-bottom: 15px;"><span class="icon-MAGALOG">calendar_month</span> Reincidência por Dia da Semana</h4>""", unsafe_allow_html=True)
                        
                        # Traduzir dias da semana
                        dias_semana = {0: 'Segunda', 1: 'Terça', 2: 'Quarta', 3: 'Quinta', 4: 'Sexta', 5: 'Sábado', 6: 'Domingo'}
                        df_filtrado['DIA_SEMANA_NUM'] = df_filtrado['DATA'].dt.weekday
                        df_filtrado['DIA_SEMANA'] = df_filtrado['DIA_SEMANA_NUM'].map(dias_semana)
                        
                        df_semana = df_filtrado.groupby(['DIA_SEMANA_NUM', 'DIA_SEMANA']).size().reset_index(name='Qtd')
                        
                        fig_sem = px.bar(df_semana, x='DIA_SEMANA', y='Qtd', text='Qtd', color_discrete_sequence=['#FF3366'])
                        fig_sem.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(l=0, r=0, t=10, b=0), height=300, xaxis_title=None, yaxis_title=None)
                        fig_sem.update_traces(textposition='outside', textfont=dict(weight='bold'))
                        fig_sem.update_xaxes(showgrid=False)
                        fig_sem.update_yaxes(showgrid=True, gridcolor='#F1F5F9', showticklabels=False)
                        st.plotly_chart(fig_sem, use_container_width=True, config={'displayModeBar': False})
                        st.markdown('</div>', unsafe_allow_html=True)
                        
                    with col_g2:
                        st.markdown("""<div class="MAGALOG-card"><h4 style="color: #334155; margin-bottom: 15px;"><span class="icon-MAGALOG">date_range</span> Reincidência por Dia do Mês</h4>""", unsafe_allow_html=True)
                        
                        df_filtrado['DIA_MES'] = df_filtrado['DATA'].dt.day
                        df_mes = df_filtrado.groupby('DIA_MES').size().reset_index(name='Qtd')
                        
                        # Preenche os dias que não tiveram falta com 0 para o gráfico não pular números
                        todos_dias = pd.DataFrame({'DIA_MES': range(1, 32)})
                        df_mes = pd.merge(todos_dias, df_mes, on='DIA_MES', how='left').fillna(0)
                        
                        fig_mes = go.Figure(go.Scatter(
                            x=df_mes['DIA_MES'], y=df_mes['Qtd'],
                            mode='lines+markers+text',
                            line=dict(color='#8B5CF6', width=3),
                            marker=dict(size=8, color='#FFFFFF', line=dict(width=2, color='#8B5CF6')),
                            text=df_mes['Qtd'].apply(lambda x: int(x) if x > 0 else ""),
                            textposition='top center',
                            textfont=dict(weight='bold', color='#8B5CF6')
                        ))
                        fig_mes.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(l=0, r=0, t=10, b=0), height=300, xaxis=dict(tickmode='linear', dtick=2))
                        fig_mes.update_yaxes(showgrid=False, showticklabels=False)
                        st.plotly_chart(fig_mes, use_container_width=True, config={'displayModeBar': False})
                        st.markdown('</div>', unsafe_allow_html=True)

                    # 5. Hall da Vergonha (Ranking Negativo)
                    st.markdown("<br><h4 style='color: #DC2626; margin-bottom: 15px;'><span class='icon-MAGALOG'>policy</span> Ranking de Absenteísmo (Maiores Ocorrências)</h4>", unsafe_allow_html=True)
                    
                    # Agrupa e pivota a tabela
                    df_rank = df_filtrado.groupby(['NOME', col_ocorrencia]).size().unstack(fill_value=0).reset_index()
                    
                    # Cria a coluna TOTAL somando as colunas de motivo
                    motivos_cols = [c for c in df_rank.columns if c != 'NOME']
                    df_rank['TOTAL'] = df_rank[motivos_cols].sum(axis=1)
                    
                    df_rank = df_rank.sort_values('TOTAL', ascending=False).reset_index(drop=True)
                    
                    # Estilização CSS Negativa
                    def estilizar_hall_vergonha(df_v):
                        estilos = pd.DataFrame('', index=df_v.index, columns=df_v.columns)
                        for idx, row in df_v.iterrows():
                            # Se a pessoa tem mais de 2 faltas/atestados, pinta de vermelho forte
                            if row['TOTAL'] >= 3:
                                estilos.loc[idx, 'TOTAL'] = 'color: #991B1B; background-color: #FEE2E2; font-weight: 900; text-align: center;'
                            elif row['TOTAL'] == 2:
                                estilos.loc[idx, 'TOTAL'] = 'color: #92400E; background-color: #FEF3C7; font-weight: 800; text-align: center;'
                            else:
                                estilos.loc[idx, 'TOTAL'] = 'color: #0F172A; font-weight: 700; text-align: center;'
                                
                            estilos.loc[idx, 'NOME'] = 'font-weight: 800; color: #1E293B;'
                        return estilos

                                        # Formata a exibição
                    st.dataframe(
                        df_rank.style.apply(estilizar_hall_vergonha, axis=None), 
                        use_container_width=True, 
                        hide_index=True, 
                        height=400,
                        column_config={"TOTAL": st.column_config.NumberColumn("TOTAL DE OCORRÊNCIAS")}
                    )

                                        # =========================================================
                    # 6. GRÁFICO DE IMPACTO DIÁRIO (CAPACIDADE PRODUTIVA PERDIDA)
                    # =========================================================
                    st.markdown("<br><h4 style='color: #334155; margin-bottom: 15px;'><span class='icon-MAGALOG'>trending_down</span> Capacidade Produtiva Perdida por Dia</h4>", unsafe_allow_html=True)
                    
                    # BLINDAGEM DE DATA: Extrai apenas a data pura, ignorando horas
                    df_filtrado['DATA_PURA'] = df_filtrado['DATA'].dt.date
                    
                    # Agrupa as faltas por data exata
                    df_perdas_dia = df_filtrado.groupby('DATA_PURA').size().reset_index(name='Qtd_Faltas')
                    
                    # Calcula a perda real de cada dia
                    df_perdas_dia['Horas_Perdidas'] = df_perdas_dia['Qtd_Faltas'] * (427 / 60.0)
                    df_perdas_dia['Pecas_Perdidas'] = df_perdas_dia['Horas_Perdidas'] * avg_pecas_h
                    df_perdas_dia['M3_Perdidos'] = df_perdas_dia['Horas_Perdidas'] * avg_m3_h
                    
                    # Cria um calendário completo com o tipo DATE exato
                    todas_datas = pd.date_range(start=dt_ini, end=dt_fim).date
                    df_todas_datas = pd.DataFrame({'DATA_PURA': todas_datas})
                    
                    # Cruzamento blindado (agora os dois lados são tipo 'date')
                    df_perdas_dia = pd.merge(df_todas_datas, df_perdas_dia, on='DATA_PURA', how='left').fillna(0)
                    df_perdas_dia = df_perdas_dia.sort_values('DATA_PURA')
                    
                    st.markdown('<div class="MAGALOG-card">', unsafe_allow_html=True)
                    
                    # Monta o gráfico de eixo duplo com visual Premium (Spline + Area)
                    fig_perdas = make_subplots(specs=[[{"secondary_y": True}]])
                    
                    # Barras Laranjas Modernas (Peças)
                    text_pecas_perdidas = df_perdas_dia['Pecas_Perdidas'].apply(lambda x: f"<b>{int(x):,}</b>".replace(',', '.') if x > 0 else "")
                    fig_perdas.add_trace(
                        go.Bar(
                            x=df_perdas_dia['DATA_PURA'], 
                            y=df_perdas_dia['Pecas_Perdidas'], 
                            name="Peças Perdidas", 
                            marker_color='#FF9F43', # Laranja Moderno (Neon/Soft)
                            opacity=0.85,
                            text=text_pecas_perdidas,
                            textposition='outside',
                            textfont=dict(color='#D97706', size=11, family="Inter"),
                            cliponaxis=False,
                            hovertemplate="Data: %{x|%d/%m/%Y}<br>Faltas no dia: %{customdata}<br>Peças Perdidas: %{y:,.0f}<extra></extra>",
                            customdata=df_perdas_dia['Qtd_Faltas']
                        ), 
                        secondary_y=False
                    )
                    
                    # Curva Suave Azul com Sombreamento (m³)
                    text_m3_perdidos = df_perdas_dia['M3_Perdidos'].apply(lambda x: f"<b>{x:.1f}</b>".replace('.', ',') if x > 0 else "")
                    fig_perdas.add_trace(
                        go.Scatter(
                            x=df_perdas_dia['DATA_PURA'], 
                            y=df_perdas_dia['M3_Perdidos'], 
                            name="m³ Perdido", 
                            mode='lines+markers+text',
                            line=dict(color='#3B82F6', width=3, shape='spline'), # Azul Moderno, Linha Curvada (Spline)
                            fill='tozeroy', # Sombreado de área premium embaixo da curva
                            fillcolor='rgba(59, 130, 246, 0.15)', # Azul translúcido
                            marker=dict(size=7, color='#FFFFFF', line=dict(width=2, color='#3B82F6')),
                            text=text_m3_perdidos,
                            textposition='top center',
                            textfont=dict(color='#2563EB', size=11, family="Inter"),
                            cliponaxis=False,
                            hovertemplate="Data: %{x|%d/%m/%Y}<br>m³ Perdido: %{y:,.1f}<extra></extra>"
                        ), 
                        secondary_y=True
                    )
                    
                    # Layout premium e clean
                    fig_perdas.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)', 
                        paper_bgcolor='rgba(0,0,0,0)', 
                        margin=dict(l=10, r=10, t=40, b=10), 
                        height=380, 
                        legend=dict(
                            orientation="h", 
                            yanchor="bottom", 
                            y=1.05, 
                            xanchor="right", 
                            x=1,
                            font=dict(size=12, color='#475569', family="Inter")
                        ),
                        hovermode="x unified",
                        font_family="Inter"
                    )
                    
                    # Eixos X e Y super limpos
                    fig_perdas.update_xaxes(showgrid=False, tickformat="%d/%m", tickfont=dict(color='#64748B', size=11, weight='bold'))
                    fig_perdas.update_yaxes(secondary_y=False, showgrid=True, gridcolor='#F1F5F9', showticklabels=False, zeroline=False)
                    fig_perdas.update_yaxes(secondary_y=True, showgrid=False, showticklabels=False, zeroline=False)
                    
                    st.plotly_chart(fig_perdas, use_container_width=True, config={'displayModeBar': False})
                    st.markdown('</div>', unsafe_allow_html=True)


    except Exception as e:
        st.error(f"Erro no módulo de Gestão de Pessoas: {e}")

# ==========================================================
# MÓDULO 7: GERADOR INTELIGENTE DE EQUIPES (I.A. ESCALAÇÃO)
# ==========================================================
elif pagina_selecionada == "Gerador de Equipes (I.A.)":
    import random 
    
    render_hero(
        'Gerador Inteligente de Equipes',
        'Defina a demanda do dia e deixe o algoritmo escalar a operação balanceando habilidades, restrições e evitando repetições de equipes do dia anterior.',
        'MAGALOG • I.A. Operacional'
    )
    
    df_equipe = carregar_equipe()
    df_matriz = carregar_matriz()
    df_historico_ia = carregar_historico_escalas()
    
    if df_equipe.empty:
        st.warning("Não foi possível carregar a equipe da base de dados.")
    else:
        # ==========================================
        # 1. FILTRO DE TURNO E MEMÓRIA DA I.A.
        # ==========================================
        col_turno = next((c for c in df_equipe.columns if 'TURNO' in c.upper()), None)
        
        st.markdown("<h4 style='color: #0F172A; margin-bottom: 15px;'><span class='icon-MAGALOG' style='color:#8B5CF6;'>schedule</span> 1. Filtro de Turno</h4>", unsafe_allow_html=True)
        
        turnos_disponiveis = sorted([str(t).strip() for t in df_equipe[col_turno].dropna().unique() if str(t).strip() != '']) if col_turno else ["Turno Único"]
        turno_sel = st.multiselect("Selecione os Turnos que deseja escalar agora:", options=turnos_disponiveis, default=turnos_disponiveis[:1] if turnos_disponiveis else None)
        
        # Filtra a equipe baseada no turno selecionado
        if col_turno and turno_sel:
            df_equipe_filtrada = df_equipe[df_equipe[col_turno].isin(turno_sel)].copy()
            turno_str = " + ".join(turno_sel)
        else:
            df_equipe_filtrada = df_equipe.copy()
            turno_str = "Geral"

        todos_colaboradores = sorted([str(n).strip() for n in df_equipe_filtrada['NOME'].dropna().unique() if str(n).strip() != ''])
        
        # --- O CÉREBRO DA I.A. (LENDO O DIA DE ONTEM) ---
        historico_restricoes = {}
        data_memoria = None
        
        if not df_historico_ia.empty and turno_sel:
            df_hist_turno = df_historico_ia[df_historico_ia['TURNO'] == turno_str].copy()
            if not df_hist_turno.empty:
                df_hist_turno['DATA_DT'] = pd.to_datetime(df_hist_turno['DATA'], format='%d/%m/%Y', errors='coerce')
                ultima_data = df_hist_turno['DATA_DT'].max()
                
                if pd.notnull(ultima_data):
                    data_memoria = ultima_data.strftime('%d/%m/%Y')
                    df_ontem = df_hist_turno[df_hist_turno['DATA_DT'] == ultima_data]
                    
                    for _, row in df_ontem.iterrows():
                        membros_str = str(row.get('MEMBROS', ''))
                        membros_lista = [m.strip() for m in membros_str.split(',') if m.strip()]
                        
                        # Cria teia de aranha de restrições (Quem trabalhou junto ontem não pode hoje)
                        for m1 in membros_lista:
                            if m1 not in historico_restricoes: historico_restricoes[m1] = set()
                            for m2 in membros_lista:
                                if m1 != m2: historico_restricoes[m1].add(m2)
        
        if data_memoria:
            st.markdown(f"<div style='background: #F0FDF4; border: 1px solid #A7F3D0; padding: 10px; border-radius: 8px; color: #065F46; font-size: 13px; font-weight: 600;'><span class='icon-MAGALOG' style='vertical-align:middle; margin-right: 5px;'>psychology</span> <b>Memória I.A. Ativada:</b> O sistema rastreou as equipes formadas no dia {data_memoria} para este turno e vai forçar a rotatividade (evitando que as mesmas pessoas trabalhem juntas hoje).</div>", unsafe_allow_html=True)
            
        st.markdown("<hr style='margin: 15px 0; border-top: 1px solid #E2E8F0;'>", unsafe_allow_html=True)
        
                # Mapeamento de Skills (Focado na galera filtrada)
        skills_ecom = []
        skills_carreg = []
        skills_full = []
        if not df_matriz.empty:
            for _, row in df_matriz.iterrows():
                nome = str(row.get('NOME', '')).strip()
                if nome in todos_colaboradores: # Só mapeia quem está no turno selecionado
                    if str(row.get('ECOM', '')).upper() in ['TRUE', '1', 'SIM']: skills_ecom.append(nome)
                    if str(row.get('CARREGAMENTO', '')).upper() in ['TRUE', '1', 'SIM']: skills_carreg.append(nome)
                    if str(row.get('FULL', '')).upper() in ['TRUE', '1', 'SIM']: skills_full.append(nome)
        
        # ==========================================
        # PAINEL DE CONFIGURAÇÃO (INPUTS)
        # ==========================================
        st.markdown("<h4 style='color: #0F172A; margin-bottom: 15px;'><span class='icon-MAGALOG' style='color:#0086FF;'>settings_suggest</span> 2. Demanda Operacional de Hoje</h4>", unsafe_allow_html=True)
        
        c_eq1, c_eq2, c_eq3, c_eq4 = st.columns(4)
        with c_eq1: qtd_geral = st.number_input("Equipes GERAIS", min_value=1, max_value=15, value=4, help="Descarga Convencional 1P")
        with c_eq2: qtd_ecom = st.number_input("Equipes E-COM", min_value=0, max_value=5, value=1)
        with c_eq3: qtd_full = st.number_input("Equipes FULL", min_value=0, max_value=5, value=1)
        with c_eq4: qtd_carreg = st.number_input("Equipes CARREGAMENTO", min_value=0, max_value=5, value=1)
            
        st.markdown("<br><h4 style='color: #0F172A; margin-bottom: 15px;'><span class='icon-MAGALOG' style='color:#F59E0B;'>rule</span> 3. Regras de Ouro e Restrições Manuais</h4>", unsafe_allow_html=True)
        
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            ausentes_sel = st.multiselect("❌ Remover Ausentes / Férias", options=todos_colaboradores, help="Estas pessoas não serão escaladas hoje.")
            fixos_ecom = st.multiselect("🔒 Equipe Fixa E-COM (Devem ficar juntos)", options=[p for p in todos_colaboradores if p not in ausentes_sel], help="Eles serão forçados na Equipe E-COM 1.")
        with col_r2:
            incompativeis = st.multiselect("⚡ Incompatíveis (NÃO podem ficar juntos)", options=[p for p in todos_colaboradores if p not in ausentes_sel], help="Além da memória de ontem, o algoritmo também separará essas pessoas escolhidas manualmente.")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # BOTÃO AZUL MATADOR
        if st.button("✨ DIVIDIR E GRAVAR EQUIPES", type="primary", use_container_width=True):
            with st.spinner("Analisando passado, embaralhando equipe e cruzando habilidades..."):
                time.sleep(1.5)
                
                # 1. Fila Única e Rotatividade Aleatória (Shuffle)
                pool_disponivel = [p for p in todos_colaboradores if p not in ausentes_sel and p not in fixos_ecom]
                random.shuffle(pool_disponivel) 
                
                equipes_ecom = [[] for _ in range(qtd_ecom)]
                equipes_carreg = [[] for _ in range(qtd_carreg)]
                equipes_full = [[] for _ in range(qtd_full)]
                equipes_gerais = [[] for _ in range(qtd_geral)]
                
                # 2. Alocação FIXA (E-COM)
                if qtd_ecom > 0 and fixos_ecom:
                    equipes_ecom[0].extend(fixos_ecom)
                elif qtd_ecom == 0 and fixos_ecom:
                    pool_disponivel.extend(fixos_ecom) # Se não tiver Ecom hoje, os fixos caem no rateio normal
                
                # 3. Função INTELIGENTE de distribuição
                def alocar_pessoas(pessoas, lista_de_equipes, limite_tamanho=99, usar_memoria_ia=True):
                    sobras = []
                    for p in pessoas:
                        alocado = False
                        lista_de_equipes.sort(key=len) # Tenta sempre a equipe mais vazia para balancear
                        
                        for equipe in lista_de_equipes:
                            if len(equipe) >= limite_tamanho: continue
                            
                            conflito = False
                            for membro in equipe:
                                # Regra A: Manual (ambos na lista de incompativeis)
                                if p in incompativeis and membro in incompativeis:
                                    conflito = True; break
                                # Regra B: Memória da I.A. (Trabalharam juntos ontem)
                                if usar_memoria_ia and p in historico_restricoes and membro in historico_restricoes.get(p, set()):
                                    conflito = True; break
                            
                            if not conflito:
                                equipe.append(p)
                                alocado = True
                                break
                                
                        if not alocado:
                            sobras.append(p)
                    return sobras

                # 4. Distribuição em Cascata (Evita clonagem de multi-skills)
                
                # A. CARREGAMENTO (Prioridade 1 - Máx 3 por equipe)
                cand_carreg = [p for p in pool_disponivel if p in skills_carreg]
                for p in cand_carreg: pool_disponivel.remove(p) 
                sobras_carreg = alocar_pessoas(cand_carreg, equipes_carreg, limite_tamanho=3, usar_memoria_ia=False)
                pool_disponivel.extend(sobras_carreg) 
                
                # B. E-COM (Prioridade 2)
                cand_ecom = [p for p in pool_disponivel if p in skills_ecom]
                for p in cand_ecom: pool_disponivel.remove(p) 
                sobras_ecom = alocar_pessoas(cand_ecom, equipes_ecom, usar_memoria_ia=False)
                pool_disponivel.extend(sobras_ecom) 
                
                # C. FULL (Prioridade 3)
                cand_full = [p for p in pool_disponivel if p in skills_full]
                for p in cand_full: pool_disponivel.remove(p) 
                sobras_full = alocar_pessoas(cand_full, equipes_full, usar_memoria_ia=False)
                pool_disponivel.extend(sobras_full)
                
                # D. Operação GERAL (O restante da fila única - Com Memória de Ontem Ligada)
                sobras_gerais = alocar_pessoas(pool_disponivel, equipes_gerais, usar_memoria_ia=True)
                
                # Fallback: Força na menor equipe caso as restrições bloqueiem todas
                if sobras_gerais:
                    for p in sobras_gerais:
                        equipes_gerais.sort(key=len)
                        equipes_gerais[0].append(p)
                
                # ==========================================
                # GRAVAÇÃO AUTOMÁTICA NO HISTÓRICO
                # ==========================================
                hoje_str = (datetime.datetime.utcnow() - datetime.timedelta(hours=3)).strftime("%d/%m/%Y")
                linhas_para_salvar = []
                
                for i, eq in enumerate(equipes_gerais):
                    if eq: linhas_para_salvar.append([hoje_str, turno_str, f"Geral {i+1}", ", ".join(eq)])
                for i, eq in enumerate(equipes_ecom):
                    if eq: linhas_para_salvar.append([hoje_str, turno_str, f"E-COM {i+1}", ", ".join(eq)])
                for i, eq in enumerate(equipes_full):
                    if eq: linhas_para_salvar.append([hoje_str, turno_str, f"FULL {i+1}", ", ".join(eq)])
                for i, eq in enumerate(equipes_carreg):
                    if eq: linhas_para_salvar.append([hoje_str, turno_str, f"Carregamento {i+1}", ", ".join(eq)])
                
                if linhas_para_salvar:
                    salvar_escala_ia(linhas_para_salvar)
                    carregar_historico_escalas.clear() # Limpa o cache para amanhã ler os novos dados
                
                # ==========================================
                # RENDERIZAÇÃO DOS CARDS (O RESULTADO)
                # ==========================================
                st.markdown("---")
                st.markdown("<h4 style='color: #0F172A; margin-bottom: 20px;'><span class='icon-MAGALOG' style='color:#10B981;'>check_circle</span> Escalação Oficial do Turno Salva</h4>", unsafe_allow_html=True)
                
                def desenhar_card_equipe(titulo, membros, cor_borda, icone):
                    html_membros = ""
                    for m in membros:
                        badges = ""
                        if m in skills_ecom: badges += "<span style='color:#0284C7; font-size:10px; background:#E0F2FE; padding:2px 4px; border-radius:4px; margin-left:4px; font-weight:800;'>ECOM</span>"
                        if m in skills_carreg: badges += "<span style='color:#EA580C; font-size:10px; background:#FFEDD5; padding:2px 4px; border-radius:4px; margin-left:4px; font-weight:800;'>CARREG</span>"
                        if m in skills_full: badges += "<span style='color:#D946EF; font-size:10px; background:#FDF4FF; padding:2px 4px; border-radius:4px; margin-left:4px; font-weight:800;'>FULL</span>"
                        if m in incompativeis: badges += "<span style='color:#DC2626; font-size:10px; background:#FEF2F2; padding:2px 4px; border-radius:4px; margin-left:4px; font-weight:800;'>⚡</span>"
                        
                        html_membros += f"<div style='padding: 6px 0; border-bottom: 1px solid #F1F5F9; font-size: 13px; color: #334155; font-weight: 600;'><span class='icon-MAGALOG' style='font-size:14px; color:#94A3B8; vertical-align: middle; margin-right:4px;'>person</span>{m} {badges}</div>"
                        
                    st.markdown(f"""
                    <div class="MAGALOG-card" style="border-top: 4px solid {cor_borda} !important; height: 100%;">
                        <div style="display:flex; align-items:center; justify-content: space-between; margin-bottom:12px;">
                            <div style="display:flex; align-items:center; gap:8px;">
                                <span class="icon-MAGALOG" style="color:{cor_borda}; font-size:24px;">{icone}</span>
                                <div style="font-weight:900; color:#0F172A; font-size:15px; text-transform: uppercase;">{titulo}</div>
                            </div>
                            <div style="background:#F1F5F9; color:#475569; font-size:11px; font-weight:900; padding:4px 8px; border-radius:12px;">{len(membros)} OP</div>
                        </div>
                        {html_membros if membros else "<div style='color:#94A3B8; font-size:12px; font-style:italic;'>Nenhum operador alocado.</div>"}
                    </div>
                    """, unsafe_allow_html=True)

                if qtd_geral > 0:
                    st.markdown("<div style='color: #64748B; font-weight: 800; font-size: 13px; margin-bottom: 10px; text-transform: uppercase;'>Operação Geral (Descarga 1P)</div>", unsafe_allow_html=True)
                    cols_g = st.columns(4)
                    for i, eq in enumerate(equipes_gerais):
                        with cols_g[i % 4]:
                            desenhar_card_equipe(f"Equipe {i+1}", eq, "#0086FF", "forklift")
                
                if qtd_ecom > 0 or qtd_full > 0 or qtd_carreg > 0:
                    st.markdown("<br><div style='color: #64748B; font-weight: 800; font-size: 13px; margin-bottom: 10px; text-transform: uppercase;'>Operações Específicas</div>", unsafe_allow_html=True)
                    cols_esp = st.columns(4)
                    idx_col = 0
                    for i, eq in enumerate(equipes_ecom):
                        with cols_esp[idx_col % 4]:
                            desenhar_card_equipe(f"E-COM {i+1}", eq, "#8B5CF6", "shopping_cart")
                        idx_col += 1
                    for i, eq in enumerate(equipes_full):
                        with cols_esp[idx_col % 4]:
                            desenhar_card_equipe(f"FULL {i+1}", eq, "#D946EF", "inventory_2") # Rosa Full
                        idx_col += 1
                    for i, eq in enumerate(equipes_carreg):
                        with cols_esp[idx_col % 4]:
                            desenhar_card_equipe(f"CARREG. {i+1}", eq, "#F59E0B", "upload")
                        idx_col += 1

# ==========================================================
# MÓDULO 8: CONTROLE DE COBRANÇA (Precificação e WhatsApp)
# ==========================================================
elif pagina_selecionada == "Controle de Cobrança":
    import urllib.parse
    
    render_hero(
        'Controle de Cobrança Carga e Descarga',
        'Precifique as cargas em tempo real, defina o valor do frete e envie cobranças via WhatsApp.',
        'MAGALOG • Automação Financeira'
    )
    
    # --- FUNÇÕES LOCAIS EXCLUSIVAS DESTE MÓDULO ---
    @st.cache_data(ttl=60)
    def carregar_tabela_frete_local():
        for tentativa in range(3):
            try:
                client = conectar_google()
                sh = client.open_by_key("1NWH9BHXgUmS-6WCQ8AjAHbt8DUHIvgQLRJ8hwUSDC7U")
                dados = sh.worksheet("FRETE").get_all_values()
                if len(dados) > 1:
                    df = pd.DataFrame(dados[1:], columns=dados[0])
                    df.columns = df.columns.astype(str).str.strip().str.upper()
                    return df
                return pd.DataFrame()
            except:
                if tentativa == 2: return pd.DataFrame()
                time.sleep(1.5)

    def gravar_precificacao(agenda_wms, concater, valor):
        try:
            client = conectar_google()
            sh = client.open_by_key("1NWH9BHXgUmS-6WCQ8AjAHbt8DUHIvgQLRJ8hwUSDC7U")
            ws = sh.worksheet("Painel de Controle")
            
            cell = ws.find(str(agenda_wms).strip())
            if cell:
                header = [str(h).strip().upper() for h in ws.row_values(1)]
                
                # Acha ou cria a coluna CAT / TP CARRO
                if 'CAT / TP CARRO' in header: col_cat = header.index('CAT / TP CARRO') + 1
                else: col_cat = len(header) + 1; ws.update_cell(1, col_cat, 'CAT / TP CARRO')
                
                # Acha ou cria a coluna VALOR
                if 'VALOR' in header: col_val = header.index('VALOR') + 1
                else: col_val = len(header) + 1; ws.update_cell(1, col_val, 'VALOR')
                
                # Atualiza as duas células
                ws.update_cell(cell.row, col_cat, concater)
                ws.update_cell(cell.row, col_val, valor)
                return True
            return False
        except Exception as e:
            st.error(f"Erro ao salvar precificação: {e}")
            return False

    # NOVA FUNÇÃO: Grava direto no Checkbox "PAG" da planilha
    def atualizar_status_pagamento(agenda, novo_status):
        try:
            client = conectar_google()
            sh = client.open_by_key("1NWH9BHXgUmS-6WCQ8AjAHbt8DUHIvgQLRJ8hwUSDC7U")
            ws = sh.worksheet("Painel de Controle")
            
            cell = ws.find(str(agenda).strip())
            if cell:
                header = [str(h).strip().upper() for h in ws.row_values(1)]
                
                # Procura a coluna PAG (Onde fica o Checkbox)
                if 'PAG' in header: col_idx = header.index('PAG') + 1
                else: col_idx = len(header) + 1; ws.update_cell(1, col_idx, 'PAG')
                
                # Traduz "Pago" para TRUE (marca o checkbox) e "Pendente" para FALSE
                val_to_write = "TRUE" if novo_status == "Pago" else "FALSE"
                ws.update_cell(cell.row, col_idx, val_to_write)
                return True
            return False
        except Exception as e:
            st.error(f"Erro ao atualizar pagamento: {e}")
            return False

    df_painel, df_lg = carregar_dados_cobranca()
    df_frete = carregar_tabela_frete_local()
    
    # Lista do Dropdown de Frete
    lista_fretes = ["SELECIONE O TIPO DE CARGA E VEÍCULO..."]
    if not df_frete.empty and 'CONCATER' in df_frete.columns and 'VALOR' in df_frete.columns:
        for _, r in df_frete.iterrows():
            conc = str(r['CONCATER']).strip()
            val = str(r['VALOR']).strip()
            if conc: lista_fretes.append(f"{conc} ➖ R$ {val}")
    
    if df_painel.empty:
        st.warning("Não foi possível carregar a base do Painel de Controle.")
    else:
        # ==========================================
        # 1. TRATAMENTO E O "PROCV" EM PYTHON (BLINDADO)
        # ==========================================
        df_painel.columns = df_painel.columns.astype(str).str.strip().str.upper()
        df_painel = df_painel.loc[:, ~df_painel.columns.duplicated()].copy() 
        
        col_ag_wms = next((c for c in df_painel.columns if 'AGENDA' in c), 'AGENDA WMS')
        col_dt_ag = next((c for c in df_painel.columns if 'DATA AGENDA' in c), 'DATA AGENDA')
        
        # Mapeando direto do Painel
        col_forn_seller = next((c for c in df_painel.columns if 'FORNECEDOR' in c or 'SELLER' in c), 'FORNECEDOR/SELLER')
        col_mot_painel = next((c for c in df_painel.columns if 'MOTOR' in c), 'NOME DO MOTORISTA')
        col_placa_painel = next((c for c in df_painel.columns if 'PLACA' in c), 'PLACA')
        col_com_painel = next((c for c in df_painel.columns if 'COMANDA' in c), 'COMANDA')
        
        col_carro_painel = next((c for c in df_painel.columns if c in ['CARRO', 'VEICULO', 'VEÍCULO']), None)
        if not col_carro_painel: col_carro_painel = next((c for c in df_painel.columns if 'CARRO' in c or 'VEIC' in c), 'CARRO')
        
        df_painel['AGENDA_CLEAN'] = df_painel[col_ag_wms].astype(str).str.strip().str.upper()
        df_painel['DATA_DT'] = pd.to_datetime(df_painel[col_dt_ag].astype(str).str.split(' ').str[0], format='%d/%m/%Y', errors='coerce').dt.date
        
        if not df_lg.empty:
            df_lg.columns = df_lg.columns.astype(str).str.strip().str.upper()
            df_lg = df_lg.loc[:, ~df_lg.columns.duplicated()].copy()
            
            col_ag_lg = next((c for c in df_lg.columns if 'AGENDA' in c), 'CODAGENDA')
            col_tel   = next((c for c in df_lg.columns if 'TELEFONE' in c or 'NUMERO' in c or 'CELULAR' in c or 'FONE' in c), 'NÚMEROTELEFONE')
            
            for coluna_necessaria in [col_ag_lg, col_tel]:
                if coluna_necessaria not in df_lg.columns: df_lg[coluna_necessaria] = ""
                
            df_lg['AGENDA_CLEAN'] = df_lg[col_ag_lg].astype(str).str.strip().str.upper()
            df_lg = df_lg.drop_duplicates(subset=['AGENDA_CLEAN'], keep='last')
            
            df_lg_renamed = df_lg.rename(columns={col_tel: 'NÚMEROTELEFONE'})
            if 'NÚMEROTELEFONE' not in df_lg_renamed.columns: df_lg_renamed['NÚMEROTELEFONE'] = ""
            
            df_merged = pd.merge(df_painel, df_lg_renamed[['AGENDA_CLEAN', 'NÚMEROTELEFONE']], on='AGENDA_CLEAN', how='left')
            df_merged = df_merged.loc[:, ~df_merged.columns.duplicated()].copy()
        else:
            df_merged = df_painel.copy()
            df_merged['NÚMEROTELEFONE'] = ""
            
        # --- BLOCO DE INTELIGÊNCIA: FILTROS LOGÍSTICOS E FINANCEIROS ---
        
        # 1. Filtra as Filiais e CDs Magalu (Agora pega em cheio os CDs numerados)
        def eh_interno(forn):
            if pd.isna(forn): return False
            f = str(forn).upper().strip()
            if f.startswith(('MAGAZINE', 'FILIAL', 'CD')): return True
            return False
        
        if col_forn_seller in df_merged.columns:
            mask_interno = df_merged[col_forn_seller].apply(eh_interno)
            df_merged = df_merged[~mask_interno]

        # 2. Filtra os Status Ausente e Comercial
        col_status_log = next((c for c in df_merged.columns if c == 'STATUS' or c == 'STATUS AGENDA'), None)
        if col_status_log:
            mask_ausente_comercial = df_merged[col_status_log].astype(str).str.upper().isin(['AUSENTE', 'COMERCIAL', 'DIVER. COMERCIAL', 'DIVER COMERCIAL'])
            df_merged = df_merged[~mask_ausente_comercial]

        # 3. Mapeia a coluna "PAG" (Checkbox) para o status financeiro do sistema
        col_pag_checkbox = next((c for c in df_merged.columns if c == 'PAG' or c == 'PAG.'), None)
        if col_pag_checkbox:
            def mapear_pagamento(val):
                v = str(val).strip().upper()
                if v in ['TRUE', '1', 'SIM', 'VERDADEIRO']: return 'Pago'
                return 'Pendente'
            df_merged['STATUS PAGAMENTO'] = df_merged[col_pag_checkbox].apply(mapear_pagamento)
        else:
            df_merged['STATUS PAGAMENTO'] = 'Pendente'

        # ==========================================
        # 2. FILTROS E KPIS (BLINDADO)

        st.markdown("### <span class='icon-MAGALOG'>filter_alt</span> Filtros da Operação", unsafe_allow_html=True)
        
        c_f1, c_f2, c_f3 = st.columns(3)
        with c_f1: data_cobranca = st.date_input("Data da Agenda", value=pd.Timestamp.now().date())
        with c_f2: status_filtro = st.selectbox("Status Pagamento", ["Todos", "Pendente", "Cobrado", "Pago"])
        with c_f3: busca_placa = st.text_input("Buscar Placa ou Agenda", placeholder="Ex: ABC1234")
            
        df_filtrado = df_merged[df_merged['DATA_DT'] == data_cobranca].copy()
        
        if status_filtro != "Todos":
            df_filtrado = df_filtrado[df_filtrado['STATUS PAGAMENTO'].str.title() == status_filtro]
            
        if busca_placa:
            busca_placa = busca_placa.upper().strip()
            df_filtrado = df_filtrado[
                df_filtrado[col_placa_painel].str.upper().str.contains(busca_placa, na=False) |
                df_filtrado['AGENDA_CLEAN'].str.contains(busca_placa, na=False)
            ]
            
        col_valor = next((c for c in df_filtrado.columns if 'VALOR' in str(c).upper() or 'R$' in str(c).upper()), None)
        if not col_valor:
            df_filtrado['VALOR'] = "0,00"
            col_valor = 'VALOR'
            
        total_agendas = len(df_filtrado)
        
        def limpar_moeda_kpi(v):
            s = str(v).upper().replace('R$', '').replace(' ', '').strip()
            if s in ['', 'NAN', 'NONE', 'NULL', 'NAOÉCOBRADO', 'NÃO É COBRADO']: return 0.0
            if '.' in s and ',' in s: s = s.replace('.', '').replace(',', '.')
            elif ',' in s: s = s.replace(',', '.')
            try: return float(s)
            except: return 0.0
            
        df_filtrado['VALOR_NUM'] = df_filtrado[col_valor].apply(limpar_moeda_kpi)
        
        # BLINDAGEM MÁXIMA (ANTI-CRASH): Tenta forçar a matemática. Se o Pandas bugar com dia vazio, vira 0.0 automático.
        try:
            total_cobrado = float(df_filtrado[df_filtrado['STATUS PAGAMENTO'].str.upper() == 'PAGO']['VALOR_NUM'].sum())
        except:
            total_cobrado = 0.0
            
        try:
            aguardando = float(df_filtrado[df_filtrado['STATUS PAGAMENTO'].str.upper().isin(['PENDENTE', 'COBRADO'])]['VALOR_NUM'].sum())
        except:
            aguardando = 0.0
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        def render_kpi_cobranca(titulo, valor, subtitulo, cor):
            st.markdown(f'''
            <div class="kpi-card" style="border-top: 4px solid {cor}; height: 130px; padding: 10px; text-align: center; background: #FFFFFF; border-radius: 16px; border: 1px solid #F1F5F9; box-shadow: 0 4px 20px rgba(0, 0, 0, 0.03); margin-bottom: 12px;">
                <div style="color: #64748B; font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px;">{titulo}</div>
                <div style="font-size:28px; font-weight: 900; letter-spacing: -0.5px; color:{cor};">{valor}</div>
                <div style="font-size:11px; color:#64748B; font-weight:700; margin-top:4px; line-height:1.4;">{subtitulo}</div>
            </div>
            ''', unsafe_allow_html=True)

        k1, k2, k3 = st.columns(3)
        with k1: render_kpi_cobranca("Total Agendas (Filtro)", total_agendas, "Cargas no pátio", "#0086FF")
        with k2: render_kpi_cobranca("Total Pago (R$)", f"R$ {total_cobrado:,.2f}".replace(',','X').replace('.',',').replace('X','.'), "Confirmados na conta", "#10B981")
        with k3: render_kpi_cobranca("Aguardando Pagto (R$)", f"R$ {aguardando:,.2f}".replace(',','X').replace('.',',').replace('X','.'), "Pendentes e Cobrados", "#F59E0B")
        
        # ==========================================
        # 3. RENDERIZAÇÃO EM LINHAS (PRECIFICAÇÃO INTERATIVA)
        # ==========================================
        st.markdown("<hr style='margin: 15px 0; border-top: 1px solid #E2E8F0;'>", unsafe_allow_html=True)
        st.markdown("<h4 style='color: #0F172A; margin-bottom: 20px;'><span class='icon-MAGALOG' style='color:#0086FF;'>request_quote</span> Fila de Precificação e Cobrança</h4>", unsafe_allow_html=True)
        
        st.markdown("""
<style>

/* BASE */
div[class*="st-key-btn_zap_"] a,
div[class*="st-key-btn_pago_"] button,
div[class*="st-key-btn_voltar_"] button,
div[class*="st-key-btn_recibo_"] button {
    height: 42px !important;
    border-radius: 12px !important;
    border: none !important;
    color: #FFFFFF !important;
    font-weight: 900 !important;
    letter-spacing: .2px !important;
    transition: all .18s ease-in-out !important;
}

/* ZAP AZUL NEON */
div[class*="st-key-btn_zap_"] a {
    background: linear-gradient(135deg, #0086FF, #00C2FF) !important;
    box-shadow: 0 0 14px rgba(0, 134, 255, .55) !important;
}

/* PAGO VERDE NEON */
div[class*="st-key-btn_pago_"] button {
    background: linear-gradient(135deg, #00E676, #00B86B) !important;
    box-shadow: 0 0 14px rgba(0, 230, 118, .55) !important;
}

/* VOLTAR VERMELHO NEON */
div[class*="st-key-btn_voltar_"] button {
    background: linear-gradient(135deg, #FF1744, #D50000) !important;
    box-shadow: 0 0 14px rgba(255, 23, 68, .55) !important;
}

/* RECIBO ROXO NEON */
div[class*="st-key-btn_recibo_"] button {
    background: linear-gradient(135deg, #A855F7, #7C3AED) !important;
    box-shadow: 0 0 14px rgba(168, 85, 247, .55) !important;
}

/* HOVER */
div[class*="st-key-btn_zap_"] a:hover,
div[class*="st-key-btn_pago_"] button:hover,
div[class*="st-key-btn_voltar_"] button:hover,
div[class*="st-key-btn_recibo_"] button:hover {
    transform: translateY(-2px) scale(1.02) !important;
    filter: brightness(1.12) !important;
}

/* MATA O EXPAND_MORE DO POPOVER */
div[class*="st-key-btn_recibo_"] button [data-testid="stIconMaterial"],
div[class*="st-key-btn_recibo_"] button span:has(+ span),
div[class*="st-key-btn_recibo_"] button svg {
    display: none !important;
}

div[class*="st-key-btn_recibo_"] button::after {
    content: "" !important;
}

</style>
""", unsafe_allow_html=True)
        
        if df_filtrado.empty:
            st.info("Nenhuma agenda localizada para a data/filtros selecionados.")
        else:
            # RESETAMOS O INDEX PARA GARANTIR QUE CADA 'IDX' SEJA ÚNICO NA TELA
            for idx, row in df_filtrado.reset_index(drop=True).iterrows():
                
                # ==========================================
                # EXTRAÇÃO DE DADOS (BLINDADA CONTRA N/D)
                # ==========================================
                def extrair_dado(coluna, padrao="N/D"):
                    val = row.get(coluna, padrao)
                    if isinstance(val, pd.Series): return str(val.iloc[0]).strip()
                    if pd.isna(val) or str(val).strip() == '': return padrao
                    return str(val).strip()

                agenda = extrair_dado('AGENDA_CLEAN')
                status_pag = str(row.get('STATUS PAGAMENTO', 'Pendente')).title()
                
                # Agora PUXA TUDO da base do próprio Painel
                motorista = extrair_dado(col_mot_painel, 'Não Informado').title()
                
                # CORREÇÃO DEFINITIVA DO FORNECEDOR E DA CATEGORIA
                forn_seller = row.get('FORNECEDOR/SELLER', row.get('FORNECEDOR', 'N/D'))
                if pd.isna(forn_seller) or str(forn_seller).strip() == '': forn_seller = 'N/D'
                forn_seller = str(forn_seller).upper().strip()
                
                col_categoria_real = next((c for c in df_painel.columns if 'CATEGORIA' in c), 'CATEGORIA')
                categoria_carga = extrair_dado(col_categoria_real, 'N/D').upper()
                
                placa = extrair_dado(col_placa_painel, 'S/ Placa').upper()
                comanda = extrair_dado(col_com_painel, 'N/D') 

                
                # Puxa veículo e corta se vier "CATEGORIA-VEICULO"
                tipo_veic_lg = extrair_dado(col_carro_painel, 'Não Informado').upper()
                if '-' in tipo_veic_lg: tipo_veic_lg = tipo_veic_lg.split('-', 1)[1].strip()
                
                # Apenas TELEFONE puxado da Base LG
                telefone = extrair_dado('NÚMEROTELEFONE', '').replace(' ', '').replace('-', '').replace('(', '').replace(')', '')
                if telefone and not telefone.startswith('55') and telefone != 'N/D': telefone = '55' + telefone
                
                # Tratamento do Valor e Categoria Atribuída
                valor_atual_banco = extrair_dado(col_valor, '')
                col_cat_tp = next((c for c in df_painel.columns if 'CAT / TP CARRO' in c.upper()), None)
                cat_atual_banco = extrair_dado(col_cat_tp, '') if col_cat_tp else ""
                
                ja_precificado = (valor_atual_banco not in ['', '0', '0,00', '0.00', 'R$ 0,00', 'N/D'])
                
                # Inteligência de Cores
                if status_pag == 'Pago':
                    cor_borda = "#10B981"; bg_badge = "#D1FAE5"; cor_badge = "#065F46"
                elif status_pag == 'Cobrado':
                    cor_borda = "#3B82F6"; bg_badge = "#DBEAFE"; cor_badge = "#1E40AF"
                else:
                    cor_borda = "#F59E0B"; bg_badge = "#FEF3C7"; cor_badge = "#92400E"
                    
                with st.container():
                    st.markdown(f"""
                    <div style="background: #FFFFFF; border: 1px solid #E2E8F0; border-left: 5px solid {cor_borda}; border-radius: 12px; padding: 20px; margin-bottom: 5px; box-shadow: 0 4px 10px rgba(0,0,0,0.02);">
                        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
                            <div style="display: flex; gap: 15px; align-items: center;">
                                <div style="background: #F1F5F9; padding: 6px 12px; border-radius: 8px; font-weight: 900; color: #0F172A; font-size: 16px;">{placa}</div>
                                <div style="font-size: 13px; color: #475569;"><b>Agenda:</b> {agenda} &nbsp;|&nbsp; <b>Comanda:</b> {comanda}</div>
                            </div>
                            <div style="background: {bg_badge}; color: {cor_badge}; font-size: 11px; font-weight: 800; padding: 6px 12px; border-radius: 999px; text-transform: uppercase;">{status_pag}</div>
                        </div>
                        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px; font-size: 13px; color: #475569; margin-bottom: 15px;">
                            <div><span class='icon-MAGALOG' style='font-size:16px; vertical-align:middle; margin-right:4px;'>storefront</span><b>Fornecedor:</b> {forn_seller}</div>
                            <div><span class='icon-MAGALOG' style='font-size:16px; vertical-align:middle; margin-right:4px;'>person</span><b>Motorista:</b> {motorista} ({tipo_veic_lg})</div>
                        </div>
                    """, unsafe_allow_html=True)
                    
                    c_in1, c_in2, c_in3 = st.columns([4, 2, 2])
                    
                    # Bloco 1: Seleção do Valor (Dropdown)
                    with c_in1:
                        if ja_precificado:
                            st.markdown(f"<div style='margin-top:2px; font-size:12px; color:#64748B; font-weight:700;'>Categoria Atribuída:</div><div style='font-size:14px; color:#0F172A; font-weight:800; background:#F8FAFC; padding:8px; border-radius:6px; border:1px solid #E2E8F0;'>{cat_atual_banco}</div>", unsafe_allow_html=True)
                        else:
                            selecao_frete = st.selectbox("Definir Categoria / Tipo de Veículo", options=lista_fretes, key=f"sel_{agenda}_{idx}", label_visibility="collapsed")
                    
                    # Bloco 2: Botão de Salvar ou Exibição do Valor
                    with c_in2:
                        if ja_precificado:
                            v_formatado = valor_atual_banco if valor_atual_banco.startswith('R$') else f"R$ {valor_atual_banco}"
                            st.markdown(f"<div style='margin-top:2px; font-size:12px; color:#64748B; font-weight:700;'>Valor do Frete:</div><div style='font-size:16px; color:#0086FF; font-weight:900; background:#E6F2FF; padding:7px; border-radius:6px; text-align:center; border:1px dashed #BAE6FD;'>{v_formatado}</div>", unsafe_allow_html=True)
                        else:
                            if st.button("💾 Gravar Valor", key=f"btn_save_{agenda}_{idx}", type="primary", use_container_width=True):
                                if selecao_frete == "SELECIONE O TIPO DE CARGA E VEÍCULO...":
                                    st.error("Selecione uma opção válida.")
                                else:
                                    with st.spinner("Gravando..."):
                                        partes = selecao_frete.split(' ➖ R$ ')
                                        concater_salvar = partes[0].strip()
                                        valor_salvar = partes[1].strip()
                                        
                                        if gravar_precificacao(agenda, concater_salvar, valor_salvar):
                                            st.success("Salvo!")
                                            carregar_dados_cobranca.clear()
                                            st.rerun()
                    
                                        # Bloco 3: Ações (WhatsApp, Pagamento e Recibo)
                    with c_in3:
                        if ja_precificado:
                            texto_wa = f"""Olá Sr(a) {motorista}

✓ A etapa de entrada das notas fiscais está tudo OK, agora só aguardar a descarga. Para adiantarmos o processo, esses são os dados para pagamento da descarga:

✦ DADOS DO VEÍCULO
• Comanda: {comanda}
• Placa: {placa}
• Agenda: {agenda}
• Tipo Veiculo: {cat_atual_banco.split('-')[1] if '-' in cat_atual_banco else cat_atual_banco}

✦ DADOS DE PAGAMENTO
• Valor: R$ {valor_atual_banco.replace('R$', '').strip()}
• CHAVE PIX: 24.230.747.0001.02
• Magalu Log Serviços Logísticos LTDA

➥ Aceitamos apenas cartão de débito e pix.
𝑝𝑎𝑟𝑎 𝑝𝑎𝑔𝑎𝑚𝑒𝑛𝑡𝑜𝑠 𝑛𝑜 𝑑𝑒𝑏𝑖𝑡𝑜 𝑠𝑒 𝑑𝑖𝑟𝑒𝑐𝑖𝑜𝑛𝑎𝑟 𝑎 𝑠𝑎𝑙𝑎 𝑑𝑜 𝑟𝑒𝑐𝑒𝑏𝑖𝑚𝑒𝑛𝑡𝑜 𝑎𝑠𝑠𝑖𝑚 𝑞𝑢𝑒 𝑒𝑛𝑐𝑜𝑠𝑡𝑎𝑟 𝑛𝑎 𝑑𝑜𝑐𝑎.
➥ Envio do comprovante de pagamento em PDF.

Agradeço e sigo à disposição! ♡"""
                            link_wa = f"https://wa.me/{telefone}?text={urllib.parse.quote(texto_wa)}" if telefone and telefone != 'N/D' else "#"
                            
                            c_btn1, c_btn2, c_btn3 = st.columns([1, 1, 1.2])
                            
                            with c_btn1:
                                with st.container(key=f"btn_zap_{agenda}_{idx}"):
                                    if telefone and telefone != 'N/D':
                                        st.link_button("Cobrar", link_wa, use_container_width=True)
                                    else:
                                        st.button("Sem Tel", key=f"d_zap_{agenda}_{idx}", disabled=True, use_container_width=True)
                            
                            with c_btn2:
                                if status_pag != 'Pago':
                                    with st.container(key=f"btn_pago_{agenda}_{idx}"):
                                        if st.button("✔ Pago", key=f"pg_{agenda}_{idx}", use_container_width=True):
                                            if atualizar_status_pagamento(agenda, "Pago"):
                                                carregar_dados_cobranca.clear()
                                                st.rerun()
                                else:
                                    with st.container(key=f"btn_voltar_{agenda}_{idx}"):
                                        if st.button("↩ Voltar", key=f"est_{agenda}_{idx}", use_container_width=True):
                                            if atualizar_status_pagamento(agenda, "Pendente"):
                                                carregar_dados_cobranca.clear()
                                                st.rerun()
                            
                            with c_btn3:
                                with st.container(key=f"btn_recibo_{agenda}_{idx}"):
                                    # A MÁGICA: Tudo que está "recuado" abaixo dessa linha fica escondido no Pop-up!
                                    with st.popover("Recibo", use_container_width=True):
                                        st.markdown("<div style='font-size:13px; font-weight:800; color:#0F172A; margin-bottom:10px;'>Emitir Recibo A4</div>", unsafe_allow_html=True)
                                        
                                        # Puxa o CNPJ da aba Transportadoras
                                        df_transp = carregar_transportadoras_local()
                                        col_cnpj = next((c for c in df_transp.columns if 'CNPJ' in c), None)
                                        col_razao = next((c for c in df_transp.columns if 'RAZ' in c or 'NOME' in c), None)
                                        
                                        opcoes_transp = ["Selecione a Transportadora..."]
                                        if not df_transp.empty and col_razao:
                                            opcoes_transp += sorted(df_transp[col_razao].dropna().astype(str).unique().tolist())
                                        
                                        transp_sel = st.selectbox("Transportadora", options=opcoes_transp, key=f"transp_{agenda}_{idx}")
                                        placa_carreta = st.text_input("Placa da Carreta", key=f"plcar_{agenda}_{idx}")
                                        forma_pag = st.selectbox("Forma de Pagamento", ["Pix", "Débito", "Carta de Crédito"], key=f"fpag_{agenda}_{idx}")
                                        nf_recibo = st.text_input("Nº Nota Fiscal (Opcional)", key=f"nf_{agenda}_{idx}")
                                        assistente_adm = st.text_input("Assistente Adm", key=f"adm_{agenda}_{idx}")
                                        obs_recibo = st.text_area("Observação (Opcional)", key=f"obs_{agenda}_{idx}", height=68)
                                        
                                        # Busca o CNPJ instantaneamente
                                        cnpj_sel = ""
                                        if transp_sel != "Selecione a Transportadora..." and col_cnpj and col_razao:
                                            df_match = df_transp[df_transp[col_razao].astype(str) == transp_sel]
                                            if not df_match.empty: cnpj_sel = str(df_match[col_cnpj].iloc[0])
                                            
                                        # GERAÇÃO DO HTML FORMATO A4 COM SEU DESIGN CUSTOMIZADO E FONTES EXATAS
                                        data_hoje = data_cobranca.strftime('%d/%m/%Y')
                                        veic_cortado = cat_atual_banco.split('-')[1] if '-' in cat_atual_banco else cat_atual_banco
                                        v_limpo = valor_atual_banco.replace('R$', '').strip()
                                        
                                        html_recibo = f"""
                                        <html>
                                        <head>
                                        <title>Recibo de Descarga - OS {agenda}</title>
                                        <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;700&display=swap" rel="stylesheet">
                                        <style>
                                            @page {{
                                                size: A4 portrait;
                                                margin: 0;
                                            }}
                                            body {{
                                                margin: 0;
                                                padding: 0;
                                                background: #fff;
                                                color: #000;
                                            }}
                                            .pagina {{
                                                width: 794px;
                                                height: 1123px;
                                                box-sizing: border-box;
                                                padding: 58px 76px 55px 76px;
                                                position: relative;
                                            }}
                                            .topo {{
                                                display: flex;
                                                justify-content: space-between;
                                                align-items: flex-start;
                                                width: 100%;
                                            }}
                                            .logo-box {{
                                                width: 220px;
                                            }}
                                            .logo {{
                                                width: 220px;
                                                height: auto;
                                                display: block;
                                            }}
                                            .ordem {{
                                                text-align: center;
                                                margin-top: 2px;
                                                margin-right: 5px;
                                                font-family: 'Montserrat', sans-serif;
                                            }}
                                            .ordem-titulo {{
                                                font-size: 25px;
                                                font-weight: 700;
                                                line-height: 25px;
                                            }}
                                            .ordem-numero {{
                                                font-size: 17px;
                                                font-weight: 700;
                                                margin-top: 4px;
                                            }}
                                            .titulo {{
                                                text-align: center;
                                                font-family: 'Montserrat', sans-serif;
                                                font-size: 24px;
                                                font-weight: 700;
                                                margin-top: 83px;
                                                margin-bottom: 34px;
                                            }}
                                            .conteudo {{
                                                margin-left: 13px;
                                                line-height: 19px;
                                            }}
                                            .linha {{
                                                margin: 0;
                                                padding: 0;
                                                margin-bottom: 2px;
                                            }}
                                            b {{
                                                font-family: 'Montserrat', sans-serif;
                                                font-size: 12px;
                                            }}
                                            span.valor {{
                                                font-family: Arial, Helvetica, sans-serif;
                                                font-size: 10px;
                                            }}
                                            .espaco {{
                                                height: 25px;
                                            }}
                                            .rodape {{
                                                text-align: center;
                                                font-family: Arial, Helvetica, sans-serif;
                                                font-size: 13px;
                                                line-height: 15px;
                                                margin-top: 31px;
                                                font-weight: 400;
                                            }}
                                            .declaracao {{
                                                text-align: center;
                                                font-family: Arial, Helvetica, sans-serif;
                                                font-size: 14px;
                                                line-height: 16px;
                                                margin-top: 29px;
                                                font-weight: 700;
                                            }}
                                            .assinatura {{
                                                margin-top: 57px;
                                                margin-left: 0px;
                                                font-family: Arial, Helvetica, sans-serif;
                                                font-size: 14px;
                                            }}
                                            @media print {{
                                                body {{
                                                    -webkit-print-color-adjust: exact;
                                                    print-color-adjust: exact;
                                                }}
                                            }}
                                        </style>
                                        </head>
                                        <body onload="setTimeout(function(){{ window.print(); }}, 700);">
                                        <div class="pagina">
                                            <div class="topo">
                                                <div class="logo-box">
                                                    <img src="https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcTU68FtjfYluSz6o-6YMgxPz9Ixvht3lK-9Fw&s" class="logo" alt="Logo Magalu">
                                                </div>
                                                <div class="ordem">
                                                    <div class="ordem-titulo">Ordem De Serviço</div>
                                                    <div class="ordem-numero">N°&nbsp;&nbsp;{agenda}</div>
                                                </div>
                                            </div>
                                            <div class="titulo">Recibo de Descarga</div>
                                            <div class="conteudo">
                                                <p class="linha"><b>- Transportador:</b> <span class="valor">{transp_sel if transp_sel != "Selecione a Transportadora..." else "-"}</span></p>
                                                <p class="linha"><b>- CNPJ:</b> <span class="valor">{cnpj_sel if cnpj_sel else "-"}</span></p>
                                                <p class="linha"><b>- Motorista:</b> <span class="valor">{motorista}</span></p>
                                                <p class="linha"><b>- Telefone:</b> <span class="valor">{telefone}</span></p>
                                                <p class="linha"><b>- Tipo de veículo:</b> <span class="valor">{tipo_veic_lg}</span></p>
                                                <p class="linha"><b>- Categoria:</b> <span class="valor">{categoria_carga}</span></p>
                                                <p class="linha"><b>- Placa do carro:</b> <span class="valor">{placa}</span></p>
                                                <p class="linha"><b>- Placa da carreta:</b> <span class="valor">{placa_carreta if placa_carreta else "-"}</span></p>
                                                <div class="espaco"></div>
                                                <p class="linha"><b>- Data da Descarga:</b> <span class="valor">{data_hoje}</span></p>
                                                <p class="linha"><b>- Serviço prestado:</b> <span class="valor">Descarga.</span></p>
                                                <p class="linha"><b>- Local de descarga:</b> <span class="valor">Magazine Luiza - CD 2900</span></p>
                                                <div class="espaco"></div>
                                                <p class="linha"><b>- N° de agenda:</b> <span class="valor">{agenda}</span></p>
                                                <p class="linha"><b>- Valor da Descarga conforme tabela:</b> <span class="valor">R$ {v_limpo}</span></p>
                                                <p class="linha"><b>- Valor Pago:</b> <span class="valor">R$ {v_limpo}</span></p>
                                                <p class="linha"><b>- Motivo do pagamento divergente:</b> <span class="valor">-</span></p>
                                                <p class="linha"><b>- Forma de pagamento:</b> <span class="valor">{forma_pag}</span></p>
                                                <p class="linha"><b>- Assistente administrativo:</b> <span class="valor">{assistente_adm.upper() if assistente_adm else "-"}</span></p>
                                                <p class="linha"><b>- Obs:</b> <span class="valor">{obs_recibo if obs_recibo else "-"}</span></p>
                                                <p class="linha"><b>- N° Nota Fiscal:</b> <span class="valor">{nf_recibo if nf_recibo else "-"}</span></p>
                                                <p class="linha"><b>N° Comanda:</b> <span class="valor">{comanda}</span></p>
                                            </div>
                                            <div class="rodape">
                                                O CNPJ disposto neste recibo será o mesmo em que será emitida a nota fiscal.<br>
                                                Em caso de divergência, entrar em contato com o assistente do CD imediatamente.
                                            </div>
                                            <div class="declaracao">
                                                Declaro que os serviços prestados acima foram efetuados pela<br>
                                                empresa Magazine Luiza conforme acordado em tabela.
                                            </div>
                                            <div class="assinatura">
                                                ASS: _____________________________________________________________________
                                            </div>
                                        </div>
                                        </body>
                                        </html>
                                        """
                                                                                                        
                                        
                                        # BOTÃO NATIVO DO STREAMLIT PARA TRAVAR A EXECUÇÃO
                                        if st.button("🖨️ GERAR, SALVAR E IMPRIMIR", key=f"print_drive_{agenda}_{idx}", type="primary", use_container_width=True):
                                            
                                            # 1. Primeiro: Executa o Python (Salva no Drive)
                                            with st.spinner("Gerando recibo e arquivando no Google Drive..."):
                                                id_arquivo = salvar_recibo_no_drive(html_recibo, agenda)
                                                if id_arquivo:
                                                    st.toast("✅ Recibo arquivado no Drive com sucesso!", icon="☁️")
                                                else:
                                                    st.toast("⚠️ O recibo não foi salvo no Drive.", icon="⚠️")
                                            
                                            # 2. Segundo: Injeta o JavaScript que "pula" na tela para imprimir
                                            import streamlit.components.v1 as components
                                            html_recibo_safe = html_recibo.replace('\\', '\\\\').replace('`', '\\`').replace('$', '\\$').replace('</script>', '<\\/script>')
                                            
                                            js_print_automatico = f"""
                                            <script>
                                            // Executa imediatamente assim que o Python termina de carregar esta parte
                                            var htmlContent = `{html_recibo_safe}`;
                                            var printWin = window.open('', '_blank');
                                            if(printWin) {{
                                                printWin.document.open();
                                                printWin.document.write(htmlContent);
                                                printWin.document.close();
                                            }} else {{
                                                alert('Seu navegador bloqueou o pop-up! Por favor, vá na barra de endereços, libere os pop-ups (ícone vermelho) e clique no botão novamente.');
                                            }}
                                            </script>
                                            """
                                            # Roda o script de forma invisível (height=0)
                                            components.html(js_print_automatico, height=0)
                                            
                        else:
                            st.markdown("<div style='color:#94A3B8; font-size:12px; text-align:center; padding-top:10px;'>Defina o valor para liberar ações.</div>", unsafe_allow_html=True)
                            
                    st.markdown("</div>", unsafe_allow_html=True)
