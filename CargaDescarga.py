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
from datetime import date

# ==========================================================
# 1. CONFIGURAÇÃO DA PÁGINA E CSS (THEME MAGALU AAAA+)
# ==========================================================
st.set_page_config(page_title="Magalu | Gestão Logística", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    /* 1. FONTE PREMIUM (Apple/Stripe Style) */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
    * { font-family: 'Inter', sans-serif !important; }

    /* --- CORREÇÃO DOS ÍCONES DA SIDEBAR E SISTEMA --- */
    .material-icons, .material-symbols-rounded, [data-testid="stSidebarCollapseButton"] * {
        font-family: 'Material Symbols Rounded', 'Material Icons', sans-serif !important;
    }

    /* 2. ANIMAÇÃO RGB LUIZALABS (Desnecessário, mas absurdamente foda) */
    @keyframes magaluGlow {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    /* Linha Tech Animada no Topo da Tela */
    .stApp::before {
        content: ""; position: fixed; top: 0; left: 0; right: 0; height: 5px;
        background: linear-gradient(90deg, #0086FF, #FF007F, #00C853, #0086FF, #FF007F);
        background-size: 300% 300%;
        animation: magaluGlow 6s linear infinite;
        z-index: 999999;
    }

    /* Fundo da Aplicação (Soft com micro-gradiente) */
    .stApp {
        background-color: #F0F4F8;
        background-image: radial-gradient(circle at 100% 0%, #E2EDF8 0%, transparent 40%);
    }

    /* 3. TÍTULOS COM DEGRADÊ METÁLICO */
    .magalu-page-title { 
        background: linear-gradient(135deg, #0086FF 0%, #001A57 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 32px; font-weight: 900; letter-spacing: -1px; margin-bottom: 2px;
    }
    .magalu-page-subtitle { color: #64748B; font-size: 15px; font-weight: 500; margin-bottom: 25px; }

    /* 4. ABAS (TABS) CORPORATIVAS - Adeus texto flutuante! */
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
    /* Aba Ativa (Azul Magalu com Glow) */
    button[data-baseweb="tab"][aria-selected="true"] {
        background-color: #0086FF !important;
        color: #FFFFFF !important;
        box-shadow: 0 4px 15px rgba(0,134,255,0.35) !important;
    }
    /* Ocultar aquela linha feia padrão do Streamlit embaixo da aba */
    [data-baseweb="tab-border"] { display: none; }

    /* 5. RIBBONS ANIMADOS */
    .magalu-ribbon {
        background: linear-gradient(90deg, #0086FF, #005BFF, #FF007F, #0086FF);
        background-size: 300% 300%;
        animation: magaluGlow 8s ease infinite;
        color: #FFFFFF; padding: 8px 24px; font-size: 13px; font-weight: 700;
        border-radius: 0px 8px 8px 0px; margin-bottom: 15px; margin-top: 10px;
        position: relative; left: -1rem; box-shadow: 0 4px 15px rgba(0,134,255,0.3);
        text-transform: uppercase; letter-spacing: 1px;
    }

    /* 6. CARDS (Docas, Formulários) - EFEITO GLASSMORPHISM */
    .magalu-card, div[data-testid="stVerticalBlock"] > div > div[data-testid="stVerticalBlockBorderWrapper"] {
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

    /* 7. CAIXAS DE INPUT E DROPDOWNS (CORRIGIDO) */
    div[data-baseweb="input"] > div, 
    div[data-baseweb="select"] > div { 
        border-radius: 12px !important; 
        min-height: 50px; /* Sem forçar altura máxima */
        border: 1px solid #E2E8F0 !important;
        background-color: #F8FAFC !important;
        color: #1E293B !important;
        font-weight: 500 !important;
        transition: all 0.3s ease !important;
    }
    div[data-baseweb="input"] > div:focus-within, 
    div[data-baseweb="select"] > div:focus-within {
        border-color: #0086FF !important;
        background-color: #FFFFFF !important;
        box-shadow: 0 0 0 3px rgba(0,134,255,0.15) !important;
    }

    /* 7.1 CHIPS DO MULTISELECT (Tags da Equipe Azul Magalu) */
    span[data-baseweb="tag"] {
        background-color: #E6F2FF !important;
        color: #0086FF !important;
        border-radius: 8px !important;
        border: 1px solid #BAE6FD !important;
        font-weight: 700 !important;
        padding: 6px 12px !important;
        margin: 4px 4px 4px 0px !important;
    }
    span[data-baseweb="tag"] svg { fill: #0086FF !important; } /* Cor do X de fechar */

    /* 8. BOTÕES PRINCIPAIS E SECUNDÁRIOS */
    .stButton>button {
        background-color: #FFFFFF; color: #0F172A; border: 1px solid #CBD5E1; 
        border-radius: 12px; font-weight: 700; font-size: 15px; padding: 0.8rem 1.2rem; 
        box-shadow: 0 2px 4px rgba(0,0,0,0.02); transition: all 0.3s ease;
    }
    .stButton>button:hover { 
        border-color: #0086FF; color: #0086FF; transform: translateY(-1px);
    }
    
    /* Botão Verde "Finalizar" com Efeito Glow/Pulse */
    button[kind="primary"] {
        background: linear-gradient(135deg, #00C853 0%, #009624 100%) !important;
        border: none !important; color: white !important; font-weight: 800 !important;
        border-radius: 10px !important; box-shadow: 0 6px 20px rgba(0,200,83,0.3) !important;
        transition: all 0.3s ease !important;
    }
    button[kind="primary"]:hover {
        box-shadow: 0 8px 25px rgba(0,200,83,0.5) !important; transform: translateY(-2px) !important;
    }

    /* 9. SCROLLBAR DE LUXO (O que separa os meninos dos homens no Front-end) */
    ::-webkit-scrollbar { width: 8px; height: 8px; }
    ::-webkit-scrollbar-track { background: transparent; }
    ::-webkit-scrollbar-thumb { background: #CBD5E1; border-radius: 10px; }
    ::-webkit-scrollbar-thumb:hover { background: #94A3B8; }

    /* 10. KPI CARDS FINANCEIROS */
    .kpi-card { 
        background: #FFFFFF; border-radius: 16px; padding: 20px 15px; 
        border: 1px solid #F1F5F9; box-shadow: 0 4px 20px rgba(0, 0, 0, 0.03);
        margin-bottom: 12px; display: flex; flex-direction: column; 
        align-items: center; justify-content: center; transition: all 0.3s ease;
    }
    .kpi-card:hover { transform: translateY(-4px); box-shadow: 0 12px 30px rgba(0, 134, 255, 0.08); }
    .kpi-title { color: #64748B; font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px;}
    .kpi-value { color: #0F172A; font-size: 24px; font-weight: 900; letter-spacing: -0.5px; }
    </style>
""", unsafe_allow_html=True)

# ==========================================================
# 2. CONEXÃO GOOGLE SHEETS & CACHES
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
    client = conectar_google()
    sh = client.open_by_key("1NWH9BHXgUmS-6WCQ8AjAHbt8DUHIvgQLRJ8hwUSDC7U")
    ws = sh.worksheet("HISTÓRICO 2025")
    return pd.DataFrame(ws.get_all_values()[1:], columns=ws.get_all_values()[0])

@st.cache_data(ttl=60)
def carregar_equipe():
    client = conectar_google()
    sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
    ws = sh.worksheet("QUADRO CARGA e DESCARGA")
    return pd.DataFrame(ws.get_all_records())

@st.cache_data(ttl=10)
def carregar_log_produtividade():
    client = conectar_google()
    sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
    try:
        ws = sh.worksheet("LOG_PRODUTIVIDADE")
        return pd.DataFrame(ws.get_all_records())
    except:
        return pd.DataFrame()

@st.cache_data(ttl=60)
def carregar_aux():
    client = conectar_google()
    sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
    try:
        ws = sh.worksheet("aux")
        return pd.DataFrame(ws.get_all_records())
    except:
        return pd.DataFrame()

@st.cache_data(ttl=60)
def carregar_matriz():
    client = conectar_google()
    sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
    try:
        ws = sh.worksheet("MATRIZ_COMPETÊNCIA")
        return pd.DataFrame(ws.get_all_records())
    except:
        return pd.DataFrame()

@st.cache_data(ttl=60)
def carregar_docas_finalizadas():
    client = conectar_google()
    sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
    try:
        # ATENÇÃO: Confirme se o nome da sua aba de finalizados é esse mesmo!
        ws = sh.worksheet("DOCAS_FINALIZADAS") 
        df = pd.DataFrame(ws.get_all_records())
        return df
    except Exception as e:
        return pd.DataFrame()

@st.cache_data(ttl=60)
def carregar_auxexp():
    client = conectar_google()
    sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
    try:
        ws = sh.worksheet("auxexp")
        return pd.DataFrame(ws.get_all_records())
    except Exception as e:
        return pd.DataFrame()

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
    try:
        ws_final = sh.worksheet("DOCAS_FINALIZADAS")
        ws_final.append_rows(linhas_conclusao) 
        ws_log = sh.worksheet("LOG_PRODUTIVIDADE")
        ws_log.append_rows([linha_encerramento_log])
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

@st.dialog("⚠️ Confirmação de Transferência")
def exibir_popup_transferencia(doca_sel, agenda_sel, conferente_sel, equipe_sel, conflitos, info_docas):
    st.write("Colaboradores já alocados em outras docas:")
    for p, d in conflitos.items():
        st.markdown(f"- **{p}** (Sairá da Doca **{d}**)")
        
    st.write(f"Confirma a transferência para a **Doca {doca_sel}**?")
    st.markdown("<br>", unsafe_allow_html=True)
    
    c1, c2 = st.columns(2)
    if c1.button("✅ Sim, Transferir", use_container_width=True):
        with st.spinner("Atualizando docas..."):
            if processar_gravacao_doca(doca_sel, agenda_sel, conferente_sel, equipe_sel, conflitos, info_docas, False):
                carregar_log_produtividade.clear()
                st.rerun() 
                
    if c2.button("❌ Cancelar", use_container_width=True):
        st.rerun()

# --- POP-UP MAGALU: JUSTIFICATIVA DE ATRASO ---
@st.dialog("📝 Justificativa de Atraso")
def exibir_popup_justificativa(dados_multiplos, linha_log_fecha, categoria_carga):
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
        
        # O PULO DO GATO: Adiciona a Justificativa e depois a CATEGORIA em todas as linhas
        for linha in dados_multiplos:
            linha.append(justificativa_final)
            linha.append(categoria_carga)
            
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
    df_h = pd.concat([df_com_wms, df_sem_wms], ignore_index=True).drop(columns=['PRIORIDADE_STATUS'])

    df_h['ANO'] = df_h['DATA AGENDADA'].dt.year.astype('Int64')
    df_h['MES_ORDENACAO'] = df_h['DATA AGENDADA'].dt.to_period('M')
    meses_pt = {1:'Jan', 2:'Fev', 3:'Mar', 4:'Abr', 5:'Mai', 6:'Jun', 7:'Jul', 8:'Ago', 9:'Set', 10:'Out', 11:'Nov', 12:'Dez'}
    df_h['MES_NOME'] = df_h['DATA AGENDADA'].dt.month.map(meses_pt) + "/" + df_h['ANO'].astype(str)

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
    df_h['LINHA'] = df_h['LINHA'].astype(str).str.strip().str.upper()
    df_h['FORNECEDOR/SELLER'] = df_h['FORNECEDOR/SELLER'].astype(str).str.strip().str.upper()
    df_h['CATEGORIA'] = df_h['CATEGORIA'].astype(str).str.strip().str.upper()

    def eh_interno(forn):
        if forn.startswith(('MAGAZINE', 'FILIAL')): return True
        if re.match(r'^C[D\d]', forn): return True 
        return False
    df_h = df_h[~df_h['FORNECEDOR/SELLER'].apply(eh_interno)]

    mask_full = df_h['LINHA'].str.contains('FULL', na=False) | df_h['CATEGORIA'].str.contains('FULL', na=False)
    df_full = df_h[mask_full].copy()
    df_main = df_h[~mask_full].copy()

    pag_l = df_main[df_main['VALOR_REAL'] > 0]['LINHA'].unique()
    pag_c = df_main[df_main['VALOR_REAL'] > 0]['CATEGORIA'].unique()
    df_main = df_main[(df_main['LINHA'].isin([l for l in pag_l if l != ''])) | (df_main['CATEGORIA'].isin([c for c in pag_c if c != '']))].copy()

    df_ok = df_main[(df_main['STATUS'] == 'OK') & (df_main['VALOR_REAL'] > 0)]
    m_linha = df_ok.groupby('LINHA')['VALOR_REAL'].mean().to_dict()
    m_cat = df_ok.groupby('CATEGORIA')['VALOR_REAL'].mean().to_dict()

    df_main['VALOR_PERDIDO'] = 0.0
    mask_aus = df_main['STATUS'] == 'AUSENTE'
    df_main.loc[mask_aus, 'VALOR_PERDIDO'] = df_main.loc[mask_aus, 'LINHA'].map(m_linha)
    mask_zero = mask_aus & (df_main['VALOR_PERDIDO'].isna() | (df_main['VALOR_PERDIDO'] == 0))
    df_main.loc[mask_zero, 'VALOR_PERDIDO'] = df_main.loc[mask_zero, 'CATEGORIA'].map(m_cat)
    df_main['VALOR_PERDIDO'] = df_main['VALOR_PERDIDO'].fillna(0).round(2)

    df_full['VALOR_ESTIMADO'] = df_full['CATEGORIA'].map(m_cat)
    df_full.loc[df_full['CATEGORIA'] == 'DIVERSOS', 'VALOR_ESTIMADO'] = 350.00
    df_full['VALOR_ESTIMADO'] = df_full['VALOR_ESTIMADO'].fillna(500.00).round(2)

    return df_main, df_full


# ==========================================================
# 5. ROTEADOR DE MÓDULOS (SIDEBAR)
# ==========================================================

# Imagem com a linha RGB animada embaixo (Estilo Moldura Tech)
# Dica: Se quiser trocar a foto para uma imagem do CD 2900, basta trocar o link do "src" abaixo!
st.sidebar.markdown("""
    <div style="text-align: center; margin-bottom: 25px; margin-top: 10px;">
        <img src="https://media.discordapp.net/attachments/877222773720772631/1492324332376952882/Gemini_Generated_Image_kemjo1kemjo1kemj.png?ex=69daead6&is=69d99956&hm=229e06b80d7a0109065a5189ebd28575a6a1bb4e6fbbc277189703e8641df52e&=&format=webp&quality=lossless&width=534&height=648" style="max-width: 65%; margin-bottom: 18px; filter: drop-shadow(0px 4px 6px rgba(0,0,0,0.05));">
        <div style="height: 10px; width: 100%; border-radius: 2px; background: linear-gradient(90deg, #0086FF, #FF007F, #00C853, #0086FF); background-size: 300% 300%; animation: magaluGlow 6s linear infinite;"></div>
    </div>
""", unsafe_allow_html=True)

# O Menu de Navegação Limpo e Direto
pagina_selecionada = st.sidebar.radio(
    "",
    ["📋 Registro Absenteísmo", "🚛 Gestão de Docas", "📅 Registro de Alinhamento","📈 Produtividade (NS & Equipe)", "📊 Financeiro (Diretoria)"]
)

st.sidebar.markdown("---")

# Botão Anti-Pânico (Forçar Atualização do Banco de Dados)
st.sidebar.markdown("<br>", unsafe_allow_html=True)
if st.sidebar.button("🔄 Sincronizar Agora", type="secondary", use_container_width=True):
    with st.spinner("📡 Puxando dados em tempo real da base..."):
        # Limpa toda a memória cache do sistema
        st.cache_data.clear()
        # Recarrega a página instantaneamente
        st.rerun()

# ==========================================================
# MÓDULO 1: ABSENTEÍSMO
# ==========================================================
if pagina_selecionada == "📋 Registro Absenteísmo":
    st.markdown('<div class="magalu-page-title">Lançamento de Ausências</div>', unsafe_allow_html=True)
    st.markdown('<div class="magalu-page-subtitle">Pátio / Docas</div>', unsafe_allow_html=True)
    
    try:
        df_equipe = carregar_equipe()
        st.markdown('<div class="magalu-card">', unsafe_allow_html=True)
        data_chamada = st.date_input("Data", date.today())
        busca = st.text_input("🔍 Buscar Colaborador", placeholder="ID ou Nome...")
        st.markdown('</div>', unsafe_allow_html=True)

        if busca:
            df_filtrado = df_equipe[df_equipe['NOME'].str.contains(busca, case=False, na=False) | df_equipe['ID'].astype(str).str.contains(busca, na=False)].copy()
        else:
            df_filtrado = df_equipe.copy()

        df_filtrado['OCORRÊNCIA'] = "PRESENTE" 
        opcoes_ocorrencia = ["PRESENTE", "FALTA", "DSR", "BH", "LICENÇA", "ATESTADO"]

        st.markdown('<div class="magalu-ribbon">Registro da Equipe</div>', unsafe_allow_html=True)
        
        df_editado = st.data_editor
            # 1. Garante que o ID seja lido como Texto puro (removendo qualquer .0 que o Pandas possa inventar)
        df_filtrado['ID'] = df_filtrado['ID'].astype(str).str.replace('\.0$', '', regex=True)

        # 2. O Editor com a correção no Column Config
        df_editado = st.data_editor
            # 1. Fazemos uma CÓPIA segura para o Pandas não ignorar nossas alterações
        df_filtrado = df_filtrado.copy()

        # 2. Limpeza Extrema (Mata NaN, tira espaços, e remove o '.0' se existir)
        df_filtrado['ID'] = df_filtrado['ID'].fillna('')
        df_filtrado['ID'] = df_filtrado['ID'].astype(str).str.strip()
        df_filtrado['ID'] = df_filtrado['ID'].replace(['nan', 'None', '<NA>'], '')
        df_filtrado['ID'] = df_filtrado['ID'].str.replace(r'\.0$', '', regex=True)

        # 3. O Editor de Dados Protegido
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
                        st.success(f"✅ {len(lista_final)} registros salvos!")
                        carregar_equipe.clear()
            else:
                st.warning("Nenhuma falta marcada.")

    except Exception as e:
        st.error(f"Erro na conexão com RH: {e}")

# ==========================================================
# MÓDULO 2: GESTÃO DE DOCAS E PRODUTIVIDADE
# ==========================================================
elif pagina_selecionada == "🚛 Gestão de Docas":
    st.markdown('<div class="magalu-page-title">Gestão de Docas</div>', unsafe_allow_html=True)
    st.markdown('<div class="magalu-page-subtitle">Controle Unificado: Recebimento & Expedição Consolidada.</div>', unsafe_allow_html=True)
    
    df_log = carregar_log_produtividade()
    df_matriz = carregar_matriz()
    df_aux_rec = carregar_aux()
    df_aux_exp = carregar_auxexp()
    
    # --- 1. PROCESSAMENTO DA EXPEDIÇÃO (CONSOLIDAÇÃO DE DOCAS) ---
    if not df_aux_exp.empty:
        # Garante que a coluna exista para não dar erro
        if 'Status Chegada' not in df_aux_exp.columns:
            df_aux_exp['Status Chegada'] = ''
            
        df_aux_exp = df_aux_exp.rename(columns={
            'ID Carga': 'AGENDA WMS',
            'Plano de Transporte': 'CATEGORIA',
            'M³ Total': 'VOLUME_M3',
            'Ped Venda': 'PEDIDOS',
            'Limit Carreg': 'LIMITE_RAW',
            'Doca': 'DOCA_ORIGINAL'
        })

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
                'PEDIDOS': 'sum',
                'LIMITE_RAW': 'min',
                'STATUS_CALC': lambda x: 'AGUARDANDO' if 'AGUARDANDO' in x.values else 'LIBERADO',
                # A MÁGICA AQUI: Puxando o status do caminhão!
                'Status Chegada': lambda x: 'AGUARD CHEGADA' if any('AGUARD' in str(v).upper() for v in x.dropna()) else 'NO PÁTIO'
            }).reset_index()

            df_exp_grouped['TIPO_OPERACAO'] = "⬆️ EXPEDIÇÃO"
            df_exp_grouped['STATUS'] = df_exp_grouped['STATUS_CALC']
            df_exp_grouped['STATUS_CHEGADA_RAW'] = df_exp_grouped['Status Chegada'] # Salva para o painel ler depois
            df_exp_grouped['LINHA'] = df_exp_grouped['CATEGORIA'] 
            df_exp_grouped['PEÇAS'] = df_exp_grouped['VOLUME_M3'] 
            df_exp_grouped['SKU'] = df_exp_grouped['PEDIDOS']    
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
        df_aux_rec['TIPO_OPERACAO'] = "⬇️ RECEBIMENTO"
        df_aux_rec_final = df_aux_rec
    else:
        df_aux_rec_final = pd.DataFrame()

    # --- 3. FUSÃO E IDENTIFICAÇÃO ---
    df_aux = pd.concat([df_aux_rec_final, df_aux_exp_final], ignore_index=True)
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
                t_badges.append("🛒 ECOM")
                h_badges.append("<span style='background:#E0F2FE; color:#0284C7; padding:2px 6px; border-radius:4px; font-size:9.5px; font-weight:800; margin-left:6px; border:1px solid #BAE6FD;'>ECOM</span>")
            if str(row.get('FULL', '')).upper() in ['TRUE', '1', 'SIM']:
                t_badges.append("📦 FULL")
                h_badges.append("<span style='background:#FAE8FF; color:#C026D3; padding:2px 6px; border-radius:4px; font-size:9.5px; font-weight:800; margin-left:6px; border:1px solid #F0ABFC;'>FULL</span>")
            if str(row.get('CARREGAMENTO', '')).upper() in ['TRUE', '1', 'SIM']:
                t_badges.append("⬆️ CARREG")
                h_badges.append("<span style='background:#FFEDD5; color:#EA580C; padding:2px 6px; border-radius:4px; font-size:9.5px; font-weight:800; margin-left:6px; border:1px solid #FDBA74;'>CARREG</span>")
            if str(row.get('DESCARGA', '')).upper() in ['TRUE', '1', 'SIM']:
                t_badges.append("⬇️ DESC")
                h_badges.append("<span style='background:#DCFCE7; color:#16A34A; padding:2px 6px; border-radius:4px; font-size:9.5px; font-weight:800; margin-left:6px; border:1px solid #86EFAC;'>DESC</span>")
            if str(row.get('ENSINAR', '')).upper() in ['TRUE', '1', 'SIM']:
                t_badges.append("⭐ MESTRE")
                h_badges.append("<span style='background:#FEF9C3; color:#CA8A04; padding:2px 6px; border-radius:4px; font-size:9.5px; font-weight:800; margin-left:6px; border:1px solid #FDE047;'>⭐ MESTRE</span>")
            if t_badges:
                dict_skills_text[nome_matriz] = " | ".join(t_badges)
                dict_skills_html[nome_matriz] = "".join(h_badges)

    def renderizar_equipe_html(string_pessoas):
        if pd.isna(string_pessoas) or str(string_pessoas).strip() == '': return ""
        lista_p = [x.strip() for x in str(string_pessoas).split(',')]
        html_final = ""
        for p in lista_p:
            habilidades = dict_skills_html.get(p, "")
            html_final += f"<div style='background:#FFFFFF; border:1px solid #CBD5E1; padding:6px 10px; border-radius:8px; display:inline-block; margin:4px 6px 4px 0px; font-size:12px; color:#1E293B; box-shadow:0 2px 4px rgba(0,0,0,0.02);'><b>{p}</b>{habilidades}</div>"
        return html_final

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

    # --- LÓGICA DE AUDITORIA DE FADIGA ERGONÔMICA (SST) ---
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

    # --- FUNÇÃO GRAVADORA MASTER (Salva Categoria no Log) ---
    def processar_gravacao_doca_v2(doca_n, agenda_n, conferente_n, equipe_n, conflitos_n, info_docas_n):
        agora_dt = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
        agora_str = agora_dt.strftime("%d/%m/%Y %H:%M:%S")
        linhas_para_gravar = []

        cat_atual = "NÃO INFORMADA"
        if not df_aux.empty:
            match_atual = df_aux[df_aux['AGENDA WMS'] == str(agenda_n).strip()]
            if not match_atual.empty: cat_atual = str(match_atual.iloc[0].get('LINHA', match_atual.iloc[0].get('CATEGORIA', ''))).upper()

        if conflitos_n:
            for pessoa, doca_antiga in conflitos_n.items():
                if doca_antiga in info_docas_n:
                    agenda_antiga = info_docas_n[doca_antiga]['agenda']
                    conf_antiga = info_docas_n[doca_antiga]['conferente']
                    cat_antiga = "NÃO INFORMADA"
                    if not df_aux.empty:
                        match_ant = df_aux[df_aux['AGENDA WMS'] == str(agenda_antiga).strip()]
                        if not match_ant.empty: cat_antiga = str(match_ant.iloc[0].get('LINHA', match_ant.iloc[0].get('CATEGORIA', ''))).upper()
                    linhas_para_gravar.append([agora_str, doca_antiga, agenda_antiga, conf_antiga, "ENCERRADO", cat_antiga])

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

    # --- POP-UP: START NA CARGA ---
    @st.dialog("🚀 START na Carga (Alocar Equipe)")
    def popup_start_carga(doca_sel, agenda_sel, conferente_sel):
        st.markdown(f"<div style='font-size:14px; margin-bottom:15px;'>Doca: <b>{doca_sel}</b> &nbsp;|&nbsp; Agenda: <b>{agenda_sel}</b> &nbsp;|&nbsp; Líder: <b>{conferente_sel}</b></div>", unsafe_allow_html=True)
        equipe_sel = st.multiselect("Selecione os colaboradores para iniciar:", options=lista_auxiliares, format_func=lambda x: f"{x}  [{dict_skills_text[x]}]" if x in dict_skills_text else x)
        
        conflitos = {}
        for pessoa in equipe_sel:
            if pessoa in mapa_pessoas:
                if mapa_pessoas[pessoa] != str(doca_sel).strip(): conflitos[pessoa] = mapa_pessoas[pessoa]
        if conflitos:
            st.warning("⚠️ Os colaboradores abaixo já estão ocupados e serão transferidos:")
            for p, d in conflitos.items(): st.markdown(f"- **{p}** (Sairá da Doca {d})")

        fadigados = checar_fadiga(equipe_sel, agenda_sel, df_log, df_aux)
        bloqueio_ergonomico = False
        
        if fadigados:
            st.markdown(f"<div style='background-color: #FEF2F2; border: 1px solid #DC2626; border-radius: 8px; padding: 15px; margin-top: 15px; margin-bottom: 15px;'><b style='color: #DC2626;'>🚨 ALERTA ERGONÔMICO (SST)</b><br><span style='color: #7F1D1D; font-size: 13px;'>Os colaboradores <b>{', '.join(fadigados)}</b> já atuaram em carga pesada (Madeira) nas últimas 24h.</span></div>", unsafe_allow_html=True)
            ciente = st.checkbox("Declaro ciência do risco e autorizo a alocação.", key="chk_fadiga_popup")
            if not ciente: bloqueio_ergonomico = True
                
        st.markdown('<br>', unsafe_allow_html=True)
        if st.button("🚀 CONFIRMAR START", type="primary", use_container_width=True):
            if not doca_sel or doca_sel == "A Definir": st.error("Esta carga precisa ter uma Doca informada antes de iniciar!")
            elif not equipe_sel: st.error("Selecione a equipe!")
            elif bloqueio_ergonomico: st.error("Você precisa assumir o risco ergonômico marcando a caixa de seleção!")
            else:
                with st.spinner("Iniciando operação..."):
                    sucesso = processar_gravacao_doca_v2(doca_sel, agenda_sel, conferente_sel, equipe_sel, conflitos, info_docas)
                    if sucesso:
                        st.success("Carga iniciada!")
                        carregar_log_produtividade.clear()
                        st.rerun()

    # --- POP-UP: GERENCIAR OPERADOR (TRANSFERIR/RETIRAR) ---
    @st.dialog("🔄 Gerenciar Operador da Doca")
    def popup_gerenciar_operador(doca_origem, equipe_atual, info_docas_global):
        doca_origem_str = str(doca_origem).strip()
        st.markdown(f"<div style='color:#64748B; margin-bottom:15px;'>Modificando a equipe da <b>Doca {doca_origem_str}</b></div>", unsafe_allow_html=True)
        operador_sel = st.selectbox("Selecione o Operador que deseja movimentar:", equipe_atual)
        acao = st.radio("O que deseja fazer com este colaborador?", ["❌ Retirar da Operação (Ficará Livre no Pátio)", "➡️ Transferir para outra Doca ativa"])
        docas_ativas = [d for d in info_docas_global.keys() if str(d).strip() != doca_origem_str]
        doca_destino = None
        if "Transferir" in acao:
            if not docas_ativas: st.warning("Não há outras docas em processo no momento para transferir.")
            else:
                opcoes_formatadas = [f"Doca {d} (Líder: {info_docas_global[d]['conferente']})" for d in docas_ativas]
                doca_destino = st.selectbox("Transferir para qual Doca?", opcoes_formatadas).split(" ")[1] 
                
        st.markdown('<br>', unsafe_allow_html=True)
        if st.button("🔄 Confirmar Alteração", type="primary", use_container_width=True):
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

    # --- FILTROS GLOBAIS DE OPERAÇÃO ---
    st.markdown("<br>", unsafe_allow_html=True)
    with st.container(border=True):
        c_f1, c_f2 = st.columns(2)
        with c_f1:
            filtro_op = st.selectbox("🔍 Filtrar por Operação:", ["Todas", "⬇️ RECEBIMENTO", "⬆️ EXPEDIÇÃO"])
        with c_f2:
            filtro_doca = st.text_input("📍 Buscar Doca ou Agenda:", placeholder="Ex: 14 ou 99999")
    st.markdown("<br>", unsafe_allow_html=True)

    # Criação das TRÊS abas
    aba1, aba2, aba3 = st.tabs([" Visão das Docas (EM PROCESSO)", " Fila de Docas (PENDENTE)", " Montar Equipes"])

    # --- ABA 1: EM PROCESSO (Ativas) ---
    with aba1:
        if not df_log.empty:
            agendas_logadas = df_log['AGENDA'].astype(str).str.strip().unique().tolist()
            df_ativos = df_ativos_bruto.copy()
            df_ativos = df_ativos.groupby(['DOCA', 'AGENDA', 'CONFERENTE', 'DATA_HORA', 'DATA_HORA_DT'])['AUXILIARES'].apply(lambda x: ', '.join(x.dropna())).reset_index()
            df_ativos = df_ativos[df_ativos['AUXILIARES'] != 'ENCERRADO']
            
            agora_dt = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
            
            # --- ALGORITMO DE ORDENAÇÃO (SLA) ---
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
                except: return 99999 # Joga pro final se der erro

            # Calcula a urgência e ORDENA (Os mais negativos = mais atrasados sobem pro topo!)
            df_ativos['URGENCIA'] = df_ativos.apply(calc_urgencia_processo, axis=1)
            df_ativos = df_ativos.sort_values('URGENCIA', ascending=True)

            if df_ativos.empty: st.info("Nenhuma doca ativa no momento. Pátio limpo!")
            else:
                cards_exibidos_aba1 = 0
                for index, row in df_ativos.iterrows():
                    agenda_str = str(row['AGENDA']).strip()
                    
                    # PEGA O TIPO DE OPERAÇÃO ANTES PARA APLICAR O FILTRO
                    aux_row = pd.Series()
                    if not df_aux.empty and agenda_str in df_aux['AGENDA WMS'].values:
                        aux_row = df_aux[df_aux['AGENDA WMS'] == agenda_str].iloc[0]
                    tipo_op = str(aux_row.get('TIPO_OPERACAO', '⬇️ RECEBIMENTO'))
                    
                    # APLICA OS FILTROS DA TELA
                    if filtro_op != "Todas" and filtro_op not in tipo_op: continue
                    if filtro_doca and filtro_doca not in str(row['DOCA']) and filtro_doca not in agenda_str: continue
                    
                    cards_exibidos_aba1 += 1
                    
                    info = {'LINHA': '-', 'SKU': '-', 'PEÇAS': '-', 'VALOR': '-', 'PAGTO': '-', 'STATUS': '-'}
                    meta_minutos = 60
                    if not aux_row.empty:
                        pagto_str = "✅ Sim" if str(aux_row.get('PAGAMENTO', '')).upper() == 'TRUE' else "⏳ Pendente"
                        valor_desc = aux_row.get('R$ DESCARGA', '-')
                        if str(valor_desc).replace('.','',1).isdigit(): valor_desc = f"R$ {float(valor_desc):,.2f}".replace(',','X').replace('.',',').replace('X','.')
                        linha_val = aux_row.get('LINHA', aux_row.get('CATEGORIA', '-'))
                        info = {'LINHA': linha_val, 'SKU': aux_row.get('SKU', '-'), 'PEÇAS': aux_row.get('PEÇAS', '-'), 'VALOR': valor_desc, 'PAGTO': pagto_str, 'STATUS': aux_row.get('STATUS', '-')}
                        try:
                            col_meta = next((c for c in df_aux.columns if 'META' in str(c).upper()), None)
                            if col_meta and pd.notna(aux_row[col_meta]): meta_minutos = int(float(str(aux_row[col_meta]).replace(',', '.')))
                        except: pass

                    inicio_dt = row['DATA_HORA_DT']
                    if pd.isna(inicio_dt): inicio_dt = agora_dt
                    decorrido_min = (agora_dt - inicio_dt).total_seconds() / 60
                    restante_min = meta_minutos - decorrido_min
                    
                    if restante_min >= 0:
                        h, m = int(restante_min // 60), int(restante_min % 60)
                        cor_timer, bg_timer, txt_timer = "#00C853", "#E6F9EC", f"⏳ Resta {h:02d}h{m:02d}m"
                    else:
                        atraso = abs(restante_min)
                        h, m = int(atraso // 60), int(atraso % 60)
                        cor_timer, bg_timer, txt_timer = "#DC2626", "#FEF2F2", f"🚨 Atraso -{h:02d}h{m:02d}m"

                    html_equipe_cards = renderizar_equipe_html(row['AUXILIARES'])
                    auxiliares_lista = [x.strip() for x in str(row['AUXILIARES']).split(',')]

                    with st.container(border=True):
                        if "EXPEDIÇÃO" in tipo_op:
                            cor_tema = "#0086FF"
                            css_hack = f"<style>div[data-testid='stVerticalBlockBorderWrapper']:has(.card-proc-{index}) {{ border: 2px solid {cor_tema} !important; box-shadow: 0 4px 15px rgba(0,134,255,0.15) !important; }}</style><div class='card-proc-{index}'></div>"
                            html_detalhes = f"<div style='font-size: 11.5px; color: #475569; background-color: #F0F9FF; padding: 10px; border-radius: 8px; margin-bottom: 8px; border: 1px solid #BAE6FD;'><b>Planos:</b> <span style='color:#0369A1; font-weight:bold;'>{info['LINHA']}</span><br><div style='margin-top: 4px;'><b>M³ Total:</b> {info['PEÇAS']} &nbsp;|&nbsp; <b>Pedidos:</b> {info['SKU']} &nbsp;|&nbsp; <b>Status:</b> <span style='color:#0284C7; font-weight:bold;'>{info['STATUS']}</span></div></div>"
                        else:
                            cor_tema = "#EA580C"
                            css_hack = f"<style>div[data-testid='stVerticalBlockBorderWrapper']:has(.card-proc-{index}) {{ border: 2px solid {cor_tema} !important; box-shadow: 0 4px 15px rgba(234,88,12,0.15) !important; }}</style><div class='card-proc-{index}'></div>"
                            html_detalhes = f"<div style='font-size: 11.5px; color: #475569; background-color: #FFF7ED; padding: 10px; border-radius: 8px; margin-bottom: 8px; border: 1px solid #FFEDD5;'><b>Linha:</b> {info['LINHA']} &nbsp;|&nbsp; <b>SKU:</b> {info['SKU']} &nbsp;|&nbsp; <b>Peças:</b> {info['PEÇAS']}<br><div style='margin-top: 4px;'><b>Valor Carga:</b> {info['VALOR']} &nbsp;|&nbsp; <b>Pagto:</b> {info['PAGTO']} &nbsp;|&nbsp; <b>Status:</b> <span style='color:#EA580C; font-weight:bold;'>{info['STATUS']}</span></div></div>"

                        st.markdown(css_hack, unsafe_allow_html=True)
                        c_title, c_time = st.columns([5, 5])
                        c_title.markdown(f"<h4 style='margin:0; color:{cor_tema};'>Doca {row['DOCA']}</h4>", unsafe_allow_html=True)
                        c_time.markdown(f"<div style='text-align:right;'><div style='font-size:11px; color:#64748B; margin-bottom: 2px;'>⌚ Início: {row['DATA_HORA']}</div><div style='display:inline-block; font-size:12.5px; font-weight:800; color:{cor_timer}; background-color:{bg_timer}; padding:3px 6px; border-radius:4px; border: 1px solid {cor_timer};'>{txt_timer} <span style='font-size:10px; font-weight:normal;'>(Meta: {meta_minutos}m)</span></div></div>", unsafe_allow_html=True)
                        st.markdown(f"<div style='font-size: 13px; margin: 4px 0px 4px 0px;'><b>Agenda:</b> {row['AGENDA']} | <b>Líder:</b> {row['CONFERENTE']}</div>", unsafe_allow_html=True)
                        
                        st.markdown(html_detalhes, unsafe_allow_html=True)
                        st.markdown(f"<div style='background-color: #F1F5F9; padding: 12px; border-radius: 8px; border: 1px dashed #CBD5E1; margin-bottom: 15px;'><div style='font-size: 11px; font-weight: 800; color: #64748B; margin-bottom: 6px; text-transform: uppercase;'>Operadores Alocados:</div>{html_equipe_cards}</div>", unsafe_allow_html=True)
                        
                        c_eq, c_btn = st.columns([7, 3])
                        with c_btn:
                            if st.button("✅ Finalizar Operação", key=f"btn_fin_{row['DOCA']}_{index}", type="primary", use_container_width=True):
                                clique_dt = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
                                duracao_final = clique_dt - row['DATA_HORA_DT']
                                total_minutos_final = int(duracao_final.total_seconds() / 60)
                                horas, mins = total_minutos_final // 60, total_minutos_final % 60
                                tempo_str = f"{horas:02d}:{mins:02d}"
                                
                                cat_final = str(info['LINHA']).upper()
                                linhas_conclusao_multiplas = []
                                for pessoa in auxiliares_lista:
                                    linhas_conclusao_multiplas.append([clique_dt.strftime("%d/%m/%Y"), str(row['DOCA']), str(row['AGENDA']), str(row['CONFERENTE']), len(auxiliares_lista), pessoa, row['DATA_HORA'], clique_dt.strftime("%H:%M:%S"), tempo_str])
                                
                                linha_log_fecha = [clique_dt.strftime("%d/%m/%Y %H:%M:%S"), str(row['DOCA']), row['AGENDA'], row['CONFERENTE'], "ENCERRADO", cat_final]
                                
                                if restante_min < 0: 
                                    exibir_popup_justificativa(linhas_conclusao_multiplas, linha_log_fecha, cat_final)
                                else:
                                    for linha in linhas_conclusao_multiplas: 
                                        linha.append("No Prazo")
                                        linha.append(cat_final)
                                    with st.spinner("Finalizando..."):
                                        if gravar_conclusao_doca(linhas_conclusao_multiplas, linha_log_fecha):
                                            st.success(f"Doca finalizada com sucesso!")
                                            carregar_log_produtividade.clear()
                                            st.rerun()
                                            
                            st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True) 
                            if st.button("🔄 Mover/Retirar Alguém", key=f"btn_mgr_{row['DOCA']}_{index}", use_container_width=True):
                                popup_gerenciar_operador(row['DOCA'], auxiliares_lista, info_docas)
                                
                if cards_exibidos_aba1 == 0: st.info("Nenhuma doca encontrada com esses filtros.")
        else:
            st.info("O Log de Operações ainda está vazio.")

    # --- ABA 2: PENDENTES (Fila de Espera) ---
    with aba2:
        if not df_aux.empty:
            df_pendentes = df_aux[~df_aux['AGENDA WMS'].isin(agendas_logadas)].copy()
            df_pendentes = df_pendentes[df_pendentes['AGENDA WMS'] != '']
            status_ignorados = ['AUSENTE', 'DEVOLVIDA', 'OK','LIBERADO']
            if 'STATUS' in df_pendentes.columns: df_pendentes = df_pendentes[~df_pendentes['STATUS'].astype(str).str.upper().isin(status_ignorados)]
            
            agora_dt = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
            
            # --- ALGORITMO DE ORDENAÇÃO DE FILA (SLA PENDENTE) ---
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
                        
                        # --- CORREÇÃO DA VIRADA DE DIA (MADRUGADA) ---
                        # Se o limite é de madrugada (ex: 01h) e agora já passou de meio dia, o limite é AMANHÃ!
                        if h_lim < 12 and agora_dt.hour >= 12:
                            limite_dt += datetime.timedelta(days=1)
                        # Se agora é madrugada (ex: 01h) e o limite era noite (ex: 23h), o limite era ONTEM!
                        elif h_lim >= 18 and agora_dt.hour < 12:
                            limite_dt -= datetime.timedelta(days=1)
                            
                        hora_max_inicio = limite_dt - datetime.timedelta(minutes=meta_minutos)
                        return (hora_max_inicio - agora_dt).total_seconds() / 60
                except: pass
                return 99999 # Joga pro fim da fila se não tiver meta

            df_pendentes['URGENCIA'] = df_pendentes.apply(calc_urgencia_pendente, axis=1)
            df_pendentes = df_pendentes.sort_values('URGENCIA', ascending=True)

            if df_pendentes.empty: st.info("Nenhuma agenda aguardando equipe. Pátio zerado!")
            else:
                cards_exibidos_aba2 = 0
                for index, row in df_pendentes.iterrows():
                    tipo_op = str(row.get('TIPO_OPERACAO', '⬇️ RECEBIMENTO'))
                    agenda_str = str(row['AGENDA WMS'])
                    doca_str = str(row[col_doca]).strip() if col_doca and pd.notna(row[col_doca]) else "A Definir"
                    
                    if filtro_op != "Todas" and filtro_op not in tipo_op: continue
                    if filtro_doca and filtro_doca not in str(doca_str) and filtro_doca not in agenda_str: continue
                    
                    cards_exibidos_aba2 += 1
                    
                    if doca_str.lower() in ['nan', 'none', '']: doca_str = "A Definir"
                    conf_str = str(row[col_conf]).strip() if col_conf and pd.notna(row[col_conf]) else "A Definir"
                    if conf_str.lower() in ['nan', 'none', '']: conf_str = "A Definir"
                    pagto_str = "✅ Sim" if str(row.get('PAGAMENTO', '')).upper() == 'TRUE' else "⏳ Pendente"
                    valor_desc = row.get('R$ DESCARGA', '-')
                    if str(valor_desc).replace('.','',1).isdigit(): valor_desc = f"R$ {float(valor_desc):,.2f}".replace(',','X').replace('.',',').replace('X','.')
                    linha_val = row.get('LINHA', row.get('CATEGORIA', '-'))
                    info = {'LINHA': linha_val, 'SKU': row.get('SKU', '-'), 'PEÇAS': row.get('PEÇAS', '-'), 'VALOR': valor_desc, 'PAGTO': pagto_str, 'STATUS': row.get('STATUS', '-')}
                    meta_minutos = 60
                    limite_str, hora_max_str = "", "-"
                    txt_timer_pend, cor_timer_pend, bg_timer_pend = "⏳ Aguardando...", "#F59E0B", "#FEF3C7"
                    try:
                        col_meta = next((c for c in df_aux.columns if 'META' in str(c).upper()), None)
                        if col_meta and pd.notna(row[col_meta]): meta_minutos = int(float(str(row[col_meta]).replace(',', '.')))
                        col_limite = next((c for c in df_aux.columns if 'LIMITE' in str(c).upper()), None)
                        if col_limite and pd.notna(row[col_limite]) and str(row[col_limite]).strip() != '':
                            limite_str = str(row[col_limite]).strip()
                            partes = limite_str.split(':')
                            h_lim, m_lim = int(partes[0]), int(partes[1][:2])
                            limite_dt = agora_dt.replace(hour=h_lim, minute=m_lim, second=0, microsecond=0)
                            
                            # --- CORREÇÃO DA VIRADA DE DIA (Repetido pro visual) ---
                            if h_lim < 12 and agora_dt.hour >= 12:
                                limite_dt += datetime.timedelta(days=1)
                            elif h_lim >= 18 and agora_dt.hour < 12:
                                limite_dt -= datetime.timedelta(days=1)
                                
                            hora_max_inicio = limite_dt - datetime.timedelta(minutes=meta_minutos)
                            hora_max_str = hora_max_inicio.strftime("%H:%M")
                            diff_min = (hora_max_inicio - agora_dt).total_seconds() / 60
                            if diff_min >= 0:
                                h, m = int(diff_min // 60), int(diff_min % 60)
                                cor_timer_pend, bg_timer_pend, txt_timer_pend = "#00C853", "#E6F9EC", f"🟢 Sobra {h:02d}h{m:02d}m p/ Iniciar"
                            else:
                                atraso = abs(diff_min)
                                h, m = int(atraso // 60), int(atraso % 60)
                                cor_timer_pend, bg_timer_pend, txt_timer_pend = "#DC2626", "#FEF2F2", f"🚨 ATRASADO HÁ {h:02d}h{m:02d}m"
                    except: pass
                    
                    with st.container(border=True):
                        except: pass
                    
                    # --- INTERCEPTAÇÃO DE ANSIEDADE (CAMINHÃO NÃO CHEGOU) ---
                    tipo_op = str(row.get('TIPO_OPERACAO', '⬇️ RECEBIMENTO'))
                    status_chegada = str(row.get('STATUS_CHEGADA_RAW', '')).upper()
                    
                    # Se for expedição e o caminhão não estiver no CD, desliga a cor vermelha de atraso
                    if "EXPEDIÇÃO" in tipo_op and "AGUARD" in status_chegada:
                        txt_timer_pend = "🚚 AGUARDANDO VEÍCULO"
                        cor_timer_pend = "#64748B" # Cinza Sutil
                        bg_timer_pend = "#F1F5F9"  # Fundo Cinza Claro
                    
                    with st.container(border=True):
                        # ... resto do código do container igual ao seu ...
                        if "EXPEDIÇÃO" in tipo_op:
                            cor_tema = "#0086FF"
                            css_hack = f"<style>div[data-testid='stVerticalBlockBorderWrapper']:has(.card-pend-{index}) {{ border: 2px solid {cor_tema} !important; box-shadow: 0 4px 15px rgba(0,134,255,0.15) !important; }}</style><div class='card-pend-{index}'></div>"
                            html_detalhes = f"<div style='font-size: 11.5px; color: #475569; background-color: #F0F9FF; padding: 10px; border-radius: 8px; margin-bottom: 8px; border: 1px solid #BAE6FD;'><b>Planos:</b> <span style='color:#0369A1; font-weight:bold;'>{info['LINHA']}</span><br><div style='margin-top: 4px;'><b>M³ Total:</b> {info['PEÇAS']} &nbsp;|&nbsp; <b>Pedidos:</b> {info['SKU']} &nbsp;|&nbsp; <b>Status:</b> <span style='color:#0284C7; font-weight:bold;'>{info['STATUS']}</span></div></div>"
                        else:
                            cor_tema = "#EA580C"
                            css_hack = f"<style>div[data-testid='stVerticalBlockBorderWrapper']:has(.card-pend-{index}) {{ border: 2px solid {cor_tema} !important; box-shadow: 0 4px 15px rgba(234,88,12,0.15) !important; }}</style><div class='card-pend-{index}'></div>"
                            html_detalhes = f"<div style='font-size: 11.5px; color: #475569; background-color: #FFF7ED; padding: 10px; border-radius: 8px; margin-bottom: 8px; border: 1px solid #FFEDD5;'><b>Linha:</b> {info['LINHA']} &nbsp;|&nbsp; <b>SKU:</b> {info['SKU']} &nbsp;|&nbsp; <b>Peças:</b> {info['PEÇAS']}<br><div style='margin-top: 4px;'><b>Valor Carga:</b> {info['VALOR']} &nbsp;|&nbsp; <b>Pagto:</b> {info['PAGTO']} &nbsp;|&nbsp; <b>Status:</b> <span style='color:#EA580C; font-weight:bold;'>{info['STATUS']}</span></div></div>"

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
                        c_eq_pend.markdown(f"""<div style='font-size: 12px; color: #DC2626; background-color: #FEF2F2; padding: 8px; border-radius: 8px; border: 1px solid #FECACA;'><b>Equipe:</b> <span style="font-weight:900;">PENDENTE ALOCAÇÃO</span></div>""", unsafe_allow_html=True)
                        
                        with c_btn_pend:
                            if st.button("➕ Adicionar Equipe", key=f"btn_add_{index}", use_container_width=True): 
                                popup_start_carga(doca_str, agenda_str, conf_str)
                                
                if cards_exibidos_aba2 == 0: st.info("Nenhuma agenda encontrada com esses filtros.")

    # --- ABA 3: APONTAR / MOVIMENTAR (FORMULÁRIO MANUAL) ---
    with aba3:
        try:
            with st.container(border=True):
                st.markdown('<h4 style="color: #0086FF; margin-top: 0px; margin-bottom: 20px;">📍 Lançamento Manual / Atualizar</h4>', unsafe_allow_html=True)
                lista_agendas = []
                if not df_aux.empty: lista_agendas = df_aux[df_aux['AGENDA WMS'] != '']['AGENDA WMS'].unique().tolist()
                opcoes_agenda = [""] + lista_agendas + ["➕ DIGITAR OUTRA AGENDA..."]
                agenda_combo = st.selectbox("Nº da Agenda (Selecione ou digite para buscar)", options=opcoes_agenda, index=0)
                if agenda_combo == "➕ DIGITAR OUTRA AGENDA...": agenda_sel = st.text_input("Digite manualmente o Nº da Agenda", placeholder="Ex: 99999")
                else: agenda_sel = agenda_combo
                doca_padrao, conf_padrao = "", ""
                if agenda_sel and agenda_sel != "➕ DIGITAR OUTRA AGENDA..." and not df_aux.empty:
                    match = df_aux[df_aux['AGENDA WMS'] == agenda_sel.strip()]
                    if not match.empty:
                        col_l = [str(c).upper().strip() for c in match.columns]
                        for cr, cu in zip(match.columns, col_l):
                            if 'DOCA' in cu:
                                v = str(match.iloc[0][cr])
                                if v.lower() != 'nan' and v.strip() != '': doca_padrao = v.strip()
                                break
                        for cr, cu in zip(match.columns, col_l):
                            if 'CONFERENTE' in cu or 'LIDER' in cu or 'LÍDER' in cu:
                                v = str(match.iloc[0][cr])
                                if v.lower() != 'nan' and v.strip() != '': conf_padrao = v.strip()
                                break
                
                col1, col2 = st.columns(2)
                with col1: doca_sel = st.text_input("Número da Doca", value=doca_padrao, placeholder="Ex: 68")
                with col2: conferente_sel = st.text_input("Nome do Conferente", value=conf_padrao, placeholder="Ex: Edson")
                st.markdown('<br>', unsafe_allow_html=True)
                
                equipe_sel = st.multiselect("Equipe Alocada Agora", options=lista_auxiliares, format_func=lambda x: f"{x}  [{dict_skills_text[x]}]" if x in dict_skills_text else x)
                
                conflitos = {}
                for pessoa in equipe_sel:
                    if pessoa in mapa_pessoas:
                        if mapa_pessoas[pessoa] != str(doca_sel).strip(): conflitos[pessoa] = mapa_pessoas[pessoa]

                fadigados = checar_fadiga(equipe_sel, agenda_sel, df_log, df_aux)
                bloqueio_ergonomico = False
                
                if fadigados:
                    st.markdown(f"<div style='background-color: #FEF2F2; border: 1px solid #DC2626; border-radius: 8px; padding: 15px; margin-top: 15px; margin-bottom: 15px;'><b style='color: #DC2626;'>🚨 ALERTA DE SAÚDE E SEGURANÇA (SST)</b><br><span style='color: #7F1D1D; font-size: 13px;'>Os colaboradores <b>{', '.join(fadigados)}</b> já atuaram em carga pesada nas últimas 24h. Risco ergonômico!</span></div>", unsafe_allow_html=True)
                    ciente = st.checkbox("Declaro ciência do risco e autorizo a alocação na carga pesada.", key="chk_fadiga_aba3")
                    if not ciente: bloqueio_ergonomico = True

                st.markdown('<br>', unsafe_allow_html=True)
                if st.button("Gravar / Atualizar Doca", use_container_width=True):
                    if not doca_sel: st.warning("Preencha o número da Doca para continuar.")
                    elif not equipe_sel: st.warning("Selecione a equipe atual.")
                    elif bloqueio_ergonomico: st.error("Você precisa confirmar a ciência do risco ergonômico para gravar!")
                    else:
                        with st.spinner("Registrando movimentação..."):
                            sucesso = processar_gravacao_doca_v2(doca_sel, agenda_sel, conferente_sel, equipe_sel, conflitos, info_docas)
                            if sucesso:
                                st.success(f"✅ Doca {doca_sel} atualizada!")
                                st.balloons()
                                carregar_log_produtividade.clear()
                                st.rerun()
        except Exception as e: st.error(f"Erro no módulo de Docas: {e}")

# ==========================================================
# MÓDULO 3: FINANCEIRO E DRE
# ==========================================================
elif pagina_selecionada == "📊 Financeiro (Diretoria)":
    try:
        with st.spinner('Sincronizando com Base de Dados Financeira...'):
            df_raw = carregar_dados_financeiros()
            df, df_full = tratar_dados(df_raw)

        # --- Filtros ---
        st.sidebar.markdown('<div class="magalu-ribbon" style="left: 0; font-size: 12px;">Parâmetros de Data</div>', unsafe_allow_html=True)
        hoje = datetime.date.today()
        d_min = df['DATA AGENDADA'].min().date() if not df.empty else datetime.date(2025, 1, 1)
        d_max_limite = max(hoje, datetime.date(2026, 12, 31))
        
        selecao = st.sidebar.date_input("Selecione o Período", value=(d_min, hoje), min_value=d_min, max_value=d_max_limite)
        if isinstance(selecao, tuple) and len(selecao) == 2:
            data_ini, data_fim = selecao
        else:
            data_ini = selecao[0] if isinstance(selecao, (tuple, list)) else selecao
            data_fim = data_ini

        mask_data_main = (df['DATA AGENDADA'].dt.date >= data_ini) & (df['DATA AGENDADA'].dt.date <= data_fim)
        df_f = df[mask_data_main].copy()

        st.markdown('<div class="magalu-page-title">Visão Oficial de Faturamento</div>', unsafe_allow_html=True)
        st.markdown(f"<div class='magalu-page-subtitle'>Período: <b>{data_ini.strftime('%d/%m/%Y')}</b> até <b>{data_fim.strftime('%d/%m/%Y')}</b></div>", unsafe_allow_html=True)

        if not df_f.empty:
            rec = df_f[df_f['STATUS'] == 'OK']
            aus = df_f[df_f['STATUS'] == 'AUSENTE']
            
            # --- Indicadores Principais ---
            total_r = rec['VALOR_REAL'].sum()
            total_p = aus['VALOR_PERDIDO'].sum()
            tkt_carga = rec['VALOR_REAL'].mean() if not rec.empty else 0
            
            dias_unicos = rec['DATA AGENDADA'].dt.date.nunique()
            tkt_dia = total_r / dias_unicos if dias_unicos > 0 else 0
            
            meses_unicos = rec['MES_ORDENACAO'].nunique()
            tkt_mes = total_r / meses_unicos if meses_unicos > 0 else 0

            # --- 5 KPIs no Topo (Design Premium Corporativo) ---
            col1, col2, col3, col4, col5 = st.columns(5)
            
            # Função limpa para gerar os cards padronizados sem poluição visual
            def render_kpi(titulo, valor, subtitulo, cor_hex):
                return f"""
                <div style="background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px; padding: 20px 15px; box-shadow: 0 4px 15px rgba(0,0,0,0.02); position: relative; overflow: hidden; margin-bottom: 20px;">
                    <div style="position: absolute; top: 0; left: 0; width: 4px; height: 100%; background-color: {cor_hex};"></div>
                    <div style="padding-left: 10px;">
                        <div style="font-size: 11px; color: #64748B; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;">{titulo}</div>
                        <div style="font-size: 24px; font-weight: 800; color: #0F172A; letter-spacing: -0.5px; line-height: 1.2;">{valor}</div>
                        <div style="font-size: 11px; color: #94A3B8; font-weight: 500; margin-top: 4px;">{subtitulo}</div>
                    </div>
                </div>
                """

            with col1: 
                st.markdown(render_kpi("Arrecadação", formatar_moeda_br(total_r), "Faturamento Bruto", "#00C853"), unsafe_allow_html=True)
            with col2: 
                st.markdown(render_kpi("Perdas (Ausentes)", formatar_moeda_br(total_p), "Custo de Oportunidade", "#EF4444"), unsafe_allow_html=True)
            with col3: 
                st.markdown(render_kpi("Ticket / Carga", formatar_moeda_br(tkt_carga), "Média por Veículo", "#0086FF"), unsafe_allow_html=True)
            with col4: 
                st.markdown(render_kpi("Ticket / Dia", formatar_moeda_br(tkt_dia), "Média Diária", "#8B5CF6"), unsafe_allow_html=True)
            with col5: 
                st.markdown(render_kpi("Ticket / Mês", formatar_moeda_br(tkt_mes), "Projeção Mensal", "#6366F1"), unsafe_allow_html=True)

            # --- Tabelas Lado a Lado ---
            col_t1, col_t2 = st.columns(2)
            
            with col_t1:
                st.markdown("<h4 style='color: #334155;'>🏆 Top 10 Arrecadação por Fornecedor</h4>", unsafe_allow_html=True)
                if not rec.empty:
                    top_rec = rec.groupby('FORNECEDOR/SELLER').agg(
                        Cargas=('VALOR_REAL', 'count'),
                        Ticket_Medio=('VALOR_REAL', 'mean'),
                        Total_Arrecadado=('VALOR_REAL', 'sum')
                    ).reset_index().sort_values('Total_Arrecadado', ascending=False).head(10)
                    
                    top_rec['% da Oper.'] = (top_rec['Total_Arrecadado'] / total_r) * 100
                    
                    st.dataframe(
                        top_rec,
                        column_config={
                            "FORNECEDOR/SELLER": "Fornecedor",
                            "Cargas": st.column_config.NumberColumn("Cargas"),
                            "Ticket_Medio": st.column_config.NumberColumn("Ticket Médio", format="R$ %.2f"),
                            "Total_Arrecadado": st.column_config.NumberColumn("Total Arrecadado", format="R$ %.2f"),
                            "% da Oper.": st.column_config.ProgressColumn("% da Oper.", format="%.1f%%", min_value=0, max_value=100)
                        },
                        hide_index=True, use_container_width=True
                    )
            
            with col_t2:
                st.markdown("<h4 style='color: #334155;'>⚠️ Top 10 Perdas por ausência</h4>", unsafe_allow_html=True)
                if not aus.empty:
                    top_aus = aus.groupby('FORNECEDOR/SELLER').agg(
                        Faltas=('VALOR_PERDIDO', 'count'),
                        Ticket=('VALOR_PERDIDO', 'mean'),
                        Prejuizo=('VALOR_PERDIDO', 'sum')
                    ).reset_index().sort_values('Prejuizo', ascending=False).head(10)
                    
                    top_aus['% Perda'] = (top_aus['Prejuizo'] / total_p) * 100 if total_p > 0 else 0
                    
                    st.dataframe(
                        top_aus,
                        column_config={
                            "FORNECEDOR/SELLER": "Fornecedor",
                            "Faltas": st.column_config.NumberColumn("Faltas"),
                            "Ticket": st.column_config.NumberColumn("Ticket", format="R$ %.2f"),
                            "Prejuizo": st.column_config.NumberColumn("Prejuízo", format="R$ %.2f"),
                            "% Perda": st.column_config.ProgressColumn("% Perda", format="%.1f%%", min_value=0, max_value=100)
                        },
                        hide_index=True, use_container_width=True
                    )

            # --- Gráfico de Barras e Linha ---
            st.markdown('<div class="magalu-card" style="margin-top: 15px;">', unsafe_allow_html=True)
            st.markdown("<h4 style='color: #334155;'>📊 Evolução de Arrecadação x Perdas</h4>", unsafe_allow_html=True)
            
            ev_mes = df_f.groupby(['MES_ORDENACAO', 'MES_NOME']).agg(
                ARRECADADO=('VALOR_REAL', 'sum'), 
                PERDIDO=('VALOR_PERDIDO', 'sum')
            ).reset_index().sort_values('MES_ORDENACAO')
            
            if not ev_mes.empty:
                # Criando as etiquetas formatadas em R$ para exibir no gráfico
                text_arrecadado = [formatar_moeda_br(v) for v in ev_mes['ARRECADADO']]
                text_perdido = [formatar_moeda_br(v) for v in ev_mes['PERDIDO']]

                fig = make_subplots(specs=[[{"secondary_y": True}]])
                
                # Barras de Faturamento (com valor dentro/em cima da barra)
                fig.add_trace(go.Bar(
                    x=ev_mes['MES_NOME'], 
                    y=ev_mes['ARRECADADO'], 
                    name="Faturado (R$)", 
                    marker_color='#0086FF', 
                    text=text_arrecadado,
                    textposition='auto', # 'auto' coloca dentro se couber, ou fora se a barra for pequena
                    textfont=dict(size=11, color='#FFFFFF', weight='bold'),
                    hovertemplate="%{x}<br>Faturado: %{text}<extra></extra>"
                ), secondary_y=False)
                
                # Linha Vermelha de Perdas (com valor acima do ponto)
                fig.add_trace(go.Scatter(
                    x=ev_mes['MES_NOME'], 
                    y=ev_mes['PERDIDO'], 
                    name="Perda por Ausência (R$)", 
                    mode='lines+markers+text', # O '+text' obriga a mostrar o rótulo
                    line=dict(color='#FF3366', width=4), 
                    marker=dict(size=8), 
                    text=text_perdido,
                    textposition='top center', # Coloca o valor acima da linha vermelha
                    textfont=dict(size=11, color='#FF3366', weight='bold'),
                    hovertemplate="%{x}<br>Perdido: %{text}<extra></extra>"
                ), secondary_y=True)
                
                fig.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)', 
                    paper_bgcolor='rgba(0,0,0,0)',
                    margin=dict(t=30, b=10, l=0, r=0),
                    legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="center", x=0.5)
                )
                
                # Tira as linhas de grade para ficar um gráfico mais limpo
                fig.update_yaxes(showgrid=False, secondary_y=False)
                fig.update_yaxes(showgrid=False, secondary_y=True)
                
                st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
            st.markdown('</div>', unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Erro no módulo financeiro: {e}")
        
# ==========================================================
# MÓDULO EXTRA: REGISTRO DE ALINHAMENTO
# ==========================================================
elif pagina_selecionada == "📅 Registro de Alinhamento":
    st.markdown('<div class="magalu-page-title">Registro de Alinhamento</div>', unsafe_allow_html=True)
    st.markdown('<div class="magalu-page-subtitle">Programe folgas, DSRs, banco de horas e férias da equipe.</div>', unsafe_allow_html=True)

    try:
        # Puxa a mesma lista de funcionários que já usamos
        df_equipe = carregar_equipe()
        lista_funcionarios = df_equipe[df_equipe['NOME'].notna()]['NOME'].unique().tolist()
        lista_funcionarios = [str(nome).strip() for nome in lista_funcionarios if str(nome).strip() != '']
        lista_funcionarios.sort()
        
        # O nosso container com efeito de Vidro Fosco
        with st.container(border=True):
            st.markdown('<h4 style="color: #0086FF; margin-top: 0px; margin-bottom: 20px;">🗓️ Novo Alinhamento</h4>', unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            with col1:
                nome_sel = st.selectbox("Selecione o Colaborador", options=[""] + lista_funcionarios)
                motivo_sel = st.selectbox("Motivo do Alinhamento", options=["DSR", "BH", "FÉRIAS", "OUTROS"])
            
            with col2:
                data_sel = st.date_input("Data do Alinhamento (Folga)", format="DD/MM/YYYY")
                
                # A Mágica: O campo 'Outros' só aparece se a opção for selecionada
                motivo_outro = ""
                if motivo_sel == "OUTROS":
                    motivo_outro = st.text_input("Descreva o motivo (Obrigatório):", placeholder="Ex: Licença paternidade, Folga prêmio...")
                    
            st.markdown('<br>', unsafe_allow_html=True)
            
            # Botão verde Premium
            if st.button("💾 Gravar Alinhamento", use_container_width=True, type="primary"):
                if not nome_sel:
                    st.warning("Selecione o colaborador.")
                elif motivo_sel == "OUTROS" and not motivo_outro.strip():
                    st.warning("Como você selecionou 'OUTROS', por favor descreva o motivo na caixa de texto.")
                else:
                    motivo_final = motivo_outro.strip().upper() if motivo_sel == "OUTROS" else motivo_sel
                    data_folga_str = data_sel.strftime("%d/%m/%Y")
                    data_registro_str = (datetime.datetime.utcnow() - datetime.timedelta(hours=3)).strftime("%d/%m/%Y %H:%M:%S")
                    
                    # Ordem dos dados: Registro (Agora), Nome, Data da Folga, Motivo
                    linha_gravar = [data_registro_str, nome_sel, data_folga_str, motivo_final]
                    
                    with st.spinner("Registrando no sistema..."):
                        if gravar_alinhamento(linha_gravar):
                            st.success(f"✅ Sucesso! Alinhamento de {nome_sel} para o dia {data_folga_str} ({motivo_final}) registrado!")
                            st.balloons()
                            
    except Exception as e:
        st.error(f"Erro no módulo de Alinhamento: {e}")

# ==========================================================
# MÓDULO 4: PRODUTIVIDADE, NS E DESEMPENHO
# ==========================================================
elif pagina_selecionada == "📈 Produtividade (NS & Equipe)":
    st.markdown('<div class="magalu-page-title">Produtividade & Nível de Serviço</div>', unsafe_allow_html=True)
    st.markdown('<div class="magalu-page-subtitle">Acompanhamento de SLA, tempo de ciclo e performance individual.</div>', unsafe_allow_html=True)

    try:
        with st.spinner("Calculando métricas de performance..."):
            df_fin = carregar_docas_finalizadas()
            
            if df_fin.empty:
                st.warning("Ainda não há dados de cargas finalizadas para gerar os indicadores.")
            else:
                # Padronizar nomes de colunas para evitar erros de digitação na planilha
                colunas_upper = {c: str(c).upper().strip() for c in df_fin.columns}
                df_fin = df_fin.rename(columns=colunas_upper)
                
                # Encontrar as colunas dinamicamente
                col_data = next((c for c in df_fin.columns if 'DATA' in c), None)
                col_agenda = next((c for c in df_fin.columns if 'AGENDA' in c), None)
                col_tempo = next((c for c in df_fin.columns if 'TEMPO' in c), None)
                col_just = next((c for c in df_fin.columns if 'JUSTIFICATIVA' in c), None)
                col_cat = next((c for c in df_fin.columns if 'CATEGORIA' in c or 'LINHA' in c), None)
                col_aux = next((c for c in df_fin.columns if 'NOME' in c or 'PESSOA' in c), None)

                # Função para converter "02:30" em 150 minutos para fazer conta matemática
                def tempo_para_minutos(t_str):
                    try:
                        h, m = map(int, str(t_str).split(':'))
                        return h * 60 + m
                    except: return 0
                
                # Função para converter 150 minutos de volta para "02h30m" bonito
                def minutos_para_texto(mins):
                    if pd.isna(mins) or mins == 0: return "00h00m"
                    h = int(mins // 60)
                    m = int(mins % 60)
                    return f"{h:02d}h{m:02d}m"

                df_fin['MINUTOS'] = df_fin[col_tempo].apply(tempo_para_minutos)
                
                # --- Filtro de Data ---
                if col_data:
                    df_fin[col_data] = pd.to_datetime(df_fin[col_data], format="%d/%m/%Y", errors='coerce')
                    min_date, max_date = df_fin[col_data].min().date(), df_fin[col_data].max().date()
                    col_f1, col_f2 = st.columns([1, 3])
                    with col_f1:
                        datas_sel = st.date_input("Filtrar Período", [min_date, max_date])
                        if len(datas_sel) == 2:
                            df_fin = df_fin[(df_fin[col_data].dt.date >= datas_sel[0]) & (df_fin[col_data].dt.date <= datas_sel[1])]

                # Remove linhas duplicadas de Agenda (para não contar o tempo da mesma agenda 5x se tiver 5 ajudantes)
                df_agendas_unicas = df_fin.drop_duplicates(subset=[col_agenda])

                # --- CÁLCULOS DOS KPIS ---
                total_cargas = len(df_agendas_unicas)
                tempo_medio_geral = df_agendas_unicas['MINUTOS'].mean()
                
                # Cálculo de SLA (Nível de Serviço)
                qtd_no_prazo = len(df_agendas_unicas[df_agendas_unicas[col_just].astype(str).str.upper().str.contains("NO PRAZO", na=False)])
                sla_percent = (qtd_no_prazo / total_cargas * 100) if total_cargas > 0 else 0
                cor_sla = "#00C853" if sla_percent >= 90 else "#F59E0B" if sla_percent >= 75 else "#DC2626"

                # Criação das Abas Internas (Visão Macro x Visão Equipe)
                aba_macro, aba_equipe = st.tabs(["📊 Visão Macro & NS", "🧑‍🔧 Desempenho Indivídual"])

                with aba_macro:
                    # --- CARDS DE KPI ---
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        st.markdown(f'<div class="kpi-card" style="border-top: 4px solid #0086FF;"><div class="kpi-title">Cargas Finalizadas</div><div class="kpi-value" style="font-size:28px;">{total_cargas}</div></div>', unsafe_allow_html=True)
                    with c2:
                        st.markdown(f'<div class="kpi-card" style="border-top: 4px solid {cor_sla};"><div class="kpi-title">SLA (Nível de Serviço)</div><div class="kpi-value" style="font-size:28px; color:{cor_sla};">{sla_percent:.1f}%</div><div style="font-size:11px; color:#64748B;">{qtd_no_prazo} cargas no prazo</div></div>', unsafe_allow_html=True)
                    with c3:
                        st.markdown(f'<div class="kpi-card" style="border-top: 4px solid #8B5CF6;"><div class="kpi-title">Tempo Médio Geral</div><div class="kpi-value" style="font-size:28px;">{minutos_para_texto(tempo_medio_geral)}</div></div>', unsafe_allow_html=True)

                    # --- GRÁFICOS ---
                    col_g1, col_g2 = st.columns(2)
                    
                    with col_g1:
                        st.markdown('<div class="magalu-card">', unsafe_allow_html=True)
                        st.markdown("<h4 style='color: #334155; margin-bottom: 15px;'>⏱️ Tempo Médio por Categoria</h4>", unsafe_allow_html=True)
                        df_cat = df_agendas_unicas.groupby(col_cat)['MINUTOS'].mean().reset_index().sort_values('MINUTOS')
                        df_cat['TEXTO_TEMPO'] = df_cat['MINUTOS'].apply(minutos_para_texto)
                        
                        fig1 = px.bar(df_cat, x='MINUTOS', y=col_cat, orientation='h', text='TEXTO_TEMPO', color_discrete_sequence=['#0086FF'])
                        fig1.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', xaxis=dict(showgrid=False, visible=False), yaxis_title=None, margin=dict(l=0, r=0, t=0, b=0), height=350)
                        fig1.update_traces(textposition='outside', textfont=dict(weight='bold'))
                        st.plotly_chart(fig1, use_container_width=True, config={'displayModeBar': False})
                        st.markdown('</div>', unsafe_allow_html=True)

                    with col_g2:
                        st.markdown('<div class="magalu-card">', unsafe_allow_html=True)
                        st.markdown("<h4 style='color: #334155; margin-bottom: 15px;'>🚨 Motivos de Atraso (Ofensores)</h4>", unsafe_allow_html=True)
                        df_atrasos = df_agendas_unicas[~df_agendas_unicas[col_just].astype(str).str.upper().str.contains("NO PRAZO", na=False)]
                        if not df_atrasos.empty:
                            df_motivos = df_atrasos[col_just].value_counts().reset_index()
                            df_motivos.columns = ['Motivo', 'Qtd']
                            fig2 = px.pie(df_motivos, values='Qtd', names='Motivo', hole=0.6, color_discrete_sequence=px.colors.sequential.Reds_r)
                            fig2.update_layout(margin=dict(l=0, r=0, t=0, b=0), height=350, showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=-0.2, xanchor="center", x=0.5))
                            st.plotly_chart(fig2, use_container_width=True, config={'displayModeBar': False})
                        else:
                            st.info("Nenhum atraso registrado no período! Operação perfeita! 🎉")
                        st.markdown('</div>', unsafe_allow_html=True)

                with aba_equipe:
                    st.markdown("<h4 style='color: #0086FF; margin-bottom: 20px;'>🏆 Performance Individual por Categoria</h4>", unsafe_allow_html=True)
                    
                    categorias_lista = df_fin[col_cat].dropna().unique().tolist()
                    cat_sel = st.selectbox("Selecione a Categoria para comparar a equipe:", ["Todas as Categorias"] + categorias_lista)

                    df_operadores = df_fin.copy()
                    if cat_sel != "Todas as Categorias":
                        df_operadores = df_operadores[df_operadores[col_cat] == cat_sel]

                    if not df_operadores.empty:
                        # Agrupando por Auxiliar para ver Produtividade
                        df_rank = df_operadores.groupby(col_aux).agg(
                            Cargas_Participadas=(col_agenda, 'nunique'),
                            Tempo_Medio_Minutos=('MINUTOS', 'mean')
                        ).reset_index().sort_values('Tempo_Medio_Minutos') # Do mais rápido pro mais lento
                        
                        df_rank = df_rank[df_rank['Cargas_Participadas'] > 0]
                        df_rank['Tempo_Medio'] = df_rank['Tempo_Medio_Minutos'].apply(minutos_para_texto)
                        
                        media_geral_cat = df_rank['Tempo_Medio_Minutos'].mean()
                        
                        c_rank1, c_rank2 = st.columns([6, 4])
                        
                        with c_rank1:
                            st.markdown('<div class="magalu-card">', unsafe_allow_html=True)
                            st.markdown(f"<h5 style='color:#334155;'>Ranking de Velocidade - {cat_sel}</h5>", unsafe_allow_html=True)
                            
                            # Gráfico de barras comparando quem é o mais rápido
                            fig_rank = px.bar(df_rank, x='Tempo_Medio_Minutos', y=col_aux, orientation='h', text='Tempo_Medio', 
                                            color='Tempo_Medio_Minutos', color_continuous_scale='blues_r')
                            
                            # Linha de Média Geral para o cara saber se tá acima ou abaixo da média da empresa
                            fig_rank.add_vline(x=media_geral_cat, line_width=2, line_dash="dash", line_color="red", annotation_text="Média da Categoria")
                            
                            fig_rank.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', xaxis_title="Minutos (Menos é Melhor)", yaxis_title=None, coloraxis_showscale=False, height=450)
                            fig_rank.update_traces(textposition='inside', textfont=dict(color='white', weight='bold'))
                            st.plotly_chart(fig_rank, use_container_width=True, config={'displayModeBar': False})
                            st.markdown('</div>', unsafe_allow_html=True)
                            
                        with c_rank2:
                            st.markdown('<div class="magalu-card">', unsafe_allow_html=True)
                            st.markdown(f"<h5 style='color:#334155;'>Matriz de Participação</h5>", unsafe_allow_html=True)
                            st.dataframe(
                                df_rank[[col_aux, 'Cargas_Participadas', 'Tempo_Medio']],
                                column_config={
                                    col_aux: "Operador",
                                    "Cargas_Participadas": st.column_config.NumberColumn("Qtd Cargas", format="%d"),
                                    "Tempo_Medio": "Tempo Médio"
                                },
                                hide_index=True, use_container_width=True, height=450
                            )
                            st.markdown('</div>', unsafe_allow_html=True)
                    else:
                        st.info("Sem dados suficientes para esta categoria.")

    except Exception as e:
        st.error(f"Erro no módulo de Produtividade: {e}")
   
