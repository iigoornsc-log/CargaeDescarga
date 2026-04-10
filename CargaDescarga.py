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
# 1. CONFIGURAÇÃO DA PÁGINA E CSS
# ==========================================================
st.set_page_config(page_title="Magalu | Gestão Logística", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    .stApp { background-color: #F4F6F9; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #E2E8F0; }
    * { font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; }
    p, div, span { color: #334155; }
    .magalu-page-title { color: #0086FF; font-size: 22px; font-weight: 800; margin-bottom: 5px; line-height: 1.2;}
    .magalu-page-subtitle { color: #64748B; font-size: 13px; margin-bottom: 20px; }
    .magalu-ribbon {
        background-color: #0086FF; color: #FFFFFF; padding: 6px 16px; font-size: 14px; font-weight: 600;
        display: inline-block; border-radius: 0px 4px 4px 0px; margin-bottom: 10px; margin-top: 15px;
        position: relative; left: -1rem; box-shadow: 0 2px 4px rgba(0,134,255,0.2);
    }
    .magalu-card {
        background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; padding: 15px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.04); margin-bottom: 15px;
    }
    .stButton>button {
        background-color: #0086FF; color: white; border: none; border-radius: 8px;
        font-weight: 700; font-size: 16px; padding: 0.8rem 1rem; height: auto;
    }
    .stButton>button:hover { background-color: #0073E6; color: white; }
    button[kind="primary"] {
        background-color: #00C853 !important; border: none !important; color: white !important;
        min-height: 32px !important; font-size: 14px !important; border-radius: 6px !important;
    }
    button[kind="primary"]:hover { background-color: #00B248 !important; }
    input, .stSelectbox div[data-baseweb="select"] { border-radius: 6px !important; min-height: 45px !important;}
    .kpi-card { background-color: #FFFFFF; border-radius: 8px; padding: 12px; border-left: 4px solid #0086FF; margin-bottom: 10px;}
    .kpi-title { color: #6B7280; font-size: 12px; font-weight: 700; text-transform: uppercase; }
    .kpi-value { color: #111827; font-size: 20px; font-weight: 800; }
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

@st.dialog("📝 Justificativa de Atraso")
def exibir_popup_justificativa(dados_multiplos, linha_log_fecha):
    st.warning("Esta carga ultrapassou o tempo de meta. Por favor, informe o motivo do atraso para finalizar:")
    
    opcoes_atraso = [
        "Problema Sistêmico (WMS/TMS)",
        "Divergência de Nota Fiscal",
        "Aguardando Conferente/Líder",
        "Falta de Auxiliares na Doca",
        "Caminhão com avaria/difícil descarga",
        "Troca de turno/refeição",
        "Outro (Descrever abaixo)"
    ]
    
    motivo = st.selectbox("Selecione o motivo principal:", opcoes_atraso)
    detalhe = st.text_area("Detalhes adicionais (opcional):", placeholder="Ex: O caminhão chegou com as caixas tombadas...")
    
    if st.button("Confirmar Finalização", use_container_width=True):
        justificativa_final = f"{motivo} - {detalhe}".strip(" - ")
        
        # Adiciona a justificativa em todas as linhas (uma para cada auxiliar)
        for linha in dados_multiplos:
            linha.append(justificativa_final)
            
        with st.spinner("Gravando justificativa e finalizando..."):
            if gravar_conclusao_doca(dados_multiplos, linha_log_fecha):
                st.cache_data.clear()
                st.rerun()

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
st.sidebar.markdown('<div class="magalu-ribbon" style="left: 0;">Módulos do App</div>', unsafe_allow_html=True)
pagina_selecionada = st.sidebar.radio(
    "",
    ["📋 Absenteísmo (Doca)", "🚛 Gestão de Docas", "📊 Financeiro (Diretoria)"]
)
st.sidebar.markdown("---")

# ==========================================================
# MÓDULO 1: ABSENTEÍSMO
# ==========================================================
if pagina_selecionada == "📋 Absenteísmo (Doca)":
    st.markdown('<div class="magalu-page-title">Lançamento de Ocorrências</div>', unsafe_allow_html=True)
    st.markdown('<div class="magalu-page-subtitle">Pátio / Docas</div>', unsafe_allow_html=True)
    
    try:
        df_equipe = carregar_equipe()
        st.markdown('<div class="magalu-card">', unsafe_allow_html=True)
        data_chamada = st.date_input("Data", date.today())
        busca = st.text_input("🔍 Buscar Colaborador", placeholder="Matrícula ou Nome...")
        st.markdown('</div>', unsafe_allow_html=True)

        if busca:
            df_filtrado = df_equipe[df_equipe['NOME'].str.contains(busca, case=False, na=False) | df_equipe['ID'].astype(str).str.contains(busca, na=False)].copy()
        else:
            df_filtrado = df_equipe.copy()

        df_filtrado['OCORRÊNCIA'] = "PRESENTE" 
        opcoes_ocorrencia = ["PRESENTE", "FALTA", "DSR", "BH", "LICENÇA", "ATESTADO"]

        st.markdown('<div class="magalu-ribbon">Registro da Equipe</div>', unsafe_allow_html=True)
        
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
    st.markdown('<div class="magalu-page-subtitle">Acompanhe e movimente a equipe em tempo real.</div>', unsafe_allow_html=True)
    
    # Carregamento de dados global para as 3 abas
    df_log = carregar_log_produtividade()
    df_aux = carregar_aux()
    
    agendas_logadas = []
    col_doca, col_conf = None, None
    
    if not df_aux.empty:
        df_aux['AGENDA WMS'] = df_aux['AGENDA WMS'].astype(str).str.strip()
        colunas_limpas = [str(c).upper().strip() for c in df_aux.columns]
        col_doca = next((c for c, cu in zip(df_aux.columns, colunas_limpas) if 'DOCA' in cu), None)
        col_conf = next((c for c, cu in zip(df_aux.columns, colunas_limpas) if 'CONFERENTE' in cu or 'LIDER' in cu or 'LÍDER' in cu), None)

    # Criação das TRÊS abas
    aba1, aba2, aba3 = st.tabs(["👀 Visão das Docas (EM PROCESSO)", "⏳ Fila de Docas (PENDENTE)", "✍️ Montar Equipes"])

    # --- ABA 1: EM PROCESSO (Ativas) ---
    with aba1:
        if not df_log.empty:
            agendas_logadas = df_log['AGENDA'].astype(str).str.strip().unique().tolist()
            
            df_log['DATA_HORA_DT'] = pd.to_datetime(df_log['DATA_HORA'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
            ultimos_horarios = df_log.groupby('DOCA')['DATA_HORA_DT'].transform('max')
            df_ativos_bruto = df_log[df_log['DATA_HORA_DT'] == ultimos_horarios]
            df_ativos = df_ativos_bruto.groupby(['DOCA', 'AGENDA', 'CONFERENTE', 'DATA_HORA', 'DATA_HORA_DT'])['AUXILIARES'].apply(lambda x: ', '.join(x.dropna())).reset_index()
            
            df_ativos = df_ativos[df_ativos['AUXILIARES'] != 'ENCERRADO']
            df_ativos = df_ativos.sort_values('DATA_HORA_DT', ascending=False)

            if df_ativos.empty:
                st.info("Nenhuma doca ativa no momento. Pátio limpo! 🍃")
            else:
                agora_dt = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
                
                for index, row in df_ativos.iterrows():
                    agenda_str = str(row['AGENDA']).strip()
                    info = {'LINHA': '-', 'SKU': '-', 'PEÇAS': '-', 'VALOR': '-', 'PAGTO': '-', 'STATUS': '-'}
                    meta_minutos = 60 # Padrão
                    
                    if not df_aux.empty and agenda_str in df_aux['AGENDA WMS'].values:
                        aux_row = df_aux[df_aux['AGENDA WMS'] == agenda_str].iloc[0]
                        pagto_str = "✅ Sim" if str(aux_row.get('PAGAMENTO', '')).upper() == 'TRUE' else "⏳ Pendente"
                        valor_desc = aux_row.get('R$ DESCARGA', '-')
                        if str(valor_desc).replace('.','',1).isdigit():
                            valor_desc = f"R$ {float(valor_desc):,.2f}".replace(',','X').replace('.',',').replace('X','.')
                            
                        # Tratando possível erro de nome da coluna Linha/Categoria
                        linha_val = aux_row.get('LINHA', aux_row.get('CATEGORIA', '-'))
                        info = {'LINHA': linha_val, 'SKU': aux_row.get('SKU', '-'), 'PEÇAS': aux_row.get('PEÇAS', '-'), 'VALOR': valor_desc, 'PAGTO': pagto_str, 'STATUS': aux_row.get('STATUS', '-')}
                        
                        try:
                            col_meta = next((c for c in df_aux.columns if 'META' in str(c).upper()), None)
                            if col_meta and pd.notna(aux_row[col_meta]):
                                meta_minutos = int(float(str(aux_row[col_meta]).replace(',', '.')))
                        except: pass

                    # CÁLCULO DO CRONÔMETRO DE SLA (Tempo de Execução)
                    inicio_dt = row['DATA_HORA_DT']
                    if pd.isna(inicio_dt): inicio_dt = agora_dt
                    
                    decorrido_min = (agora_dt - inicio_dt).total_seconds() / 60
                    restante_min = meta_minutos - decorrido_min
                    
                    if restante_min >= 0:
                        h = int(restante_min // 60)
                        m = int(restante_min % 60)
                        cor_timer = "#00C853"
                        bg_timer = "#E6F9EC"
                        txt_timer = f"⏳ Resta {h:02d}h{m:02d}m"
                    else:
                        atraso = abs(restante_min)
                        h = int(atraso // 60)
                        m = int(atraso % 60)
                        cor_timer = "#DC2626"
                        bg_timer = "#FEF2F2"
                        txt_timer = f"🚨 Atraso -{h:02d}h{m:02d}m"

                    with st.container(border=True):
                        c_title, c_time = st.columns([5, 5])
                        c_title.markdown(f"<h4 style='margin:0; color:#0086FF;'>Doca {row['DOCA']}</h4>", unsafe_allow_html=True)
                        
                        c_time.markdown(f"""
                        <div style='text-align:right;'>
                            <div style='font-size:11px; color:#64748B; margin-bottom: 2px;'>⌚ Início: {row['DATA_HORA']}</div>
                            <div style='display:inline-block; font-size:12.5px; font-weight:800; color:{cor_timer}; background-color:{bg_timer}; padding:3px 6px; border-radius:4px; border: 1px solid {cor_timer};'>
                                {txt_timer} <span style="font-size:10px; font-weight:normal;">(Meta: {meta_minutos}m)</span>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        st.markdown(f"<div style='font-size: 13px; margin: 4px 0px 4px 0px;'><b>Agenda:</b> {row['AGENDA']} | <b>Líder:</b> {row['CONFERENTE']}</div>", unsafe_allow_html=True)
                        
                        st.markdown(f"""
                        <div style='font-size: 11.5px; color: #475569; background-color: #F8FAFC; padding: 6px; border-radius: 4px; margin-bottom: 8px; border: 1px solid #E2E8F0;'>
                            <b>Linha:</b> {info['LINHA']} &nbsp;|&nbsp; <b>SKU:</b> {info['SKU']} &nbsp;|&nbsp; <b>Peças:</b> {info['PEÇAS']}<br>
                            <b>Valor Carga:</b> {info['VALOR']} &nbsp;|&nbsp; <b>Pagto:</b> {info['PAGTO']} &nbsp;|&nbsp; <b>Status:</b> <span style="color:#0086FF; font-weight:bold;">{info['STATUS']}</span>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        c_eq, c_btn = st.columns([7, 3])
                        c_eq.markdown(f"<div style='font-size: 12px; color: #0086FF; background-color: #E6F2FF; padding: 6px; border-radius: 4px;'><b>Equipe:</b> {row['AUXILIARES']}</div>", unsafe_allow_html=True)
                        
                        with c_btn:
                            if st.button("✅ Finalizar", key=f"btn_fin_{row['DOCA']}_{index}", type="primary", use_container_width=True):
                                clique_dt = datetime.datetime.now()
                                duracao_final = clique_dt - row['DATA_HORA_DT']
                                total_minutos_final = int(duracao_final.total_seconds() / 60)
                                horas, mins = total_minutos_final // 60, total_minutos_final % 60
                                tempo_str = f"{horas:02d}:{mins:02d}"
                                
                                auxiliares_lista = [x.strip() for x in str(row['AUXILIARES']).split(',')]
                                linhas_conclusao_multiplas = []
                                
                                for pessoa in auxiliares_lista:
                                    linhas_conclusao_multiplas.append([
                                        clique_dt.strftime("%d/%m/%Y"), str(row['DOCA']), str(row['AGENDA']), 
                                        str(row['CONFERENTE']), len(auxiliares_lista), pessoa, 
                                        row['DATA_HORA'], clique_dt.strftime("%H:%M:%S"), tempo_str
                                    ])
                                
                                linha_log_fecha = [clique_dt.strftime("%d/%m/%Y %H:%M:%S"), str(row['DOCA']), row['AGENDA'], row['CONFERENTE'], "ENCERRADO"]
                                
                                if restante_min < 0:
                                    exibir_popup_justificativa(linhas_conclusao_multiplas, linha_log_fecha)
                                else:
                                    for linha in linhas_conclusao_multiplas: linha.append("No Prazo")
                                    with st.spinner("Finalizando..."):
                                        if gravar_conclusao_doca(linhas_conclusao_multiplas, linha_log_fecha):
                                            st.success(f"Doca finalizada com sucesso!")
                                            carregar_log_produtividade.clear()
                                            st.rerun()
        else:
            st.info("O Log de Operações ainda está vazio.")

    # --- ABA 2: PENDENTES (Fila de Espera) ---
    with aba2:
        if not df_aux.empty:
            df_pendentes = df_aux[~df_aux['AGENDA WMS'].isin(agendas_logadas)].copy()
            df_pendentes = df_pendentes[df_pendentes['AGENDA WMS'] != '']
            
            status_ignorados = ['AUSENTE', 'DEVOLVIDA', 'OK']
            if 'STATUS' in df_pendentes.columns:
                df_pendentes = df_pendentes[~df_pendentes['STATUS'].astype(str).str.upper().isin(status_ignorados)]
            
            if df_pendentes.empty:
                st.info("Nenhuma agenda aguardando equipe. Pátio zerado! 🎉")
            else:
                agora_dt = datetime.datetime.utcnow() - datetime.timedelta(hours=3)
                
                for index, row in df_pendentes.iterrows():
                    agenda_str = str(row['AGENDA WMS'])
                    
                    doca_str = str(row[col_doca]).strip() if col_doca and pd.notna(row[col_doca]) else "A Definir"
                    if doca_str.lower() in ['nan', 'none', '']: doca_str = "A Definir"
                    
                    conf_str = str(row[col_conf]).strip() if col_conf and pd.notna(row[col_conf]) else "A Definir"
                    if conf_str.lower() in ['nan', 'none', '']: conf_str = "A Definir"
                    
                    pagto_str = "✅ Sim" if str(row.get('PAGAMENTO', '')).upper() == 'TRUE' else "⏳ Pendente"
                    valor_desc = row.get('R$ DESCARGA', '-')
                    if str(valor_desc).replace('.','',1).isdigit():
                        valor_desc = f"R$ {float(valor_desc):,.2f}".replace(',','X').replace('.',',').replace('X','.')

                    linha_val = row.get('LINHA', row.get('CATEGORIA', '-'))
                    info = {
                        'LINHA': linha_val, 'SKU': row.get('SKU', '-'), 
                        'PEÇAS': row.get('PEÇAS', '-'), 'VALOR': valor_desc, 
                        'PAGTO': pagto_str, 'STATUS': row.get('STATUS', '-')
                    }
                    
                    # CÁLCULO DE TEMPO LIMITE PARA INICIAR (Drop-Dead Time)
                    meta_minutos = 60
                    limite_str = ""
                    hora_max_str = "-"
                    txt_timer_pend = "⏳ Aguardando..."
                    cor_timer_pend = "#F59E0B"
                    bg_timer_pend = "#FEF3C7"
                    
                    try:
                        col_meta = next((c for c in df_aux.columns if 'META' in str(c).upper()), None)
                        if col_meta and pd.notna(row[col_meta]):
                            meta_minutos = int(float(str(row[col_meta]).replace(',', '.')))
                            
                        col_limite = next((c for c in df_aux.columns if 'LIMITE' in str(c).upper()), None)
                        if col_limite and pd.notna(row[col_limite]) and str(row[col_limite]).strip() != '':
                            limite_str = str(row[col_limite]).strip()
                            
                            h_lim, m_lim = map(int, limite_str.split(':'))
                            limite_dt = agora_dt.replace(hour=h_lim, minute=m_lim, second=0, microsecond=0)
                            
                            hora_max_inicio = limite_dt - datetime.timedelta(minutes=meta_minutos)
                            hora_max_str = hora_max_inicio.strftime("%H:%M")
                            
                            diff_min = (hora_max_inicio - agora_dt).total_seconds() / 60
                            
                            if diff_min >= 0:
                                h = int(diff_min // 60)
                                m = int(diff_min % 60)
                                cor_timer_pend = "#00C853"
                                bg_timer_pend = "#E6F9EC"
                                txt_timer_pend = f"🟢 Sobra {h:02d}h{m:02d}m p/ Iniciar"
                            else:
                                atraso = abs(diff_min)
                                h = int(atraso // 60)
                                m = int(atraso % 60)
                                cor_timer_pend = "#DC2626"
                                bg_timer_pend = "#FEF2F2"
                                txt_timer_pend = f"🚨 ATRASADO HÁ {h:02d}h{m:02d}m"
                    except Exception as e:
                        pass

                    with st.container(border=True):
                        st.markdown(f"""
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <h4 style='margin:0; color:#475569;'>Doca {doca_str}</h4>
                            <div style='display:inline-block; font-size:12px; font-weight:800; color:{cor_timer_pend}; background-color:{bg_timer_pend}; padding:3px 6px; border-radius:4px; border: 1px solid {cor_timer_pend};'>
                                {txt_timer_pend}
                            </div>
                        </div>
                        <div style='font-size: 13px; margin: 8px 0px 4px 0px; display: flex; justify-content: space-between;'>
                            <span><b>Agenda:</b> {agenda_str} | <b>Líder:</b> {conf_str}</span>
                        </div>
                        
                        <div style='font-size: 12px; display: flex; gap: 15px; margin-bottom: 8px;'>
                            <span style='color:#64748B;'>Meta Operação: <b>{meta_minutos}m</b></span>
                            <span style='color:#0086FF;'>Iniciar até: <b>{hora_max_str}</b></span>
                            <span style='color:#DC2626;'>Fim Máximo: <b>{limite_str if limite_str else '-'}</b></span>
                        </div>
                        
                        <div style='font-size: 11.5px; color: #475569; background-color: #F8FAFC; padding: 6px; border-radius: 4px; margin-bottom: 8px; border: 1px solid #E2E8F0;'>
                            <b>Linha:</b> {info['LINHA']} &nbsp;|&nbsp; <b>SKU:</b> {info['SKU']} &nbsp;|&nbsp; <b>Peças:</b> {info['PEÇAS']}<br>
                            <b>Valor Carga:</b> {info['VALOR']} &nbsp;|&nbsp; <b>Pagto:</b> {info['PAGTO']} &nbsp;|&nbsp; <b>Status:</b> <span style="color:#F59E0B; font-weight:bold;">{info['STATUS']}</span>
                        </div>
                        
                        <div style='font-size: 12px; color: #DC2626; background-color: #FEF2F2; padding: 6px; border-radius: 4px; border: 1px solid #FECACA;'>
                            <b>Equipe:</b> <span style="font-weight:900;">PENDENTE ALOCAÇÃO</span>
                        </div>
                        """, unsafe_allow_html=True)
        else:
            st.info("A base auxiliar de Agendas não foi localizada ou está vazia.")

    # --- ABA 3: APONTAR / MOVIMENTAR ---
    with aba3:
        try:
            df_equipe = carregar_equipe()
            lista_auxiliares = df_equipe[df_equipe['NOME'].notna()]['NOME'].unique().tolist()
            lista_auxiliares = [nome for nome in lista_auxiliares if str(nome).strip() != '']
            
            mapa_pessoas = {}
            info_docas = {}
            
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

            st.markdown('<div class="magalu-card">', unsafe_allow_html=True)
            st.markdown('<b style="color: #0086FF;">📍 Nova Alocação / Atualizar Doca</b>', unsafe_allow_html=True)
            
            lista_agendas = []
            if not df_aux.empty:
                lista_agendas = df_aux[df_aux['AGENDA WMS'] != '']['AGENDA WMS'].unique().tolist()
            
            opcoes_agenda = [""] + lista_agendas + ["➕ DIGITAR OUTRA AGENDA..."]
            agenda_combo = st.selectbox("Nº da Agenda (Selecione ou digite para buscar)", options=opcoes_agenda, index=0)
            
            if agenda_combo == "➕ DIGITAR OUTRA AGENDA...":
                agenda_sel = st.text_input("Digite manualmente o Nº da Agenda", placeholder="Ex: 99999")
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
            equipe_sel = st.multiselect("Equipe Alocada Agora", options=lista_auxiliares)
            
            conflitos = {}
            for pessoa in equipe_sel:
                if pessoa in mapa_pessoas:
                    if mapa_pessoas[pessoa] != str(doca_sel).strip(): conflitos[pessoa] = mapa_pessoas[pessoa]
            
            st.markdown('</div>', unsafe_allow_html=True)

            if st.button("Gravar / Atualizar Doca", use_container_width=True):
                if not doca_sel: st.warning("Preencha o número da Doca para continuar.")
                elif not equipe_sel: st.warning("Selecione a equipe atual.")
                else:
                    if conflitos:
                        exibir_popup_transferencia(doca_sel, agenda_sel, conferente_sel, equipe_sel, conflitos, info_docas)
                    else:
                        with st.spinner("Registrando movimentação..."):
                            sucesso = processar_gravacao_doca(doca_sel, agenda_sel, conferente_sel, equipe_sel, conflitos, info_docas, False)
                            if sucesso:
                                st.success(f"✅ Doca {doca_sel} atualizada!")
                                st.balloons()
                                carregar_log_produtividade.clear()
                                st.rerun()

        except Exception as e:
            st.error(f"Erro no módulo de Docas: {e}")

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

            # --- 5 KPIs no Topo (Design Original) ---
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1: st.markdown(f'<div class="magalu-card" style="border-bottom: 4px solid #00C853; text-align:center; padding: 15px 5px;"><div style="font-size:11px; color:#64748B; font-weight:bold; text-transform:uppercase;">💰 Arrecadação</div><div style="font-size:22px; font-weight:800; color:#111827;">{formatar_moeda_br(total_r)}</div></div>', unsafe_allow_html=True)
            with col2: st.markdown(f'<div class="magalu-card" style="border-bottom: 4px solid #FF3366; text-align:center; padding: 15px 5px;"><div style="font-size:11px; color:#64748B; font-weight:bold; text-transform:uppercase;">📉 Perda Ausentes</div><div style="font-size:22px; font-weight:800; color:#111827;">{formatar_moeda_br(total_p)}</div></div>', unsafe_allow_html=True)
            with col3: st.markdown(f'<div class="magalu-card" style="border-bottom: 4px solid #0086FF; text-align:center; padding: 15px 5px;"><div style="font-size:11px; color:#64748B; font-weight:bold; text-transform:uppercase;">🚛 Ticket / Carga</div><div style="font-size:22px; font-weight:800; color:#111827;">{formatar_moeda_br(tkt_carga)}</div></div>', unsafe_allow_html=True)
            with col4: st.markdown(f'<div class="magalu-card" style="border-bottom: 4px solid #0086FF; text-align:center; padding: 15px 5px;"><div style="font-size:11px; color:#64748B; font-weight:bold; text-transform:uppercase;">📅 Ticket / Dia</div><div style="font-size:22px; font-weight:800; color:#111827;">{formatar_moeda_br(tkt_dia)}</div></div>', unsafe_allow_html=True)
            with col5: st.markdown(f'<div class="magalu-card" style="border-bottom: 4px solid #0086FF; text-align:center; padding: 15px 5px;"><div style="font-size:11px; color:#64748B; font-weight:bold; text-transform:uppercase;">📆 Ticket / Mês</div><div style="font-size:22px; font-weight:800; color:#111827;">{formatar_moeda_br(tkt_mes)}</div></div>', unsafe_allow_html=True)

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
