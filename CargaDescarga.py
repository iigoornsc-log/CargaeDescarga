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

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Magalu | Gestão de Carga e Descarga", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# --- 2. CSS DE ALTO NÍVEL (CLONE UI MAGALU) ---
st.markdown("""
    <style>
    /* Fundo da aplicação (Cinza Claro) */
    .stApp { background-color: #F4F6F9; }
    
    /* Barra lateral */
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #E2E8F0; }
    [data-testid="stSidebar"] hr { margin: 10px 0; border-color: #E2E8F0; }
    
    /* Tipografia e Textos */
    * { font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; }
    p, div, span { color: #334155; }
    
    /* Títulos da Página */
    .magalu-page-title { color: #0086FF; font-size: 24px; font-weight: 700; margin-bottom: 5px; display: flex; align-items: center; gap: 10px;}
    .magalu-page-subtitle { color: #64748B; font-size: 14px; margin-bottom: 25px; }
    
    /* Etiqueta Azul (Ribbon) idêntica à imagem */
    .magalu-ribbon {
        background-color: #0086FF;
        color: #FFFFFF;
        padding: 6px 16px;
        font-size: 14px;
        font-weight: 600;
        display: inline-block;
        border-radius: 0px 4px 4px 0px;
        margin-bottom: 15px;
        margin-top: 20px;
        position: relative;
        left: -1rem; /* Puxa para a borda esquerda nativa do Streamlit */
        box-shadow: 0 2px 4px rgba(0,134,255,0.2);
    }
    
    /* Containers Brancos (Cards) */
    .magalu-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 6px;
        padding: 20px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.02);
        margin-bottom: 15px;
    }
    
    /* Estilização de Botões (Azul Magalu) */
    .stButton>button {
        background-color: #0086FF;
        color: white;
        border: none;
        border-radius: 4px;
        font-weight: 600;
        padding: 0.5rem 1rem;
        transition: all 0.2s;
    }
    .stButton>button:hover { background-color: #0073E6; color: white; box-shadow: 0 4px 6px rgba(0, 134, 255, 0.3); }
    
    /* Tabelas (DataFrames) */
    .stDataFrame { border: 1px solid #E2E8F0; border-radius: 4px; overflow: hidden; }
    
    /* Inputs */
    input, .stSelectbox div[data-baseweb="select"] { border-radius: 4px !important; border-color: #CBD5E1 !important; }
    </style>
""", unsafe_allow_html=True)

# --- 3. CONEXÃO GOOGLE SHEETS ---
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

def gravar_absenteismo(dados_para_gravar):
    client = conectar_google()
    sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
    try:
        ws_log = sh.worksheet("LOG_ABSENTEISMO")
        ws_log.append_rows(dados_para_gravar)
        return True
    except:
        st.error("Erro: A aba 'LOG_ABSENTEISMO' não foi encontrada na planilha de equipe.")
        return False

# --- 4. TRATAMENTO DE DADOS FINANCEIROS ---
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


# ==========================================
# 🚀 MENU DE NAVEGAÇÃO LATERAL (ROTEADOR)
# ==========================================
st.sidebar.markdown('<div class="magalu-ribbon" style="left: 0;">Módulos do Sistema</div>', unsafe_allow_html=True)
pagina_selecionada = st.sidebar.radio(
    "",
    ["📋 Lista de Chamada (Absenteísmo)", "📊 DRE e Oportunidades (Financeiro)"]
)
st.sidebar.markdown("---")

# ==========================================
# PÁGINA 1: MÓDULO DE ABSENTEÍSMO
# ==========================================
if pagina_selecionada == "📋 Lista de Chamada (Absenteísmo)":
    st.markdown('<div class="magalu-page-title">📄 Lançamento de Ocorrências</div>', unsafe_allow_html=True)
    st.markdown('<div class="magalu-page-subtitle">Consulte o status da equipe e acompanhe o fluxo de absenteísmo diário.</div>', unsafe_allow_html=True)
    
    try:
        df_equipe = carregar_equipe()
        
        # Filtros no padrão superior da imagem
        st.markdown('<div class="magalu-ribbon">▼ Filtros da Chamada</div>', unsafe_allow_html=True)
        st.markdown('<div class="magalu-card">', unsafe_allow_html=True)
        col_data, col_busca, col_vazia = st.columns([2, 4, 2])
        with col_data:
            data_chamada = st.date_input("Data da Ocorrência", date.today())
        with col_busca:
            busca = st.text_input("Id do Auxiliar ou Nome", placeholder="Informe o ID ou Nome")
        st.markdown('</div>', unsafe_allow_html=True)

        if busca:
            df_filtrado = df_equipe[df_equipe['NOME'].str.contains(busca, case=False, na=False) | df_equipe['ID'].astype(str).str.contains(busca, na=False)].copy()
        else:
            df_filtrado = df_equipe.copy()

        df_filtrado['OCORRÊNCIA'] = "PRESENTE" 

        st.markdown('<div class="magalu-ribbon">Lista de Auxiliares Para Registro</div>', unsafe_allow_html=True)
        opcoes_ocorrencia = ["PRESENTE", "FALTA", "DSR", "BH", "LICENÇA", "ATESTADO"]

        # Tabela (Data Editor)
        df_editado = st.data_editor(
            df_filtrado[['ID', 'NOME', 'CARGO', 'TURNO', 'OCORRÊNCIA']],
            column_config={
                "OCORRÊNCIA": st.column_config.SelectColumn("Status da Integração (Motivo)", options=opcoes_ocorrencia, required=True),
                "ID": st.column_config.TextColumn("ID da Comanda (Matrícula)", disabled=True),
                "NOME": st.column_config.TextColumn("Colaborador", disabled=True),
                "CARGO": st.column_config.TextColumn("Função", disabled=True),
                "TURNO": st.column_config.TextColumn("Turno", disabled=True),
            },
            hide_index=True,
            use_container_width=True,
            key="editor_chamada"
        )
        
        st.markdown("<br>", unsafe_allow_html=True)
        col_btn1, col_btn2 = st.columns([8, 2])
        with col_btn2:
            if st.button("Gravar Ocorrências"):
                ocorrencias = df_editado[df_editado['OCORRÊNCIA'] != "PRESENTE"]
                if not ocorrencias.empty:
                    lista_final = []
                    data_str = data_chamada.strftime("%d/%m/%Y")
                    for index, row in ocorrencias.iterrows():
                        lista_final.append([data_str, row['ID'], row['NOME'], row['OCORRÊNCIA'], row['TURNO']])
                    with st.spinner("Gravando no Banco de Dados..."):
                        sucesso = gravar_absenteismo(lista_final)
                        if sucesso:
                            st.success(f"✅ Sucesso! {len(lista_final)} ocorrências registradas para o dia {data_str}.")
                            st.cache_data.clear()
                else:
                    st.warning("Nenhuma falta ou ocorrência foi selecionada.")

    except Exception as e:
        st.error(f"Erro ao carregar equipe: {e}")

# ==========================================
# PÁGINA 2: DASHBOARD FINANCEIRO E DRE
# ==========================================
elif pagina_selecionada == "📊 DRE e Oportunidades (Financeiro)":
    try:
        with st.spinner('Sincronizando com Base de Dados Financeira...'):
            df_raw = carregar_dados_financeiros()
            df, df_full = tratar_dados(df_raw)

        # Filtros Específicos do Financeiro na Sidebar
        st.sidebar.markdown('<div class="magalu-ribbon" style="left: 0; padding: 4px 12px; font-size: 12px;">Filtros de Data</div>', unsafe_allow_html=True)
        hoje = datetime.date.today()
        ontem = hoje - datetime.timedelta(days=1)
        d_min = df['DATA AGENDADA'].min().date() if not df.empty else datetime.date(2025, 1, 1)
        d_max_limite = max(hoje, datetime.date(2026, 12, 31))
        
        if st.sidebar.button("Filtrar Apenas Ontem"):
            data_ini = data_fim = ontem
        else:
            selecao = st.sidebar.date_input("Selecione o Período", value=(d_min, hoje), min_value=d_min, max_value=d_max_limite)
            if isinstance(selecao, tuple) and len(selecao) == 2:
                data_ini, data_fim = selecao
            else:
                data_ini = selecao[0] if isinstance(selecao, (tuple, list)) else selecao
                data_fim = data_ini

        st.sidebar.markdown('<div class="magalu-ribbon" style="left: 0; padding: 4px 12px; font-size: 12px;">Custos (Folha)</div>', unsafe_allow_html=True)
        qtd_pessoas = st.sidebar.number_input("Tamanho da Equipe", min_value=1, value=40, step=1)
        salario_base = st.sidebar.number_input("Custo/Pessoa (R$)", min_value=0.0, value=2100.0, step=100.0)
        custo_mensal_equipe = qtd_pessoas * salario_base

        mask_data_main = (df['DATA AGENDADA'].dt.date >= data_ini) & (df['DATA AGENDADA'].dt.date <= data_fim)
        mask_data_full = (df_full['DATA AGENDADA'].dt.date >= data_ini) & (df_full['DATA AGENDADA'].dt.date <= data_fim)
        
        df_f = df[mask_data_main].copy()
        df_full_f = df_full[mask_data_full].copy()

        st.markdown('<div class="magalu-page-title">📈 Demonstrativo de Resultados (DRE)</div>', unsafe_allow_html=True)
        st.markdown(f"<div class='magalu-page-subtitle'>Consulte o faturamento consolidado do período selecionado.</div>", unsafe_allow_html=True)

        if not df_f.empty:
            rec = df_f[df_f['STATUS'] == 'OK']
            aus = df_f[df_f['STATUS'] == 'AUSENTE']
            
            total_r = rec['VALOR_REAL'].sum()
            total_p = aus['VALOR_PERDIDO'].sum()
            tkt_carga = rec['VALOR_REAL'].mean()

            st.markdown('<div class="magalu-ribbon">Indicadores Oficiais</div>', unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)
            with col1: st.markdown(f'<div class="magalu-card" style="border-left: 4px solid #00C853;"><div>Faturamento Bruto</div><div style="font-size:24px; font-weight:bold;">{formatar_moeda_br(total_r)}</div></div>', unsafe_allow_html=True)
            with col2: st.markdown(f'<div class="magalu-card" style="border-left: 4px solid #FF3366;"><div>Perda Ausentes</div><div style="font-size:24px; font-weight:bold;">{formatar_moeda_br(total_p)}</div></div>', unsafe_allow_html=True)
            with col3: st.markdown(f'<div class="magalu-card" style="border-left: 4px solid #0086FF;"><div>Ticket Médio (Carga)</div><div style="font-size:24px; font-weight:bold;">{formatar_moeda_br(tkt_carga)}</div></div>', unsafe_allow_html=True)

            layout_clean = dict(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font=dict(family="Segoe UI, sans-serif", color='#334155'))

            # DRE E TABELA VIP
            st.markdown('<div class="magalu-ribbon">Análise Gráfica DRE e Cargas VIP</div>', unsafe_allow_html=True)
            st.markdown('<div class="magalu-card">', unsafe_allow_html=True)
            col_grafico, col_tabela = st.columns([7, 3])
            
            with col_grafico:
                ev_mes = df_f.groupby(['MES_ORDENACAO', 'MES_NOME']).agg(ARRECADADO=('VALOR_REAL', 'sum'), PERDIDO=('VALOR_PERDIDO', 'sum')).reset_index().sort_values('MES_ORDENACAO')
                
                if not ev_mes.empty:
                    ev_mes['TOTAL_POTENCIAL'] = ev_mes['ARRECADADO'] + ev_mes['PERDIDO']
                    ev_mes['PERDA_PERCENTUAL'] = (ev_mes['PERDIDO'] / ev_mes['TOTAL_POTENCIAL'].replace(0, 1)) * 100
                    ev_mes['FOLHA_CUSTO'] = custo_mensal_equipe
                    ev_mes['LUCRO_BRUTO'] = ev_mes['ARRECADADO'] - ev_mes['FOLHA_CUSTO']
                    
                    ev_mes['TXT_ARRECADADO'] = ev_mes['ARRECADADO'].apply(formatar_moeda_br)
                    ev_mes['TXT_CUSTO'] = ev_mes['FOLHA_CUSTO'].apply(formatar_moeda_br)
                    ev_mes['TXT_LUCRO'] = ev_mes['LUCRO_BRUTO'].apply(formatar_moeda_br)
                    ev_mes['TXT_PERDA'] = ev_mes['PERDIDO'].apply(formatar_moeda_br)
                    ev_mes['TXT_PERCENT'] = ev_mes['PERDA_PERCENTUAL'].apply(lambda x: f"{x:.1f}%")

                    fig3 = make_subplots(specs=[[{"secondary_y": True}]])
                    fig3.add_trace(go.Bar(x=ev_mes['MES_NOME'], y=ev_mes['ARRECADADO'], name="Faturamento", marker_color='#0086FF', text=ev_mes['TXT_ARRECADADO'], textposition='outside', hovertemplate="<b>Mês:</b> %{x}<br><b>Faturado:</b> %{text}<br><b>Lucro (Pós-Folha):</b> %{customdata}<extra></extra>", customdata=ev_mes['TXT_LUCRO']), secondary_y=False)
                    fig3.add_trace(go.Bar(x=ev_mes['MES_NOME'], y=ev_mes['FOLHA_CUSTO'], name="Custo Folha", marker_color='#94A3B8', text=ev_mes['TXT_CUSTO'], textposition='outside', hovertemplate="<b>Mês:</b> %{x}<br><b>Custo Equipe:</b> %{text}<extra></extra>"), secondary_y=False)
                    fig3.add_trace(go.Scatter(x=ev_mes['MES_NOME'], y=ev_mes['PERDA_PERCENTUAL'], name="Taxa de Perda (%)", mode='lines+markers+text', marker=dict(color='#FF3366', size=10), line=dict(color='#FF3366', width=4, shape='spline'), text=ev_mes['TXT_PERCENT'], textposition='top center', textfont=dict(color='#FF3366', size=14, weight='bold'), customdata=ev_mes['TXT_PERDA'], hovertemplate="<b>Mês:</b> %{x}<br><b>Taxa de Perda:</b> %{text}<br><b>Valor Perdido:</b> %{customdata}<extra></extra>"), secondary_y=True)
                    
                    fig3.update_layout(**layout_clean, barmode='group', showlegend=True, legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="center", x=0.5), margin=dict(t=50, b=10, l=0, r=0))
                    
                    max_y = max(ev_mes['ARRECADADO'].max(), ev_mes['FOLHA_CUSTO'].max())
                    fig3.update_yaxes(visible=False, secondary_y=False, range=[0, max_y * 1.3]) 
                    fig3.update_yaxes(visible=False, secondary_y=True, range=[0, max(ev_mes['PERDA_PERCENTUAL'].max() * 1.5, 10)])
                    
                    st.plotly_chart(fig3, use_container_width=True, config={'displayModeBar': False})
                    
            with col_tabela:
                tkt_forn = rec.groupby('FORNECEDOR/SELLER').agg(QTD_CARGAS=('VALOR_REAL', 'count'), TICKET_MEDIO=('VALOR_REAL', 'mean')).reset_index().sort_values('TICKET_MEDIO', ascending=False).head(10)
                if not tkt_forn.empty:
                    tkt_forn['Ticket'] = tkt_forn['TICKET_MEDIO'].apply(formatar_moeda_br)
                    st.dataframe(tkt_forn[['FORNECEDOR/SELLER', 'QTD_CARGAS', 'Ticket']].rename(columns={'FORNECEDOR/SELLER':'Fornecedor', 'QTD_CARGAS':'Cargas'}), hide_index=True, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # OPORTUNIDADE FULL
        st.markdown('<div class="magalu-ribbon" style="background-color: #334155;">Estudo de Oportunidade: Malha FULL</div>', unsafe_allow_html=True)
        st.markdown('<div class="magalu-card">', unsafe_allow_html=True)
        if df_full_f.empty:
            st.info("Nenhum veículo da malha FULL agendado neste período.")
        else:
            full_ok = df_full_f[df_full_f['STATUS'] == 'OK']
            qtd_full = len(full_ok)
            receita_potencial = full_ok['VALOR_ESTIMADO'].sum()

            c1, c2 = st.columns(2)
            with c1: st.markdown(f'<div style="font-size:16px;">Veículos Recebidos: <b>{qtd_full} cargas</b></div>', unsafe_allow_html=True)
            with c2: st.markdown(f'<div style="font-size:16px;">Receita Potencial (Oportunidade): <b>{formatar_moeda_br(receita_potencial)}</b></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Erro no módulo Financeiro: {e}")
