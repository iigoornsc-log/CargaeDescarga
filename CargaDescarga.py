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
        if re.match(r'^C[D\d]', forn): return True 
        return False
        
    df_h = df_h[~df_h['FORNECEDOR/SELLER'].apply(eh_interno)]

    # ==========================================
    # 2º PASSO: SEPARAÇÃO DA MALHA FULL (OPORTUNIDADE EXTERNA)
    # ==========================================
    mask_full = df_h['LINHA'].str.contains('FULL', na=False) | df_h['CATEGORIA'].str.contains('FULL', na=False)
    
    df_full = df_h[mask_full].copy() # Base separada só com os FULL Externos
    df_main = df_h[~mask_full].copy() # Base oficial de faturamento

    # Só mantém pagantes na base principal
    pag_l = df_main[df_main['VALOR_REAL'] > 0]['LINHA'].unique()
    pag_c = df_main[df_main['VALOR_REAL'] > 0]['CATEGORIA'].unique()
    pag_l = [l for l in pag_l if l != '']
    pag_c = [c for c in pag_c if c != '']
    df_main = df_main[(df_main['LINHA'].isin(pag_l)) | (df_main['CATEGORIA'].isin(pag_c))].copy()

    # ==========================================
    # CÁLCULO DAS MÉDIAS E IMPUTAÇÃO
    # ==========================================
    df_ok = df_main[(df_main['STATUS'] == 'OK') & (df_main['VALOR_REAL'] > 0)]
    m_linha = df_ok.groupby('LINHA')['VALOR_REAL'].mean().to_dict()
    m_cat = df_ok.groupby('CATEGORIA')['VALOR_REAL'].mean().to_dict()

    # Imputação de Ausentes na Base Principal (Faturamento Real)
    df_main['VALOR_PERDIDO'] = 0.0
    mask_aus = df_main['STATUS'] == 'AUSENTE'
    df_main.loc[mask_aus, 'VALOR_PERDIDO'] = df_main.loc[mask_aus, 'LINHA'].map(m_linha)
    mask_zero = mask_aus & (df_main['VALOR_PERDIDO'].isna() | (df_main['VALOR_PERDIDO'] == 0))
    df_main.loc[mask_zero, 'VALOR_PERDIDO'] = df_main.loc[mask_zero, 'CATEGORIA'].map(m_cat)
    df_main['VALOR_PERDIDO'] = df_main['VALOR_PERDIDO'].fillna(0).round(2)

    # ==========================================
    # 🚀 CALIBRAGEM INTELIGENTE DO TICKET MALHA FULL
    # ==========================================
    # 1. Tenta puxar a média real da categoria (Vai funcionar perfeitamente para MADEIRA, etc)
    df_full['VALOR_ESTIMADO'] = df_full['CATEGORIA'].map(m_cat)
    
    # 2. Regra Comercial: Se a categoria for DIVERSOS (Vans menores), trava o ticket em R$ 350
    df_full.loc[df_full['CATEGORIA'] == 'DIVERSOS', 'VALOR_ESTIMADO'] = 350.00
    
    # 3. O que não achou categoria (vazio), joga a média combinada de R$ 500 (range 400-600)
    df_full['VALOR_ESTIMADO'] = df_full['VALOR_ESTIMADO'].fillna(500.00).round(2)

    return df_main, df_full

# --- 5. RENDERIZAÇÃO E FILTROS ---
try:
    with st.spinner('Sincronizando com Base de Dados e Calculando Oportunidades...'):
        df_raw = carregar_dados()
        df, df_full = tratar_dados(df_raw)

    # --- SIDEBAR ---
    st.sidebar.markdown('<h2 style="color: #0086FF;">🔍 Filtros</h2>', unsafe_allow_html=True)
    hoje = datetime.date.today()
    ontem = hoje - datetime.timedelta(days=1)
    
    d_min = df['DATA AGENDADA'].min().date() if not df.empty else datetime.date(2025, 1, 1)
    d_max_limite = max(hoje, datetime.date(2026, 12, 31))
    
    if st.sidebar.button("📅 Filtrar Apenas Ontem"):
        data_ini = data_fim = ontem
    else:
        selecao = st.sidebar.date_input("Selecione o Período", value=(d_min, hoje), min_value=d_min, max_value=d_max_limite)
        if isinstance(selecao, tuple) and len(selecao) == 2:
            data_ini, data_fim = selecao
        else:
            data_ini = selecao[0] if isinstance(selecao, (tuple, list)) else selecao
            data_fim = data_ini

    mask_data_main = (df['DATA AGENDADA'].dt.date >= data_ini) & (df['DATA AGENDADA'].dt.date <= data_fim)
    mask_data_full = (df_full['DATA AGENDADA'].dt.date >= data_ini) & (df_full['DATA AGENDADA'].dt.date <= data_fim)
    
    df_f = df[mask_data_main].copy()
    df_full_f = df_full[mask_data_full].copy()

    # --- CORPO DO DASHBOARD (FATURAMENTO OFICIAL) ---
    st.markdown('<div class="main-title">Logística Magazine Luiza 📦</div>', unsafe_allow_html=True)
    st.markdown(f"<div style='color: #6B7280; margin-bottom: 20px;'>Visão Oficial de Faturamento | Período: <b>{data_ini.strftime('%d/%m/%Y')}</b> até <b>{data_fim.strftime('%d/%m/%Y')}</b></div>", unsafe_allow_html=True)

    if not df_f.empty:
        rec = df_f[df_f['STATUS'] == 'OK']
        aus = df_f[df_f['STATUS'] == 'AUSENTE']
        
        total_r = rec['VALOR_REAL'].sum()
        total_p = aus['VALOR_PERDIDO'].sum()
        tkt_carga = rec['VALOR_REAL'].mean()
        m_dia = rec.groupby(rec['DATA AGENDADA'].dt.date)['VALOR_REAL'].sum().mean()
        m_mes = rec.groupby('MES_ORDENACAO')['VALOR_REAL'].sum().mean()

        col1, col2, col3, col4, col5 = st.columns(5)
        metrics = [
            ("💰 Arrecadação", total_r, "success"), ("📉 Perda Ausentes", total_p, "danger"),
            ("🚛 Ticket / Carga", tkt_carga, ""), ("📅 Ticket / Dia", m_dia, ""), ("📆 Ticket / Mês", m_mes, "")
        ]
        
        for i, (title, val, style) in enumerate(metrics):
            with [col1, col2, col3, col4, col5][i]:
                st.markdown(f'<div class="kpi-card {style}"><div class="kpi-title">{title}</div><div class="kpi-value">{formatar_moeda_br(val)}</div></div>', unsafe_allow_html=True)

        # Tabelas TOP 10 Oficiais
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        g1, g2 = st.columns(2)
        
        with g1:
            st.markdown('<h4 style="color: #333;">Fornecedores mais cobrados</h4>', unsafe_allow_html=True)
            top_f = rec.groupby('FORNECEDOR/SELLER').agg(QTD=('STATUS', 'count'), TICKET=('VALOR_REAL', 'mean'), TOTAL=('VALOR_REAL', 'sum')).reset_index()
            top_f = top_f[top_f['TOTAL'] > 0].sort_values('TOTAL', ascending=False).head(10)
            
            if not top_f.empty:
                total_geral_rec = max(rec['VALOR_REAL'].sum(), 1)
                top_f['% REC'] = (top_f['TOTAL'] / total_geral_rec) * 100
                top_f['TICKET_FMT'] = top_f['TICKET'].apply(formatar_moeda_br)
                top_f['TOTAL_FMT'] = top_f['TOTAL'].apply(formatar_moeda_br)
                st.dataframe(top_f[['FORNECEDOR/SELLER', 'QTD', 'TICKET_FMT', 'TOTAL_FMT', '% REC']], column_config={"FORNECEDOR/SELLER": "Fornecedor", "QTD": "Cargas", "TICKET_FMT": "Ticket Médio", "TOTAL_FMT": "Total Arrecadado", "% REC": st.column_config.ProgressColumn("% da Oper.", format="%.1f%%", min_value=0, max_value=100)}, hide_index=True, use_container_width=True)
        with g2:
            st.markdown('<h4 style="color: #333;">Maiores índices de Ausencia</h4>', unsafe_allow_html=True)
            top_p = aus.groupby('FORNECEDOR/SELLER').agg(QTD=('STATUS', 'count'), TICKET=('VALOR_PERDIDO', 'mean'), TOTAL=('VALOR_PERDIDO', 'sum')).reset_index()
            top_p = top_p[top_p['TOTAL'] > 0].sort_values('TOTAL', ascending=False).head(10)
            
            if not top_p.empty:
                total_geral_aus = max(aus['VALOR_PERDIDO'].sum(), 1)
                top_p['% PERDA'] = (top_p['TOTAL'] / total_geral_aus) * 100
                top_p['TICKET_FMT'] = top_p['TICKET'].apply(formatar_moeda_br)
                top_p['TOTAL_FMT'] = top_p['TOTAL'].apply(formatar_moeda_br)
                st.dataframe(top_p[['FORNECEDOR/SELLER', 'QTD', 'TICKET_FMT', 'TOTAL_FMT', '% PERDA']], column_config={"FORNECEDOR/SELLER": "Fornecedor", "QTD": "Faltas", "TICKET_FMT": "Ticket", "TOTAL_FMT": "Prejuízo", "% PERDA": st.column_config.ProgressColumn("% Perda", format="%.1f%%", min_value=0, max_value=100, color="#FF3366")}, hide_index=True, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # Gráfico e Tabela VIP
        layout_clean = dict(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font=dict(family="Inter, sans-serif", color='#4B5563'), margin=dict(t=40, b=10, l=0, r=0))
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        col_grafico, col_tabela = st.columns([7, 3])
        
        with col_grafico:
            st.markdown('<h4 style="color: #333;">📊 Evolução Mensal: Arrecadação vs % de Perda</h4>', unsafe_allow_html=True)
            ev_mes = df_f.groupby(['MES_ORDENACAO', 'MES_NOME']).agg(ARRECADADO=('VALOR_REAL', 'sum'), PERDIDO=('VALOR_PERDIDO', 'sum')).reset_index().sort_values('MES_ORDENACAO')
            if not ev_mes.empty:
                ev_mes['TOTAL_POTENCIAL'] = ev_mes['ARRECADADO'] + ev_mes['PERDIDO']
                ev_mes['PERDA_PERCENTUAL'] = (ev_mes['PERDIDO'] / ev_mes['TOTAL_POTENCIAL'].replace(0, 1)) * 100
                
                fig3 = make_subplots(specs=[[{"secondary_y": True}]])
                fig3.add_trace(go.Bar(x=ev_mes['MES_NOME'], y=ev_mes['ARRECADADO'], name="Arrecadado", marker_color='#0086FF', text=ev_mes['ARRECADADO'].apply(formatar_moeda_br), textposition='outside'), secondary_y=False)
                fig3.add_trace(go.Scatter(x=ev_mes['MES_NOME'], y=ev_mes['PERDA_PERCENTUAL'], name="Taxa de Perda (%)", mode='lines+markers+text', marker=dict(color='#FF3366', size=10), line=dict(color='#FF3366', width=4, shape='spline'), text=ev_mes['PERDA_PERCENTUAL'].apply(lambda x: f"{x:.1f}%"), textposition='top center'), secondary_y=True)
                
                fig3.update_layout(**layout_clean)
                fig3.update_layout(showlegend=False, margin=dict(t=30, b=10, l=0, r=0))
                fig3.update_yaxes(visible=False)
                
                st.plotly_chart(fig3, use_container_width=True, config={'displayModeBar': False})
                
        with col_tabela:
            st.markdown('<h4 style="color: #333;">💎 Cargas com Maior valor agregado (Top 10)</h4>', unsafe_allow_html=True)
            tkt_forn = rec.groupby('FORNECEDOR/SELLER').agg(QTD_CARGAS=('VALOR_REAL', 'count'), TICKET_MEDIO=('VALOR_REAL', 'mean')).reset_index().sort_values('TICKET_MEDIO', ascending=False).head(10)
            if not tkt_forn.empty:
                tkt_forn['Ticket'] = tkt_forn['TICKET_MEDIO'].apply(formatar_moeda_br)
                st.dataframe(tkt_forn[['FORNECEDOR/SELLER', 'QTD_CARGAS', 'Ticket']].rename(columns={'FORNECEDOR/SELLER':'Fornecedor', 'QTD_CARGAS':'Cargas'}), hide_index=True, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    # ==========================================
    # 🚀 SIMULAÇÃO DE MALHA FULL (C-LEVEL)
    # ==========================================
    st.markdown("---")
    st.markdown('<h2 style="color: #111827; font-weight: 800;"> Estudo de Oportunidade: Recebimento cargas FULL</h2>', unsafe_allow_html=True)
    st.markdown("<p style='color: #6B7280;'>Projeção financeira considerando ticket médio histórico para categorias de linha pesada (ex: Madeira) e ticket calibrado em R$ 350,00 para categoria Diversos (Vans).</p>", unsafe_allow_html=True)
    
    if df_full_f.empty:
        st.info("Nenhum veículo da malha FULL agendado neste período.")
    else:
        full_ok = df_full_f[df_full_f['STATUS'] == 'OK']
        
        qtd_full = len(full_ok)
        receita_potencial = full_ok['VALOR_ESTIMADO'].sum()
        ticket_medio_full = full_ok['VALOR_ESTIMADO'].mean() if qtd_full > 0 else 0

        st.markdown('<div class="chart-container gold-container">', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f'<div class="kpi-card gold"><div class="kpi-title">🚛 Veículos FULL (Recebidos)</div><div class="kpi-value">{qtd_full} cargas</div></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="kpi-card gold"><div class="kpi-title">💰 Receita estimada (Deixou de Cobrar)</div><div class="kpi-value">{formatar_moeda_br(receita_potencial)}</div></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="kpi-card gold"><div class="kpi-title">🎯 Ticket Médio </div><div class="kpi-value">{formatar_moeda_br(ticket_medio_full)}</div></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.markdown('<div class="chart-container">', unsafe_allow_html=True)
        st.markdown('<h4>Detalhamento de Fornecedores FULL</h4>', unsafe_allow_html=True)
        
        full_agrupado = full_ok.groupby('FORNECEDOR/SELLER').agg(
            QTD=('STATUS', 'count'),
            TICKET_MED_EST=('VALOR_ESTIMADO', 'mean'),
            RECEITA_ESTIMADA=('VALOR_ESTIMADO', 'sum')
        ).reset_index().sort_values('RECEITA_ESTIMADA', ascending=False)
        
        full_agrupado['Ticket Médio (R$)'] = full_agrupado['TICKET_MED_EST'].apply(formatar_moeda_br)
        full_agrupado['Receita Estimada (R$)'] = full_agrupado['RECEITA_ESTIMADA'].apply(formatar_moeda_br)
        
        st.dataframe(
            full_agrupado[['FORNECEDOR/SELLER', 'QTD', 'Ticket Médio (R$)', 'Receita Estimada (R$)']].rename(columns={'FORNECEDOR/SELLER':'Fornecedor FULL', 'QTD': 'Volume de Cargas'}),
            hide_index=True, use_container_width=True
        )
        st.markdown('</div>', unsafe_allow_html=True)

    # ==========================================
    # AUDITORIA
    # ==========================================
    with st.expander("🔍 Inspecionar Base de Dados Oficial"):
        st.dataframe(df_f[['DATA AGENDA', 'AGENDA WMS', 'FORNECEDOR/SELLER', 'LINHA', 'CATEGORIA', 'STATUS', 'VALOR_REAL', 'VALOR_PERDIDO']].sort_values('DATA AGENDA', ascending=False), use_container_width=True, hide_index=True)

except Exception as e:
    st.error(f"Erro Crítico: {e}")
