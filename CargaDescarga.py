import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import json
import datetime
from datetime import date

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gestão Operacional | Magalu", layout="wide")

# --- CSS SENIOR ---
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stButton>button { width: 100%; background-color: #0086FF; color: white; border-radius: 8px; height: 3em; font-weight: bold; }
    .kpi-box { background-color: white; padding: 20px; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-left: 5px solid #8A2BE2; }
    </style>
""", unsafe_allow_html=True)

# --- CONEXÃO GOOGLE SHEETS ---
def conectar_google():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    try:
        cred_dict = json.loads(st.secrets["google_json"])
        creds = Credentials.from_service_account_info(cred_dict, scopes=scopes)
    except:
        creds = Credentials.from_service_account_file(r'C:\Users\IIGOORNSC\Documents\CargaDescarga\credential_key.json', scopes=scopes)
    return gspread.authorize(creds)

@st.cache_data(ttl=60)
def carregar_equipe():
    client = conectar_google()
    sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
    ws = sh.worksheet("QUADRO CARGA e DESCARGA")
    df = pd.DataFrame(ws.get_all_records())
    return df

def gravar_absenteismo(dados_para_gravar):
    client = conectar_google()
    sh = client.open_by_key("1lrX3wQ41ncVMLzCaqGIQlbwvd_0n-AYOyU-NH1ge5oI")
    # Tenta conectar na aba de log, se não existir, avisa
    try:
        ws_log = sh.worksheet("LOG_ABSENTEISMO")
        ws_log.append_rows(dados_para_gravar)
        return True
    except:
        st.error("Erro: A aba 'LOG_ABSENTEISMO' não foi encontrada na planilha.")
        return False

# --- INTERFACE PRINCIPAL ---
st.title("📋 Lista de Chamada - Carga e Descarga")
st.markdown("Selecione os auxiliares que possuem ocorrências hoje.")

# 1. Carregar Dados
try:
    df_equipe = carregar_equipe()
    
    # Criar colunas para filtros no topo
    col_busca, col_data = st.columns([3, 1])
    
    with col_busca:
        busca = st.text_input("🔍 Pesquisar Auxiliar (Nome ou ID)", placeholder="Ex: Daniel...")
    
    with col_data:
        data_chamada = st.date_input("Data da Ocorrência", date.today())

    # 2. Filtrar a lista com base na busca
    if busca:
        df_filtrado = df_equipe[
            df_equipe['NOME'].str.contains(busca, case=False, na=False) | 
            df_equipe['ID'].astype(str).str.contains(busca, na=False)
        ].copy()
    else:
        df_filtrado = df_equipe.copy()

    # 3. Preparar o Data Editor (A "Mágica" do Front-end)
    # Adicionamos uma coluna de Ocorrência que não existe na planilha original
    df_filtrado['OCORRÊNCIA'] = "PRESENTE" # Valor padrão

    st.markdown("### Selecione o Motivo da Ocorrência")
    st.info("Apenas auxiliares com motivo diferente de 'PRESENTE' serão gravados no banco de dados.")

    # Opções do Dropdown
    opcoes_ocorrencia = ["PRESENTE", "FALTA", "DSR", "BH", "LICENÇA", "ATESTADO"]

    # Renderiza a tabela editável
    df_editado = st.data_editor(
        df_filtrado[['ID', 'NOME', 'CARGO', 'TURNO', 'OCORRÊNCIA']],
        column_config={
            "OCORRÊNCIA": st.column_config.SelectColumn(
                "STATUS / MOTIVO",
                help="Selecione o motivo da ausência",
                options=opcoes_ocorrencia,
                required=True,
            ),
            "ID": st.column_config.TextColumn("ID", disabled=True),
            "NOME": st.column_config.TextColumn("Nome Completo", disabled=True),
            "CARGO": st.column_config.TextColumn("Cargo", disabled=True),
            "TURNO": st.column_config.TextColumn("Turno", disabled=True),
        },
        hide_index=True,
        use_container_width=True,
        key="editor_chamada"
    )

    # 4. Lógica de Gravação (Back-end)
    st.markdown("---")
    if st.button("🚀 Gravar Ocorrências selecionadas"):
        # Filtra apenas quem NÃO está presente para gravar no log
        ocorrencias = df_editado[df_editado['OCORRÊNCIA'] != "PRESENTE"]
        
        if not ocorrencias.empty:
            # Prepara a lista de listas para o gspread (append_rows)
            lista_final = []
            data_str = data_chamada.strftime("%d/%m/%Y")
            
            for index, row in ocorrencias.iterrows():
                lista_final.append([
                    data_str, 
                    row['ID'], 
                    row['NOME'], 
                    row['OCORRÊNCIA'], 
                    row['TURNO']
                ])
            
            with st.spinner("Gravando no Banco de Dados..."):
                sucesso = gravar_absenteismo(lista_final)
                if sucesso:
                    st.success(f"✅ Sucesso! {len(lista_final)} ocorrências registradas para o dia {data_str}.")
                    # Limpa o cache para atualizar
                    st.cache_data.clear()
        else:
            st.warning("Nenhuma falta ou ocorrência foi selecionada (Todos estão como 'PRESENTE').")

except Exception as e:
    st.error(f"Erro ao carregar equipe: {e}")

# --- DASHBOARD DE APOIO (VISÃO SENIOR) ---
st.markdown("---")
st.subheader("💡 Visão do Analista")
c1, c2 = st.columns(2)
with c1:
    st.write("**Dica de Back-end:** Os dados são gravados na aba `LOG_ABSENTEISMO`. Isso evita que você perca o histórico se alguém mudar de turno no futuro.")
with c2:
    st.write("**Dica de Front-end:** Use a barra de pesquisa para filtrar rapidamente. Você pode editar múltiplos auxiliares antes de clicar em gravar.")
