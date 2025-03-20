import streamlit as st
import pandas as pd
from io import BytesIO
import unidecode
import re

# Funções utilitárias
def extrair_uf(local):
    match = re.search(r'\b[A-Z]{2}\b', local)
    return match.group(0) if match else "Outros"

def recarregar_dados():
    if st.session_state.get("dados_empresas") is not None:
        del st.session_state["dados_empresas"]

def filtrar_dados_empresas():
    if "dados_empresas" not in st.session_state:
        st.session_state["dados_empresas"] = empresas.copy()
    dados_filtrados = st.session_state["dados_empresas"]
    if st.session_state["filtro_tipo"] == "Única Data":
        data_unica = st.session_state["data_unica"]
        if data_unica:
            dados_filtrados = dados_filtrados[dados_filtrados['data do mapeamento'] == pd.Timestamp(data_unica)]
    elif st.session_state["filtro_tipo"] == "Intervalo de Datas":
        data_inicial, data_final = st.session_state["data_intervalo"]
        if data_inicial and data_final:
            dados_filtrados = dados_filtrados[
                (dados_filtrados['data do mapeamento'] >= pd.Timestamp(data_inicial)) &
                (dados_filtrados['data do mapeamento'] <= pd.Timestamp(data_final))
            ]
    st.session_state["dados_filtrados"] = dados_filtrados

def categorizar_completo(local):
    if ("CEARÁ" in local.upper() or "FORTALEZA" in local.upper() or "- CE" in local.upper() or 
        "CE " in local.upper() or local.upper().strip() == "CE" or 
        "AQUIRAZ" in local.upper() or "CAUCAIA" in local.upper()):
        return "CE"
    else:
        return extrair_uf(local)

# Configuração da página
st.set_page_config(page_title="APP de contagem de status das empresas", layout="wide")
st.title("Ferramenta para contagem de status das empresas")

# Exemplos de colunas necessárias para cada arquivo
colunas_mapeamento = pd.DataFrame({
    'Job': [], 'Elementar': [], 'CNPJ': [], 'Status do Item2': [],
})
colunas_elementar = pd.DataFrame({
    'Elementar': [], 'Descricao do Item': [], 'Unidade': [], 'Simples/Composto': [], 'Status de Pesquisa': []
})
def dataframe_to_markdown(df):
    return df.to_markdown(index=False)

@st.cache_data
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='ContagemStatus')
    return output.getvalue()

# Funções para carregar os arquivos (aba específica)
@st.cache_data
def carregar_mapeamento(arquivo):
    excel_data = pd.ExcelFile(arquivo)
    if "Mapeamento_Ajustada" not in excel_data.sheet_names:
        st.error("O arquivo de mapeamento deve conter a aba 'Mapeamento_Ajustada'.")
        st.stop()
    df = pd.read_excel(arquivo, sheet_name="Mapeamento_Ajustada").astype(str)
    df.columns = df.columns.str.strip().str.lower()
    df['elementar'] = df['elementar'].str.split('.').str[0]
    return df

@st.cache_data
def carregar_elementares(arquivo):
    excel_data = pd.ExcelFile(arquivo)
    if "Estudo do Varejo" not in excel_data.sheet_names:
        st.error("O arquivo de elementares deve conter a aba 'Estudo do Varejo'.")
        st.stop()
    elementares = pd.read_excel(arquivo, sheet_name="Estudo do Varejo").astype(str)
    elementares.columns = elementares.columns.str.replace('\n', '').str.strip().str.lower()
    return elementares

with st.expander("Dados necessários"):

    st.markdown("Para carregar os dados relacionados ao **mapeamento** é necessário que a planilha contenha as seguintes colunas:")
    st.markdown(dataframe_to_markdown(colunas_mapeamento))
    st.markdown("Além disso, os dados que serão carregados relacionados ao **mapeamento** devem estar na aba **Mapeamento_Ajustada**.")
    st.markdown("Para carregar os dados relacionados ao **elementar** é necessário que a planilha contenha as seguintes colunas:")
    st.markdown(dataframe_to_markdown(colunas_elementar))
    st.markdown("Além disso, os dados que serão carregados relacionados ao **elementar** devem estar na aba **Estudo do Varejo**.")
    
    st.markdown("**Importante:** as colunas devem conter exatamente os nomes apresentados acima.")

# Upload dos arquivos
st.subheader('Selecione o arquivo de mapeamento:', divider="red")
mapeamento_arquivo = st.file_uploader("arquivo de mapeamento", type="xlsx", label_visibility="hidden")
if mapeamento_arquivo is not None:
    mapeamento = carregar_mapeamento(mapeamento_arquivo)
    st.subheader('Selecione o arquivo de elementares:', divider="red")
    elementares_arquivo = st.file_uploader("arquivo de elementares", type="xlsx", label_visibility="hidden")
    
    @st.cache_data
    def processar_ufs(df):
        uf_pricing = pd.DataFrame({'Local': df['uf do preço'].unique()})
        uf_pricing['UF'] = uf_pricing['Local'].apply(categorizar_completo)
        return uf_pricing.explode("Local").reset_index(drop=True)

    df_locais_por_uf = processar_ufs(mapeamento)
    df_locais_por_uf['UF'] = df_locais_por_uf['UF'].str.upper()
    
    if elementares_arquivo is not None:
        # Carrega e pré-processa os dados de elementares
        elementares = carregar_elementares(elementares_arquivo)
        if 'status de pesquisa' in elementares.columns:
            elementares['status de pesquisa'] = elementares['status de pesquisa'].apply(lambda x: unidecode.unidecode(x.lower()))
        else:
            st.error("A coluna 'status de pesquisa' não foi encontrada no arquivo carregado. Verifique o arquivo e tente novamente.")
            st.stop()
        elementares = elementares[elementares['status de pesquisa'] == 'item pesquisavel']
        if 'simples/composto' not in elementares.columns:
            elementares['simples/composto'] = None
        
        # Normaliza nomes e converte coluna 'elementar'
        mapeamento.columns = mapeamento.columns.str.strip().str.lower()
        elementares.columns = elementares.columns.str.strip().str.lower()
        mapeamento['elementar'] = pd.to_numeric(mapeamento['elementar'], errors='coerce').fillna(0).astype(int)
        elementares['elementar'] = pd.to_numeric(elementares['elementar'], errors='coerce').fillna(0).astype(int)
        
        # Seleciona as colunas necessárias e junta os dataframes
        mapeamento = mapeamento[['elementar', 'cnpj', 'status do item2', 'job', 'data do mapeamento']].copy()
        elementares = elementares[['elementar', 'descricao do item', 'unidade', 'simples/composto']].copy()
        planilha = pd.merge(elementares, mapeamento, on='elementar', how='left')
        
        # Define o status com base em 'status do item2'
        planilha['status'] = planilha['status do item2'].apply(
            lambda x: 'Sem status' if pd.isna(x) else ('PO' if x == 'PO' else ('AG' if x == 'AG' else ('RE' if x == 'RE' else 'NG')))
        )
        planilha['job'] = planilha['job'].fillna('Sem Job')
        
        # Cria dataframe de prioridades e ordena
        prioridade = pd.DataFrame({
            'status': ['PO', 'AG', 'RE', 'NG', 'Sem status'],
            'prioridade': [1, 2, 3, 4, 5]
        })
        planilha = pd.merge(planilha, prioridade, on='status', how='left')
        planilha = planilha.sort_values(by=['job', 'elementar', 'cnpj', 'prioridade'])
        planilha['cnpj'] = planilha['cnpj'].fillna('CNPJ_DESCONHECIDO')
        planilha['data do mapeamento'] = planilha['data do mapeamento'].fillna('Sem data')
        
        # Mantém apenas a primeira linha por grupo
        planilha = planilha.groupby(['job', 'elementar', 'cnpj']).head(1)
        
        # Agrupa e conta a quantidade mapeada por status
        qtd_mapeada_por_status = (
            planilha.groupby(['job', 'elementar', 'data do mapeamento', 'status']).size().reset_index(name='count')
        )
        qtd_mapeada_por_status = qtd_mapeada_por_status.pivot_table(
            index=['job', 'elementar', 'data do mapeamento'],
            columns='status', values='count', fill_value=0
        ).reset_index()
        if 'Sem status' in qtd_mapeada_por_status.columns:
            qtd_mapeada_por_status.drop(columns=['Sem status'], inplace=True)
        
        # Junta os dados e garante as colunas de status
        empresas = pd.merge(qtd_mapeada_por_status, elementares, on='elementar', how='left')
        status_colunas = ['PO', 'NG', 'RE', 'AG']
        for status in status_colunas:
            if status not in empresas.columns:
                empresas[status] = 0
        empresas['Empresas Mapeadas'] = empresas['PO'] + empresas['AG'] + empresas['RE'] + empresas['NG']
        
        # Reordena as colunas fixas (removendo 'job') e preserva as demais dinâmicas
        colunas_reordenadas = [
            'data do mapeamento', 'elementar', 'descricao do item', 'unidade', 'simples/composto', 'Empresas Mapeadas', 'PO', 'RE', 'NG', 'AG'
        ]
        colunas_existentes = empresas.columns.tolist()
        colunas_completas = colunas_reordenadas + [col for col in colunas_existentes if col not in colunas_reordenadas]
        empresas = empresas[colunas_completas]
        
        # Gera colunas específicas para cada job
        jobs_unicos = empresas['job'].unique()
        for job in jobs_unicos:
            for status in status_colunas:
                nova_coluna = f"{job.upper()}-{status}"
                empresas[nova_coluna] = empresas.apply(
                    lambda row: row[status] if job.upper() in row['job'].upper() else 0, axis=1
                )
            col_empresas = f"{job.upper()}-TOTAL"
            empresas[col_empresas] = empresas.apply(
                lambda row: row.get(f"{job.upper()}-PO", 0) +
                            row.get(f"{job.upper()}-AG", 0) +
                            row.get(f"{job.upper()}-RE", 0) +
                            row.get(f"{job.upper()}-NG", 0),
                axis=1
            )
        
        # Filtro por data
        empresas['data do mapeamento'] = pd.to_datetime(empresas['data do mapeamento'], errors='coerce')
        st.subheader("Filtrar por Data do Mapeamento", divider="red")
        if 'filtro_tipo' not in st.session_state:
            st.session_state['filtro_tipo'] = "Nenhum"
        if 'data_unica' not in st.session_state:
            st.session_state['data_unica'] = None
        if 'data_intervalo' not in st.session_state:
            st.session_state['data_intervalo'] = (None, None)
        
        @st.cache_data
        def filtrar_por_intervalo(df, data_inicial=None, data_final=None):
            if data_inicial and data_final:
                return df[
                    (df['data do mapeamento'] >= pd.Timestamp(data_inicial)) &
                    (df['data do mapeamento'] <= pd.Timestamp(data_final))
                ]
            return df
        
        filtro_tipo = st.radio("Tipo de Filtro:", ("Nenhum", "Única Data", "Intervalo de Datas"), horizontal=True)
        empresas['data do mapeamento'] = pd.to_datetime(empresas['data do mapeamento'], errors='coerce')
        data_default = pd.Timestamp("2000-01-01")
        empresas['data do mapeamento'] = empresas['data do mapeamento'].fillna(data_default)
        if filtro_tipo == "Única Data":
            data_selecionada = st.date_input("Escolha uma data", value=None, key="data_unica")
            if data_selecionada:
                data_selecionada = pd.Timestamp(data_selecionada).date()
                empresas['data do mapeamento'] = empresas['data do mapeamento'].fillna(data_selecionada)
                empresas_filtradas = empresas[empresas['data do mapeamento'] == pd.Timestamp(data_selecionada)]
            else:
                empresas_filtradas = empresas.copy()
        elif filtro_tipo == "Intervalo de Datas":
            data_inicial, data_final = st.date_input(
                "Escolha o intervalo de datas",
                value=(empresas['data do mapeamento'].min().date(), empresas['data do mapeamento'].max().date()),
                key="data_intervalo"
            )
            if data_inicial and data_final:
                empresas['data do mapeamento'] = empresas['data do mapeamento'].fillna(pd.Timestamp(data_inicial))
                empresas_filtradas = filtrar_por_intervalo(empresas, data_inicial, data_final)
            else:
                empresas_filtradas = empresas.copy()
        else:
            empresas['data do mapeamento'] = empresas['data do mapeamento'].fillna(data_default)
            empresas_filtradas = empresas.copy()
        
        # Formatação da data para exibição
        empresas_filtradas['data do mapeamento'] = pd.to_datetime(empresas_filtradas['data do mapeamento'], errors='coerce')
        if 'data do mapeamento' in empresas_filtradas.columns:
            empresas_filtradas['data do mapeamento'] = empresas_filtradas['data do mapeamento'].dt.strftime('%d/%m/%Y')
        
        # Agrupamento final: remova a coluna 'job' e capitalize os nomes fixos
        empresas_filtradas_sem_data = empresas_filtradas.copy()
        if empresas_filtradas_sem_data.empty:
            st.warning("Nenhum dado encontrado após o filtro.")
        else:
            # Agrupa por 'elementar' e soma as colunas dinâmicas
            colunas_dinamicas = [
                col for col in empresas_filtradas_sem_data.columns 
                if '-' in col and (col.split('-')[-1] in status_colunas or col.split('-')[-1] == 'TOTAL')
            ]
            empresas_filtradas_sem_data = empresas_filtradas_sem_data.groupby('elementar').agg({
                'job': lambda x: ', '.join(sorted(set(x))),
                'data do mapeamento': 'max',
                'descricao do item': 'first',
                'unidade': 'first',
                'simples/composto': 'first',
                'Empresas Mapeadas': 'sum',
                **{col: 'sum' for col in colunas_dinamicas}
            }).reset_index()
            # Remove a coluna 'job'
            if 'job' in empresas_filtradas_sem_data.columns:
                empresas_filtradas_sem_data.drop(columns=['job'], inplace=True)
            empresas_filtradas_sem_data = empresas_filtradas_sem_data.sort_values(by='elementar')
            
            # Renomeia as colunas fixas para ter a primeira letra maiúscula
            fixed_cols = ['data do mapeamento', 'elementar', 'descricao do item', 'unidade', 'simples/composto']
            rename_mapping = {col: col.capitalize() for col in fixed_cols if col in empresas_filtradas_sem_data.columns}
            empresas_filtradas_sem_data.rename(columns=rename_mapping, inplace=True)
            
            st.write(empresas_filtradas_sem_data)
        
        excel_data = to_excel(empresas_filtradas_sem_data)
        st.subheader('Clique para baixar o resultado:', divider="red")
        st.download_button(label="Baixar", data=excel_data, file_name="Empresas.xlsx")
    else:
        st.info("Por favor, carregue o arquivo de elementares para continuar.")
else:
    st.info("Por favor, carregue o arquivo de mapeamento para continuar.")
