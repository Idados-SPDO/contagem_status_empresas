import streamlit as st
import pandas as pd
from io import BytesIO
import unidecode
import re


def extrair_uf(local):
    match = re.search(r'\b[A-Z]{2}\b', local)
    return match.group(0) if match else "Outros"

def recarregar_dados():
    """Função para recalcular os dados ao alterar os filtros."""
    if st.session_state.get("dados_empresas") is not None:
        del st.session_state["dados_empresas"]

def filtrar_dados_empresas():
    """Função para aplicar filtros aos dados."""
    if "dados_empresas" not in st.session_state:
        # Carregue os dados iniciais
        st.session_state["dados_empresas"] = empresas.copy()

    dados_filtrados = st.session_state["dados_empresas"]

    # Filtro por data
    if st.session_state["filtro_tipo"] == "Única Data":
        data_unica = st.session_state["data_unica"]
        if data_unica:
            dados_filtrados = dados_filtrados[dados_filtrados['Data Do Mapeamento'] == pd.Timestamp(data_unica)]

    elif st.session_state["filtro_tipo"] == "Intervalo de Datas":
        data_inicial, data_final = st.session_state["data_intervalo"]
        if data_inicial and data_final:
            dados_filtrados = dados_filtrados[
                (dados_filtrados['Data Do Mapeamento'] >= pd.Timestamp(data_inicial)) &
                (dados_filtrados['Data Do Mapeamento'] <= pd.Timestamp(data_final))
            ]

    st.session_state["dados_filtrados"] = dados_filtrados


def categorizar_completo(local):
    if "CEARÁ" in local.upper() or "FORTALEZA" in local.upper() or "- CE" in local.upper() or "CE " in local.upper() or local.upper().strip() == "CE" or "AQUIRAZ" in local.upper() or "CAUCAIA" in local.upper():
        return "CE"
    else:
        return extrair_uf(local)

st.set_page_config(
    page_title = "APP de contagem de status das empresas",
    layout = "wide"
)

st.title("Ferramenta para contagem de status das empresas")

colunas_mapeamento = pd.DataFrame({
            'Job': [],
            'Elementar': [],
            'CNPJ': [],
            'Status do Item': [],
        })

colunas_elementar = pd.DataFrame({
            'Elementar': [],
            'Descricao do Item': [],
            'Unidade': [],
            'Simples/Composto': [],
            'Status de Pesquisa': []
        })

def dataframe_to_markdown(df):
    return df.to_markdown(index=False)

@st.cache_data
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='ContagemStatus')
    return output.getvalue()
    

n_sheet1 = ['Mapeamento-EPE.xlsx', ## 1job
'Mapeamento-SCO.xlsx', ## 1job
'Mapeamento-SIMG ÍNDICE.xlsx', ## 4job
'Mapeamento-IPA e INCC.xlsx'
]
n_sheet2 = ['Mapeamento_INFRAES.xlsx', ## 27jobs
'Mapeamento_SICFER.xlsx', 
'Mapeamento_SABESP.xlsx', 
'Mapeamento_DER_MG_2023.xlsx', ## 6jobs 
'Mapeamento_FGV_Transportes.xlsx', ## 1job
'Mapeamento_CAGECE.xlsx', ## 1job
'Mapeamento_SANEAGO.xlsx', ## 1job
'Mapeamento_DER_SP.xlsx', ## 15jobs
'Mapeamento_GOINFRA.xlsx' ## 1job
]
n_sheet3 = [
    'Mapeamento_ECON_DNIT.xlsx', ## 1job
    'Mapeamento_SICRO (2018 a 2022).xlsx', ## 27jobs
    'Mapeamento_SICRO (2023).xlsx', ## 27jobs
    'Mapeamento_EST_VAREJO.xlsx' ## 27jobs
    ]
n_sheet4 = ['Mapeamento_DER_MG_2022.xlsx'] ## 6jobs
n_sheet6 = ['2- CONTROLE DE MAPEAMENTO - APOIO DNIT-ANTT 2020-2021.xlsx']


## Filtrar por job

job_1 = []
@st.cache_data
def carregar_mapeamento(arquivo, sheets):
    dataframes = []
    for sheet in sheets:
        df = pd.read_excel(arquivo, sheet_name=sheet).astype(str)
        df.columns = df.columns.str.strip().str.lower()
        df['elementar'] = df['elementar'].str.split('.').str[0]
        dataframes.append(df)
    return pd.concat(dataframes, ignore_index=True)

@st.cache_data
def carregar_elementares(arquivo):
    elementares = pd.read_excel(arquivo).astype(str)
    elementares.columns = elementares.columns.str.replace('\n', '').str.strip().str.lower() # Remove a quebra de linha
    return elementares    

with st.expander("Dados necessários"):
    st.markdown('''
    Para carregar os dados relacionados ao **mapeamento** é necessário que a planilha contenha as seguintes colunas:
    ''')
    st.markdown(dataframe_to_markdown(colunas_mapeamento))
    st.markdown('''
    ''')
    st.markdown('''
    ''')
    st.markdown('''
    Para carregar os dados relacionados ao **elementar** é necessário que a planilha contenha as seguintes colunas:
    ''')
    st.markdown(dataframe_to_markdown(colunas_elementar))
    st.markdown('''
    ''')
    st.markdown('''
    ''')
    st.markdown('''          
    **Importante:** as colunas devem conter exatamente os nomes apresentados acima.
    ''')

st.subheader('Selecione o arquivo de mapeamento:', divider = "red")
mapeamento_arquivo = st.file_uploader("arquivo de mapeamento", type="xlsx", label_visibility = "hidden")

if mapeamento_arquivo is not None:
    # Lendo de acordo com a quantidade de abas predeterminada em cada contrato
    if mapeamento_arquivo.name in n_sheet1:
        mapeamento = carregar_mapeamento(mapeamento_arquivo, [0])

    elif mapeamento_arquivo.name in n_sheet2:        
        mapeamento = carregar_mapeamento(mapeamento_arquivo, [0, 1])
    elif mapeamento_arquivo.name in n_sheet3:        
        mapeamento = carregar_mapeamento(mapeamento_arquivo, [0, 1, 2])
    elif mapeamento_arquivo.name in n_sheet4:        
        mapeamento = carregar_mapeamento(mapeamento_arquivo, [0, 1, 2, 3])
    # elif mapeamento_arquivo.name in n_sheet5:    
    #   mapeamento = carregar_mapeamento(mapeamento_arquivo, [0, 1, 2, 3, 4])   
    elif mapeamento_arquivo.name in n_sheet6:        
         mapeamento = carregar_mapeamento(mapeamento_arquivo, [0, 1, 2, 3, 4, 5])
    else:
        st.warning('O arquivo não é compatível!')
        st.stop()
    
    st.subheader('Selecione o arquivo de elementares:', divider = "red")
    elementares_arquivo = st.file_uploader("arquivo de elementares", type="xlsx", label_visibility = "hidden")
    
    @st.cache_data
    def processar_ufs(df):
        uf_pricing = pd.DataFrame({'Local': df['uf do preço'].unique()})
        uf_pricing['UF'] = uf_pricing['Local'].apply(categorizar_completo)
        return uf_pricing.explode("Local").reset_index(drop=True)

    df_locais_por_uf = processar_ufs(mapeamento)
    df_locais_por_uf['UF'] = df_locais_por_uf['UF'].str.upper()
    if elementares_arquivo is not None:

        # Lendo o arquivo de elementares
        elementares = carregar_elementares(elementares_arquivo)
        if 'status de pesquisa' in elementares.columns:
            # Padronizando os valores da coluna
            elementares['status de pesquisa'] = elementares['status de pesquisa'].apply(lambda x: unidecode.unidecode(x.lower()))
        else:
            st.error("A coluna 'status de pesquisa' não foi encontrada no arquivo carregado. Verifique o arquivo e tente novamente.")
            st.stop()
        # Filtrando apenas o que é item pesquisável
        elementares = elementares[elementares['status de pesquisa'] == 'item pesquisavel']
        if 'simples/composto' not in elementares.columns:
            elementares['simples/composto'] = None  # Adiciona a coluna com valores padrão
    
        
        # Normalizando os nomes das colunas
        mapeamento.columns = mapeamento.columns.str.strip().str.lower()
        elementares.columns = elementares.columns.str.strip().str.lower()
        mapeamento['elementar'] = pd.to_numeric(mapeamento['elementar'], errors='coerce').fillna(0).astype(int)
        elementares['elementar'] = pd.to_numeric(elementares['elementar'], errors='coerce').fillna(0).astype(int)


        # Selecionando as variáveis necessárias
        mapeamento = mapeamento[['elementar', 'cnpj', 'status do item', 'job', 'data do mapeamento']].copy()
        elementares = elementares[['elementar', 'descricao do item', 'unidade', 'simples/composto']].copy()


        planilha = pd.merge(elementares, mapeamento, left_on = 'elementar', right_on = 'elementar', how = 'left')

        planilha['status'] = planilha['status do item'].apply(lambda x: 'Sem status' if pd.isna(x) else ('PO' if x == 'PO' else ('AG' if x == 'AG' else ('RE' if x == 'RE' else 'NG'))))

        # Preenchendo o job quando não aparece para não sumir quando fizer a ligação com a tabela de prioridades
        planilha['job'] = planilha['job'].fillna('Sem Job')

        # Criando o DataFrame de prioridades
        prioridade = pd.DataFrame({
            'status': ['PO', 'AG', 'RE','NG', 'Sem status'],
            'prioridade': [1, 2, 3, 4, 5]
        })

        # Realizando a junção dos dados com a ordem de prioridades
        planilha = pd.merge(planilha, prioridade, on = 'status', how = 'left')

        # Ordenando os dados
        planilha = planilha.sort_values(by = ['job', 'elementar', 'cnpj', 'prioridade'])
        

        # Substituindo os cnpjs faltantes
        planilha['cnpj'] = planilha['cnpj'].fillna('CNPJ_DESCONHECIDO')

        planilha['data do mapeamento'] = planilha['data do mapeamento'].fillna('Sem data')
        # Agrupando e filtrando para manter apenas a primeira linha por grupo
        planilha = planilha.groupby(['job', 'elementar', 'cnpj']).head(1)
        
        # Agrupando e contando a quantidade mapeada por status
        qtd_mapeada_por_status = (
            planilha.groupby(['job', 'elementar', 'data do mapeamento', 'status']).size().reset_index(name='count')
        )
        
        # Pivotando para o formato wide
        qtd_mapeada_por_status = qtd_mapeada_por_status.pivot_table(
            index=['job', 'elementar', 'data do mapeamento'],
            columns='status',
            values='count',
            fill_value=0
        ).reset_index()

        # Verificando se a coluna 'Sem status' existe no DataFrame antes de remover
        if 'Sem status' in qtd_mapeada_por_status.columns:
            qtd_mapeada_por_status.drop(columns = ['Sem status'], inplace=True)

        # Trazendo as informações para a tabela 
        empresas = pd.merge(qtd_mapeada_por_status, elementares, on='elementar', how='left')
        status_colunas = ['PO', 'NG', 'RE','AG']
        jobs_unicos = empresas['job'].unique()
        ufs_unicas = df_locais_por_uf['UF'].str.upper().unique()

        # Garantindo que todas as colunas estejam presentes no DataFrame
        for status in status_colunas:
            if status not in empresas.columns:
                empresas[status] = 0  

        # Gerando o resultado da quantidade de empresas mapeadas
        empresas['Empresas Mapeadas'] = empresas['PO'] + empresas['AG'] + empresas['RE']+ empresas['NG']
        # Definindo a ordem das colunas
        colunas_reordenadas = [
            'job', 'data do mapeamento' ,'elementar', 'descricao do item', 'unidade', 'simples/composto' ,'Empresas Mapeadas', 'PO','RE', 'NG', 'AG'
        ]
        colunas_existentes = empresas.columns.tolist()  # Todas as colunas do DataFrame
        colunas_completas = colunas_reordenadas + [col for col in colunas_existentes if col not in colunas_reordenadas]
        # Reordenando as colunas 
        empresas = empresas[colunas_completas]
        for job in jobs_unicos:
            for status in status_colunas:
                # Nome da nova coluna, incluindo apenas o job e o status
                nova_coluna = f"{job.strip().upper()}-{status}".upper()
                # Atribuir diretamente os valores com base na combinação de job e STATUS
                empresas[nova_coluna] = empresas.apply(
                    lambda row: row[status] if row['job'] == job else 0, axis=1
                )
        # Remover colunas originais de UFs, se necessário
        empresas = empresas.drop(columns=status_colunas, errors='ignore')

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

        

        filtro_tipo = st.radio(
            "Tipo de Filtro:",
            ("Nenhum","Única Data", "Intervalo de Datas"),
            horizontal=True
        )
        if filtro_tipo == "Única Data":
            # Seleção de uma única data
            data_selecionada = st.date_input("Escolha uma data", value=None, key="data_unica")
            if data_selecionada:
                data_selecionada = pd.Timestamp(data_selecionada).date()
                empresas_filtradas = empresas[empresas['data do mapeamento'] == pd.Timestamp(data_selecionada)]
            else:
                empresas_filtradas = empresas.copy()  # Sem filtro de data

        elif filtro_tipo == "Intervalo de Datas":
            # Seleção de intervalo de datas
            data_inicial, data_final = st.date_input(
                "Escolha o intervalo de datas",
                value=(empresas['data do mapeamento'].min().date(), empresas['data do mapeamento'].max().date()),
                key="data_intervalo"
            )
            if data_inicial and data_final:
                empresas_filtradas = filtrar_por_intervalo(empresas, data_inicial, data_final)
            else:
                empresas_filtradas = empresas.copy()  # Sem filtro de data
        else:
            # Sem filtro de data
            empresas_filtradas = empresas.copy()

        if 'data do mapeamento' in empresas_filtradas.columns:
            empresas_filtradas['data do mapeamento'] = empresas_filtradas['data do mapeamento'].dt.strftime('%d/%m/%Y')

        empresas_filtradas.columns = [
            col.title() if col in ufs_unicas or col in ['job', 'elementar','descricao do item', 'unidade', 'simples/composto'] else col
            for col in empresas_filtradas.columns
        ]

        empresas_filtradas['Elementar'] = empresas_filtradas['Elementar'].astype(str).str.replace(",", "")

        if 'data do mapeamento' in empresas_filtradas.columns:
            empresas_filtradas_sem_data = empresas_filtradas.drop(columns=['data do mapeamento'])
        else:
            empresas_filtradas_sem_data = empresas_filtradas.copy()
        

        if empresas_filtradas_sem_data.empty:
            st.warning("Nenhum dado encontrado após o filtro.")
        else:
            st.write(empresas_filtradas_sem_data)

        excel_data = to_excel(empresas_filtradas_sem_data)

        # Download do arquivo
        st.subheader('Clique para baixar o resultado:', divider="red")
        st.download_button(label="Baixar", data=excel_data, file_name="Empresas.xlsx")
    else:
     st.info("Por favor, carregue o arquivo de elementares para continuar.")

else:
     st.info("Por favor, carregue o arquivo de mapeamento para continuar.")
