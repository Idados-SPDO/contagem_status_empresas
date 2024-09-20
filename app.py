import streamlit as st
import pandas as pd
from io import BytesIO
import unidecode

st.set_page_config(
    page_title = "APP de contagem de status das empresas",
    layout = "wide"
)

st.title("Ferramenta para contagem de status das empresas")

colunas_mapeamento = pd.DataFrame({
            'Job': [],
            'Elementar': [],
            'CNPJ': [],
            'Status do Item': []
        })

colunas_elementar = pd.DataFrame({
            'Elementar': [],
            'Descricao do Item': [],
            'Unidade': [],
            'Simples/Composto': [],
            'Status de pesquisa': []
        })

def dataframe_to_markdown(df):
    return df.to_markdown(index=False)

n_sheet1 = ['Mapeamento-EPE.xlsx', 'Mapeamento-SCO.xlsx', 'Mapeamento-SIMG ÍNDICE.xlsx', 'Mapeamento-IPA e INCC.xlsx']
n_sheet2 = ['Mapeamento_INFRAES.xlsx', 'Mapeamento_SICFER.xlsx', 'Mapeamento_SABESP.xlsx', 'Mapeamento_DER_MG_2023.xlsx', 
            'Mapeamento_FGV_Transportes.xlsx', 'Mapeamento_CAGECE.xlsx', 'Mapeamento_SANEAGO.xlsx', 'Mapeamento_DER_SP.xlsx',
            'Mapeamento_GOINFRA.xlsx']
n_sheet3 = ['Mapeamento_ECON_DNIT.xlsx', 'Mapeamento_SICRO (2018 a 2022).xlsx', 'Mapeamento_SICRO (2023).xlsx', 
            'Mapeamento_EST_VAREJO.xlsx']
n_sheet4 = ['Mapeamento_DER_MG_2022.xlsx']
n_sheet6 = ['2- CONTROLE DE MAPEAMENTO - APOIO DNIT-ANTT 2020-2021.xlsx']

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
            mapeamento = pd.read_excel(mapeamento_arquivo, sheet_name = 0).astype(str)
            
            # Garantindo que a coluna 'Elementar' seja convertida para string e sem decimal  
            mapeamento['Elementar'] = mapeamento['Elementar'].str.split('.').str[0]  

    elif mapeamento_arquivo.name in n_sheet2:        
        mapeamento1 = pd.read_excel(mapeamento_arquivo, sheet_name = 0).astype(str)  
        mapeamento2 = pd.read_excel(mapeamento_arquivo, sheet_name = 1).astype(str)  

        # Garantindo que a coluna 'Elementar' seja convertida para string e sem decimal
        mapeamento1['Elementar'] = mapeamento1['Elementar'].str.split('.').str[0]  
        mapeamento2['Elementar'] = mapeamento2['Elementar'].str.split('.').str[0]  

        # Concatenando as abas
        mapeamento = pd.concat([mapeamento1, mapeamento2], ignore_index=True)

    elif mapeamento_arquivo.name in n_sheet3:        
        mapeamento1 = pd.read_excel(mapeamento_arquivo, sheet_name = 0).astype(str)  
        mapeamento2 = pd.read_excel(mapeamento_arquivo, sheet_name = 1).astype(str)  
        mapeamento3 = pd.read_excel(mapeamento_arquivo, sheet_name = 2).astype(str)  

        # Garantindo que a coluna 'Elementar' seja convertida para string e sem decimal
        mapeamento1['Elementar'] = mapeamento1['Elementar'].str.split('.').str[0]  
        mapeamento2['Elementar'] = mapeamento2['Elementar'].str.split('.').str[0]  
        mapeamento3['Elementar'] = mapeamento3['Elementar'].str.split('.').str[0]  

        # Concatenando as abas
        mapeamento = pd.concat([mapeamento1, mapeamento2, mapeamento3], ignore_index=True)

    elif mapeamento_arquivo.name in n_sheet4:        
        mapeamento1 = pd.read_excel(mapeamento_arquivo, sheet_name = 0).astype(str)  
        mapeamento2 = pd.read_excel(mapeamento_arquivo, sheet_name = 1).astype(str)  
        mapeamento3 = pd.read_excel(mapeamento_arquivo, sheet_name = 2).astype(str)  
        mapeamento4 = pd.read_excel(mapeamento_arquivo, sheet_name = 3).astype(str)  

        # Garantindo que a coluna 'Elementar' seja convertida para string e sem decimal
        mapeamento1['Elementar'] = mapeamento1['Elementar'].str.split('.').str[0]  
        mapeamento2['Elementar'] = mapeamento2['Elementar'].str.split('.').str[0]  
        mapeamento3['Elementar'] = mapeamento3['Elementar'].str.split('.').str[0]  
        mapeamento4['Elementar'] = mapeamento4['Elementar'].str.split('.').str[0]  

        # Concatenando as abas
        mapeamento = pd.concat([mapeamento1, mapeamento2, mapeamento3, mapeamento4], ignore_index=True)

    # elif mapeamento_arquivo.name in n_sheet5:        
    #     mapeamento1 = pd.read_excel(mapeamento_arquivo, sheet_name = 0).astype(str)  
    #     mapeamento2 = pd.read_excel(mapeamento_arquivo, sheet_name = 1).astype(str)  
    #     mapeamento3 = pd.read_excel(mapeamento_arquivo, sheet_name = 2).astype(str)  
    #     mapeamento4 = pd.read_excel(mapeamento_arquivo, sheet_name = 3).astype(str)  
    #     mapeamento5 = pd.read_excel(mapeamento_arquivo, sheet_name = 4).astype(str)  


    #     # Garantindo que a coluna 'Elementar' seja convertida para string e sem decimal
    #     mapeamento1['Elementar'] = mapeamento1['Elementar'].str.split('.').str[0]  
    #     mapeamento2['Elementar'] = mapeamento2['Elementar'].str.split('.').str[0]  
    #     mapeamento3['Elementar'] = mapeamento3['Elementar'].str.split('.').str[0]  
    #     mapeamento4['Elementar'] = mapeamento4['Elementar'].str.split('.').str[0]  
    #     mapeamento5['Elementar'] = mapeamento5['Elementar'].str.split('.').str[0]  

    #     # Concatenando as abas
    #     mapeamento = pd.concat([mapeamento1, mapeamento2, mapeamento3, mapeamento4, mapeamento5], ignore_index=True)

    elif mapeamento_arquivo.name in n_sheet6:        
        mapeamento1 = pd.read_excel(mapeamento_arquivo, sheet_name = 0).astype(str)  
        mapeamento2 = pd.read_excel(mapeamento_arquivo, sheet_name = 1).astype(str)  
        mapeamento3 = pd.read_excel(mapeamento_arquivo, sheet_name = 2).astype(str)  
        mapeamento4 = pd.read_excel(mapeamento_arquivo, sheet_name = 3).astype(str)  
        mapeamento5 = pd.read_excel(mapeamento_arquivo, sheet_name = 4).astype(str)  
        mapeamento6 = pd.read_excel(mapeamento_arquivo, sheet_name = 5).astype(str) 

        # Garantindo que a coluna 'Elementar' seja convertida para string e sem decimal
        mapeamento1['Elementar'] = mapeamento1['Elementar'].str.split('.').str[0]  
        mapeamento2['Elementar'] = mapeamento2['Elementar'].str.split('.').str[0]  
        mapeamento3['Elementar'] = mapeamento3['Elementar'].str.split('.').str[0]  
        mapeamento4['Elementar'] = mapeamento4['Elementar'].str.split('.').str[0]  
        mapeamento5['Elementar'] = mapeamento5['Elementar'].str.split('.').str[0]  
        mapeamento6['Elementar'] = mapeamento6['Elementar'].str.split('.').str[0]  

        # Concatenando as abas
        mapeamento = pd.concat([mapeamento1, mapeamento2, mapeamento3, mapeamento4, mapeamento5, mapeamento6], ignore_index=True)
    
    else:
        st.warning('O arquivo não é compatível!')
        st.stop()

    st.subheader('Selecione o arquivo de elementares:', divider = "red")
    elementares_arquivo = st.file_uploader("arquivo de elementares", type="xlsx", label_visibility = "hidden")

    if elementares_arquivo is not None:
        
        # Lendo o arquivo de elementares
        elementares = pd.read_excel(elementares_arquivo).astype(str)
        elementares.columns = elementares.columns.str.replace('\n', '').str.strip() # Remove a quebra de linha

        # Padronizando a coluna de status
        elementares['Status de pesquisa'] = elementares['Status de pesquisa'].apply(lambda x: unidecode.unidecode(x.lower()))

        # Filtrando apenas o que é item pesquisável
        elementares = elementares[elementares['Status de pesquisa'] == 'item pesquisavel']

        # Selecionando as variáveis
        mapeamento = mapeamento[['Elementar', 'CNPJ', 'Status do Item', 'Job']]
        elementares = elementares[['Elementar', 'Descricao do Item', 'Unidade', 'Simples/Composto']]

        # Realizando a junção dos dados
        planilha = pd.merge(elementares, mapeamento, left_on = 'Elementar', right_on = 'Elementar', how = 'left')

        # Redefinindo os status
        planilha['status'] = planilha['Status do Item'].apply(lambda x: 'Sem status' if pd.isna(x) else ('PO' if x == 'PO' else ('AG' if x == 'AG' else 'NG')))

        # Preenchendo o job quando não aparece para não sumir quando fizer a ligação com a tabela de prioridades
        planilha['Job'] = planilha['Job'].fillna('Sem Job')

        # Criando o DataFrame de prioridades
        prioridade = pd.DataFrame({
            'status': ['PO', 'AG', 'NG', 'Sem status'],
            'prioridade': [1, 2, 3, 4]
        })

        # Realizando a junção dos dados com a ordem de prioridades
        planilha = pd.merge(planilha, prioridade, on = 'status', how = 'left')

        # Ordenando os dados
        planilha = planilha.sort_values(by = ['Job', 'Elementar', 'CNPJ', 'prioridade'])

        # Substituindo os CNPJs faltantes
        planilha['CNPJ'] = planilha['CNPJ'].fillna('CNPJ_DESCONHECIDO')

        # Agrupando e filtrando para manter apenas a primeira linha por grupo
        planilha = planilha.groupby(['Job', 'Elementar', 'CNPJ']).head(1)

        # Agrupando e contando a quantidade mapeada por status
        qtd_mapeada_por_status = (
            planilha.groupby(['Job', 'Elementar', 'status']).size().reset_index(name = 'count')
        )

        # Pivotando para o formato wide
        qtd_mapeada_por_status = qtd_mapeada_por_status.pivot_table(
            index = ['Job', 'Elementar'],
            columns = 'status',
            values = 'count',
            fill_value = 0
        ).reset_index()

        # Verificando se a coluna 'Sem status' existe no DataFrame antes de remover
        if 'Sem status' in qtd_mapeada_por_status.columns:
            qtd_mapeada_por_status.drop(columns = ['Sem status'], inplace=True)

        # Trazendo as informações para a tabela 
        empresas = pd.merge(qtd_mapeada_por_status, elementares, left_on = 'Elementar', right_on = 'Elementar', how = 'left')

        # Garantindo que todas as colunas estejam presentes no DataFrame
        for status in ['PO', 'AG', 'NG']:
            if status not in empresas.columns:
                empresas[status] = 0  

        # Gerando o resultado da quantidade de empresas mapeadas
        empresas['Empresas Mapeadas'] = empresas['PO'] + empresas['AG'] + empresas['NG']

        # Definindo a ordem das colunas
        colunas_reordenadas = [
            'Job', 'Elementar', 'Descricao do Item', 'Unidade', 'Simples/Composto', 'Empresas Mapeadas', 'PO', 'NG', 'AG'
        ]

        # Reordenando as colunas 
        empresas = empresas.reindex(columns = colunas_reordenadas)

        # Obtendo a lista de jobs
        jobs = planilha['Job'].unique()

        # Selecionando qual job será mostrado
        job = st.selectbox(label="JOB", options=jobs)

        # Filtrando a planilha
        empresas_filtradas = empresas[empresas['Job'] == job]

        # Apresentando a tabela com os resultados
        st.write(empresas_filtradas)

        # Criando um buffer para o arquivo Excel
        buffer = BytesIO()

        # Usando pd.ExcelWriter para escrever no buffer
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            # Iterando sobre cada job e criar uma aba correspondente
            for job in jobs:
                # Filtrando o DataFrame para o job 
                df_job = empresas[empresas['Job'] == job]
                
                # Adicionando uma aba para cada job
                df_job.to_excel(writer, index=False, sheet_name = job)

        # Colocando o buffer no início
        buffer.seek(0)

        # Fazendo o download do arquivo
        st.subheader('Clique para baixar o resultado:', divider = "red")
        st.download_button(label="Baixar", data = buffer, file_name = "Empresas.xlsx")