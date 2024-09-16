import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="APP de contagem de status das empresas",
    layout="wide"
)

st.title("Ferramenta para contagem de status das empresas")

colunas_mapeamento = pd.DataFrame({
            'Elementar': [],
            'CNPJ': [],
            'Status do Item': []
        })

colunas_elementar = pd.DataFrame({
            'Elementar': [],
            'Descricao do Item': [],
            'Unidade': [],
            'Simples/Composto': []
        })

def dataframe_to_markdown(df):
    return df.to_markdown(index=False)

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
    **Importante:** as colunas devem conter exatamente os mesmos nomes apresentados acima.
    ''')

st.subheader('Selecione o arquivo de mapeamento:', divider="red")
mapeamento_arquivo = st.file_uploader("arquivo de mapeamento", type="xlsx", label_visibility="hidden")

if mapeamento_arquivo is not None:

    st.subheader('Selecione o arquivo de elementares:', divider="red")
    elementares_arquivo = st.file_uploader("arquivo de elementares", type="xlsx", label_visibility="hidden")

    if elementares_arquivo is not None:

        # Lendo as abas 1 e 2 do arquivo de mapeamento
        mapeamento1 = pd.read_excel(mapeamento_arquivo, sheet_name=0).astype(str)  # Aba 1
        mapeamento2 = pd.read_excel(mapeamento_arquivo, sheet_name=1).astype(str)  # Aba 2
        # Concatenando as abas
        mapeamento = pd.concat([mapeamento1, mapeamento2], ignore_index=True)

        # Lendo o arquivo de elementares
        elementares = pd.read_excel(elementares_arquivo).astype(str)
        elementares.columns = elementares.columns.str.replace('\n', '').str.strip()

        # Selecionando as variáveis
        mapeamento = mapeamento[['Elementar', 'CNPJ', 'Status do Item']]
        elementares = elementares[['Elementar', 'Descricao do Item', 'Unidade', 'Simples/Composto']]

        planilha = pd.merge(elementares, mapeamento, left_on='Elementar', right_on='Elementar', how='left')

        planilha['status'] = planilha['Status do Item'].apply(lambda x: 'Sem status' if pd.isna(x) else ('PO' if x == 'PO' else ('AG' if x == 'AG' else 'NG')))

        # Criar o DataFrame de prioridades
        prioridade = pd.DataFrame({
            'status': ['PO', 'AG', 'NG', 'Sem status'],
            'prioridade': [1, 2, 3, 4]
        })

        # Realizar a junção à esquerda com o DataFrame de prioridades
        planilha = pd.merge(planilha, prioridade, on='status', how='left')

        # Ordenar os dados
        planilha = planilha.sort_values(by=['Elementar', 'CNPJ', 'prioridade'])

        planilha['CNPJ'] = planilha['CNPJ'].fillna('CNPJ_DESCONHECIDO')

        # Agrupar e filtrar para manter apenas a primeira linha por grupo
        planilha = planilha.groupby(['Elementar', 'CNPJ']).head(1)

        # Agrupar e contar a quantidade mapeada por status
        qtd_mapeada_por_status = (
            planilha.groupby(['Elementar', 'status']).size().reset_index(name='count')
        )

        # Pivotar para o formato wide
        qtd_mapeada_por_status = qtd_mapeada_por_status.pivot_table(
            index='Elementar',
            columns='status',
            values='count',
            fill_value=0
        ).reset_index()

        qtd_mapeada_por_status.drop(columns=['Sem status'], inplace=True)

        empresas = pd.merge(qtd_mapeada_por_status, elementares, left_on='Elementar', right_on='Elementar', how='left')

        empresas['Empresas Mapeadas'] = empresas['PO'] + empresas['AG'] + empresas['NG']

        # Definir a ordem desejada das colunas
        colunas_reordenadas = [
            'Elementar', 'Descricao do Item', 'Unidade', 'Simples/Composto', 'Empresas Mapeadas', 'PO', 'NG', 'AG'
        ]

        # Reordenar as colunas usando reindex
        empresas = empresas.reindex(columns=colunas_reordenadas)

        # Apresentando a tabela com os resultados
        st.write(empresas)

        # Download da tabela
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            empresas.to_excel(writer, index=False, sheet_name="Contagem")
        buffer.seek(0)  

        st.subheader('Clique para baixar o resultado:', divider="red")
        st.download_button(label="Baixar", data=buffer, file_name="Empresas.xlsx")

