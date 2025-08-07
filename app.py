import pandas as pd
import plotly.express as px
import streamlit as st
import io

# Configurar a página do Streamlit
st.set_page_config(layout="wide")
st.title("Dashboard de Frotas")

# Carregar o arquivo CSV com a codificação 'latin1'
# Altere o caminho se necessário
file_path = 'frotas.csv'
df = pd.read_csv(file_path, delimiter=';', encoding='latin1')

# Renomear colunas com caracteres especiais
df.rename(columns={
    'DescriÃ§Ã£o': 'Descrição',
    'CÃ³digo': 'Código',
    'SÃ©rie': 'Série',
    'ProprietÃ¡rio': 'Proprietário',
    'Filial ProprietÃ¡rio': 'Filial Proprietário'
}, inplace=True)

# Preencher valores ausentes e padronizar
df['Placa'] = df['Placa'].fillna('Não Informado')
df['Patrimonio'] = df['Patrimonio'].fillna('Não Informado')
df['Chassis'] = df['Chassis'].fillna('Não Informado')
df['Série'] = df['Série'].fillna('Não Informado')
df['Proprietário'] = df['Proprietário'].str.replace(' S/A', '', regex=False).str.strip().str.upper()

# Contar a frequência de cada descrição
descricao_counts = df['Descrição'].value_counts()

# Manter apenas as top 10 descrições mais frequentes, para melhor visualização
top_10_descricoes = descricao_counts.index[:10]

# Criar a coluna 'Categoria' com as top 10 descrições e 'Outros' para as demais
df['Categoria'] = df['Descrição'].apply(lambda x: x if x in top_10_descricoes else 'Outros')


# Preparar dados para o gráfico de pizza
df_por_categoria = df['Categoria'].value_counts().reset_index()
df_por_categoria.columns = ['Categoria', 'Número de Itens']


# --- Construir a interface com Streamlit ---

st.subheader("Distribuição da Frota por Categoria")

# Criar um gráfico de barras em vez de pizza para ter uma visualização mais clara das muitas categorias
fig_bar = px.bar(df_por_categoria, x='Categoria', y='Número de Itens',
                 title='Distribuição da Frota por Categoria (Top 10)')
st.plotly_chart(fig_bar, use_container_width=True)


st.subheader("Dados da Frota Completa")

# Adicionar a caixa de pesquisa
search_term = st.text_input("Pesquisar na frota (ex: Trator, Placa, Proprietário)")

# Filtrar o DataFrame com base no termo de pesquisa
if search_term:
    df_filtered = df[df.apply(lambda row: row.astype(str).str.contains(search_term, case=False).any(), axis=1)]
    st.dataframe(df_filtered, use_container_width=True)
else:
    # Exibir a tabela completa se não houver termo de pesquisa
    st.dataframe(df, use_container_width=True)
