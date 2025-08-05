import streamlit as st
import pandas as pd
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder
import calendar

# Função auxiliar para formatação brasileira
def formatar_brasileiro(valor):
    return "{:,.2f}".format(valor).replace(",", "X").replace(".", ",").replace("X", ".")

# Carregar os dados
@st.cache_data
def load_data():
    df = pd.read_excel("Acompto_Abast.xlsx", sheet_name="BD", skiprows=2)
    df.columns = [
        "Data", "Cod_Equip", "Descricao_Equip", "Qtde_Litros", "Km_Hs_Rod",
        "Media", "Media_P", "Perc_Media", "Ton_Cana", "Litros_Ton", 
        "Ref1", "Ref2", "Unidade", "Safra", "Mes", "Semana",
        "Classe", "Classe_Operacional", "Descricao_Proprietario", "Potencia_CV"
    ]
    df = df[pd.to_datetime(df["Data"], errors='coerce').notna()]
    df["Data"] = pd.to_datetime(df["Data"])
    df["Ano"] = df["Data"].dt.year
    df["AnoMes"] = df["Data"].dt.to_period("M").astype(str)
    df["AnoSemana"] = df["Data"].dt.strftime("%Y-%U")
    df["Qtde_Litros"] = pd.to_numeric(df["Qtde_Litros"], errors="coerce")
    df["Media"] = pd.to_numeric(df["Media"], errors="coerce")
    df["Media_P"] = pd.to_numeric(df["Media_P"], errors="coerce")
    df["NomeMes"] = df["Data"].dt.month.apply(lambda x: calendar.month_name[x])
    df["AnoMesLabel"] = df["Data"].dt.to_period("M").dt.strftime('%b %Y')
    return df

df = load_data()

st.title("📊 Dashboard de Consumo de Abastecimentos")

# Filtros
st.sidebar.header("Filtros")
classes_op = st.sidebar.checkbox("Todas as Classes Operacionais", value=True)
selected_classes_op = df["Classe_Operacional"].dropna().unique() if classes_op else st.sidebar.multiselect("Classe Operacional", options=df["Classe_Operacional"].dropna().unique())

safras_check = st.sidebar.checkbox("Todas as Safras", value=True)
selected_safras = df["Safra"].dropna().unique() if safras_check else st.sidebar.multiselect("Safra", options=df["Safra"].dropna().unique())

anos_check = st.sidebar.checkbox("Todos os Anos", value=True)
selected_anos = df["Ano"].dropna().unique() if anos_check else st.sidebar.multiselect("Ano", options=df["Ano"].dropna().unique())

meses_check = st.sidebar.checkbox("Todos os Meses", value=True)
meses_ordenados = list(calendar.month_name)[1:]
selected_meses = meses_ordenados if meses_check else st.sidebar.multiselect("Mês", options=meses_ordenados)

semanas_check = st.sidebar.checkbox("Todas as Semanas", value=True)
selected_semanas = sorted(df["Semana"].dropna().unique()) if semanas_check else st.sidebar.multiselect("Semana", options=sorted(df["Semana"].dropna().unique()))

periodo = st.sidebar.date_input("Período", [df["Data"].min(), df["Data"].max()])

filtro = (
    df["Classe_Operacional"].isin(selected_classes_op) &
    df["Safra"].isin(selected_safras) &
    df["Ano"].isin(selected_anos) &
    df["NomeMes"].isin(selected_meses) &
    df["Semana"].isin(selected_semanas) &
    (df["Data"] >= pd.to_datetime(periodo[0])) &
    (df["Data"] <= pd.to_datetime(periodo[1]))
)
df_filtrado = df[filtro]

# Adicionar campo legível para equipamento
# Agora Equipamento_Label é o Cod_Equip de forma limpa
#
df_filtrado["Equipamento_Label"] = df_filtrado["Cod_Equip"].astype(str)

# KPIs
col1, col2, col3 = st.columns(3)
col1.metric("Total de Litros Abastecidos", formatar_brasileiro(df_filtrado['Qtde_Litros'].sum()))
col2.metric("Média de Consumo (todos)", formatar_brasileiro(df_filtrado['Media'].mean()))
col3.metric("Qtd. Equipamentos Únicos", df_filtrado["Cod_Equip"].nunique())

# Alertas: veículos com consumo abaixo de 1.5 ou acima de 5 (exemplo)
with st.expander("🚨 Alertas de Consumo Fora do Padrão", expanded=True):
    alertas = df_filtrado[(df_filtrado['Media'] < 1.5) | (df_filtrado['Media'] > 5)]
    if alertas.empty:
        st.success("Nenhum veículo com consumo fora do padrão identificado.")
    else:
        st.warning(f"{alertas['Cod_Equip'].nunique()} veículos com consumo fora do padrão.")
        st.dataframe(alertas[["Data", "Cod_Equip", "Classe_Operacional", "Media"]])

# Tendência por equipamento (melhorada com Nome Mês)
with st.expander("📉 Tendência de Consumo por Equipamento", expanded=True):
    tendencia = df_filtrado.groupby(["AnoMesLabel", "Cod_Equip"])["Media"].mean().reset_index()
    fig_tend = px.line(tendencia, x="AnoMesLabel", y="Media", color="Cod_Equip",
                       title="Tendência de Consumo Médio por Equipamento")
    fig_tend.update_layout(xaxis_title="Mês", yaxis_title="Média de Consumo")
    st.plotly_chart(fig_tend, use_container_width=True)

# Ranking por Equipamento (com Cod_Equip claro no eixo X)
with st.expander("🚜 Ranking de Veículos por Consumo Médio", expanded=True):
    ranking_media = df_filtrado.groupby("Cod_Equip")["Media"].mean().reset_index()
    ranking_media = ranking_media.sort_values("Media", ascending=False).head(10)
    fig_rank = px.bar(ranking_media, x="Cod_Equip", y="Media", text="Media",
                      title="Top 10 Veículos mais Econômicos")
    fig_rank.update_traces(texttemplate='%{text:.2f}', textposition="outside")
    fig_rank.update_layout(xaxis_title="Código do Equipamento", xaxis_tickangle=-45)
    st.plotly_chart(fig_rank, use_container_width=True)

# Comparativo por Ano
with st.expander("🧾 Comparativo de Consumo Total por Ano", expanded=True):
    consumo_ano = df_filtrado.groupby("Ano")["Qtde_Litros"].sum().reset_index()
    fig_ano = px.bar(consumo_ano, x="Ano", y="Qtde_Litros", text="Qtde_Litros",
                     labels={"Qtde_Litros": "Litros Abastecidos"}, title="Consumo Total por Ano")
    fig_ano.update_traces(texttemplate='%{text:.2f}', textposition='outside')
    st.plotly_chart(fig_ano, use_container_width=True)

# Os demais gráficos continuam os mesmos...
