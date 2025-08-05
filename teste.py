import streamlit as st
import pandas as pd
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder

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
    df["AnoMes"] = df["Data"].dt.to_period("M").astype(str)
    df["AnoSemana"] = df["Data"].dt.strftime("%Y-%U")
    df["Qtde_Litros"] = pd.to_numeric(df["Qtde_Litros"], errors="coerce")
    df["Media"] = pd.to_numeric(df["Media"], errors="coerce")
    df["Media_P"] = pd.to_numeric(df["Media_P"], errors="coerce")
    return df

df = load_data()

st.title("📊 Dashboard de Consumo de Abastecimentos")

# Filtros
st.sidebar.header("Filtros")
classes_op = st.sidebar.multiselect("Classe Operacional", options=df["Classe_Operacional"].dropna().unique(), default=df["Classe_Operacional"].dropna().unique())
safras = st.sidebar.multiselect("Safra", options=df["Safra"].dropna().unique(), default=df["Safra"].dropna().unique())
periodo = st.sidebar.date_input("Período", [df["Data"].min(), df["Data"].max()])

filtro = (
    df["Classe_Operacional"].isin(classes_op) &
    df["Safra"].isin(safras) &
    (df["Data"] >= pd.to_datetime(periodo[0])) &
    (df["Data"] <= pd.to_datetime(periodo[1]))
)
df_filtrado = df[filtro]

# KPIs
col1, col2, col3 = st.columns(3)
col1.metric("Total de Litros Abastecidos", f"{df_filtrado['Qtde_Litros'].sum():,.2f} L")
col2.metric("Média de Consumo (todos)", f"{df_filtrado['Media'].mean():.2f} km/L")
col3.metric("Qtd. Equipamentos Únicos", df_filtrado["Cod_Equip"].nunique())

# Visão geral interativa
with st.expander("🔄 Visão Geral por Classe Operacional", expanded=True):
    media_por_classe_op = df_filtrado.groupby("Classe_Operacional")["Media"].mean().reset_index()
    fig_media_op = px.bar(media_por_classe_op, x="Classe_Operacional", y="Media", text="Media",
                          title="Média de Consumo por Classe Operacional",
                          labels={"Media": "Média (km/l ou equivalente)"})
    fig_media_op.update_traces(texttemplate='%{text:.2f}', textposition='outside')
    fig_media_op.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')
    st.plotly_chart(fig_media_op, use_container_width=True)

# 2. Consumo por Classe
with st.expander("📈 Consumo Total por Classe", expanded=True):
    classe_group = df_filtrado.groupby("Classe")["Qtde_Litros"].sum().reset_index().sort_values("Qtde_Litros", ascending=False)
    fig_classe = px.bar(classe_group, x="Classe", y="Qtde_Litros", text="Qtde_Litros",
                        labels={"Qtde_Litros": "Litros Abastecidos"}, title="Consumo Total por Classe")
    fig_classe.update_traces(texttemplate='%{text:.2s}', textposition='outside')
    fig_classe.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')
    st.plotly_chart(fig_classe, use_container_width=True)

# 3. Consumo Semanal (pizza)
with st.expander("🗓️ Consumo Semanal (Pizza)", expanded=True):
    semana_group = df_filtrado.groupby("AnoSemana")["Qtde_Litros"].sum().reset_index()
    fig_semana = px.pie(semana_group, names="AnoSemana", values="Qtde_Litros", title="Distribuição de Consumo Semanal")
    st.plotly_chart(fig_semana, use_container_width=True)

# 4. Consumo Mensal (barras)
with st.expander("📅 Consumo Mensal (Barras)", expanded=True):
    mes_group = df_filtrado.groupby("AnoMes")["Qtde_Litros"].sum().reset_index()
    fig_mes = px.bar(mes_group, x="AnoMes", y="Qtde_Litros", text="Qtde_Litros",
                     title="Consumo Mensal", labels={"Qtde_Litros": "Litros"})
    fig_mes.update_traces(texttemplate='%{text:.2s}', textposition='outside')
    st.plotly_chart(fig_mes, use_container_width=True)

# 5. Ranking por Equipamento
with st.expander("🚜 Ranking de Veículos por Consumo Médio", expanded=True):
    ranking_media = df_filtrado.groupby("Cod_Equip")["Media"].mean().reset_index()
    ranking_media = ranking_media.sort_values("Media", ascending=False).head(10)
    fig_rank = px.bar(ranking_media, x="Media", y="Cod_Equip", orientation="h",
                      title="Top 10 Veículos mais Econômicos", text="Media")
    fig_rank.update_traces(texttemplate='%{text:.2f}', textposition="outside")
    fig_rank.update_layout(yaxis=dict(autorange="reversed"))
    st.plotly_chart(fig_rank, use_container_width=True)

# 6. Comparativo por Classe Operacional
with st.expander("📊 Comparativo de Classes Operacionais", expanded=False):
    comparativo = df_filtrado.groupby(["Classe_Operacional", "Cod_Equip"])["Media"].mean().reset_index()
    fig_comp = px.box(comparativo, x="Classe_Operacional", y="Media", points="outliers",
                      title="Boxplot da Média de Consumo por Classe Operacional")
    fig_comp.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig_comp, use_container_width=True)

# 7. Tabela interativa com AgGrid
with st.expander("📋 Tabela Detalhada com Filtros", expanded=False):
    gb = GridOptionsBuilder.from_dataframe(df_filtrado)
    gb.configure_pagination()
    gb.configure_default_column(filterable=True, sortable=True, resizable=True)
    grid_options = gb.build()
    AgGrid(df_filtrado.drop(columns=["Descricao_Equip"]), gridOptions=grid_options, enable_enterprise_modules=True, height=400)

# 8. Exportação de dados
with st.expander("⬇️ Exportar Dados", expanded=False):
    csv = df_filtrado.to_csv(index=False).encode("utf-8")
    st.download_button("📥 Baixar CSV", data=csv, file_name="dados_filtrados.csv", mime="text/csv")
