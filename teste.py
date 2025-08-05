import streamlit as st
import pandas as pd
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder

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
    df["Fazenda"] = df["Ref1"]  # Assumindo que Ref1 representa a fazenda
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
selected_meses = sorted(df["Mes"].dropna().unique()) if meses_check else st.sidebar.multiselect("Mês", options=sorted(df["Mes"].dropna().unique()))

semanas_check = st.sidebar.checkbox("Todas as Semanas", value=True)
selected_semanas = sorted(df["Semana"].dropna().unique()) if semanas_check else st.sidebar.multiselect("Semana", options=sorted(df["Semana"].dropna().unique()))

fazendas_check = st.sidebar.checkbox("Todas as Fazendas", value=True)
selected_fazendas = df["Fazenda"].dropna().unique() if fazendas_check else st.sidebar.multiselect("Fazenda", options=sorted(df["Fazenda"].dropna().unique()))

periodo = st.sidebar.date_input("Período", [df["Data"].min(), df["Data"].max()])

filtro = (
    df["Classe_Operacional"].isin(selected_classes_op) &
    df["Safra"].isin(selected_safras) &
    df["Ano"].isin(selected_anos) &
    df["Mes"].isin(selected_meses) &
    df["Semana"].isin(selected_semanas) &
    df["Fazenda"].isin(selected_fazendas) &
    (df["Data"] >= pd.to_datetime(periodo[0])) &
    (df["Data"] <= pd.to_datetime(periodo[1]))
)
df_filtrado = df[filtro]

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

# Visão geral interativa
with st.expander("🔄 Visão Geral por Classe Operacional", expanded=True):
    media_por_classe_op = df_filtrado.groupby("Classe_Operacional")["Media"].mean().reset_index()
    fig_media_op = px.bar(media_por_classe_op, x="Classe_Operacional", y="Media", text="Media",
                          title="Média de Consumo por Classe Operacional",
                          labels={"Media": "Média (km/l ou equivalente)"})
    fig_media_op.update_traces(texttemplate='%{text:.2f}', textposition='outside')
    fig_media_op.update_layout(uniformtext_minsize=8, uniformtext_mode='hide', xaxis_tickangle=-45)
    st.plotly_chart(fig_media_op, use_container_width=True)

# Consumo por Classe
with st.expander("📈 Consumo Total por Classe", expanded=True):
    classe_group = df_filtrado.groupby("Classe")["Qtde_Litros"].sum().reset_index().sort_values("Qtde_Litros", ascending=False)
    fig_classe = px.bar(classe_group, x="Classe", y="Qtde_Litros", text="Qtde_Litros",
                        labels={"Qtde_Litros": "Litros Abastecidos"}, title="Consumo Total por Classe")
    fig_classe.update_traces(texttemplate='%{text:.2f}', textposition='outside')
    fig_classe.update_layout(uniformtext_minsize=8, uniformtext_mode='hide', xaxis_tickangle=-45)
    st.plotly_chart(fig_classe, use_container_width=True)

# Consumo Semanal (pizza)
with st.expander("🗓️ Consumo Semanal (Pizza)", expanded=True):
    semana_group = df_filtrado.groupby("AnoSemana")["Qtde_Litros"].sum().reset_index()
    fig_semana = px.pie(semana_group, names="AnoSemana", values="Qtde_Litros", title="Distribuição de Consumo Semanal")
    st.plotly_chart(fig_semana, use_container_width=True)

# Consumo Mensal (barras)
with st.expander("📅 Consumo Mensal (Barras)", expanded=True):
    mes_group = df_filtrado.groupby("AnoMes")["Qtde_Litros"].sum().reset_index()
    fig_mes = px.bar(mes_group, x="AnoMes", y="Qtde_Litros", text="Qtde_Litros",
                     title="Consumo Mensal", labels={"Qtde_Litros": "Litros"})
    fig_mes.update_traces(texttemplate='%{text:.2f}', textposition='outside')
    st.plotly_chart(fig_mes, use_container_width=True)

# Tendência de Consumo por Equipamento (Top 10 em Litros)
with st.expander("📊 Tendência de Consumo por Equipamento (Top 10)", expanded=True):
    if df_filtrado.empty:
        st.info("Nenhum dado disponível para o período e filtros selecionados.")
    else:
        # Identificar os 10 equipamentos com maior consumo total
        top10_equipamentos = (
            df_filtrado.groupby("Cod_Equip")["Qtde_Litros"]
            .sum()
            .sort_values(ascending=False)
            .head(10)
            .index
        )

        # Filtrar apenas os dados desses equipamentos
        df_top10 = df_filtrado[df_filtrado["Cod_Equip"].isin(top10_equipamentos)]

        # Agrupar média mensal por equipamento
        tendencia = df_top10.groupby(["AnoMes", "Cod_Equip"])["Media"].mean().reset_index()

        # Criar gráfico de linha
        fig_linha = px.line(
            tendencia,
            x="AnoMes",
            y="Media",
            color="Cod_Equip",
            markers=True,
            title="Tendência de Consumo por Equipamento (Top 10 em Litros)",
            labels={
                "AnoMes": "Ano/Mês",
                "Media": "Média de Consumo",
                "Cod_Equip": "Código do Equipamento"
            }
        )

        # Estilizar gráfico
        fig_linha.update_layout(
            xaxis_tickangle=-45,
            legend_title_text="Equipamento",
            hovermode="x unified"
        )

        st.plotly_chart(fig_linha, use_container_width=True)

# Ranking por Equipamento
with st.expander("🚜 Ranking de Veículos por Consumo Médio", expanded=True):
    # Agrupar por código e descrição
    ranking_media = df_filtrado.groupby(["Cod_Equip", "Descricao_Equip"])["Media"].mean().reset_index()

    # Ordenar pela média decrescente e pegar o Top 10
    ranking_media = ranking_media.sort_values("Media", ascending=False).head(10)

    # Criar rótulo combinando código e descrição
    ranking_media["Label"] = ranking_media["Cod_Equip"].astype(str) + " - " + ranking_media["Descricao_Equip"]

    # Criar gráfico
    fig_rank = px.bar(
        ranking_media,
        x="Label",
        y="Media",
        text="Media",
        title="Top 10 Veículos mais Econômicos",
        labels={"Media": "Média de Consumo"}
    )

    # Formatando texto no gráfico
    fig_rank.update_traces(texttemplate='%{text:.2f}', textposition="outside")

    # Ajustes visuais
    fig_rank.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig_rank, use_container_width=True)

# Comparativo por Classe Operacional com rótulos personalizados
with st.expander("📦 Distribuição de Consumo por Classe Operacional", expanded=True):
    fig_box = px.box(
        df_filtrado,
        x="Classe_Operacional",
        y="Media",
        points="outliers",  # para destacar valores fora da curva
        title="Distribuição de Consumo por Classe Operacional",
        labels={"Media": "Média de Consumo", "Classe_Operacional": "Classe Operacional"}
    )

    fig_box.update_layout(xaxis_tickangle=-45)
    st.plotly_chart(fig_box, use_container_width=True)

# Tabela interativa com AgGrid
with st.expander("📋 Tabela Detalhada com Filtros", expanded=False):
    gb = GridOptionsBuilder.from_dataframe(df_filtrado)
    gb.configure_pagination()
    gb.configure_default_column(filterable=True, sortable=True, resizable=True)
    grid_options = gb.build()
    AgGrid(df_filtrado.drop(columns=["Descricao_Equip"]), gridOptions=grid_options, enable_enterprise_modules=True, height=400)

# Exportação de dados
with st.expander("⬇️ Exportar Dados", expanded=False):
    csv = df_filtrado.to_csv(index=False).encode("utf-8")
    st.download_button("📥 Baixar CSV", data=csv, file_name="dados_filtrados.csv", mime="text/csv")
