import streamlit as st
import pandas as pd
import plotly.express as px

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
classes = st.sidebar.multiselect("Classe do Veículo", options=df["Classe"].dropna().unique(), default=df["Classe"].dropna().unique())
periodo = st.sidebar.date_input("Período", [df["Data"].min(), df["Data"].max()])

filtro = (df["Classe"].isin(classes)) & (df["Data"] >= pd.to_datetime(periodo[0])) & (df["Data"] <= pd.to_datetime(periodo[1]))
df_filtrado = df[filtro]

# 1. Equipamentos com consumo acima da média
with st.expander("⬆️ Equipamentos com Consumo Acima da Média", expanded=True):
    acima_media = df_filtrado[df_filtrado["Media"] > df_filtrado["Media_P"]]
    st.dataframe(acima_media[["Data", "Descricao_Equip", "Media", "Media_P", "Classe"]])

# 2. Consumo por Classe
with st.expander("📈 Consumo Total por Classe", expanded=True):
    classe_group = df_filtrado.groupby("Classe")["Qtde_Litros"].sum().reset_index().sort_values("Qtde_Litros", ascending=False)
    fig_classe = px.bar(classe_group, x="Classe", y="Qtde_Litros", text="Qtde_Litros",
                        labels={"Qtde_Litros": "Litros Abastecidos"}, title="Consumo Total por Classe")
    fig_classe.update_traces(texttemplate='%{text:.2s}', textposition='outside')
    fig_classe.update_layout(uniformtext_minsize=8, uniformtext_mode='hide')
    st.plotly_chart(fig_classe, use_container_width=True)

# 3. Consumo Semanal
with st.expander("🗓️ Consumo Semanal", expanded=True):
    semana_group = df_filtrado.groupby("AnoSemana")["Qtde_Litros"].sum().reset_index()
    media_semanal = semana_group["Qtde_Litros"].mean()
    fig_semana = px.line(semana_group, x="AnoSemana", y="Qtde_Litros", markers=True,
                         labels={"Qtde_Litros": "Litros"}, title="Consumo Semanal")
    fig_semana.add_scatter(x=semana_group["AnoSemana"], y=[media_semanal]*len(semana_group),
                           mode="lines", name="Média")
    st.plotly_chart(fig_semana, use_container_width=True)
    st.metric("Média Semanal", f"{media_semanal:.2f} litros")

# 4. Consumo Mensal
with st.expander("📅 Consumo Mensal", expanded=True):
    mes_group = df_filtrado.groupby("AnoMes")["Qtde_Litros"].sum().reset_index()
    media_mensal = mes_group["Qtde_Litros"].mean()
    fig_mes = px.line(mes_group, x="AnoMes", y="Qtde_Litros", markers=True,
                      labels={"Qtde_Litros": "Litros"}, title="Consumo Mensal")
    fig_mes.add_scatter(x=mes_group["AnoMes"], y=[media_mensal]*len(mes_group),
                        mode="lines", name="Média")
    st.plotly_chart(fig_mes, use_container_width=True)
    st.metric("Média Mensal", f"{media_mensal:.2f} litros")
