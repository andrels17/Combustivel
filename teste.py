import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

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

st.title("Dashboard de Consumo de Abastecimentos")

# Filtros
classes = st.multiselect("Filtrar por Classe", options=df["Classe"].dropna().unique(), default=df["Classe"].dropna().unique())
periodo = st.date_input("Período", [df["Data"].min(), df["Data"].max()])

filtro = (df["Classe"].isin(classes)) & (df["Data"] >= pd.to_datetime(periodo[0])) & (df["Data"] <= pd.to_datetime(periodo[1]))
df_filtrado = df[filtro]

# 1. Equipamentos com consumo acima da média
st.subheader("Equipamentos com Consumo Acima da Média")
acima_media = df_filtrado[df_filtrado["Media"] > df_filtrado["Media_P"]]
st.dataframe(acima_media[["Data", "Descricao_Equip", "Media", "Media_P", "Classe"]])

# 2. Consumo por Classe
st.subheader("Consumo Total por Classe")
classe_group = df_filtrado.groupby("Classe")["Qtde_Litros"].sum().sort_values(ascending=False)
st.bar_chart(classe_group)

# 3. Consumo Semanal
st.subheader("Consumo Semanal")
semana_group = df_filtrado.groupby("AnoSemana")["Qtde_Litros"].sum()
media_semanal = semana_group.mean()
st.line_chart(semana_group)
st.metric("Média Semanal", f"{media_semanal:.2f} litros")

# 4. Consumo Mensal
st.subheader("Consumo Mensal")
mes_group = df_filtrado.groupby("AnoMes")["Qtde_Litros"].sum()
media_mensal = mes_group.mean()
st.line_chart(mes_group)
st.metric("Média Mensal", f"{media_mensal:.2f} litros")
