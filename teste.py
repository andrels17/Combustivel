# app.py

import streamlit as st
import pandas as pd
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder

# --------------- Configurações ---------------

EXCEL_PATH  = "Acompto_Abast.xlsx"
SHEET_NAME  = "BD"
ALERTA_MIN  = 1.5
ALERTA_MAX  = 5.0

# --------------- Utilitários ---------------

def formatar_brasileiro(valor: float) -> str:
    """Formata número no padrão brasileiro com duas casas decimais."""
    return (
        "{:,.2f}".format(valor)
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )

@st.cache_data(show_spinner=False)
def load_data(path: str, sheet: str) -> pd.DataFrame:
    """
    Carrega e prepara o DataFrame:
     - Lê Excel, renomeia colunas
     - Converte Data
     - Extrai mês e semana
     - Gera colunas Ano, AnoMes, AnoSemana
     - Converte numéricos
     - Define Fazenda
    """
    try:
        df = pd.read_excel(path, sheet_name=sheet, skiprows=2)
    except FileNotFoundError:
        st.error(f"Arquivo não encontrado em `{path}`")
        st.stop()

    df.columns = [
        "Data", "Cod_Equip", "Descricao_Equip", "Qtde_Litros", "Km_Hs_Rod",
        "Media", "Media_P", "Perc_Media", "Ton_Cana", "Litros_Ton",
        "Ref1", "Ref2", "Unidade", "Safra", "Mes_Excel", "Semana_Excel",
        "Classe", "Classe_Operacional", "Descricao_Proprietario", "Potencia_CV"
    ]

    # Converte Data e filtra nulos
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df = df[df["Data"].notna()]

    # Extrai mês (1–12) e semana ISO (1–53)
    df["Mes"]     = df["Data"].dt.month
    df["Semana"]  = df["Data"].dt.isocalendar().week

    # Períodos auxiliares
    df["Ano"]      = df["Data"].dt.year
    df["AnoMes"]   = df["Data"].dt.to_period("M").astype(str)
    df["AnoSemana"] = df["Data"].dt.strftime("%Y-%U")

    # Converte colunas numéricas
    df["Qtde_Litros"] = pd.to_numeric(df["Qtde_Litros"], errors="coerce")
    df["Media"]       = pd.to_numeric(df["Media"], errors="coerce")
    df["Media_P"]     = pd.to_numeric(df["Media_P"], errors="coerce")

    # Define Fazenda
    df["Fazenda"] = df["Ref1"].astype(str)

    return df

Esse erro acontece porque o Streamlit gera um ID interno para cada widget com base no rótulo (label) e, quando você usa o mesmo texto duas vezes (ou mais), acaba gerando IDs duplicados. A solução é atribuir a cada widget uma chave (key) única.
Abaixo segue a versão ajustada da sua função sidebar_filters, com key explícito em todos os checkboxes, multiselects e no date_input. Basta substituir sua função atual por esta:
def sidebar_filters(df: pd.DataFrame) -> dict:
    """
    Constrói a barra lateral de filtros, com dependência entre eles.
    Cada widget recebe uma key única para evitar StreamlitDuplicateElementId.
    """
    st.sidebar.header("📅 Filtros")

    ano_max    = int(df["Ano"].max())
    mes_max    = int(df[df["Ano"] == ano_max]["Mes"].max())
    semana_max = int(df[df["Ano"] == ano_max]["Semana"].max())
    safra_max  = sorted(df["Safra"].dropna().unique())[-1]

    # Safra
    todas_safras = st.sidebar.checkbox(
        "Todas as Safras", value=False, key="cb_todas_safras"
    )
    safras_opts = sorted(df["Safra"].dropna().unique())
    sel_safras = safras_opts if todas_safras else st.sidebar.multiselect(
        "Safra", safras_opts, default=[safra_max], key="ms_safras"
    )

    # Ano → Mês → Semana (dependentes)
    todos_anos = st.sidebar.checkbox(
        "Todos os Anos", value=False, key="cb_todos_anos"
    )
    anos_opts  = sorted(df["Ano"].unique())
    sel_anos   = anos_opts if todos_anos else st.sidebar.multiselect(
        "Ano", anos_opts, default=[ano_max], key="ms_anos"
    )

    todos_meses = st.sidebar.checkbox(
        "Todos os Meses", value=False, key="cb_todos_meses"
    )
    meses_opts  = sorted(df[df["Ano"].isin(sel_anos)]["Mes"].unique())
    sel_meses   = meses_opts if todos_meses else st.sidebar.multiselect(
        "Mês", meses_opts, default=[mes_max], key="ms_meses"
    )

    todos_semanas = st.sidebar.checkbox(
        "Todas as Semanas", value=False, key="cb_todas_semanas"
    )
    semanas_opts = sorted(
        df[(df["Ano"].isin(sel_anos)) & (df["Mes"].isin(sel_meses))]["Semana"].unique()
    )
    sel_semanas = semanas_opts if todos_semanas else st.sidebar.multiselect(
        "Semana", semanas_opts, default=[semana_max], key="ms_semanas"
    )

    # Classe Operacional
    todas_classes = st.sidebar.checkbox(
        "Todas as Classes Operacionais", value=True, key="cb_todas_classes"
    )
    classes_opts = sorted(df["Classe_Operacional"].dropna().unique())
    sel_classes  = classes_opts if todas_classes else st.sidebar.multiselect(
        "Classe Operacional", classes_opts, default=classes_opts, key="ms_classes"
    )

    # Período
    dt_min, dt_max = df["Data"].min(), df["Data"].max()
    sel_periodo = st.sidebar.date_input(
        "Período", [dt_min, dt_max], key="di_periodo"
    )

    return {
        "safras":     sel_safras,
        "anos":       sel_anos,
        "meses":      sel_meses,
        "semanas":    sel_semanas,
        "classes_op": sel_classes,
        "periodo":    sel_periodo
    }

def filtrar_dados(df: pd.DataFrame, opts: dict) -> pd.DataFrame:
    """Aplica todos os filtros no DataFrame e retorna o subset."""
    mask = (
        df["Safra"].isin(opts["safras"])
        & df["Ano"].isin(opts["anos"])
        & df["Mes"].isin(opts["meses"])
        & df["Semana"].isin(opts["semanas"])
        & df["Classe_Operacional"].isin(opts["classes_op"])
        & (df["Data"] >= pd.to_datetime(opts["periodo"][0]))
        & (df["Data"] <= pd.to_datetime(opts["periodo"][1]))
    )
    return df.loc[mask].copy()

def calcular_kpis(df: pd.DataFrame) -> dict:
    """
    Calcula KPIs principais e variação percentual em relação ao período anterior.
    """
    total_litros   = df["Qtde_Litros"].sum()
    media_consumo  = df["Media"].mean()
    eqp_unicos     = df["Cod_Equip"].nunique()

    # Período anterior (mesmo intervalo)
    inicio, fim = df["Data"].min(), df["Data"].max()
    delta = fim - inicio
    prev = df[(df["Data"] >= inicio - delta) & (df["Data"] < inicio)]
    prev_litros = prev["Qtde_Litros"].sum() or 1
    delta_pct   = (total_litros - prev_litros) / prev_litros * 100

    return {
        "total_litros":     total_litros,
        "media_consumo":    media_consumo,
        "eqp_unicos":       eqp_unicos,
        "delta_litros_pct": delta_pct
    }

# --------------- App ---------------

st.title("Dashboard de Consumo de Combustível")

# Carrega e filtra dados
df           = load_data(EXCEL_PATH, SHEET_NAME)
opcoes       = sidebar_filters(df)
df_filtrado  = filtrar_dados(df, opcoes)

# Exibe KPIs
kpis = calcular_kpis(df_filtrado)
col1, col2, col3, col4 = st.columns(4)
col1.metric("Litros Consumidos", formatar_brasileiro(kpis["total_litros"]))
col2.metric("Média de Consumo", formatar_brasileiro(kpis["media_consumo"]))
col3.metric("Equipamentos Únicos", kpis["eqp_unicos"])
col4.metric("Variação Litros (%)", f"{kpis['delta_litros_pct']:.1f}%")

st.markdown("---")

#  Gráfico: Média de Consumo por Equipamento
dados_plot = (
    df_filtrado
    .groupby(["Cod_Equip", "Descricao_Equip"])["Media"]
    .mean()
    .reset_index()
)

# Rótulo combinado e arredondamento
dados_plot["Equip_Label"] = (
    dados_plot["Cod_Equip"].astype(str)
    + " - "
    + dados_plot["Descricao_Equip"]
)
dados_plot["Media"] = dados_plot["Media"].round(1)
dados_plot = dados_plot.sort_values("Media", ascending=False)

# Cria gráfico de barras
fig = px.bar(
    dados_plot,
    x="Equip_Label",
    y="Media",
    text="Media",
    title="Média de Consumo por Equipamento",
    labels={
        "Equip_Label": "Equipamento",
        "Media": "Média de Consumo (L)"
    }
)

fig.update_traces(
    textposition="outside",
    marker=dict(line=dict(color="black", width=0.5))
)
fig.update_layout(
    xaxis_tickangle=-45,
    margin=dict(l=20, r=20, t=50, b=80),
    yaxis=dict(title="Média de Consumo (L)")
)

# Exibe no Streamlit
st.plotly_chart(fig, use_container_width=True)
st.markdown("---")

# --------------- Montagem do Dashboard ---------------

def main():
    st.set_page_config(
        page_title="Dashboard Consumo Abastecimentos",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    df   = load_data(EXCEL_PATH, SHEET_NAME)
    st.title("📊 Dashboard de Consumo de Abastecimentos")

    # 1) Filtros
    opts = sidebar_filters(df)
    df_f = filtrar_dados(df, opts)
    if df_f.empty:
        st.error("Sem dados no período/filtros selecionados.")
        st.stop()

    # 2) KPIs
    kpis       = calcular_kpis(df_f)
    c1, c2, c3 = st.columns(3)
    c1.metric(
        "Total de Litros",
        formatar_brasileiro(kpis["total_litros"]),
        f"{kpis['delta_litros_pct']:.1f}%"
    )
    c2.metric("Média de Consumo", formatar_brasileiro(kpis["media_consumo"]))
    c3.metric("Equipamentos Únicos", kpis["eqp_unicos"])

    st.markdown("---")

    # 3) Alertas
    with st.expander("🚨 Alertas de Consumo Fora do Padrão", expanded=True):
        fora = df_f[
            (df_f["Media"] < ALERTA_MIN) | (df_f["Media"] > ALERTA_MAX)
        ]
        if fora.empty:
            st.success("Nenhum consumo fora do padrão.")
        else:
            st.warning(f"{fora['Cod_Equip'].nunique()} veículos fora do padrão")
            st.dataframe(fora[["Data", "Cod_Equip", "Classe_Operacional", "Media"]])

    st.markdown("---")

    # 4.1) Média por Classe Operacional
    media_op = (
        df_f.groupby("Classe_Operacional")["Media"]
        .mean()
        .reset_index()
    )
    fig1 = px.bar(
        media_op,
        x="Classe_Operacional",
        y="Media",
        text="Media",
        title="Média de Consumo por Classe Operacional",
        labels={"Media": "km/l ou equiv."}
    )
    fig1.update_traces(texttemplate="%{text:.2f}", textposition="outside")
    fig1.update_layout(xaxis_tickangle=-45, uniformtext_mode="hide")
    st.plotly_chart(fig1, use_container_width=True)

    # 4.2) Consumo Mensal vs Média (dropdown)
    agg = (
        df_f.groupby("AnoMes")[["Qtde_Litros", "Media"]]
        .mean()
        .reset_index()
    )
    agg["AnoMes"] = agg["AnoMes"].astype(str)

    fig2 = px.bar(
        agg,
        x="AnoMes",
        y="Qtde_Litros",
        text="Qtde_Litros",
        title="Consumo Mensal / Média",
        labels={"Qtde_Litros": "Litros", "AnoMes": "Período"}
    )
    fig2.update_traces(texttemplate="%{text:.1f}", textposition="outside")
    fig2.update_layout(
        xaxis=dict(
            tickmode="array",
            tickvals=agg["AnoMes"],
            ticktext=agg["AnoMes"],
            tickangle=-45
        ),
        updatemenus=[{
            "buttons": [
                {
                    "label": "Litros",
                    "method": "update",
                    "args": [
                        {"y": ["Qtde_Litros"]},
                        {"yaxis": {"title": "Litros"}}
                    ]
                },
                {
                    "label": "Média",
                    "method": "update",
                    "args": [
                        {"y": ["Media"]},
                        {"yaxis": {"title": "Média (km/l)"}}
                    ]
                }
            ],
            "direction": "down"
        }]
    )
    st.plotly_chart(fig2, use_container_width=True)

    # 4.3) Consumo Mensal (Top 10 Equipamentos) com rótulo aprimorado e dados visíveis
    top10 = (
        df_f.groupby("Cod_Equip")["Qtde_Litros"]
        .sum()
        .nlargest(10)
        .index
    )
    
    trend = (
        df_f[df_f["Cod_Equip"].isin(top10)]
        .groupby(["Cod_Equip", "Descricao_Equip"])["Media"]
        .mean()
        .reset_index()
    )
    
    # Criar rótulo no formato "2042 - TRATOR DE PNEUS (4X4)"
    trend["Equip_Label"] = trend.apply(
        lambda row: f"{row['Cod_Equip']} - {row['Descricao_Equip']}", axis=1
    )
    
    # Arredondar para 1 casa decimal
    trend["Media"] = trend["Media"].round(1)
    
    # Criar gráfico
    fig = px.bar(
        trend,
        x="Equip_Label",
        y="Media",
        text="Media",  # Isso exibe o valor acima da barra
        title="Média de Consumo por Equipamento (Top 10)",
        labels={
            "Equip_Label": "Equipamento",
            "Media": "Média de Consumo"
        }
    )
    
    fig.update_traces(
        textposition="outside",  # Mostra os valores acima da barra
        marker=dict(line=dict(color="black", width=0.5))  # Borda sutil nas barras
    )
    
    fig.update_layout(
        xaxis_tickangle=-45,
        margin=dict(l=20, r=20, t=50, b=80),
        yaxis=dict(title="Média de Consumo (L)")
    )
    
    # Exibe no Streamlit
    st.plotly_chart(fig, use_container_width=True)
    st.markdown("---")
    # 5) Tabela detalhada com AgGrid
    with st.expander("📋 Tabela Interativa", expanded=False):
        gb = GridOptionsBuilder.from_dataframe(df_f)
        gb.configure_default_column(filterable=True, sortable=True, resizable=True)
        gb.configure_pagination(paginationAutoPageSize=True)
        AgGrid(
            df_f.drop(columns=["Descricao_Equip"]),
            gridOptions=gb.build(),
            height=400
        )

    # 6) Exportar CSV
    with st.expander("⬇️ Exportar Dados", expanded=False):
        csv = df_f.to_csv(index=False).encode("utf-8")
        st.download_button(
            "📥 Baixar CSV",
            data=csv,
            file_name="dados_filtrados.csv",
            mime="text/csv"
        )

if __name__ == "__main__":
    main()
