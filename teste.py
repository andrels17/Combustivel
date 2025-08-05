# app.py

import streamlit as st
import pandas as pd
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder

# --------------- Configurações ---------------

EXCEL_PATH = "Acompto_Abast.xlsx"
SHEET_NAME = "BD"
ALERTA_MIN = 1.5
ALERTA_MAX = 5.0

# --------------- Utilitários ---------------

def formatar_brasileiro(valor: float) -> str:
    """Formata número no padrão brasileiro com duas casas decimais."""
    return "{:,.2f}".format(valor).replace(",", "X").replace(".", ",").replace("X", ".")

@st.cache_data(show_spinner=False)
def load_data(path: str, sheet: str) -> pd.DataFrame:
    """
    Carrega e prepara o DataFrame:
     - Lê Excel, renomeia colunas
     - Converte Data
     - Gera colunas Ano, AnoMes, AnoSemana
     - Converte numerics
     - Define Fazendas
    """
    try:
        df = pd.read_excel(path, sheet_name=sheet, skiprows=2)
    except FileNotFoundError:
        st.error(f"Arquivo não encontrado em `{path}`")
        st.stop()

    df.columns = [
        "Data", "Cod_Equip", "Descricao_Equip", "Qtde_Litros", "Km_Hs_Rod",
        "Media", "Media_P", "Perc_Media", "Ton_Cana", "Litros_Ton",
        "Ref1", "Ref2", "Unidade", "Safra", "Mes", "Semana",
        "Classe", "Classe_Operacional", "Descricao_Proprietario", "Potencia_CV"
    ]
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df = df[df["Data"].notna()]

    # Períodos
    df["Ano"] = df["Data"].dt.year
    df["AnoMes"] = df["Data"].dt.to_period("M").astype(str)
    df["AnoSemana"] = df["Data"].dt.strftime("%Y-%U")

    # Números
    df["Qtde_Litros"] = pd.to_numeric(df["Qtde_Litros"], errors="coerce")
    df["Media"] = pd.to_numeric(df["Media"], errors="coerce")
    df["Media_P"] = pd.to_numeric(df["Media_P"], errors="coerce")

    # Fazenda
    df["Fazenda"] = df["Ref1"].astype(str)
    return df


def sidebar_filters(df: pd.DataFrame) -> dict:
    """
    Constrói a barra lateral de filtros, com dependência entre eles.
    Retorna dicionário com as seleções.
    """
    st.sidebar.header("📅 Filtros")
    # valores mais recentes
    ano_max = int(df["Ano"].max())
    mes_max = int(df[df["Ano"] == ano_max]["Mes"].max())
    semana_max = sorted(df[df["Ano"] == ano_max]["Semana"].unique())[-1]
    safra_max = sorted(df["Safra"].dropna().unique())[-1]

    # Safra
    todas_safras = st.sidebar.checkbox("Todas as Safras", value=False)
    safras_opts = sorted(df["Safra"].dropna().unique())
    sel_safras = safras_opts if todas_safras else st.sidebar.multiselect(
        "Safra", safras_opts, default=[safra_max]
    )

    # Ano → Mês → Semana (dependentes)
    todos_anos = st.sidebar.checkbox("Todos os Anos", value=False)
    anos_opts = sorted(df["Ano"].unique())
    sel_anos = anos_opts if todos_anos else st.sidebar.multiselect(
        "Ano", anos_opts, default=[ano_max]
    )

    meses_opts = sorted(df[df["Ano"].isin(sel_anos)]["Mes"].unique())
    todos_meses = st.sidebar.checkbox("Todos os Meses", value=False)
    sel_meses = meses_opts if todos_meses else st.sidebar.multiselect(
        "Mês", meses_opts, default=[mes_max]
    )

    semanas_opts = sorted(df[
        (df["Ano"].isin(sel_anos)) & (df["Mes"].isin(sel_meses))
    ]["Semana"].unique())
    todos_semanas = st.sidebar.checkbox("Todas as Semanas", value=False)
    sel_semanas = semanas_opts if todos_semanas else st.sidebar.multiselect(
        "Semana", semanas_opts, default=[semana_max]
    )

    # Classe Operacional
    todas_classes = st.sidebar.checkbox("Todas as Classes Operacionais", value=True)
    classes_opts = sorted(df["Classe_Operacional"].dropna().unique())
    sel_classes = classes_opts if todas_classes else st.sidebar.multiselect(
        "Classe Operacional", classes_opts, default=classes_opts
    )

    # Período
    dt_min, dt_max = df["Data"].min(), df["Data"].max()
    sel_periodo = st.sidebar.date_input("Período", [dt_min, dt_max])

    return {
        "safras": sel_safras,
        "anos": sel_anos,
        "meses": sel_meses,
        "semanas": sel_semanas,
        "classes_op": sel_classes,
        "periodo": sel_periodo
    }


def filtrar_dados(df: pd.DataFrame, opts: dict) -> pd.DataFrame:
    """Aplica todos os filtros no DataFrame e retorna o subset."""
    mask = (
        df["Safra"].isin(opts["safras"]) &
        df["Ano"].isin(opts["anos"]) &
        df["Mes"].isin(opts["meses"]) &
        df["Semana"].isin(opts["semanas"]) &
        df["Classe_Operacional"].isin(opts["classes_op"]) &
        (df["Data"] >= pd.to_datetime(opts["periodo"][0])) &
        (df["Data"] <= pd.to_datetime(opts["periodo"][1]))
    )
    result = df.loc[mask].copy()
    return result


def calcular_kpis(df: pd.DataFrame) -> dict:
    """
    Calcula KPIs principais e variação percentual em relação ao período anterior.
    Retorna dicionário com totais, médias e deltas.
    """
    total_litros = df["Qtde_Litros"].sum()
    media_consumo = df["Media"].mean()
    eqp_unicos = df["Cod_Equip"].nunique()

    # periodo anterior (mesmo intervalo de dias imediatamente anterior)
    inicio, fim = df["Data"].min(), df["Data"].max()
    delta = fim - inicio
    prev = df[
        (df["Data"] >= inicio - delta) &
        (df["Data"] < inicio)
    ]
    prev_litros = prev["Qtde_Litros"].sum() or 1  # evita zero
    delta_pct = (total_litros - prev_litros) / prev_litros * 100

    return {
        "total_litros": total_litros,
        "media_consumo": media_consumo,
        "eqp_unicos": eqp_unicos,
        "delta_litros_pct": delta_pct
    }


# --------------- Montagem do Dashboard ---------------

def main():
    st.set_page_config(
        page_title="Dashboard Consumo Abastecimentos",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    df = load_data(EXCEL_PATH, SHEET_NAME)
    st.title("📊 Dashboard de Consumo de Abastecimentos")

    # 1) Filtros
    opts = sidebar_filters(df)
    df_f = filtrar_dados(df, opts)

    if df_f.empty:
        st.error("Sem dados no período/filtros selecionados.")
        st.stop()

    # 2) KPIs
    kpis = calcular_kpis(df_f)
    c1, c2, c3 = st.columns(3)
    c1.metric(
        "Total de Litros", formatar_brasileiro(kpis["total_litros"]),
        f"{kpis['delta_litros_pct']:.1f}%"
    )
    c2.metric("Média de Consumo", formatar_brasileiro(kpis["media_consumo"]))
    c3.metric("Equipamentos Únicos", kpis["eqp_unicos"])

    st.markdown("---")

    # 3) Alertas
    with st.expander("🚨 Alertas de Consumo Fora do Padrão", expanded=True):
        fora = df_f[(df_f["Media"] < ALERTA_MIN) | (df_f["Media"] > ALERTA_MAX)]
        if fora.empty:
            st.success("Nenhum consumo fora do padrão.")
        else:
            st.warning(f"{fora['Cod_Equip'].nunique()} veículos fora do padrão")
            st.dataframe(fora[["Data", "Cod_Equip", "Classe_Operacional", "Media"]])

    st.markdown("---")

    # 4) Gráficos Dinâmicos

    # 4.1 - Média por Classe Operacional
    media_op = df_f.groupby("Classe_Operacional")["Media"].mean().reset_index()
    fig1 = px.bar(
        media_op, x="Classe_Operacional", y="Media", text="Media",
        title="Média de Consumo por Classe Operacional",
        labels={"Media": "km/l ou equiv."}
    )
    fig1.update_traces(texttemplate="%{text:.2f}", textposition="outside")
    fig1.update_layout(xaxis_tickangle=-45, uniformtext_mode="hide")
    st.plotly_chart(fig1, use_container_width=True)

    # 4.2 - Consumo Mensal vs Média (dropdown)
    agg = df_f.groupby("AnoMes")[["Qtde_Litros", "Media"]].mean().reset_index()
    fig2 = px.bar(
        agg, x="AnoMes", y="Qtde_Litros", text="Qtde_Litros",
        title="Consumo Mensal / Média",
        labels={"Qtde_Litros": "Litros"}
    )
    fig2.update_layout(
        updatemenus=[{
            "buttons": [
                {"label": "Litros", "method": "update",
                 "args": [{"y": ["Qtde_Litros"]}, {"yaxis": {"title": "Litros"}}]},
                {"label": "Média", "method": "update",
                 "args": [{"y": ["Media"]}, {"yaxis": {"title": "Média (km/l)"}}]}
            ],
            "direction": "down"
        }]
    )
    st.plotly_chart(fig2, use_container_width=True)

    # 4.3 - Tendência Top 10 Equipamentos
    top10 = (
        df_f.groupby("Cod_Equip")["Qtde_Litros"]
        .sum().nlargest(10).index
    )
    trend = (
        df_f[df_f["Cod_Equip"].isin(top10)]
        .groupby(["AnoMes", "Cod_Equip"])["Media"]
        .mean().reset_index()
    )
    fig3 = px.line(
        trend, x="AnoMes", y="Media", color="Cod_Equip", markers=True,
        title="Tendência de Consumo (Top 10 Equip.)",
        labels={"Media": "Média de Consumo"}
    )
    fig3.update_layout(xaxis_tickangle=-45, hovermode="x unified")
    st.plotly_chart(fig3, use_container_width=True)

    st.markdown("---")

    # 5) Tabela detalhada com AgGrid
    with st.expander("📋 Tabela Interativa", expanded=False):
        gb = GridOptionsBuilder.from_dataframe(df_f)
        gb.configure_default_column(filterable=True, sortable=True, resizable=True)
        gb.configure_pagination(paginationAutoPageSize=True)
        AgGrid(df_f.drop(columns=["Descricao_Equip"]), gb.build(), height=400)

    # 6) Exportar CSV
    with st.expander("⬇️ Exportar Dados", expanded=False):
        csv = df_f.to_csv(index=False).encode("utf-8")
        st.download_button("📥 Baixar CSV", data=csv,
                           file_name="dados_filtrados.csv",
                           mime="text/csv")


if __name__ == "__main__":
    main()
