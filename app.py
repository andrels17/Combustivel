# app.py

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# --------------- Configurações Gerais ---------------
EXCEL_PATH  = "Acompto_Abast.xlsx"
SHEET_NAME  = "BD"

# --------------- Funções Utilitárias ---------------

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
    Carrega o Excel e renomeia colunas.
    Se só vierem 20 cols (sem Fazenda/Frente), cria colunas padrão.
    """
    try:
        df = pd.read_excel(path, sheet_name=sheet, skiprows=2)
    except FileNotFoundError:
        st.error(f"Arquivo não encontrado em `{path}`")
        st.stop()

    # Caso a planilha contenha apenas 20 colunas originais
    if df.shape[1] == 20:
        df.columns = [
            "Data", "Cod_Equip", "Descricao_Equip", "Qtde_Litros", "Km_Hs_Rod",
            "Media", "Media_P", "Perc_Media", "Ton_Cana", "Litros_Ton",
            "Ref1", "Ref2", "Unidade", "Safra", "Mes_Excel", "Semana_Excel",
            "Classe", "Classe_Operacional", "Descricao_Proprietario", "Potencia_CV"
        ]
        # cria colunas padrão
        df["Fazenda"] = df["Ref1"].astype(str)  # talvez código de fazenda
        df["Frente"]  = "Geral"
    else:
        # planilha já veio com Fazenda e Frente (22 colunas)
        df.columns = [
            "Data", "Cod_Equip", "Descricao_Equip", "Qtde_Litros", "Km_Hs_Rod",
            "Media", "Media_P", "Perc_Media", "Ton_Cana", "Litros_Ton",
            "Ref1", "Ref2", "Unidade", "Safra", "Mes_Excel", "Semana_Excel",
            "Classe", "Classe_Operacional", "Descricao_Proprietario",
            "Potencia_CV", "Fazenda", "Frente"
        ]

    # Conversão de datas e extração de períodos
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df = df[df["Data"].notna()]
    df["Mes"]       = df["Data"].dt.month
    df["Semana"]    = df["Data"].dt.isocalendar().week
    df["Ano"]       = df["Data"].dt.year
    df["AnoMes"]    = df["Data"].dt.to_period("M").astype(str)
    df["AnoSemana"] = df["Data"].dt.strftime("%Y-%U")

    # Conversão numérica
    df["Qtde_Litros"] = pd.to_numeric(df["Qtde_Litros"], errors="coerce")
    df["Media"]       = pd.to_numeric(df["Media"], errors="coerce")
    df["Media_P"]     = pd.to_numeric(df["Media_P"], errors="coerce")

    return df

def sidebar_filters(df: pd.DataFrame) -> dict:
    """Cria filtros na sidebar, incluindo fazenda e frente."""
    st.sidebar.header("📅 Filtros")

    # Fazenda
    faz_opts = sorted(df["Fazenda"].dropna().unique())
    sel_faz  = st.sidebar.multiselect(
        "Fazenda", faz_opts, default=faz_opts, key="ms_fazendas"
    )

    # Frente (depende de Fazenda)
    frente_opts = sorted(df[df["Fazenda"].isin(sel_faz)]["Frente"].dropna().unique())
    sel_frente  = st.sidebar.multiselect(
        "Frente", frente_opts, default=frente_opts, key="ms_frentes"
    )

    # Safra
    safra_opts  = sorted(df["Safra"].dropna().unique())
    todas_safras = st.sidebar.checkbox("Todas as Safras", False, key="cb_safras")
    sel_safras = (
        safra_opts if todas_safras
        else st.sidebar.multiselect("Safra", safra_opts,
                                    default=[max(safra_opts)], key="ms_safras")
    )

    # Ano
    anos_opts   = sorted(df["Ano"].unique())
    todos_anos  = st.sidebar.checkbox("Todos os Anos", False, key="cb_anos")
    sel_anos    = (
        anos_opts if todos_anos
        else st.sidebar.multiselect("Ano", anos_opts,
                                    default=[max(anos_opts)], key="ms_anos")
    )

    # Mês
    meses_opts  = sorted(df[df["Ano"].isin(sel_anos)]["Mes"].unique())
    todos_meses = st.sidebar.checkbox("Todos os Meses", False, key="cb_meses")
    sel_meses   = (
        meses_opts if todos_meses
        else st.sidebar.multiselect("Mês", meses_opts,
                                    default=[max(meses_opts)], key="ms_meses")
    )

    # Semana
    semanas_opts  = sorted(
        df[(df["Ano"].isin(sel_anos)) & (df["Mes"].isin(sel_meses))]["Semana"].unique()
    )
    todos_semanas = st.sidebar.checkbox("Todas as Semanas", False, key="cb_semanas")
    sel_semanas   = (
        semanas_opts if todos_semanas
        else st.sidebar.multiselect("Semana", semanas_opts,
                                    default=[max(semanas_opts)], key="ms_semanas")
    )

    # Classe Operacional
    classes_opts  = sorted(df["Classe_Operacional"].dropna().unique())
    todas_classes = st.sidebar.checkbox("Todas as Classes", True, key="cb_classes")
    sel_classes   = (
        classes_opts if todas_classes
        else st.sidebar.multiselect("Classe Operacional",
                                    classes_opts, default=classes_opts, key="ms_classes")
    )

    # Período de datas
    dt_min, dt_max = df["Data"].min(), df["Data"].max()
    sel_periodo    = st.sidebar.date_input(
        "Período", [dt_min, dt_max], key="ms_periodo"
    )

    return {
        "fazendas":   sel_faz,
        "frentes":    sel_frente,
        "safras":     sel_safras,
        "anos":       sel_anos,
        "meses":      sel_meses,
        "semanas":    sel_semanas,
        "classes_op": sel_classes,
        "periodo":    sel_periodo,
    }

def filtrar_dados(df: pd.DataFrame, opts: dict) -> pd.DataFrame:
    """Aplica todos os filtros e retorna o DataFrame resultante."""
    mask = (
        df["Fazenda"].isin(opts["fazendas"])
        & df["Frente"].isin(opts["frentes"])
        & df["Safra"].isin(opts["safras"])
        & df["Ano"].isin(opts["anos"])
        & df["Mes"].isin(opts["meses"])
        & df["Semana"].isin(opts["semanas"])
        & df["Classe_Operacional"].isin(opts["classes_op"])
        & (df["Data"] >= pd.to_datetime(opts["periodo"][0]))
        & (df["Data"] <= pd.to_datetime(opts["periodo"][1]))
    )
    return df.loc[mask].copy()

def calcular_kpis(df: pd.DataFrame) -> dict:
    """Calcula KPIs e variação percentual no período anterior."""
    total_litros   = df["Qtde_Litros"].sum()
    media_consumo  = df["Media"].mean()
    eqp_unicos     = df["Cod_Equip"].nunique()

    inicio, fim = df["Data"].min(), df["Data"].max()
    delta       = fim - inicio
    prev        = df[(df["Data"] >= inicio - delta) & (df["Data"] < inicio)]
    prev_litros = prev["Qtde_Litros"].sum() or 1
    delta_pct   = (total_litros - prev_litros) / prev_litros * 100

    return {
        "total_litros":     total_litros,
        "media_consumo":    media_consumo,
        "eqp_unicos":       eqp_unicos,
        "delta_litros_pct": delta_pct
    }

# --------------- Função Principal ---------------

def main():
    st.set_page_config(
        page_title="Dashboard Consumo Abastecimentos",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    st.title("📊 Dashboard de Consumo de Abastecimentos")

    # Carrega e filtra dados
    df   = load_data(EXCEL_PATH, SHEET_NAME)
    opts = sidebar_filters(df)
    df_f = filtrar_dados(df, opts)

    if df_f.empty:
        st.error("Sem dados no período/filtros selecionados.")
        st.stop()

    # KPIs
    kpis = calcular_kpis(df_f)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Litros Consumidos", formatar_brasileiro(kpis["total_litros"]))
    c2.metric("Média de Consumo", formatar_brasileiro(kpis["media_consumo"]))
    c3.metric("Equipamentos Únicos", kpis["eqp_unicos"])
    c4.metric("Δ Litros (%)", f"{kpis['delta_litros_pct']:.1f}%")

    # Abas
    tab1, tab2, tab3 = st.tabs(["📊 Gráficos", "📋 Tabela", "⚙️ Configurações"])

    with tab3:
        st.warning("Edite os thresholds por Classe Operacional no painel anterior.")

    with tab1:
        # Exemplo de gráfico: Média por classe, já respeitando os filtros de fazenda/frente
        media_op = df_f.groupby("Classe_Operacional")["Media"].mean().reset_index()
        fig = px.bar(
            media_op,
            x="Classe_Operacional",
            y="Media",
            text="Media",
            title="Média de Consumo por Classe Operacional",
            labels={"Media": "km/l"}
        )
        fig.update_traces(texttemplate="%{text:.2f}", textposition="outside")
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig, use_container_width=True)

    with tab2:
        st.header("📋 Tabela Detalhada")
        gb = GridOptionsBuilder.from_dataframe(df_f)
        gb.configure_default_column(filterable=True, sortable=True, resizable=True)
        gb.configure_column("Media", type=["numericColumn"], precision=1,
                            header_name="Média (L/km)")
        gb.configure_column("Qtde_Litros", type=["numericColumn"], precision=1,
                            header_name="Litros")
        gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=10)
        gb.configure_selection(selection_mode="multiple",
                               use_checkbox=True, groupSelectsChildren=True)
        grid_opts     = gb.build()
        grid_response = AgGrid(df_f, gridOptions=grid_opts,
                               height=400, allow_unsafe_jscode=True,
                               update_mode=GridUpdateMode.SELECTION_CHANGED)
        sel_rows = grid_response["selected_rows"]
        if sel_rows:
            df_sel  = pd.DataFrame(sel_rows).drop("_selectedRowNodeInfo", axis=1)
            st.write(f"Linhas selecionadas: {len(df_sel)}")
            csv_sel = df_sel.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Baixar selecionadas",
                               data=csv_sel,
                               file_name="selecionadas.csv",
                               mime="text/csv",
                               key="download_selected")

if __name__ == "__main__":
    main()
