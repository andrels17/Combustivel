# app.py

import os
from typing import Dict

import pandas as pd
import plotly.express as px
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode


# --------------- Configurações Gerais ---------------

EXCEL_PATH = "Acompto_Abast.xlsx"
SHEET_NAME = "BD"


# --------------- Utilitários de Formatação ---------------

def formatar_brasileiro(valor: float) -> str:
    """Formata número no padrão brasileiro com duas casas decimais."""
    return (
        "{:,.2f}".format(valor)
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


# --------------- Carregamento e Preparação de Dados ---------------

@st.cache_data(show_spinner=False)
def load_data(path: str, sheet: str) -> pd.DataFrame:
    """
    Lê um Excel e entrega um DataFrame pronto para análise:
      - Renomeia colunas
      - Converte Datas e extrai Ano, Mês, Semana, etc.
      - Garante tipos numéricos
      - Cria coluna 'Fazenda'
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

    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df = df[df["Data"].notna()]

    df["Mes"]       = df["Data"].dt.month
    df["Semana"]    = df["Data"].dt.isocalendar().week
    df["Ano"]       = df["Data"].dt.year
    df["AnoMes"]    = df["Data"].dt.to_period("M").astype(str)
    df["AnoSemana"] = df["Data"].dt.strftime("%Y-%U")

    for col in ["Qtde_Litros", "Media", "Media_P"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["Fazenda"] = df["Ref1"].astype(str)
    return df


# --------------- Filtragem pela Sidebar ---------------

def sidebar_filters(df: pd.DataFrame) -> Dict[str, any]:
    """
    Adiciona controles à sidebar e retorna um dict com as seleções:
     - Safra, Ano, Mês, Semana, Classe Operacional, Intervalo de Datas
    """
    st.sidebar.header("📅 Filtros")

    ano_max    = int(df["Ano"].max())
    mes_max    = int(df[df["Ano"] == ano_max]["Mes"].max())
    sem_max    = int(df[df["Ano"] == ano_max]["Semana"].max())
    safra_max  = sorted(df["Safra"].dropna().unique())[-1]

    todas_safras = st.sidebar.checkbox("Todas as Safras", False)
    safras_opts  = sorted(df["Safra"].dropna().unique())
    sel_safras   = (
        safras_opts
        if todas_safras
        else st.sidebar.multiselect("Safra", safras_opts, [safra_max])
    )

    todos_anos = st.sidebar.checkbox("Todos os Anos", False)
    anos_opts  = sorted(df["Ano"].unique())
    sel_anos   = (
        anos_opts
        if todos_anos
        else st.sidebar.multiselect("Ano", anos_opts, [ano_max])
    )

    todos_meses = st.sidebar.checkbox("Todos os Meses", False)
    meses_opts  = sorted(df[df["Ano"].isin(sel_anos)]["Mes"].unique())
    sel_meses   = (
        meses_opts
        if todos_meses
        else st.sidebar.multiselect("Mês", meses_opts, [mes_max])
    )

    todas_semanas = st.sidebar.checkbox("Todas as Semanas", False)
    semanas_opts  = sorted(
        df[(df["Ano"].isin(sel_anos)) & (df["Mes"].isin(sel_meses))]["Semana"].unique()
    )
    sel_semanas   = (
        semanas_opts
        if todas_semanas
        else st.sidebar.multiselect("Semana", semanas_opts, [sem_max])
    )

    todas_classes = st.sidebar.checkbox("Todas as Classes Operacionais", True)
    classes_opts  = sorted(df["Classe_Operacional"].dropna().unique())
    sel_classes   = (
        classes_opts
        if todas_classes
        else st.sidebar.multiselect("Classe Operacional", classes_opts, classes_opts)
    )

    dt_min, dt_max = df["Data"].min(), df["Data"].max()
    sel_periodo    = st.sidebar.date_input("Período", [dt_min, dt_max])

    return {
        "safras":     sel_safras,
        "anos":       sel_anos,
        "meses":      sel_meses,
        "semanas":    sel_semanas,
        "classes_op": sel_classes,
        "periodo":    sel_periodo,
    }


# --------------- Filtra e Calcula KPI ---------------

def filtrar_dados(df: pd.DataFrame, opts: Dict[str, any]) -> pd.DataFrame:
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


def calcular_kpis(df: pd.DataFrame) -> Dict[str, float]:
    total_litros = df["Qtde_Litros"].sum()
    media_consumo = df["Media"].mean()
    eqp_unicos    = df["Cod_Equip"].nunique()

    inicio, fim = df["Data"].min(), df["Data"].max()
    delta       = fim - inicio
    prev        = df[(df["Data"] >= inicio - delta) & (df["Data"] < inicio)]
    prev_litros = prev["Qtde_Litros"].sum() or 1
    delta_pct   = (total_litros - prev_litros) / prev_litros * 100

    return {
        "total_litros":     total_litros,
        "media_consumo":    media_consumo,
        "eqp_unicos":       eqp_unicos,
        "delta_litros_pct": delta_pct,
    }


# --------------- Função Principal ---------------

def main():
    st.set_page_config(
        page_title="Dashboard Consumo Abastecimentos",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    st.title("📊 Dashboard de Consumo de Abastecimentos")

    df   = load_data(EXCEL_PATH, SHEET_NAME)
    opts = sidebar_filters(df)
    df_f = filtrar_dados(df, opts)
    if df_f.empty:
        st.error("Sem dados no período/filtros selecionados.")
        st.stop()

    # KPI Metrics
    kpis = calcular_kpis(df_f)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Litros Consumidos",  formatar_brasileiro(kpis["total_litros"]))
    c2.metric("Média de Consumo",   formatar_brasileiro(kpis["media_consumo"]))
    c3.metric("Equipamentos Únicos", kpis["eqp_unicos"])
    c4.metric("Δ Litros (%)",       f"{kpis['delta_litros_pct']:.1f}%")

    # Abas
    tab1, tab2, tab3 = st.tabs([
        "📊 Gráficos",
        "📋 Tabela",
        "⚙️ Configurações"
    ])

    # --- Aba de Configurações ---
    with tab3:
        st.header("⚙️ Ajustes")
        alerta_min = st.number_input(
            "Limite mínimo de consumo (km/l)",
            min_value=0.0,
            max_value=100.0,
            value=1.5,
            step=0.1
        )
        alerta_max = st.number_input(
            "Limite máximo de consumo (km/l)",
            min_value=0.0,
            max_value=100.0,
            value=5.0,
            step=0.1
        )
        paletas = {
            "Plotly":  px.colors.qualitative.Plotly,
            "Viridis": px.colors.sequential.Viridis,
            "Cividis": px.colors.sequential.Cividis,
            "Inferno": px.colors.sequential.Inferno
        }
        paleta_nome = st.selectbox(
            "Paleta de cores para Top10",
            options=list(paletas.keys()),
            index=0
        )
        palette_seq = paletas[paleta_nome]

    # --- Aba de Gráficos ---
    with tab1:
        # 4.1 Média por Classe Operacional
        media_op = (
            df_f
            .groupby("Classe_Operacional")["Media"]
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

        # 4.2 Consumo Mensal vs Média
        agg = (
            df_f
            .groupby("AnoMes")[["Qtde_Litros", "Media"]]
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
        fig2.add_hline(
            y=agg["Qtde_Litros"].mean(),
            line_dash="dash",
            line_color="gray",
            annotation_text="Média Global",
            annotation_position="top left"
        )
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
                "direction": "down",
                "showactive": True,
                "pad": {"r": 10, "t": 10},
                "x": 0,
                "xanchor": "left",
                "y": 1.1,
                "yanchor": "top"
            }]
        )
        st.plotly_chart(fig2, use_container_width=True)

        # 4.3 Top 10 Equipamentos
        top10 = (
            df_f
            .groupby("Cod_Equip")["Qtde_Litros"]
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
        trend["Equip_Label"] = trend.apply(
            lambda r: f"{r['Cod_Equip']} - {r['Descricao_Equip']}", axis=1
        )
        trend["Media"] = trend["Media"].round(1)
        trend = trend.sort_values("Media", ascending=False)

        fig3 = px.bar(
            trend,
            x="Equip_Label",
            y="Media",
            text="Media",
            color_discrete_sequence=palette_seq,
            title="Média de Consumo por Equipamento (Top 10)",
            labels={"Equip_Label": "Equipamento", "Media": "Média de Consumo (L)"}
        )
        fig3.update_traces(
            textposition="outside",
            marker=dict(line=dict(color="black", width=0.5))
        )
        fig3.update_layout(
            xaxis_tickangle=-45,
            margin=dict(l=20, r=20, t=50, b=80)
        )
        st.plotly_chart(
            fig3,
            use_container_width=True,
            config={
                "modeBarButtonsToAdd": ["toImage"],
                "toImageButtonOptions": {
                    "format": "png",
                    "filename": "top10",
                    "height": 600,
                    "width": 800
                }
            }
        )

with tab2:
    st.header("📋 Tabela Detalhada")

    # Configurações do Grid
    gb = GridOptionsBuilder.from_dataframe(df_f)
    gb.configure_default_column(filterable=True, sortable=True, resizable=True)

    # cellStyle simplificado (sem quebras de linha)
    js_cond = (
        f"function(params) {{"
        f"  if (params.value < {alerta_min} || params.value > {alerta_max}) {{"
        f"    return {{color: 'white', backgroundColor: 'red'}};"
        f"  }}"
        f"  return null;"
        f"}}"
    )
    gb.configure_column(
        "Media",
        type=["numericColumn"],
        precision=1,
        cellStyle=js_cond,
        header_name="Média (L/km)"
    )
    gb.configure_column(
        "Qtde_Litros",
        type=["numericColumn"],
        precision=1,
        header_name="Litros"
    )
    gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=10)
    gb.configure_selection(
        selection_mode="multiple",
        use_checkbox=True,
        groupSelectsChildren=True
    )

    grid_opts = gb.build()

    # Renderiza com key para evitar mismatch
    try:
        grid_response = AgGrid(
            df_f,
            gridOptions=grid_opts,
            height=400,
            allow_unsafe_jscode=True,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            key="table_1"
        )
    except Exception as e:
        st.error(f"Não foi possível renderizar a tabela: {e}")
        return

    sel_rows = grid_response["selected_rows"]
    if sel_rows:
        df_sel = (
            pd.DataFrame(sel_rows)
              .drop("_selectedRowNodeInfo", axis=1, errors="ignore")
        )
        st.write(f"Linhas selecionadas: {len(df_sel)}")
        csv_sel = df_sel.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Baixar selecionadas",
            data=csv_sel,
            file_name="selecionadas.csv",
            mime="text/csv",
            key="download_selected"
        )
    main()
