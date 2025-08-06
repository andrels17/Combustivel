import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# --------------- Configurações Gerais ---------------

EXCEL_PATH = "Acompto_Abast.xlsx"
SHEET_NAME = "BD"

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
    Carrega e prepara o DataFrame:
      - Lê Excel, renomeia colunas
      - Converte Data
      - Extrai mês e semana ISO
      - Cria colunas Ano, AnoMes, AnoSemana
      - Converte colunas numéricas
      - Define campo Fazenda
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

    df["Qtde_Litros"] = pd.to_numeric(df["Qtde_Litros"], errors="coerce")
    df["Media"]       = pd.to_numeric(df["Media"], errors="coerce")
    df["Media_P"]     = pd.to_numeric(df["Media_P"], errors="coerce")

    df["Fazenda"] = df["Ref1"].astype(str)

    return df

def sidebar_filters(df: pd.DataFrame) -> dict:
    """Constrói a barra lateral de filtros, garantindo keys únicas."""
    st.sidebar.header("📅 Filtros")

    ano_max    = int(df["Ano"].max())
    mes_max    = int(df[df["Ano"] == ano_max]["Mes"].max())
    semana_max = int(df[df["Ano"] == ano_max]["Semana"].max())
    safra_max  = sorted(df["Safra"].dropna().unique())[-1]

    todas_safras = st.sidebar.checkbox(
        "Todas as Safras", value=False, key="sidebar_todas_safras"
    )
    safras_opts = sorted(df["Safra"].dropna().unique())
    sel_safras = (
        safras_opts if todas_safras
        else st.sidebar.multiselect(
            "Safra", safras_opts, default=[safra_max], key="sidebar_ms_safras"
        )
    )

    todos_anos = st.sidebar.checkbox(
        "Todos os Anos", value=False, key="sidebar_todos_anos"
    )
    anos_opts = sorted(df["Ano"].unique())
    sel_anos = (
        anos_opts if todos_anos
        else st.sidebar.multiselect(
            "Ano", anos_opts, default=[ano_max], key="sidebar_ms_anos"
        )
    )

    todos_meses = st.sidebar.checkbox(
        "Todos os Meses", value=False, key="sidebar_todos_meses"
    )
    meses_opts = sorted(df[df["Ano"].isin(sel_anos)]["Mes"].unique())
    sel_meses = (
        meses_opts if todos_meses
        else st.sidebar.multiselect(
            "Mês", meses_opts, default=[mes_max], key="sidebar_ms_meses"
        )
    )

    todos_semanas = st.sidebar.checkbox(
        "Todas as Semanas", value=False, key="sidebar_todas_semanas"
    )
    semanas_opts = sorted(
        df[(df["Ano"].isin(sel_anos)) & (df["Mes"].isin(sel_meses))]["Semana"].unique()
    )
    sel_semanas = (
        semanas_opts if todos_semanas
        else st.sidebar.multiselect(
            "Semana", semanas_opts, default=[semana_max], key="sidebar_ms_semanas"
        )
    )

    todas_classes = st.sidebar.checkbox(
        "Todas as Classes Operacionais", value=True, key="sidebar_todas_classes"
    )
    classes_opts = sorted(df["Classe_Operacional"].dropna().unique())
    sel_classes = (
        classes_opts if todas_classes
        else st.sidebar.multiselect(
            "Classe Operacional",
            classes_opts,
            default=classes_opts,
            key="sidebar_ms_classes"
        )
    )

    dt_min, dt_max = df["Data"].min(), df["Data"].max()
    sel_periodo = st.sidebar.date_input(
        "Período", [dt_min, dt_max], key="sidebar_di_periodo"
    )

    return {
        "safras":     sel_safras,
        "anos":       sel_anos,
        "meses":      sel_meses,
        "semanas":    sel_semanas,
        "classes_op": sel_classes,
        "periodo":    sel_periodo,
    }

def filtrar_dados(df: pd.DataFrame, opts: dict) -> pd.DataFrame:
    """Aplica filtros e retorna subset."""
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

    # Carrega dados e aplica filtros
    df   = load_data(EXCEL_PATH, SHEET_NAME)
    opts = sidebar_filters(df)
    df_f = filtrar_dados(df, opts)
    if df_f.empty:
        st.error("Sem dados no período/filtros selecionados.")
        st.stop()

    # KPI Metrics
    kpis     = calcular_kpis(df_f)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Litros Consumidos", formatar_brasileiro(kpis["total_litros"]))
    c2.metric("Média de Consumo", formatar_brasileiro(kpis["media_consumo"]))
    c3.metric("Equipamentos Únicos", kpis["eqp_unicos"])
    c4.metric("Δ Litros (%)", f"{kpis['delta_litros_pct']:.1f}%")

    # Criação das abas
    tab1, tab2, tab3 = st.tabs(["📊 Gráficos", "📋 Tabela", "⚙️ Configurações"])

    # ─────────────── TAB 3: Padrões por Classe Operacional ───────────────
    with tab3:
        st.header("⚙️ Padrões por Classe Operacional")

        classes = sorted(df["Classe_Operacional"].dropna().unique())

        # Inicializa thresholds defaults
        if "thr" not in st.session_state:
            st.session_state.thr = {
                cls: {"min": 1.5, "max": 5.0} for cls in classes
            }

        # Inputs para cada classe
        for cls in classes:
            col_min, col_max = st.columns(2)
            with col_min:
                mn = st.number_input(
                    f"{cls} → Mínimo (km/l)",
                    min_value=0.0, max_value=100.0,
                    value=st.session_state.thr[cls]["min"],
                    step=0.1, key=f"min_{cls}"
                )
            with col_max:
                mx = st.number_input(
                    f"{cls} → Máximo (km/l)",
                    min_value=0.0, max_value=100.0,
                    value=st.session_state.thr[cls]["max"],
                    step=0.1, key=f"max_{cls}"
                )
            st.session_state.thr[cls]["min"] = mn
            st.session_state.thr[cls]["max"] = mx

    # ─────────────── TAB 1: Gráficos com thresholds por classe ───────────────
    with tab1:
        df_alerta = df_f.copy()
        df_alerta["thr_min"] = df_alerta["Classe_Operacional"].map(
            lambda c: st.session_state.thr[c]["min"]
        )
        df_alerta["thr_max"] = df_alerta["Classe_Operacional"].map(
            lambda c: st.session_state.thr[c]["max"]
        )

        df_alerta["Status"] = np.where(
            (df_alerta["Media"] >= df_alerta["thr_min"])
            & (df_alerta["Media"] <= df_alerta["thr_max"]),
            "Dentro do padrão", "Fora do padrão"
        )

        total_fora = (df_alerta["Status"] == "Fora do padrão").sum()
        st.warning(f"Total de equipamentos fora do padrão: {total_fora}")

        df_fora = (
            df_alerta.query("Status=='Fora do padrão'")
            .assign(
                Equip_Label=lambda d: (
                    d.Cod_Equip.astype(str) + " – " + d.Descricao_Equip
                )
            )
            .sort_values("Media", ascending=True)
        )

        fig_hbar = px.bar(
            df_fora,
            x="Media",
            y="Equip_Label",
            orientation="h",
            title="Consumo dos Equipamentos Fora do Padrão (km/l)",
            labels={"Media": "Consumo (km/l)", "Equip_Label": "Equipamento"}
        )
        fig_hbar.update_layout(height=600, yaxis={"automargin": True})
        st.plotly_chart(fig_hbar, use_container_width=True)

        # Média por Classe Operacional
        media_op = df_f.groupby("Classe_Operacional")["Media"].mean().reset_index()
        fig1 = px.bar(
            media_op, x="Classe_Operacional", y="Media", text="Media",
            title="Média de Consumo por Classe Operacional",
            labels={"Media": "km/l ou equiv."}
        )
        fig1.update_traces(texttemplate="%{text:.2f}", textposition="outside")
        fig1.update_layout(xaxis_tickangle=-45, uniformtext_mode="hide")
        st.plotly_chart(fig1, use_container_width=True)

        # Consumo Mensal vs Média
        agg = df_f.groupby("AnoMes")[["Qtde_Litros", "Media"]].mean().reset_index()
        agg["AnoMes"] = agg["AnoMes"].astype(str)
        fig2 = px.bar(
            agg, x="AnoMes", y="Qtde_Litros", text="Qtde_Litros",
            title="Consumo Mensal / Média",
            labels={"Qtde_Litros": "Litros", "AnoMes": "Período"}
        )
        fig2.update_traces(texttemplate="%{text:.1f}", textposition="outside")
        fig2.add_hline(
            y=agg["Qtde_Litros"].mean(),
            line_dash="dash", line_color="gray",
            annotation_text="Média Global", annotation_position="top left"
        )
        fig2.update_layout(
            xaxis=dict(tickmode="array", tickvals=agg["AnoMes"],
                       ticktext=agg["AnoMes"], tickangle=-45),
            updatemenus=[{
                "buttons": [
                    {"label": "Litros", "method": "update",
                     "args":[{"y":["Qtde_Litros"]},
                             {"yaxis":{"title":"Litros"}}]},
                    {"label": "Média", "method": "update",
                     "args":[{"y":["Media"]},
                             {"yaxis":{"title":"Média (km/l)"}}]}
                ],
                "direction":"down","showactive":True,
                "pad":{"r":10,"t":10},
                "x":0,"xanchor":"left","y":1.1,"yanchor":"top"
            }]
        )
        st.plotly_chart(fig2, use_container_width=True)

        # Top 10 Equipamentos
        top10 = df_f.groupby("Cod_Equip")["Qtde_Litros"].sum().nlargest(10).index
        trend = (
            df_f[df_f["Cod_Equip"].isin(top10)]
            .groupby(["Cod_Equip","Descricao_Equip"])["Media"]
            .mean().reset_index()
        )
        trend["Equip_Label"] = trend.apply(
            lambda r: f"{r['Cod_Equip']} - {r['Descricao_Equip']}", axis=1
        )
        trend["Media"] = trend["Media"].round(1)
        trend = trend.sort_values("Media", ascending=False)

        fig3 = px.bar(
            trend, x="Equip_Label", y="Media", text="Media",
            color_discrete_sequence=st.session_state.thr.get("palette_seq",
                                   px.colors.qualitative.Plotly),
            title="Média de Consumo por Equipamento (Top 10)",
            labels={"Equip_Label":"Equipamento","Media":"Média de Consumo (L)"}
        )
        fig3.update_traces(textposition="outside",
                           marker=dict(line=dict(color="black",width=0.5)))
        fig3.update_layout(xaxis_tickangle=-45,
                           margin=dict(l=20,r=20,t=50,b=80))
        st.plotly_chart(fig3, use_container_width=True)

        img_bytes = fig3.to_image(format="png")
        st.download_button(
            "📷 Exportar Top10 (PNG)",
            data=img_bytes,
            file_name="top10.png",
            mime="image/png",
            key="download_top10"
        )

    # ─────────────── TAB 2: Tabela Detalhada ───────────────
    with tab2:
        st.header("📋 Tabela Detalhada")

        cell_style_rules = {
            'fora_padrao': {
                'condition': (
                    f"x.value < {st.session_state.thr[c]['min']} || "
                    f"x.value > {st.session_state.thr[c]['max']}"
                ),
                'style': {'backgroundColor':'red','color':'white'}
            } for c in classes
        }

        gb = GridOptionsBuilder.from_dataframe(df_f)
        gb.configure_default_column(filterable=True, sortable=True,
                                    resizable=True)
        gb.configure_column(
            "Media",
            type=["numericColumn"], precision=1,
            cellStyleRules=cell_style_rules,
            header_name="Média (L/km)"
        )
        gb.configure_column("Qtde_Litros", type=["numericColumn"],
                            precision=1, header_name="Litros")
        gb.configure_pagination(paginationAutoPageSize=False,
                                paginationPageSize=10)
        gb.configure_selection(selection_mode="multiple",
                               use_checkbox=True, groupSelectsChildren=True)

        grid_opts     = gb.build()
        grid_response = AgGrid(
            df_f,
            gridOptions=grid_opts,
            height=400,
            allow_unsafe_jscode=True,
            update_mode=GridUpdateMode.SELECTION_CHANGED
        )

        sel_rows = grid_response["selected_rows"]
        if sel_rows:
            df_sel = pd.DataFrame(sel_rows).drop(
                "_selectedRowNodeInfo", axis=1, errors="ignore"
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

if __name__ == "__main__":
    main()
