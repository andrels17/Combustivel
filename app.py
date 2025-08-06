import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# --------------- Configurações Gerais ---------------

EXCEL_PATH = "Acompto_Abast.xlsx"
SHEET_NAME  = "BD"

# --------------- Funções Utilitárias ---------------

def formatar_brasileiro(valor: float) -> str:
    """Formata número no padrão brasileiro com duas casas decimais."""
    if pd.isna(valor):
        return "–"
    return (
        "{:,.2f}".format(valor)
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )

@st.cache_data(show_spinner=False)
def load_data(path: str, sheet: str) -> pd.DataFrame:
    """Carrega e prepara o DataFrame."""
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
    st.sidebar.header("📅 Filtros")
    ano_max    = int(df["Ano"].max())
    mes_max    = int(df[df["Ano"] == ano_max]["Mes"].max())
    semana_max = int(df[df["Ano"] == ano_max]["Semana"].max())
    safra_max  = sorted(df["Safra"].dropna().unique())[-1]

    todas_safras = st.sidebar.checkbox("Todas as Safras", False, key="todas_safras")
    safra_opts   = sorted(df["Safra"].dropna().unique())
    sel_safras   = safra_opts if todas_safras else st.sidebar.multiselect(
        "Safra", safra_opts, default=[safra_max], key="ms_safras"
    )

    todos_anos = st.sidebar.checkbox("Todos os Anos", False, key="todos_anos")
    anos_opts  = sorted(df["Ano"].unique())
    sel_anos   = anos_opts if todos_anos else st.sidebar.multiselect(
        "Ano", anos_opts, default=[ano_max], key="ms_anos"
    )

    todos_meses = st.sidebar.checkbox("Todos os Meses", False, key="todos_meses")
    meses_opts  = sorted(df[df["Ano"].isin(sel_anos)]["Mes"].unique())
    sel_meses   = meses_opts if todos_meses else st.sidebar.multiselect(
        "Mês", meses_opts, default=[mes_max], key="ms_meses"
    )

    todas_semanas = st.sidebar.checkbox("Todas as Semanas", False, key="todas_semanas")
    semanas_opts  = sorted(
        df[(df["Ano"].isin(sel_anos)) & (df["Mes"].isin(sel_meses))]["Semana"].unique()
    )
    sel_semanas   = semanas_opts if todas_semanas else st.sidebar.multiselect(
        "Semana", semanas_opts, default=[semana_max], key="ms_semanas"
    )

    todas_classes = st.sidebar.checkbox("Todas as Classes", True, key="todas_classes")
    classes_opts  = sorted(df["Classe_Operacional"].dropna().unique())
    sel_classes   = classes_opts if todas_classes else st.sidebar.multiselect(
        "Classe Operacional", classes_opts,
        default=classes_opts, key="ms_classes"
    )

    dt_min, dt_max = df["Data"].min(), df["Data"].max()
    sel_periodo    = st.sidebar.date_input("Período", [dt_min, dt_max], key="periodo")

    return {
        "safras":     sel_safras,
        "anos":       sel_anos,
        "meses":      sel_meses,
        "semanas":    sel_semanas,
        "classes_op": sel_classes,
        "periodo":    sel_periodo,
    }

def filtrar_dados(df: pd.DataFrame, opts: dict) -> pd.DataFrame:
    mask = (
        df["Safra"].isin(opts["safras"]) &
        df["Ano"].isin(opts["anos"]) &
        df["Mes"].isin(opts["meses"]) &
        df["Semana"].isin(opts["semanas"]) &
        df["Classe_Operacional"].isin(opts["classes_op"]) &
        (df["Data"] >= pd.to_datetime(opts["periodo"][0])) &
        (df["Data"] <= pd.to_datetime(opts["periodo"][1]))
    )
    return df.loc[mask].copy()

def calcular_kpis(df: pd.DataFrame) -> dict:
    total_litros  = df["Qtde_Litros"].sum()
    media_consumo = df["Media"].mean()
    eqp_unicos    = df["Cod_Equip"].nunique()

    inicio, fim = df["Data"].min(), df["Data"].max()
    delta       = fim - inicio
    prev        = df[(df["Data"] >= inicio - delta) & (df["Data"] < inicio)]
    prev_sum    = prev["Qtde_Litros"].sum()
    delta_pct   = None if prev_sum == 0 else (total_litros - prev_sum) / prev_sum * 100
    diff_litros = total_litros - prev_sum

    return {
        "total_litros":    total_litros,
        "media_consumo":   media_consumo,
        "eqp_unicos":      eqp_unicos,
        "delta_litros_pct": delta_pct,
        "diff_litros":     diff_litros
    }

def main():
    st.set_page_config(
        page_title="Dashboard Consumo Abastecimentos",
        layout="wide"
    )
    st.title("📊 Dashboard de Consumo de Abastecimentos")

    df   = load_data(EXCEL_PATH, SHEET_NAME)
    opts = sidebar_filters(df)
    df_f = filtrar_dados(df, opts)
    if df_f.empty:
        st.error("Sem dados no período/filtros selecionados.")
        st.stop()

    kpis     = calcular_kpis(df_f)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric(
        "Litros Consumidos",
        formatar_brasileiro(kpis["total_litros"]),
        delta=f"{kpis['diff_litros']:.0f} L"
    )
    c2.metric(
        "Média de Consumo",
        formatar_brasileiro(kpis["media_consumo"])
    )
    c3.metric("Equipamentos Únicos", kpis["eqp_unicos"])
    pct = f"{kpis['delta_litros_pct']:.1f}%" if kpis["delta_litros_pct"] is not None else "–"
    c4.metric("Δ Litros (%)", pct)

    tab1, tab2, tab3 = st.tabs([
        "📊 Gráficos",
        "📋 Tabela",
        "⚙️ Configurações"
    ])

    # TAB 1: gráficos e comparativo
    with tab1:
        # thresholds e alertas
        if "thr" not in st.session_state:
            classes = df["Classe_Operacional"].dropna().unique()
            st.session_state.thr = {
                cls: {"min": 1.5, "max": 5.0} for cls in classes
            }
        thr_df = pd.DataFrame.from_dict(
            st.session_state.thr, orient="index"
        ).rename_axis("Classe_Operacional").reset_index()
        df_alerta = df_f.merge(thr_df, on="Classe_Operacional", how="left")
        df_alerta["Status"] = np.where(
            (df_alerta["Media"] >= df_alerta["min"]) &
            (df_alerta["Media"] <= df_alerta["max"]),
            "Dentro do padrão", "Fora do padrão"
        )
        total_fora = (df_alerta["Status"] == "Fora do padrão").sum()
        st.warning(f"Equipamentos fora do padrão: {total_fora}")

        # 1) Gráfico barras: média por classe
        media_op = df_f.groupby("Classe_Operacional")["Media"].mean().reset_index()
        media_op["Media"] = media_op["Media"].round(1)
        fig1 = px.bar(
            media_op,
            x="Classe_Operacional",
            y="Media",
            text="Media",
            title="Média de Consumo por Classe Operacional",
            labels={"Media": "Média (km/l)", "Classe_Operacional": "Classe"}
        )
        fig1.update_traces(texttemplate="%{text:.1f}", textposition="outside")
        fig1.update_layout(xaxis_tickangle=-45, height=500)
        st.plotly_chart(fig1, use_container_width=True)

        # 2) Gráfico barras: consumo mensal
        agg = df_f.groupby("AnoMes")["Qtde_Litros"].mean().reset_index()
        # converte "YYYY-MM" em nome do mês
        agg["Mes"] = pd.to_datetime(agg["AnoMes"] + "-01").dt.strftime("%b %Y")
        agg["Qtde_Litros"] = agg["Qtde_Litros"].round(1)
        fig2 = px.bar(
            agg,
            x="Mes",
            y="Qtde_Litros",
            text="Qtde_Litros",
            title="Consumo Mensal",
            labels={"Qtde_Litros": "Litros", "Mes": "Mês"}
        )
        fig2.update_traces(texttemplate="%{text:.1f}", textposition="outside")
        fig2.update_layout(xaxis_tickangle=-45, height=450)
        st.plotly_chart(fig2, use_container_width=True)

        # 3) Comparativo consumo acumulado por safra
        st.header("📈 Comparativo de Consumo Acumulado por Safra")
        safras_disp = sorted(df["Safra"].dropna().unique())
        sel_safras = st.multiselect(
            "Selecione safras", safras_disp,
            default=safras_disp[-2:] if len(safras_disp) > 1 else safras_disp
        )
        if sel_safras:
            df_cmp = df[df["Safra"].isin(sel_safras)].copy()
            iniciais = df_cmp.groupby("Safra")["Data"].min().to_dict()
            df_cmp["Dias_Uteis"] = (
                df_cmp["Data"] - df_cmp["Safra"].map(iniciais)
            ).dt.days + 1

            df_cmp = (
                df_cmp
                .groupby(["Safra", "Dias_Uteis"])["Qtde_Litros"]
                .sum()
                .groupby(level=0)
                .cumsum()
                .reset_index()
            )

            fig_acum = px.line(
                df_cmp,
                x="Dias_Uteis",
                y="Qtde_Litros",
                color="Safra",
                markers=True,
                labels={
                    "Dias_Uteis": "Dia desde início da safra",
                    "Qtde_Litros": "Consumo acumulado (L)"
                },
                title="Consumo Acumulado por Safra"
            )
            ultima = sel_safras[-1]
            df_u = df_cmp[df_cmp["Safra"] == ultima]
            fig_acum.add_scatter(
                x=[df_u["Dias_Uteis"].max()],
                y=[df_u["Qtde_Litros"].max()],
                mode="markers+text",
                text=[f"Hoje: {formatar_brasileiro(df_u['Qtde_Litros'].max())} L"],
                textposition="top right",
                marker=dict(size=8, color="black"),
                showlegend=False
            )
            st.plotly_chart(fig_acum, use_container_width=True)

    # TAB 2: tabela
    with tab2:
        st.header("📋 Tabela Detalhada")
        classes = df_f["Classe_Operacional"].dropna().unique()
        cell_rules = {}
        for cls in classes:
            mn = st.session_state.thr[cls]["min"]
            mx = st.session_state.thr[cls]["max"]
            cell_rules[cls] = {
                'condition': f"x.value < {mn} || x.value > {mx}",
                'style': {'backgroundColor': 'red', 'color': 'white'}
            }

        gb = GridOptionsBuilder.from_dataframe(df_f)
        gb.configure_default_column(filterable=True, sortable=True, resizable=True)
        gb.configure_column("Media", type=["numericColumn"], precision=1,
                            cellStyleRules=cell_rules, header_name="Média (km/l)")
        gb.configure_column("Qtde_Litros", type=["numericColumn"], precision=1,
                            header_name="Litros")
        gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=10)
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

        sel = grid_response["selected_rows"]
        if sel:
            df_sel = pd.DataFrame(sel).drop("_selectedRowNodeInfo", axis=1, errors="ignore")
            st.write(f"Linhas selecionadas: {len(df_sel)}")
            csv = df_sel.to_csv(index=False).encode("utf-8")
            st.download_button(
                "⬇️ Baixar selecionadas",
                data=csv, file_name="selecionadas.csv", mime="text/csv"
            )

    # TAB 3: configuração de thresholds
    with tab3:
        st.header("⚙️ Padrões por Classe Operacional")
        classes = sorted(df["Classe_Operacional"].dropna().unique())
        for cls in classes:
            c_min, c_max = st.columns(2)
            with c_min:
                mn = st.number_input(
                    f"{cls} → Mínimo (km/l)", min_value=0.0, max_value=100.0,
                    value=st.session_state.thr[cls]["min"], step=0.1, key=f"min_{cls}"
                )
            with c_max:
                mx = st.number_input(
                    f"{cls} → Máximo (km/l)", min_value=0.0, max_value=100.0,
                    value=st.session_state.thr[cls]["max"], step=0.1, key=f"max_{cls}"
                )
            st.session_state.thr[cls]["min"] = mn
            st.session_state.thr[cls]["max"] = mx

if __name__ == "__main__":
    main()
