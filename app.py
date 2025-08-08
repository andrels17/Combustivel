import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder
from streamlit_plotly_events import plotly_events
from datetime import datetime
from fpdf import FPDF
import io

# --------------- Configurações Gerais ---------------
EXCEL_PATH = "Acompto_Abast.xlsx"

# --------------- Funções Utilitárias ---------------

def formatar_brasileiro(valor: float) -> str:
    if pd.isna(valor) or not np.isfinite(valor):
        return "–"
    return (
        "{:,.2f}".format(valor)
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )

@st.cache_data(show_spinner="Carregando e processando dados...")
def load_data(path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    try:
        df_abast = pd.read_excel(path, sheet_name="BD", skiprows=2)
        df_frotas = pd.read_excel(path, sheet_name="FROTAS", skiprows=1)
    except FileNotFoundError:
        st.error(f"Arquivo não encontrado em `{path}`")
        st.stop()
    except ValueError:
        st.error("Verifique se as planilhas 'BD' e 'FROTAS' existem no arquivo.")
        st.stop()

    df_frotas = (
        df_frotas
        .rename(columns={"COD_EQUIPAMENTO": "Cod_Equip"})
        .drop_duplicates(subset=["Cod_Equip"])
    )
    df_frotas["ANOMODELO"] = pd.to_numeric(df_frotas["ANOMODELO"], errors="coerce")

    df_abast.columns = [
        "Data", "Cod_Equip", "Descricao_Equip", "Qtde_Litros", "Km_Hs_Rod",
        "Media", "Media_P", "Perc_Media", "Ton_Cana", "Litros_Ton",
        "Ref1", "Ref2", "Unidade", "Safra", "Mes_Excel", "Semana_Excel",
        "Classe_Original", "Classe_Operacional", "Descricao_Proprietario_Original",
        "Potencia_CV_Abast"
    ]
    df = pd.merge(df_abast, df_frotas, on="Cod_Equip", how="left")
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df.dropna(subset=["Data"], inplace=True)

    df["Mes"] = df["Data"].dt.month
    df["Semana"] = df["Data"].dt.isocalendar().week
    df["Ano"] = df["Data"].dt.year
    df["AnoMes"] = df["Data"].dt.to_period("M").astype(str)

    for col in ["Qtde_Litros", "Media", "Media_P"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["DESCRICAOMARCA"] = df["Ref2"].astype(str)
    df["Fazenda"] = df["Ref1"].astype(str)

    return df, df_frotas

def sidebar_filters(df: pd.DataFrame) -> dict:
    st.sidebar.header("📅 Filtros de Consumo")
    ano_max = int(df["Ano"].max())
    mes_max = int(df[df["Ano"] == ano_max]["Mes"].max())
    safra_max = sorted(df["Safra"].dropna().unique())[-1]

    todas_safras = st.sidebar.checkbox("Todas as Safras", False)
    safra_opts = sorted(df["Safra"].dropna().unique())
    sel_safras = safra_opts if todas_safras else st.sidebar.multiselect(
        "Safra", safra_opts, default=[safra_max]
    )

    todos_anos = st.sidebar.checkbox("Todos os Anos", False)
    anos_opts = sorted(df["Ano"].unique())
    sel_anos = anos_opts if todos_anos else st.sidebar.multiselect(
        "Ano", anos_opts, default=[ano_max]
    )

    todos_meses = st.sidebar.checkbox("Todos os Meses", False)
    meses_opts = sorted(df[df["Ano"].isin(sel_anos)]["Mes"].unique())
    sel_meses = meses_opts if todos_meses else st.sidebar.multiselect(
        "Mês", meses_opts, default=[mes_max]
    )

    st.sidebar.markdown("---")
    todas_marcas = st.sidebar.checkbox("Todas as Marcas", True)
    marcas_opts = sorted(df["DESCRICAOMARCA"].dropna().unique())
    sel_marcas = marcas_opts if todas_marcas else st.sidebar.multiselect(
        "Marca", marcas_opts, default=marcas_opts
    )

    todas_classes = st.sidebar.checkbox("Todas as Classes", True)
    classes_opts = sorted(df["Classe_Operacional"].dropna().unique())
    sel_classes = classes_opts if todas_classes else st.sidebar.multiselect(
        "Classe Operacional", classes_opts, default=classes_opts
    )

    return {
        "safras": sel_safras,
        "anos": sel_anos,
        "meses": sel_meses,
        "marcas": sel_marcas,
        "classes_op": sel_classes,
    }

def filtrar_dados(df: pd.DataFrame, opts: dict) -> pd.DataFrame:
    mask = (
        df["Safra"].isin(opts["safras"]) &
        df["Ano"].isin(opts["anos"]) &
        df["Mes"].isin(opts["meses"]) &
        df["DESCRICAOMARCA"].isin(opts["marcas"]) &
        df["Classe_Operacional"].isin(opts["classes_op"])
    )
    return df.loc[mask].copy()

def calcular_kpis_consumo(df: pd.DataFrame) -> dict:
    return {
        "total_litros": df["Qtde_Litros"].sum(),
        "media_consumo": df["Media"].mean(),
        "eqp_unicos": df["Cod_Equip"].nunique(),
    }

def generate_pdf(images: list[bytes]) -> bytes:
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    for img in images:
        pdf.add_page()
        pdf.image(io.BytesIO(img), x=10, y=20, w=190)
    return pdf.output(dest="S").encode("latin1")

def main():
    st.set_page_config(page_title="Dashboard de Frotas e Abastecimentos", layout="wide")
    st.title("📊 Dashboard de Frotas e Abastecimentos")

    df, df_frotas = load_data(EXCEL_PATH)
    tab_principal, tab_consulta, tab_tabela, tab_config = st.tabs([
        "📊 Análise de Consumo",
        "🔎 Consulta de Frota",
        "📋 Tabela Detalhada",
        "⚙️ Configurações"
    ])

    # --- ABA 1: Análise de Consumo ---
    with tab_principal:
        opts = sidebar_filters(df)
        df_f = filtrar_dados(df, opts)
        if df_f.empty:
            st.error("Sem dados para os filtros selecionados.")
            st.stop()

        # KPIs principais
        kpis = calcular_kpis_consumo(df_f)
        total_eq = df_frotas.shape[0]
        ativos = df_frotas.query("ATIVO == 'ATIVO'").shape[0]
        idade_media = datetime.now().year - df_frotas["ANOMODELO"].median()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Litros Consumidos", formatar_brasileiro(kpis["total_litros"]))
        c2.metric("Média de Consumo", formatar_brasileiro(kpis["media_consumo"]))
        c3.metric("Veículos Ativos", f"{ativos} / {total_eq}")
        c4.metric("Idade Média da Frota", f"{idade_media:.0f} anos")

        # 1) Média por Classe Operacional
        media_op = df_f.groupby("Classe_Operacional")["Media"].mean().reset_index()
        media_op["Media"] = media_op["Media"].round(1)
        fig1 = px.bar(
            media_op, x="Classe_Operacional", y="Media", text="Media",
            title="Média de Consumo por Classe Operacional",
            labels={"Media": "Média (km/l)", "Classe_Operacional": "Classe"}
        )
        fig1.update_traces(textposition="outside")
        st.plotly_chart(fig1, use_container_width=True)

        # 2) Consumo Mensal
        agg = df_f.groupby("AnoMes")["Qtde_Litros"].mean().reset_index()
        agg["Mes"] = pd.to_datetime(agg["AnoMes"] + "-01").dt.strftime("%b %Y")
        agg["Qtde_Litros"] = agg["Qtde_Litros"].round(1)
        fig2 = px.bar(
            agg, x="Mes", y="Qtde_Litros", text="Qtde_Litros",
            title="Consumo Mensal",
            labels={"Qtde_Litros": "Litros", "Mes": "Mês"}
        )
        fig2.update_traces(texttemplate="%{text:.1f}", textposition="outside")
        fig2.update_layout(xaxis_tickangle=-45, height=450)
        st.plotly_chart(fig2, use_container_width=True)

        # 3) Top 10 Equipamentos por Consumo Médio
        top10 = df_f.groupby("Cod_Equip")["Qtde_Litros"].sum().nlargest(10).index
        trend = (
            df_f[df_f["Cod_Equip"].isin(top10)]
            .groupby(["Cod_Equip", "Descricao_Equip"])["Media"].mean()
            .reset_index()
            .sort_values("Media", ascending=False)
        )
        trend["Equip_Label"] = trend.apply(
            lambda r: f"{r['Cod_Equip']} - {r['Descricao_Equip']}", axis=1
        )
        trend["Media"] = trend["Media"].round(1)
        fig3 = px.bar(
            trend, x="Equip_Label", y="Media", text="Media",
            title="Média de Consumo por Equipamento (Top 10)",
            labels={"Equip_Label": "Equipamento", "Media": "Média (km/l)"}
        )
        fig3.update_traces(
            textposition="outside",
            marker=dict(line=dict(color="black", width=0.5))
        )
        fig3.update_layout(xaxis_tickangle=-45, margin=dict(l=20, r=20, t=50, b=80))
        st.plotly_chart(fig3, use_container_width=True)

        # 4) Consumo Acumulado por Safra
        st.header("📈 Comparativo de Consumo Acumulado por Safra")
        safras = sorted(df["Safra"].dropna().unique())
        sel_safras = st.multiselect(
            "Selecione safras", safras,
            default=safras[-2:] if len(safras) > 1 else safras
        )
        if sel_safras:
            df_cmp = df[df["Safra"].isin(sel_safras)].copy()
            iniciais = df_cmp.groupby("Safra")["Data"].min().to_dict()
            df_cmp["Dias_Uteis"] = (
                df_cmp["Data"] - df_cmp["Safra"].map(iniciais)
            ).dt.days + 1
            df_cmp = (
                df_cmp
                .groupby(["Safra", "Dias_Uteis"])["Qtde_Litros"].sum()
                .groupby(level=0).cumsum()
                .reset_index()
            )
            fig_acum = px.line(
                df_cmp, x="Dias_Uteis", y="Qtde_Litros", color="Safra",
                markers=True,
                labels={
                    "Dias_Uteis": "Dias desde início da safra",
                    "Qtde_Litros": "Consumo acumulado (L)"
                },
                title="Consumo Acumulado por Safra"
            )
            ultima = sel_safras[-1]
            df_u = df_cmp[df_cmp["Safra"] == ultima]
            if not df_u.empty:
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

        # ────────────────────────────
        # Gráficos adicionais sugeridos
        # 5) Consumo Mensal YoY
        df_month = df_f.set_index("Data").resample("M")["Qtde_Litros"].sum().to_frame()
        df_month["Ano"] = df_month.index.year
        df_month["Mes"] = df_month.index.strftime("%b")
        pivot = df_month.pivot(index="Mes", columns="Ano", values="Qtde_Litros")
        # ordenar meses
        meses_ord = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
        pivot = pivot.reindex(meses_ord).dropna(how="all")
        fig_yoy = px.bar(
            pivot,
            title="Consumo Mensal YoY",
            labels={"value": "Litros", "Mes": "Mês"},
            barmode="group"
        )
        st.plotly_chart(fig_yoy, use_container_width=True)

        # 6) Projeção Próximos 3 Meses (Média Móvel)
        mov3 = df_month["Qtde_Litros"].rolling(3).mean()
        last_val = mov3.dropna().iloc[-1]
        proj_idx = pd.date_range(df_month.index.max()+pd.offsets.MonthBegin(), periods=3, freq="M")
        proj = pd.Series([last_val]*3, index=proj_idx)
        df_proj = pd.concat([df_month["Qtde_Litros"].rename("Consumo"), proj.rename("Consumo")])
        fig_proj = px.line(
            df_proj,
            title="Consumo Real + Projeção Próximos 3 Meses",
            labels={"index": "Data", "Consumo": "Litros"}
        )
        st.plotly_chart(fig_proj, use_container_width=True)

        # 7) Pareto: Top 20% dos Equipamentos
        cons_eq = df_f.groupby("Cod_Equip")["Qtde_Litros"].sum().sort_values(ascending=False)
        cumperc = cons_eq.cumsum() / cons_eq.sum()
        df_pareto = pd.DataFrame({"Consumo": cons_eq, "Acumulado": cumperc})
        fig_par = px.bar(
            df_pareto,
            x=df_pareto.index.astype(str),
            y="Consumo",
            title="Pareto: Top Equipamentos por Consumo",
            labels={"x": "Equipamento", "Consumo": "Litros"}
        )
        fig_par.add_scatter(
            x=df_pareto.index.astype(str),
            y=df_pareto["Acumulado"] * cons_eq.max(),
            mode="lines+markers",
            name="Acumulado",
            yaxis="y2"
        )
        fig_par.update_layout(
            yaxis2=dict(
                overlaying="y", side="right",
                title="Acumulado (%)",
                tickvals=[0, cons_eq.max()],
                ticktext=["0%", "100%"]
            )
        )
        st.plotly_chart(fig_par, use_container_width=True)

        # 8) Drill‐down por Classe Operacional
        fig1_dd = px.bar(
            media_op, x="Classe_Operacional", y="Media",
            title="Clique na Classe para Detalhar",
            labels={"Media": "Média (km/l)", "Classe_Operacional": "Classe"},
            hover_data=["Media"]
        )
        sel = plotly_events(fig1_dd, click_event=True, key="drill_cls")
        st.plotly_chart(fig1_dd, use_container_width=True)
        if sel:
            idx = sel[0]["pointIndex"]
            cls = media_op.loc[idx, "Classe_Operacional"]
            st.subheader(f"Detalhamento: Classe {cls}")
            st.dataframe(df_f[df_f["Classe_Operacional"] == cls])

        # 9) Exportar Relatório em PDF
        imgs = [
            fig1.to_image(format="png"),
            fig2.to_image(format="png"),
            fig3.to_image(format="png"),
            fig_acum.to_image(format="png"),
            fig_yoy.to_image(format="png"),
            fig_proj.to_image(format="png"),
            fig_par.to_image(format="png")
        ]
        pdf_bytes = generate_pdf(imgs)
        st.download_button(
            "📄 Baixar Relatório PDF",
            pdf_bytes,
            file_name="dashboard_frotas.pdf",
            mime="application/pdf"
        )

    # --- ABA 2: Consulta de Frota ---
    with tab_consulta:
        st.header("🔎 Ficha Individual do Equipamento")
        df_frotas["label"] = (
            df_frotas["Cod_Equip"].astype(str)
            + " - "
            + df_frotas["DESCRICAO_EQUIPAMENTO"].fillna("")
            + " ("
            + df_frotas["PLACA"].fillna("Sem Placa")
            + ")"
        )
        equip = st.selectbox(
            "Selecione o Equipamento",
            df_frotas.sort_values("Cod_Equip")["label"]
        )
        if equip:
            cod = int(equip.split(" - ")[0])
            dados_eq = df_frotas.query("Cod_Equip == @cod").iloc[0]
            cons_eq = df.query("Cod_Equip == @cod").sort_values("Data", ascending=False)

            st.subheader(f"{dados_eq['DESCRICAO_EQUIPAMENTO']} ({dados_eq['PLACA']})")
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Status", dados_eq["ATIVO"])
            col2.metric("Placa", dados_eq["PLACA"])
            col3.metric("Média Geral", formatar_brasileiro(cons_eq["Media"].mean()))
            col4.metric("Total Consumido (L)", formatar_brasileiro(cons_eq["Qtde_Litros"].sum()))

            if not cons_eq.empty:
                ult = cons_eq.iloc[0]["Km_Hs_Rod"]
                km_display = str(int(ult)) if pd.notna(ult) else "–"
                safra_ult = cons_eq["Safra"].max()
                df_saf = cons_eq[cons_eq["Safra"] == safra_ult]
                tot_saf = df_saf["Qtde_Litros"].sum()
                med_saf = df_saf["Media"].mean()
            else:
                km_display = "–"
                safra_ult = None
                tot_saf = med_saf = None

            col5, col6, col7 = st.columns(3)
            col5.metric("KM/Hr Último Registro", km_display)
            col6.metric(
                f"Total Última Safra{f' ({safra_ult})' if safra_ult else ''}",
                formatar_brasileiro(tot_saf) if tot_saf is not None else "–"
            )
            col7.metric(
                "Média Última Safra",
                formatar_brasileiro(med_saf) if med_saf is not None else "–"
            )

            st.markdown("---")
            st.subheader("Informações Cadastrais")
            st.dataframe(dados_eq.drop("label").to_frame("Valor"), use_container_width=True)

    # --- ABA 3: Tabela Detalhada ---
    with tab_tabela:
        st.header("📋 Tabela Detalhada de Abastecimentos")
        cols = [
            "Data", "Cod_Equip", "Descricao_Equip", "PLACA",
            "DESCRICAOMARCA", "ANOMODELO", "Qtde_Litros",
            "Media", "Media_P", "Classe_Operacional"
        ]
        df_tab = df_f[[c for c in cols if c in df_f.columns]]
        gb = GridOptionsBuilder.from_dataframe(df_tab)
        gb.configure_default_column(filterable=True, sortable=True, resizable=True)
        gb.configure_column("Media", type=["numericColumn"], precision=1)
        gb.configure_column("Qtde_Litros", type=["numericColumn"], precision=1)
        gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=15)
        gb.configure_selection("multiple", use_checkbox=True)
        AgGrid(df_tab, gridOptions=gb.build(), height=500, allow_unsafe_jscode=True)

    # --- ABA 4: Configurações ---
    with tab_config:
        st.header("⚙️ Padrões por Classe Operacional (Alertas)")
        if "thr" not in st.session_state:
            classes = df["Classe_Operacional"].dropna().unique()
            st.session_state.thr = {cls: {"min": 1.5, "max": 5.0} for cls in classes}
        for cls in sorted(st.session_state.thr.keys()):
            c_min, c_max = st.columns(2)
            mn = c_min.number_input(
                f"{cls} → Mínimo (km/l)",
                min_value=0.0, max_value=100.0,
                value=st.session_state.thr[cls]["min"],
                step=0.1, key=f"min_{cls}"
            )
            mx = c_max.number_input(
                f"{cls} → Máximo (km/l)",
                min_value=0.0, max_value=100.0,
                value=st.session_state.thr[cls]["max"],
                step=0.1, key=f"max_{cls}"
            )
            st.session_state.thr[cls]["min"] = mn
            st.session_state.thr[cls]["max"] = mx

if __name__ == "__main__":
    main()
