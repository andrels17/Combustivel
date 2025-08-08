import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder
from datetime import datetime
import textwrap

# ---------------- Configurações ----------------
EXCEL_PATH = "Acompto_Abast.xlsx"

# Paletas (clara e escura)
PALETTE_LIGHT = px.colors.sequential.Blues_r
PALETTE_DARK = px.colors.sequential.Plasma_r

# ---------------- Utilitários ----------------
def formatar_brasileiro(valor: float) -> str:
    """Formata número no padrão brasileiro com duas casas decimais."""
    if pd.isna(valor) or not np.isfinite(valor):
        return "–"
    return "{:,.2f}".format(valor).replace(",", "X").replace(".", ",").replace("X", ".")

def wrap_labels(s: str, width: int = 18) -> str:
    """Quebra um rótulo em múltiplas linhas usando <br> para Plotly.
    width: número aproximado de caracteres por linha antes de quebrar."""
    if pd.isna(s):
        return ""
    parts = textwrap.wrap(str(s), width=width)
    return "<br>".join(parts) if parts else str(s)

@st.cache_data(show_spinner="Carregando e processando dados...")
def load_data(path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Carrega e prepara DataFrames (Abastecimento e Frotas)."""
    try:
        df_abast = pd.read_excel(path, sheet_name="BD", skiprows=2)
        df_frotas = pd.read_excel(path, sheet_name="FROTAS", skiprows=1)
    except FileNotFoundError:
        st.error(f"Arquivo não encontrado em `{path}`")
        st.stop()
    except ValueError as e:
        if "Sheet name" in str(e):
            st.error("Verifique se as planilhas 'BD' e 'FROTAS' existem no arquivo.")
            st.stop()
        else:
            raise

    # Normaliza frotas
    df_frotas = df_frotas.rename(columns={"COD_EQUIPAMENTO": "Cod_Equip"}).drop_duplicates(subset=["Cod_Equip"])
    df_frotas["ANOMODELO"] = pd.to_numeric(df_frotas.get("ANOMODELO", pd.Series()), errors="coerce")
    df_frotas["label"] = (
        df_frotas["Cod_Equip"].astype(str)
        + " - "
        + df_frotas["DESCRICAO_EQUIPAMENTO"].fillna("")
        + " ("
        + df_frotas["PLACA"].fillna("Sem Placa")
        + ")"
    )

    # Normaliza abastecimento (mantendo nomes originais)
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

    # Campos de tempo/derivados
    df["Mes"] = df["Data"].dt.month
    df["Semana"] = df["Data"].dt.isocalendar().week
    df["Ano"] = df["Data"].dt.year
    df["AnoMes"] = df["Data"].dt.to_period("M").astype(str)
    df["AnoSemana"] = df["Data"].dt.strftime("%Y-%U")

    # Numéricos
    for col in ["Qtde_Litros", "Media", "Media_P", "Km_Hs_Rod"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Marca / Fazenda (mantém coluna, mas não será usada em filtros)
    df["DESCRICAOMARCA"] = df["Ref2"].astype(str)
    df["Fazenda"] = df["Ref1"].astype(str)

    # --- Cálculo seguro de Consumo km/l ---
    # Usa Km_Hs_Rod (km) e Qtde_Litros (litros). Se Qtde_Litros <= 0, coloca NaN.
    if "Km_Hs_Rod" in df.columns and "Qtde_Litros" in df.columns:
        df["Consumo_km_l"] = np.where(df["Qtde_Litros"] > 0, df["Km_Hs_Rod"] / df["Qtde_Litros"], np.nan)
        # Sobrescreve 'Media' para manter compatibilidade com o restante do código
        df["Media"] = df["Consumo_km_l"]
    else:
        df["Consumo_km_l"] = np.nan

    return df, df_frotas

@st.cache_data
def filtrar_dados(df: pd.DataFrame, opts: dict) -> pd.DataFrame:
    """Filtra o DataFrame conforme opções selecionadas.
    OBS: filtro de Marca foi removido conforme solicitado."""
    mask = (
        df["Safra"].isin(opts["safras"]) &
        df["Ano"].isin(opts["anos"]) &
        df["Mes"].isin(opts["meses"]) &
        df["Classe_Operacional"].isin(opts["classes_op"])
    )
    return df.loc[mask].copy()

@st.cache_data
def calcular_kpis_consumo(df: pd.DataFrame) -> dict:
    """Calcula KPIs principais (total, média, equipamentos únicos)."""
    return {
        "total_litros": float(df["Qtde_Litros"].sum()) if "Qtde_Litros" in df.columns else 0.0,
        "media_consumo": float(df["Media"].mean()) if "Media" in df.columns else 0.0,
        "eqp_unicos": int(df["Cod_Equip"].nunique()) if "Cod_Equip" in df.columns else 0,
    }

def make_bar(fig_df, x, y, title, labels, palette, rotate_x=-45, ticksize=10, height=None, hoverfmt=None, wrap_width=18):
    """Helper para criar barras padronizadas com hovertemplate e rótulos de X legíveis.
    - rotate_x: ângulo dos ticks (ex: -45)
    - ticksize: tamanho da fonte dos ticks
    - wrap_width: se labels são longos, quebra após esse nº de caracteres
    """
    # Se formos usar labels longas (equipamentos/classe), aplicamos wrap
    df_local = fig_df.copy()
    if x in df_local.columns:
        df_local[x] = df_local[x].astype(str).apply(lambda s: wrap_labels(s, width=wrap_width))

    fig = px.bar(df_local, x=x, y=y, text=y, title=title, labels=labels, color_discrete_sequence=palette)
    # texto das barras menor para não sobrepor
    fig.update_traces(textposition="outside", texttemplate="%{text:.1f}", textfont=dict(size=10))
    # layout do eixo X para legibilidade
    fig.update_layout(
        xaxis=dict(tickangle=rotate_x, tickfont=dict(size=ticksize), automargin=True),
        margin=dict(l=40, r=20, t=60, b=140),  # aumenta bottom para espaço de rótulos
        title=dict(x=0.01, xanchor="left"),
        font=dict(size=13)
    )
    if height:
        fig.update_layout(height=height)
    # hovertemplate customizado
    if hoverfmt:
        fig.update_traces(hovertemplate=hoverfmt)
    else:
        fig.update_traces(hovertemplate=None)
    return fig

# ---------------- Layout / CSS moderno ----------------
def apply_modern_css(dark: bool):
    """Aplica CSS leve para um visual mais moderno."""
    # notas: CSS inline para melhorar títulos e KPIs
    bg = "#0e1117" if dark else "#FFFFFF"
    card_bg = "#111318" if dark else "#f8f9fa"
    text_color = "#f0f0f0" if dark else "#111111"
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-color: {bg};
            color: {text_color};
        }}
        .kpi-card {{
            background: {card_bg};
            padding: 12px;
            border-radius: 10px;
            box-shadow: 0 2px 6px rgba(0,0,0,0.08);
        }}
        .kpi-title {{ font-size:14px; color: {text_color}; opacity:0.9 }}
        .kpi-value {{ font-size:20px; font-weight:700; color: {text_color} }}
        .section-title {{ font-size:18px; font-weight:700; color: {text_color} }}
        .small-muted {{ color: #8a8a8a; font-size:12px; }}
        </style>
        """,
        unsafe_allow_html=True
    )

# ---------------- App principal ----------------
def main():
    st.set_page_config(page_title="Dashboard de Frotas e Abastecimentos", layout="wide")
    st.title("📊 Dashboard de Frotas e Abastecimentos — Visual Moderno (Light Premium)")

    # Carrega dados
    df, df_frotas = load_data(EXCEL_PATH)

    # Sidebar: tema e filtros
    with st.sidebar:
        st.header("Configurações")
        dark_mode = st.checkbox("🕶️ Dark Mode (aplica visual escuro)", value=False)
        st.markdown("---")
        st.header("📅 Filtros")
        # Limpar filtros
        if st.button("🔄 Limpar Filtros"):
            st.session_state.clear()
            st.rerun()

    # Aplica CSS leve
    apply_modern_css(dark_mode)

    # Paleta ativa
    palette = PALETTE_DARK if dark_mode else PALETTE_LIGHT
    plotly_template = "plotly_dark" if dark_mode else "plotly"

    # Filtro (função reorganizada para UX moderno) - **sem filtro de marca**
    def sidebar_filters_local(df: pd.DataFrame) -> dict:
        # Defaults defensivos
        safra_opts = sorted(df["Safra"].dropna().unique()) if "Safra" in df.columns else []
        ano_opts = sorted(df["Ano"].dropna().unique()) if "Ano" in df.columns else []
        classe_opts = sorted(df["Classe_Operacional"].dropna().unique()) if "Classe_Operacional" in df.columns else []

        # Seletores
        sel_safras = st.sidebar.multiselect("Safra", safra_opts, default=safra_opts[-1:] if safra_opts else [])
        sel_anos = st.sidebar.multiselect("Ano", ano_opts, default=ano_opts[-1:] if ano_opts else [])
        sel_meses = st.sidebar.multiselect("Mês (num)", sorted(df["Mes"].dropna().unique()) if "Mes" in df.columns else [], default=[datetime.now().month])
        st.sidebar.markdown("---")
        sel_classes = st.sidebar.multiselect("Classe Operacional", classe_opts, default=classe_opts)

        # Garantias: se vazio, usa todos
        if not sel_safras:
            sel_safras = safra_opts
        if not sel_anos:
            sel_anos = ano_opts
        if not sel_meses:
            sel_meses = sorted(df["Mes"].dropna().unique()) if "Mes" in df.columns else []
        if not sel_classes:
            sel_classes = classe_opts

        return {
            "safras": sel_safras or [],
            "anos": sel_anos or [],
            "meses": sel_meses or [],
            "classes_op": sel_classes or [],
        }

    opts = sidebar_filters_local(df)
    df_f = filtrar_dados(df, opts)

    # Abas
    tab_principal, tab_consulta, tab_tabela, tab_config = st.tabs([
        "📊 Análise de Consumo",
        "🔎 Consulta de Frota",
        "📋 Tabela Detalhada",
        "⚙️ Configurações"
    ])

    # ----- Aba Principal -----
    with tab_principal:
        # Se não tem dados
        if df_f.empty:
            st.warning("Sem dados para os filtros selecionados.")
            st.stop()

        # Cabeçalho com KPIs destacados
        kpis = calcular_kpis_consumo(df_f)
        total_eq = df_frotas.shape[0]
        ativos = int(df_frotas.query("ATIVO == 'ATIVO'").shape[0]) if "ATIVO" in df_frotas.columns else 0
        idade_media = (datetime.now().year - df_frotas["ANOMODELO"].median()) if "ANOMODELO" in df_frotas.columns else 0

        # KPI cards (light premium style)
        k1, k2, k3, k4 = st.columns([1.6,1.6,1.4,1.4])
        with k1:
            st.markdown(
                '<div class="kpi-card"><div class="kpi-title">Litros Consumidos</div>'
                f'<div class="kpi-value">{formatar_brasileiro(kpis["total_litros"])}</div></div>',
                unsafe_allow_html=True
            )
        with k2:
            st.markdown(
                '<div class="kpi-card"><div class="kpi-title">Média de Consumo</div>'
                f'<div class="kpi-value">{formatar_brasileiro(kpis["media_consumo"])} km/l</div></div>',
                unsafe_allow_html=True
            )
        with k3:
            st.markdown(
                '<div class="kpi-card"><div class="kpi-title">Veículos Ativos</div>'
                f'<div class="kpi-value">{ativos} / {total_eq}</div></div>',
                unsafe_allow_html=True
            )
        with k4:
            st.markdown(
                '<div class="kpi-card"><div class="kpi-title">Idade Média da Frota</div>'
                f'<div class="kpi-value">{idade_media:.0f} anos</div></div>',
                unsafe_allow_html=True
            )

        st.markdown("### ")
        st.info(f"🔍 {len(df_f):,} registros após aplicação dos filtros")

        # Gráfico 1 - Média por Classe Operacional (usando Media = km/l)
        media_op = df_f.groupby("Classe_Operacional")["Media"].mean().reset_index()
        media_op["Media"] = media_op["Media"].round(1)
        # aplica wrap nas labels para a legibilidade do eixo X
        media_op["Classe_wrapped"] = media_op["Classe_Operacional"].astype(str).apply(lambda s: wrap_labels(s, width=18))
        hover_template_media = "Classe: %{x}<br>Média: %{y:.1f} km/l<extra></extra>"
        fig1 = make_bar(media_op, "Classe_wrapped", "Media",
                        "Média de Consumo por Classe Operacional",
                        {"Media": "Média (km/l)", "Classe_wrapped": "Classe"},
                        palette, rotate_x=-45, ticksize=10, height=520, hoverfmt=hover_template_media, wrap_width=18)
        fig1.update_traces(marker_line_width=0.3)
        fig1.update_layout(template=plotly_template)
        st.plotly_chart(fig1, use_container_width=True, theme=None)

        # Gráfico 2 - Consumo Mensal
        agg = df_f.groupby("AnoMes")["Qtde_Litros"].mean().reset_index()
        agg["Mes"] = pd.to_datetime(agg["AnoMes"] + "-01").dt.strftime("%b %Y")
        agg["Qtde_Litros"] = agg["Qtde_Litros"].round(1)
        hover_template_month = "Mês: %{x}<br>Litros: %{y:.1f} L<extra></extra>"
        fig2 = make_bar(agg, "Mes", "Qtde_Litros", "Consumo Mensal", {"Qtde_Litros": "Litros", "Mes": "Mês"}, palette, rotate_x=-45, ticksize=10, height=420, hoverfmt=hover_template_month)
        fig2.update_layout(template=plotly_template)
        st.plotly_chart(fig2, use_container_width=True, theme=None)

        # Gráfico 3 - Top 10 Equipamentos por consumo total (mostra média de consumo)
        top10 = df_f.groupby("Cod_Equip")["Qtde_Litros"].sum().nlargest(10).index
        trend = (
            df_f[df_f["Cod_Equip"].isin(top10)]
            .groupby(["Cod_Equip", "Descricao_Equip"])["Media"].mean()
            .reset_index()
            .sort_values("Media", ascending=False)
        )
        if not trend.empty:
            # cria label amigável e wrapped
            trend["Equip_Label"] = trend.apply(lambda r: f"{r['Cod_Equip']} - {r['Descricao_Equip']}", axis=1)
            trend["Equip_Label_wrapped"] = trend["Equip_Label"].apply(lambda s: wrap_labels(s, width=18))
            trend["Media"] = trend["Media"].round(1)
            hover_template_top = "Equipamento: %{x}<br>Média: %{y:.1f} km/l<extra></extra>"
            fig3 = make_bar(trend, "Equip_Label_wrapped", "Media", "Média de Consumo por Equipamento (Top 10)",
                            {"Equip_Label_wrapped": "Equipamento", "Media": "Média (km/l)"},
                            palette, rotate_x=-45, ticksize=10, height=420, hoverfmt=hover_template_top, wrap_width=18)
            fig3.update_traces(marker_line=dict(color="#000000", width=0.5))
            fig3.update_layout(template=plotly_template)
            st.plotly_chart(fig3, use_container_width=True, theme=None)

            # Download do gráfico (quando suportado)
            @st.cache_data(show_spinner=False)
            def get_fig_png(fig):
                return fig.to_image(format="png", scale=2)

            try:
                img = get_fig_png(fig3)
                st.download_button("📷 Exportar Top10 (PNG)", data=img, file_name="top10.png", mime="image/png")
            except Exception:
                # ambientes sem kaleido/plotly image export não quebram a execução
                st.caption("Exportação de imagem não disponível no ambiente atual.")

        # Gráfico 4 - Consumo acumulado por safra (usa Qtde_Litros)
        st.markdown("---")
        st.header("📈 Comparativo de Consumo Acumulado por Safra")
        safras = sorted(df["Safra"].dropna().unique())
        sel_safras = st.multiselect("Selecione safras", safras, default=safras[-2:] if len(safras)>1 else safras)
        if sel_safras:
            df_cmp = df[df["Safra"].isin(sel_safras)].copy()
            iniciais = df_cmp.groupby("Safra")["Data"].min().to_dict()
            df_cmp["Dias_Uteis"] = (df_cmp["Data"] - df_cmp["Safra"].map(iniciais)).dt.days + 1
            df_cmp = df_cmp.groupby(["Safra", "Dias_Uteis"])["Qtde_Litros"].sum().groupby(level=0).cumsum().reset_index()
            hover_template_acum = "Dia: %{x}<br>Acumulado: %{y:.0f} L<extra></extra>"
            fig_acum = px.line(df_cmp, x="Dias_Uteis", y="Qtde_Litros", color="Safra", markers=True,
                               labels={"Dias_Uteis":"Dias desde início da safra","Qtde_Litros":"Consumo acumulado (L)"},
                               color_discrete_sequence=palette)
            fig_acum.update_layout(title="Consumo Acumulado por Safra", margin=dict(l=20,r=20,t=50,b=50), template=plotly_template, font=dict(size=13))
            fig_acum.update_traces(hovertemplate=hover_template_acum)
            # marcar valor atual da última safra selecionada
            ultima = sel_safras[-1]
            df_u = df_cmp[df_cmp["Safra"] == ultima]
            if not df_u.empty:
                fig_acum.add_scatter(x=[df_u["Dias_Uteis"].max()], y=[df_u["Qtde_Litros"].max()],
                                     mode="markers+text", text=[f"Hoje: {formatar_brasileiro(df_u['Qtde_Litros'].max())} L"],
                                     textposition="top right", marker=dict(size=8, color="#000000"), showlegend=False)
            st.plotly_chart(fig_acum, use_container_width=True, theme=None)

    # ----- Aba Consulta de Frota -----
    with tab_consulta:
        st.header("🔎 Ficha Individual do Equipamento")
        # usa label pré-calculada
        equip_label = st.selectbox("Selecione o Equipamento", options=df_frotas.sort_values("Cod_Equip")["label"])
        if equip_label:
            cod_sel = int(equip_label.split(" - ")[0])
            dados_eq = df_frotas.query("Cod_Equip == @cod_sel").iloc[0]
            consumo_eq = df.query("Cod_Equip == @cod_sel").sort_values("Data", ascending=False)

            st.subheader(f"{dados_eq['DESCRICAO_EQUIPAMENTO']} ({dados_eq.get('PLACA','–')})")
            # KPIs por equipamento
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Status", dados_eq.get("ATIVO", "–"))
            col2.metric("Placa", dados_eq.get("PLACA", "–"))
            col3.metric("Média Geral", formatar_brasileiro(consumo_eq["Media"].mean()))
            col4.metric("Total Consumido (L)", formatar_brasileiro(consumo_eq["Qtde_Litros"].sum()))

            # Último registro e última safra
            if not consumo_eq.empty:
                ultimo = consumo_eq.iloc[0]
                km_hs = ultimo.get("Km_Hs_Rod", np.nan)
                km_hs_display = str(int(km_hs)) if pd.notna(km_hs) else "–"
                safra_ult = consumo_eq["Safra"].max()
                df_safra = consumo_eq[consumo_eq["Safra"] == safra_ult]
                total_ult_safra = df_safra["Qtde_Litros"].sum()
                media_ult_safra = df_safra["Media"].mean()
            else:
                km_hs_display = "–"
                safra_ult = None
                total_ult_safra = None
                media_ult_safra = None

            c5, c6, c7 = st.columns(3)
            c5.metric("KM/Hr Último Registro", km_hs_display)
            c6.metric(f"Total Última Safra{f' ({safra_ult})' if safra_ult else ''}", formatar_brasileiro(total_ult_safra) if total_ult_safra is not None else "–")
            c7.metric("Média Última Safra", formatar_brasileiro(media_ult_safra) if media_ult_safra is not None else "–")

            st.markdown("---")
            st.subheader("Informações Cadastrais")
            st.dataframe(dados_eq.drop("label").to_frame("Valor"), use_container_width=True)

    # ----- Aba Tabela Detalhada -----
    with tab_tabela:
        st.header("📋 Tabela Detalhada de Abastecimentos")
        cols = ["Data", "Cod_Equip", "Descricao_Equip", "PLACA", "DESCRICAOMARCA", "ANOMODELO", "Qtde_Litros", "Media", "Media_P", "Classe_Operacional"]
        df_tab = df[[c for c in cols if c in df.columns]]

        # Export CSV rápido
        csv_bytes = df_tab.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Exportar CSV da Tabela", csv_bytes, "abastecimentos.csv", "text/csv")

        # AgGrid com barra lateral
        gb = GridOptionsBuilder.from_dataframe(df_tab)
        gb.configure_default_column(filterable=True, sortable=True, resizable=True)
        if "Media" in df_tab.columns:
            gb.configure_column("Media", type=["numericColumn"], precision=1)
        if "Qtde_Litros" in df_tab.columns:
            gb.configure_column("Qtde_Litros", type=["numericColumn"], precision=1)
        gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=15)
        gb.configure_selection("multiple", use_checkbox=True)
        gb.configure_side_bar()
        grid_options = gb.build()
        AgGrid(df_tab, gridOptions=grid_options, height=520, allow_unsafe_jscode=True)

    # ----- Aba Configurações -----
    with tab_config:
        st.header("⚙️ Padrões por Classe Operacional (Alertas)")
        if "thr" not in st.session_state:
            classes = df["Classe_Operacional"].dropna().unique() if "Classe_Operacional" in df.columns else []
            st.session_state.thr = {cls: {"min": 1.5, "max": 5.0} for cls in classes}

        for cls in sorted(st.session_state.thr.keys()):
            c_min, c_max = st.columns(2)
            mn = c_min.number_input(f"{cls} → Mínimo (km/l)", min_value=0.0, max_value=100.0, value=st.session_state.thr[cls]["min"], step=0.1, key=f"min_{cls}")
            mx = c_max.number_input(f"{cls} → Máximo (km/l)", min_value=0.0, max_value=100.0, value=st.session_state.thr[cls]["max"], step=0.1, key=f"max_{cls}")
            st.session_state.thr[cls]["min"] = mn
            st.session_state.thr[cls]["max"] = mx

if __name__ == "__main__":
    main()
