import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from datetime import datetime

# --------------- Configurações Gerais ---------------
EXCEL_PATH = "Acompto_Abast.xlsx"

# --------------- Funções Utilitárias ---------------

def formatar_brasileiro(valor: float) -> str:
    """Formata número no padrão brasileiro com duas casas decimais."""
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
    """Carrega, mescla e prepara os DataFrames a partir do mesmo arquivo Excel."""
    try:
        df_abastecimento = pd.read_excel(path, sheet_name="BD", skiprows=2)
        df_frotas_completo = pd.read_excel(path, sheet_name="FROTAS", skiprows=1)
    except FileNotFoundError:
        st.error(f"Arquivo não encontrado em `{path}`")
        st.stop()
    except ValueError as e:
        if "Sheet name" in str(e):
            st.error(f"Erro ao carregar o arquivo. Verifique se as planilhas 'BD' e 'FROTAS' existem em `{path}`.")
            st.stop()
        else:
            raise e

    # Prepara o DataFrame de Frotas
    df_frotas_completo = df_frotas_completo.rename(columns={"COD_EQUIPAMENTO": "Cod_Equip"})
    df_frotas_completo.drop_duplicates(subset=["Cod_Equip"], inplace=True)
    df_frotas_completo['ANOMODELO'] = pd.to_numeric(df_frotas_completo['ANOMODELO'], errors='coerce')


    # Prepara o DataFrame de Abastecimento
    df_abastecimento.columns = [
        "Data", "Cod_Equip", "Descricao_Equip", "Qtde_Litros", "Km_Hs_Rod",
        "Media", "Media_P", "Perc_Media", "Ton_Cana", "Litros_Ton",
        "Ref1", "Ref2", "Unidade", "Safra", "Mes_Excel", "Semana_Excel",
        "Classe_Original", "Classe_Operacional_Original", "Descricao_Proprietario_Original", "Potencia_CV_Abast"
    ]

    # Mescla os dois DataFrames
    df = pd.merge(df_abastecimento, df_frotas_completo, on="Cod_Equip", how="left")

    # Limpeza e preparação dos dados já mesclados
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df.dropna(subset=["Data"], inplace=True)

    df["Mes"] = df["Data"].dt.month
    df["Semana"] = df["Data"].dt.isocalendar().week
    df["Ano"] = df["Data"].dt.year
    df["AnoMes"] = df["Data"].dt.to_period("M").astype(str)
    df["AnoSemana"] = df["Data"].dt.strftime("%Y-%U")

    numeric_cols = ["Qtde_Litros", "Media", "Media_P"]
    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["Fazenda"] = df["Ref1"].astype(str)
    
    return df, df_frotas_completo


def sidebar_filters(df: pd.DataFrame) -> dict:
    st.sidebar.header("📅 Filtros de Consumo")
    ano_max = int(df["Ano"].max())
    mes_max = int(df[df["Ano"] == ano_max]["Mes"].max())
    semana_max = int(df[df["Ano"] == ano_max]["Semana"].max())
    safra_max = sorted(df["Safra"].dropna().unique())[-1]

    todas_safras = st.sidebar.checkbox("Todas as Safras", False, key="todas_safras")
    safra_opts = sorted(df["Safra"].dropna().unique())
    sel_safras = safra_opts if todas_safras else st.sidebar.multiselect(
        "Safra", safra_opts, default=[safra_max], key="ms_safras"
    )

    todos_anos = st.sidebar.checkbox("Todos os Anos", False, key="todos_anos")
    anos_opts = sorted(df["Ano"].unique())
    sel_anos = anos_opts if todos_anos else st.sidebar.multiselect(
        "Ano", anos_opts, default=[ano_max], key="ms_anos"
    )

    todos_meses = st.sidebar.checkbox("Todos os Meses", False, key="todos_meses")
    meses_opts = sorted(df[df["Ano"].isin(sel_anos)]["Mes"].unique())
    sel_meses = meses_opts if todos_meses else st.sidebar.multiselect(
        "Mês", meses_opts, default=[mes_max], key="ms_meses"
    )

    # --- NOVO FILTRO POR MARCA ---
    st.sidebar.markdown("---")
    todas_marcas = st.sidebar.checkbox("Todas as Marcas", True, key="todas_marcas")
    marcas_opts = sorted(df["DESCRICAOMARCA"].dropna().unique())
    sel_marcas = marcas_opts if todas_marcas else st.sidebar.multiselect(
        "Marca", marcas_opts, default=marcas_opts, key="ms_marcas"
    )

    todas_classes = st.sidebar.checkbox("Todas as Classes", True, key="todas_classes")
    classes_opts = sorted(df["Classe Operacional"].dropna().unique())
    sel_classes = classes_opts if todas_classes else st.sidebar.multiselect(
        "Classe Operacional", classes_opts,
        default=classes_opts, key="ms_classes"
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
        df["Classe Operacional"].isin(opts["classes_op"])
    )
    return df.loc[mask].copy()


def calcular_kpis_consumo(df: pd.DataFrame) -> dict:
    total_litros = df["Qtde_Litros"].sum()
    media_consumo = df["Media"].mean()
    eqp_unicos = df["Cod_Equip"].nunique()
    
    return {
        "total_litros": total_litros,
        "media_consumo": media_consumo,
        "eqp_unicos": eqp_unicos,
    }

def main():
    st.set_page_config(
        page_title="Dashboard de Frotas e Abastecimentos",
        layout="wide"
    )
    st.title("📊 Dashboard de Frotas e Abastecimentos")

    df, df_frotas_completo = load_data(EXCEL_PATH)

    # --- Novas abas ---
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

        kpis_consumo = calcular_kpis_consumo(df_f)
        
        # --- Novos KPIs de Frota ---
        total_veiculos = len(df_frotas_completo)
        veiculos_ativos = df_frotas_completo[df_frotas_completo['ATIVO'] == 'ATIVO'].shape[0]
        idade_media = datetime.now().year - df_frotas_completo['ANOMODELO'].median()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Litros Consumidos", formatar_brasileiro(kpis_consumo["total_litros"]))
        c2.metric("Média de Consumo", formatar_brasileiro(kpis_consumo["media_consumo"]))
        c3.metric("Veículos Ativos", f"{veiculos_ativos} / {total_veiculos}")
        c4.metric("Idade Média da Frota", f"{idade_media:.0f} anos")
        
        # ... (O resto dos gráficos da tab1 permanecem aqui) ...
        media_op = df_f.groupby("Classe Operacional")["Media"].mean().reset_index()
        media_op["Media"] = media_op["Media"].round(1)
        fig1 = px.bar(
            media_op, x="Classe Operacional", y="Media", text="Media",
            title="Média de Consumo por Classe Operacional",
            labels={"Media": "Média (km/l)", "Classe Operacional": "Classe"}
        )
        st.plotly_chart(fig1, use_container_width=True)


    # --- ABA 2: Consulta de Frota (NOVO) ---
    with tab_consulta:
        st.header("🔎 Ficha Individual do Equipamento")

        # Criar uma lista amigável para seleção
        df_frotas_completo['label'] = (
            df_frotas_completo['Cod_Equip'].astype(str) + " - " + 
            df_frotas_completo['DESCRICAO_EQUIPAMENTO'].fillna('') + " (" + 
            df_frotas_completo['PLACA'].fillna('Sem Placa') + ")"
        )
        
        equip_selecionado_label = st.selectbox(
            "Selecione o Equipamento",
            options=df_frotas_completo.sort_values(by='Cod_Equip')['label']
        )
        
        if equip_selecionado_label:
            # Extrair o código do equipamento da label selecionada
            cod_equip_selecionado = int(equip_selecionado_label.split(" - ")[0])
            
            # Filtrar dados para o equipamento selecionado
            dados_equipamento = df_frotas_completo[df_frotas_completo['Cod_Equip'] == cod_equip_selecionado].iloc[0]
            consumo_equipamento = df[df['Cod_Equip'] == cod_equip_selecionado]
            
            st.subheader(f"Detalhes de: {dados_equipamento['DESCRICAO_EQUIPAMENTO']}")

            # KPIs específicos do equipamento
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Status", dados_equipamento['ATIVO'])
            col2.metric("Placa", dados_equipamento['PLACA'])
            col3.metric("Média Geral", formatar_brasileiro(consumo_equipamento['Media'].mean()))
            col4.metric("Total Consumido (L)", formatar_brasileiro(consumo_equipamento['Qtde_Litros'].sum()))

            st.markdown("---")
            st.subheader("Informações Cadastrais")
            
            # Exibir todas as informações de forma transposta
            st.dataframe(dados_equipamento.drop('label').to_frame('Valor'), use_container_width=True)


    # --- ABA 3: Tabela Detalhada ---
    with tab_tabela:
        st.header("📋 Tabela Detalhada de Abastecimentos")
        # Adicionar mais colunas à visualização da tabela
        colunas_tabela = [
            "Data", "Cod_Equip", "Descricao_Equip", "PLACA", "DESCRICAOMARCA", "ANOMODELO",
            "Qtde_Litros", "Media", "Media_P", "Classe Operacional"
        ]
        
        df_tabela = df_f[[col for col in colunas_tabela if col in df_f.columns]]

        gb = GridOptionsBuilder.from_dataframe(df_tabela)
        gb.configure_default_column(filterable=True, sortable=True, resizable=True)
        gb.configure_column("Media", type=["numericColumn"], precision=1, header_name="Média (km/l)")
        gb.configure_column("Qtde_Litros", type=["numericColumn"], precision=1, header_name="Litros")
        gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=15)
        gb.configure_selection(selection_mode="multiple", use_checkbox=True)

        grid_opts = gb.build()
        AgGrid(df_tabela, gridOptions=grid_opts, height=500, allow_unsafe_jscode=True)


    # --- ABA 4: Configurações ---
    with tab_config:
        st.header("⚙️ Padrões por Classe Operacional (Configuração de Alertas)")
        # Lógica de configuração permanece a mesma
        if "thr" not in st.session_state:
            classes = df["Classe Operacional"].dropna().unique()
            st.session_state.thr = {
                cls: {"min": 1.5, "max": 5.0} for cls in classes
            }

        classes_operacionais = sorted(df["Classe Operacional"].dropna().unique())
        for cls in classes_operacionais:
            c_min, c_max = st.columns(2)
            if cls not in st.session_state.thr:
                st.session_state.thr[cls] = {"min": 1.5, "max": 5.0}
            
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
