# app.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder
from datetime import datetime
import textwrap
import os
from pathlib import Path
from openpyxl import load_workbook

# ---------------- Configurações ----------------
EXCEL_PATH = "Acompto_Abast.xlsx"
MANUT_LOG_SHEET = "MANUTENCAO_LOG"
COL_KM_HR_ATUAL = "KM_HR_Atual"  # se criar na aba BD, o app usa este valor como leitura atual
OUTROS_CLASSES = {"Motocicletas", "Mini Carregadeira", "Usina", "Veiculos Leves"}

# Paletas
PALETTE_LIGHT = px.colors.sequential.Blues_r
PALETTE_DARK = px.colors.sequential.Plasma_r

# ---------------- Utilitários ----------------
def formatar_brasileiro(valor: float) -> str:
    if pd.isna(valor) or not np.isfinite(valor):
        return "–"
    return "{:,.2f}".format(valor).replace(",", "X").replace(".", ",").replace("X", ".")

def wrap_labels(s: str, width: int = 18) -> str:
    if pd.isna(s):
        return ""
    parts = textwrap.wrap(str(s), width=width)
    return "<br>".join(parts) if parts else str(s)

def norm_col_name(name: str) -> str:
    """Normaliza um nome de coluna (minúsculas, sem acentos / espaços extras)."""
    if not isinstance(name, str):
        return ""
    s = name.strip().lower()
    s = s.replace(" ", "_")
    s = s.replace("-", "_")
    s = s.replace(".", "")
    return s

# mapeamentos de aliases: chave = padrão usado no app, valores possíveis no Excel
COLUMN_ALIASES = {
    "Data": ["data", "date"],
    "Cod_Equip": ["cod_equipamento", "cod_equip", "codigo_equip", "cod", "codigo"],
    "Descricao_Equip": ["descricao_equipamento", "descricao_equip", "descricao"],
    "Qtde_Litros": ["qtde_litros", "litros", "qtde", "quantidade_litros"],
    "Km_Hs_Rod": ["km_hs_rod", "km_hs", "km", "kmrod", "km_hs_rodacao", "quilometros", "quilometro"],
    "Media": ["media", "consumo", "consumo_km_l"],
    "Media_P": ["media_p", "media_percentual"],
    "Ref1": ["ref1", "referencia1"],
    "Ref2": ["ref2", "referencia2"],
    "Unidade": ["unidade", "unid", "unid."],
    "Safra": ["safra"],
    "Classe_Operacional": ["classe_operacional", "classe_operacao", "classe_operacional", "classe_operacional".lower(), "classe operacional", "classeoperacional"],
    "DESCRICAOMARCA": ["ref2", "marca", "marca_descricao", "descricaomarca"],
    # Caso a planilha FROTAS use variações:
    "DESCRICAO_EQUIPAMENTO": ["descricao_equipamento", "descricaoequipamento", "descricao_equ"], 
    "PLACA": ["placa", "plate"],
    "ANOMODELO": ["ano_modelo", "ano_modelo", "anomodelo", "ano", "ano_modelo"],
    # Horimetro / Hodometro single col alternative
    COL_KM_HR_ATUAL: [norm_col_name(COL_KM_HR_ATUAL).lower(), "valor_atual", "km_hr_atual", "km_hr", "hodometro_horimetro", "hodometro", "horimetro"]
}

def find_column_by_alias(cols: list[str], aliases: list[str]) -> str | None:
    """Procura na lista cols algum que bata com aliases (normalizados). Retorna nome original da coluna."""
    norm = {norm_col_name(c): c for c in cols}
    for a in aliases:
        na = norm_col_name(a)
        if na in norm:
            return norm[na]
    return None

@st.cache_data(show_spinner="Carregando e processando dados...")
def load_data(path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Carrega planilhas BD e FROTAS. Faz tentativas de mapear colunas por aliases
    para evitar erros de KeyError por pequenas diferenças de nomenclatura.
    """
    # tenta abrir
    try:
        df_bd = pd.read_excel(path, sheet_name="BD", skiprows=2)
        df_frotas = pd.read_excel(path, sheet_name="FROTAS", skiprows=1)
    except FileNotFoundError:
        st.error(f"Arquivo não encontrado em `{path}`")
        st.stop()
    except ValueError as e:
        # sheet not found ou outro
        st.error("Erro ao abrir o arquivo ou planilhas. Verifique se as abas 'BD' e 'FROTAS' existem.")
        st.stop()

    # Normalizar nomes de colunas de df_bd através de aliases
    bd_cols = list(df_bd.columns)
    rename_map = {}

    # lista de colunas padrão que esperamos no BD (tentaremos mapear)
    expected_bd = ["Data", "Cod_Equip", "Descricao_Equip", "Qtde_Litros", "Km_Hs_Rod",
                   "Media", "Media_P", "Perc_Media", "Ton_Cana", "Litros_Ton",
                   "Ref1", "Ref2", "Unidade", "Safra", "Mes_Excel", "Semana_Excel",
                   "Classe_Original", "Classe_Operacional", "Descricao_Proprietario_Original",
                   "Potencia_CV_Abast", COL_KM_HR_ATUAL]

    for std in expected_bd:
        aliases = COLUMN_ALIASES.get(std, [std])
        found = find_column_by_alias(bd_cols, aliases)
        if found:
            rename_map[found] = std

    # aplica rename (se mapear algo)
    if rename_map:
        df_bd = df_bd.rename(columns=rename_map)

    # Para FROTAS
    f_cols = list(df_frotas.columns)
    rename_map_f = {}
    # mapeia COD_EQUIPAMENTO -> Cod_Equip
    cod_found = find_column_by_alias(f_cols, ["cod_equipamento", "cod_equip", "codigo_equip", "cod"])
    if cod_found:
        rename_map_f[cod_found] = "Cod_Equip"
    # mapeia desalinhadas
    desc_found = find_column_by_alias(f_cols, ["descricao_equipamento", "descricao", "descricao_equip"])
    if desc_found:
        rename_map_f[desc_found] = "DESCRICAO_EQUIPAMENTO"
    placa_found = find_column_by_alias(f_cols, ["placa"])
    if placa_found:
        rename_map_f[placa_found] = "PLACA"
    ano_found = find_column_by_alias(f_cols, ["ano_modelo", "anomodelo", "ano"])
    if ano_found:
        rename_map_f[ano_found] = "ANOMODELO"
    # aplica
    if rename_map_f:
        df_frotas = df_frotas.rename(columns=rename_map_f)

    # garante Cod_Equip presente em frotas
    if "Cod_Equip" not in df_frotas.columns:
        # tenta mapear a partir de BD
        if "Cod_Equip" in df_bd.columns:
            # cria frotas minimal a partir do BD
            unique = df_bd[["Cod_Equip"]].dropna().drop_duplicates()
            df_frotas = unique.rename(columns={"Cod_Equip": "Cod_Equip"})
            df_frotas["DESCRICAO_EQUIPAMENTO"] = ""
            df_frotas["PLACA"] = ""
            df_frotas["ANOMODELO"] = np.nan
        else:
            st.error("Não foi possível identificar coluna de código do equipamento (Cod_Equip) em FROTAS nem em BD.")
            st.stop()

    # padroniza frota: remove duplicados
    df_frotas = df_frotas.drop_duplicates(subset=["Cod_Equip"])
    df_frotas["ANOMODELO"] = pd.to_numeric(df_frotas.get("ANOMODELO", pd.Series()), errors="coerce")

    # cria label para seleção
    df_frotas["label"] = (
        df_frotas["Cod_Equip"].astype(str)
        + " - "
        + df_frotas.get("DESCRICAO_EQUIPAMENTO", "").fillna("")
        + " ("
        + df_frotas.get("PLACA", "").fillna("Sem Placa")
        + ")"
    )

    # agora prepara df_bd: se não existe 'Data' tenta encontrar
    if "Data" not in df_bd.columns:
        possible_date = find_column_by_alias(bd_cols, ["data", "date"])
        if possible_date:
            df_bd = df_bd.rename(columns={possible_date: "Data"})
        else:
            st.error("Coluna de data não encontrada na aba 'BD'. Verifique cabeçalhos.")
            st.stop()

    # converte Data e filtra nulos
    df_bd["Data"] = pd.to_datetime(df_bd["Data"], errors="coerce")
    df_bd = df_bd.dropna(subset=["Data"])

    # Garantir que Cod_Equip exista no BD (tentar mapear)
    if "Cod_Equip" not in df_bd.columns:
        possible = find_column_by_alias(bd_cols, ["cod_equipamento", "cod_equip", "codigo_equip", "cod"])
        if possible:
            df_bd = df_bd.rename(columns={possible: "Cod_Equip"})
        else:
            st.error("Coluna do código do equipamento (Cod_Equip) não encontrada na aba 'BD'.")
            st.stop()

    # Merge BD + Frotas
    df = pd.merge(df_bd, df_frotas, on="Cod_Equip", how="left", suffixes=("", "_frota"))

    # campos derivados de data
    df["Mes"] = df["Data"].dt.month
    df["Semana"] = df["Data"].dt.isocalendar().week
    df["Ano"] = df["Data"].dt.year
    df["AnoMes"] = df["Data"].dt.to_period("M").astype(str)
    df["AnoSemana"] = df["Data"].dt.strftime("%Y-%U")

    # numéricos seguros
    for col in ["Qtde_Litros", "Media", "Media_P", "Km_Hs_Rod", COL_KM_HR_ATUAL]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Marca / Fazenda
    if "Ref2" in df.columns:
        df["DESCRICAOMARCA"] = df["Ref2"].astype(str)
    elif "DESCRICAOMARCA" not in df.columns:
        df["DESCRICAOMARCA"] = ""

    if "Ref1" in df.columns:
        df["Fazenda"] = df["Ref1"].astype(str)
    elif "Fazenda" not in df.columns:
        df["Fazenda"] = ""

    # cria coluna unificada de valor atual (usada em manutenção)
    if COL_KM_HR_ATUAL in df.columns:
        df["Valor_Atual"] = df[COL_KM_HR_ATUAL]
    else:
        # fallback: pega último Km_Hs_Rod por equipamento
        if "Km_Hs_Rod" in df.columns:
            last_km = df.sort_values(["Cod_Equip", "Data"]).groupby("Cod_Equip")["Km_Hs_Rod"].last()
            df = df.merge(last_km.rename("Km_Last_from_hist"), on="Cod_Equip", how="left")
            df["Valor_Atual"] = df["Km_Last_from_hist"]
        else:
            df["Valor_Atual"] = np.nan

    # Garantir coluna Unid: tentar mapear
    if "Unidade" not in df.columns:
        # procurar um alias em bd_cols
        possible_unid = find_column_by_alias(bd_cols, ["unidade", "unid", "tipo_unidade", "unidad"])
        if possible_unid:
            df = df.rename(columns={possible_unid: "Unidade"})
    if "Unidade" not in df.columns:
        df["Unidade"] = np.nan

    return df, df_frotas

def save_maintenance_log(excel_path: str, entries_df: pd.DataFrame, sheet_name: str = MANUT_LOG_SHEET):
    """Salva/concatena entradas de manutenção em MANUTENCAO_LOG no Excel."""
    if entries_df.empty:
        return True
    if "Timestamp" in entries_df.columns:
        entries_df["Timestamp"] = pd.to_datetime(entries_df["Timestamp"])
    # cria arquivo se não existir
    if not Path(excel_path).exists():
        try:
            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                entries_df.to_excel(writer, sheet_name=sheet_name, index=False)
            return True
        except Exception as e:
            st.error(f"Erro ao criar arquivo Excel: {e}")
            return False
    # se existir arquivo, tenta ler e concatenar
    try:
        book = load_workbook(excel_path)
        if sheet_name in book.sheetnames:
            existing = pd.read_excel(excel_path, sheet_name=sheet_name)
            combined = pd.concat([existing, entries_df], ignore_index=True)
        else:
            combined = entries_df
        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            combined.to_excel(writer, sheet_name=sheet_name, index=False)
        return True
    except Exception as e:
        st.error(f"Erro ao gravar log de manutenção no Excel: {e}")
        return False

# ---------------- Manutenção / cálculos ---------------
def build_maintenance_dataframe(df_frotas: pd.DataFrame, df_abast: pd.DataFrame, class_intervals: dict,
                                km_default: int, hr_default: int):
    """
    Cria dataframe de manutenções: usa o último Valor_Atual do BD por equipamento.
    class_intervals: {classe: {"rev_km":[r1,r2,r3], "rev_hr":[r1,r2,r3]}}
    """
    mf = df_frotas.copy()
    # pega último Valor_Atual por Cod_Equip do BD
    last_vals = pd.Series(dtype=float)
    if "Valor_Atual" in df_abast.columns:
        last_vals = df_abast.sort_values(["Cod_Equip", "Data"]).groupby("Cod_Equip")["Valor_Atual"].last()
    mf = mf.set_index("Cod_Equip")
    mf["Km_Hr_Atual"] = last_vals.reindex(mf.index)
    mf = mf.reset_index()

    # pega Unid último do BD
    last_unid = pd.Series(dtype=object)
    if "Unidade" in df_abast.columns:
        last_unid = df_abast.sort_values(["Cod_Equip", "Data"]).groupby("Cod_Equip")["Unidade"].last()
    mf = mf.set_index("Cod_Equip")
    mf["Unid_Last"] = last_unid.reindex(mf.index)
    mf = mf.reset_index()

    # se existe Unid em frotas, prioriza
    if "Unid" in df_frotas.columns:
        mf["Unid"] = df_frotas.set_index("Cod_Equip").reindex(mf["Cod_Equip"])["Unid"].values
        mf["Unid"] = np.where(pd.isna(mf["Unid"]), mf["Unid_Last"], mf["Unid"])
    else:
        mf["Unid"] = mf["Unid_Last"]

    # configura 3 revisões por equipamento com intervalos por classe
    for r in [1, 2, 3]:
        def get_km_interval(cls):
            v = class_intervals.get(cls, {})
            revs = v.get("rev_km")
            if isinstance(revs, (list, tuple)) and len(revs) >= r:
                return revs[r-1]
            # se não for lista, retornar single or default
            if isinstance(revs, (int, float)):
                return revs
            return km_default * r

        def get_hr_interval(cls):
            v = class_intervals.get(cls, {})
            revs = v.get("rev_hr")
            if isinstance(revs, (list, tuple)) and len(revs) >= r:
                return revs[r-1]
            if isinstance(revs, (int, float)):
                return revs
            return hr_default * r

        mf[f"Rev{r}_Interval"] = mf["Classe_Operacional"].apply(lambda cls: get_km_interval(cls))
        mf[f"Rev{r}_Interval_HR"] = mf["Classe_Operacional"].apply(lambda cls: get_hr_interval(cls))

    # calcula próximos vencimentos (considera Unid para decidir KM vs HR)
    def calc_next(row, r):
        cur = row.get("Km_Hr_Atual", np.nan)
        unit = str(row.get("Unid", "")).strip().upper() if pd.notna(row.get("Unid")) else ""
        if pd.isna(cur):
            return (np.nan, np.nan)
        if "QUIL" in unit or unit.startswith("KM"):
            interval = row.get(f"Rev{r}_Interval", np.nan)
            next_due = cur + (interval if not pd.isna(interval) else np.nan)
            return (next_due, next_due - cur)
        if "HOR" in unit or unit.startswith("HR"):
            interval = row.get(f"Rev{r}_Interval_HR", np.nan)
            next_due = cur + (interval if not pd.isna(interval) else np.nan)
            return (next_due, next_due - cur)
        # sem unidade clara: assumir KM por padrão
        interval = row.get(f"Rev{r}_Interval", np.nan)
        next_due = cur + (interval if not pd.isna(interval) else np.nan)
        return (next_due, next_due - cur)

    for r in [1, 2, 3]:
        mf[[f"Rev{r}_Next", f"Rev{r}_To_Go"]] = mf.apply(lambda row: pd.Series(calc_next(row, r)), axis=1)

    mf["Due_Rev"] = False
    mf["Due_Oil"] = False
    return mf

# ---------------- Layout / CSS -------------
def apply_modern_css(dark: bool):
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
        </style>
        """,
        unsafe_allow_html=True
    )

# ----------------- App principal -----------------
def main():
    st.set_page_config(page_title="Dashboard de Frotas e Abastecimentos", layout="wide")
    st.title("📊 Dashboard de Frotas e Abastecimentos — Manutenção Integrada")

    df, df_frotas = load_data(EXCEL_PATH)

    # Sidebar
    with st.sidebar:
        st.header("Configurações")
        dark_mode = st.checkbox("🕶️ Dark Mode", value=False)
        st.markdown("---")
        st.header("Visual")
        top_n = st.slider("Top N classes antes de agrupar 'Outros'", 3, 30, 10)
        hide_text_threshold = st.slider("Esconder valores nas barras quando categorias >", 5, 40, 8)
        st.markdown("---")
        st.header("Manutenção - Defaults")
        km_interval_default = st.number_input("Intervalo padrão (km) revisão", 100, 200000, 10000, step=100)
        hr_interval_default = st.number_input("Intervalo padrão (horas) lubrificação", 1, 5000, 250, step=1)
        km_due_threshold = st.number_input("Alerta revisão se faltar <= (km)", 10, 5000, 500, step=10)
        hr_due_threshold = st.number_input("Alerta lubrificação se faltar <= (horas)", 1, 500, 20, step=1)

    apply_modern_css(dark_mode)
    palette = PALETTE_DARK if dark_mode else PALETTE_LIGHT
    plotly_template = "plotly_dark" if dark_mode else "plotly"

    # session state para intervalos por classe
    if "manut_by_class" not in st.session_state:
        st.session_state.manut_by_class = {}
        classes = sorted(df["Classe_Operacional"].dropna().unique()) if "Classe_Operacional" in df.columns else []
        for cls in classes:
            st.session_state.manut_by_class[cls] = {
                "rev_km": [km_interval_default, km_interval_default * 2, km_interval_default * 3],
                "rev_hr": [hr_interval_default, hr_interval_default * 2, hr_interval_default * 3],
            }

    # abas
    tabs = st.tabs(["📊 Análise de Consumo", "🔎 Consulta de Frota", "📋 Tabela Detalhada", "⚙️ Configurações", "🛠️ Manutenção"])

    # ---------- Aba 1 ----------
    with tabs[0]:
        st.header("Análise de Consumo")
        df_plot = df.copy()
        # garantir coluna de classe presente (mapeamento robusto já feito)
        if "Classe_Operacional" not in df_plot.columns:
            df_plot["Classe_Operacional"] = "Sem Classe"
        df_plot["Classe_Operacional"] = df_plot["Classe_Operacional"].fillna("Sem Classe")
        df_plot["Classe_Grouped"] = df_plot["Classe_Operacional"].apply(lambda s: "Outros" if s in OUTROS_CLASSES else s)
        media_op_full = df_plot.groupby("Classe_Grouped")["Media"].mean().reset_index()
        media_op_full["Media"] = media_op_full["Media"].round(1)
        media_sorted = media_op_full.sort_values("Media", ascending=False)
        if media_sorted.shape[0] > top_n:
            top_keep = media_sorted.head(top_n)["Classe_Grouped"].tolist()
            df_plot["Classe_TopN"] = df_plot["Classe_Grouped"].apply(lambda s: s if s in top_keep else "Outros")
            media_op = df_plot.groupby("Classe_TopN")["Media"].mean().reset_index().rename(columns={"Classe_TopN":"Classe_Grouped"})
            media_op["Media"] = media_op["Media"].round(1)
        else:
            media_op = media_sorted
        media_op["Classe_wrapped"] = media_op["Classe_Grouped"].apply(lambda s: wrap_labels(s, width=16))
        fig = px.bar(media_op, x="Classe_wrapped", y="Media", text="Media", color_discrete_sequence=palette)
        fig.update_layout(template=plotly_template, margin=dict(b=140))
        st.plotly_chart(fig, use_container_width=True)

    # ---------- Aba 2 ----------
    with tabs[1]:
        st.header("Ficha Individual do Equipamento")
        equip_label = st.selectbox("Selecione o Equipamento", options=df_frotas.sort_values("Cod_Equip")["label"])
        if equip_label:
            cod_sel = int(str(equip_label).split(" - ")[0])
            dados_eq = df_frotas.query("Cod_Equip == @cod_sel").iloc[0]
            # último valor atual
            last_val_series = df.sort_values(["Cod_Equip", "Data"]).query("Cod_Equip == @cod_sel").groupby("Cod_Equip")["Valor_Atual"].last()
            val = last_val_series.iloc[0] if not last_val_series.empty else np.nan
            unid_series = df.sort_values(["Cod_Equip", "Data"]).query("Cod_Equip == @cod_sel").groupby("Cod_Equip")["Unidade"].last()
            unidade = unid_series.iloc[0] if not unid_series.empty else ""
            st.subheader(f"{dados_eq.get('DESCRICAO_EQUIPAMENTO','–')} ({dados_eq.get('PLACA','–')})")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Status", dados_eq.get("ATIVO", "–"))
            c2.metric("Placa", dados_eq.get("PLACA", "–"))
            c3.metric("Medida Atual", f"{formatar_brasileiro(val)} {unidade}")
            consumo_eq = df.query("Cod_Equip == @cod_sel")
            c4.metric("Média Geral", formatar_brasileiro(consumo_eq["Media"].mean()))

    # ---------- Aba 3 ----------
    with tabs[2]:
        st.header("Tabela Detalhada de Abastecimentos")
        cols = ["Data", "Cod_Equip", "Descricao_Equip", "PLACA", "DESCRICAOMARCA", "ANOMODELO", "Qtde_Litros", "Media", "Media_P", "Classe_Operacional", COL_KM_HR_ATUAL, "Unidade"]
        df_tab = df[[c for c in cols if c in df.columns]]
        st.download_button("⬇️ Exportar CSV da Tabela", df_tab.to_csv(index=False).encode("utf-8"), "abastecimentos.csv", "text/csv")
        gb = GridOptionsBuilder.from_dataframe(df_tab)
        gb.configure_default_column(filterable=True, sortable=True, resizable=True)
        if "Media" in df_tab.columns:
            gb.configure_column("Media", type=["numericColumn"], precision=1)
        if "Qtde_Litros" in df_tab.columns:
            gb.configure_column("Qtde_Litros", type=["numericColumn"], precision=1)
        gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=15)
        gb.configure_selection("single", use_checkbox=True)
        AgGrid(df_tab, gridOptions=gb.build(), height=520, allow_unsafe_jscode=True)

    # ---------- Aba 4 ----------
    with tabs[3]:
        st.header("Padrões por Classe Operacional (Alertas & Intervalos)")
        classes = sorted(df["Classe_Operacional"].dropna().unique()) if "Classe_Operacional" in df.columns else []
        st.markdown("Configure intervalos (km e horas) para as 3 revisões por classe.")
        for cls in classes:
            st.subheader(str(cls))
            col1, col2 = st.columns(2)
            rev_km = st.session_state.manut_by_class.get(cls, {}).get("rev_km", [km_interval_default, km_interval_default*2, km_interval_default*3])
            rev_hr = st.session_state.manut_by_class.get(cls, {}).get("rev_hr", [hr_interval_default, hr_interval_default*2, hr_interval_default*3])
            with col1:
                nk1 = st.number_input(f"{cls} → Rev1 (km)", min_value=0, max_value=1000000, value=int(rev_km[0]), key=f"{cls}_r1km")
                nk2 = st.number_input(f"{cls} → Rev2 (km)", min_value=0, max_value=1000000, value=int(rev_km[1]), key=f"{cls}_r2km")
                nk3 = st.number_input(f"{cls} → Rev3 (km)", min_value=0, max_value=1000000, value=int(rev_km[2]), key=f"{cls}_r3km")
            with col2:
                nh1 = st.number_input(f"{cls} → Rev1 (hr)", min_value=0, max_value=1000000, value=int(rev_hr[0]), key=f"{cls}_r1hr")
                nh2 = st.number_input(f"{cls} → Rev2 (hr)", min_value=0, max_value=1000000, value=int(rev_hr[1]), key=f"{cls}_r2hr")
                nh3 = st.number_input(f"{cls} → Rev3 (hr)", min_value=0, max_value=1000000, value=int(rev_hr[2]), key=f"{cls}_r3hr")
            st.session_state.manut_by_class[cls] = {"rev_km":[int(nk1), int(nk2), int(nk3)], "rev_hr":[int(nh1), int(nh2), int(nh3)]}

    # ---------- Aba 5 - Manutenção ----------
    with tabs[4]:
        st.header("Controle de Revisões e Lubrificação")
        st.markdown("O sistema usa a coluna `KM_HR_Atual` (na aba BD) quando disponível; senão usa o último Km/Hs do histórico. A coluna `Unidade` em BD deve indicar `QUILÔMETROS` ou `HORAS` (ou similar).")

        # montar class_intervals
        class_intervals = {}
        for k,v in st.session_state.manut_by_class.items():
            class_intervals[k] = {"rev_km": v.get("rev_km", []), "rev_hr": v.get("rev_hr", [])}

        mf = build_maintenance_dataframe(df_frotas, df, class_intervals, int(km_interval_default), int(hr_interval_default))

        # calcular flags de proximidade
        def set_due_flags(row):
            due_km = False
            due_hr = False
            unit = str(row.get("Unid","")).upper() if pd.notna(row.get("Unid")) else ""
            for r in [1,2,3]:
                to_go = row.get(f"Rev{r}_To_Go", np.nan)
                if pd.isna(to_go):
                    continue
                if "QUIL" in unit or unit.startswith("KM"):
                    if to_go <= km_due_threshold:
                        due_km = True
                elif "HOR" in unit or unit.startswith("HR"):
                    if to_go <= hr_due_threshold:
                        due_hr = True
                else:
                    if to_go <= km_due_threshold:
                        due_km = True
            return pd.Series({"Due_Rev": due_km, "Due_Oil": due_hr})

        if not mf.empty:
            flags = mf.apply(set_due_flags, axis=1)
            mf["Due_Rev"] = flags["Due_Rev"]
            mf["Due_Oil"] = flags["Due_Oil"]
            mf["Any_Due"] = mf["Due_Rev"] | mf["Due_Oil"]
        else:
            mf["Any_Due"] = False

        df_due = mf[mf["Any_Due"]].copy().sort_values(["Due_Rev", "Due_Oil"], ascending=False)

        st.subheader("Equipamentos com revisão/lubrificação próxima ou vencida")
        st.write(f"Total: {len(df_due)}")
        if not df_due.empty:
            display_cols = ["Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Unid", "Km_Hr_Atual",
                            "Rev1_To_Go", "Rev2_To_Go", "Rev3_To_Go", "Due_Rev", "Due_Oil"]
            available = [c for c in display_cols if c in df_due.columns]
            st.dataframe(df_due[available].reset_index(drop=True), use_container_width=True)

            st.markdown("### Marcar revisões como concluídas")
            st.markdown("Selecione o que foi feito e clique em **Salvar ações** — isso gravará um registro na aba `MANUTENCAO_LOG` do Excel.")

            actions = []
            for idx, row in df_due.reset_index(drop=True).iterrows():
                cod = row["Cod_Equip"]
                name = row.get("DESCRICAO_EQUIPAMENTO", "")
                unit = row.get("Unid", "")
                cur_val = row.get("Km_Hr_Atual", np.nan)
                st.markdown(f"**{cod} - {name}** — Atual: {formatar_brasileiro(cur_val)} {unit}")
                cols = st.columns([1,1,1,1])
                cb1 = cols[0].checkbox(f"Rev1 (cod {cod})", key=f"r1_{cod}")
                cb2 = cols[1].checkbox(f"Rev2 (cod {cod})", key=f"r2_{cod}")
                cb3 = cols[2].checkbox(f"Rev3 (cod {cod})", key=f"r3_{cod}")
                cbd = cols[3].checkbox(f"Lubr. (cod {cod})", key=f"lub_{cod}")
                if cb1:
                    actions.append({"Cod_Equip": cod, "Tipo":"Rev1", "Valor_Atual": cur_val, "Unid":unit})
                if cb2:
                    actions.append({"Cod_Equip": cod, "Tipo":"Rev2", "Valor_Atual": cur_val, "Unid":unit})
                if cb3:
                    actions.append({"Cod_Equip": cod, "Tipo":"Rev3", "Valor_Atual": cur_val, "Unid":unit})
                if cbd:
                    actions.append({"Cod_Equip": cod, "Tipo":"Lubrificacao", "Valor_Atual": cur_val, "Unid":unit})

            if st.button("💾 Salvar ações de manutenção"):
                if not actions:
                    st.info("Nenhuma ação selecionada.")
                else:
                    now = datetime.now()
                    rows = []
                    for a in actions:
                        rows.append({
                            "Timestamp": now,
                            "Cod_Equip": a["Cod_Equip"],
                            "Tipo": a["Tipo"],
                            "Valor_Atual": a["Valor_Atual"],
                            "Unid": a["Unid"],
                            "Usuario": st.session_state.get("usuario","(anon)")
                        })
                    entries_df = pd.DataFrame(rows)
                    ok = save_maintenance_log(EXCEL_PATH, entries_df, MANUT_LOG_SHEET)
                    if ok:
                        st.success(f"{len(rows)} ação(ões) registrada(s) em `{MANUT_LOG_SHEET}`.")
                        st.experimental_rerun()
                    else:
                        st.error("Falha ao gravar ações. Verifique permissões/arquivo.")
        else:
            st.info("Nenhum equipamento com alerta dentro dos thresholds configurados.")

        st.markdown("---")
        st.subheader("Visão geral da frota - manutenção planejada")
        overview_cols = ["Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Km_Hr_Atual", "Unid",
                         "Rev1_Next", "Rev1_To_Go", "Rev2_Next", "Rev2_To_Go", "Rev3_Next", "Rev3_To_Go"]
        available_over = [c for c in overview_cols if c in mf.columns]
        st.dataframe(mf[available_over].sort_values("Cod_Equip").reset_index(drop=True), use_container_width=True)
        st.download_button("⬇️ Exportar CSV - Plano de Manutenção (Visão Geral)", mf[available_over].to_csv(index=False).encode("utf-8"), "manutencao_overview.csv", "text/csv")

if __name__ == "__main__":
    main()
