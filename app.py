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

# Paletas
PALETTE_LIGHT = px.colors.sequential.Blues_r
PALETTE_DARK = px.colors.sequential.Plasma_r

# Classes agrupadas em 'Outros'
OUTROS_CLASSES = {"Motocicletas", "Mini Carregadeira", "Usina", "Veiculos Leves"}

# Nome exato da coluna única que conterá hodômetro / horímetro atual na aba BD (opcional)
COL_KM_HR_ATUAL = "KM_HR_Atual"  # se você criar essa coluna na planilha BD, o app a usará diretamente

# Nome da aba de log de manutenção
MANUT_LOG_SHEET = "MANUTENCAO_LOG"

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

def find_col_like(df: pd.DataFrame, keywords: list[str]) -> str | None:
    """Procura primeira coluna cujo nome contém qualquer palavra-chave (case-insensitive)."""
    cols = df.columns.astype(str)
    low = [c.lower() for c in cols]
    for kw in keywords:
        for i, c in enumerate(low):
            if kw.lower() in c:
                return cols[i]
    return None

def find_first_numeric_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Retorna primeiro candidato que exista em df e contenha valores numéricos na maioria."""
    for cand in candidates:
        if cand in df.columns:
            ser = pd.to_numeric(df[cand], errors="coerce")
            # se pelo menos alguns valores forem numéricos, assume-se que é coluna válida
            if ser.notna().sum() > 0:
                return cand
    # fallback: se nenhuma correspondência por nome, procura colunas numéricas que pareçam "km" ou "hor" no nome
    for c in df.columns:
        name = str(c).lower()
        if any(k in name for k in ["km", "quil", "hor", "hr", "hora", "kms"]) and pd.to_numeric(df[c], errors="coerce").notna().sum() > 0:
            return c
    return None

@st.cache_data(show_spinner="Carregando e processando dados...")
def load_data(path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Carrega BD (sheet 'BD') e FROTAS (sheet 'FROTAS'). Ajusta nomes e tipos com segurança."""
    try:
        # carrega as duas abas; se sheet não existir, aborta com mensagem
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
        + df_frotas.get("DESCRICAO_EQUIPAMENTO", "").fillna("")
        + " ("
        + df_frotas.get("PLACA", "").fillna("Sem Placa")
        + ")"
    )

    # Tenta mapear as colunas do BD sem forçar nome por posição — por causa de variações na planilha
    # Se BD tiver exatamente as colunas esperadas por posição, podemos renomear com segurança.
    expected = [
        "Data", "Cod_Equip", "Descricao_Equip", "Qtde_Litros", "Km_Hs_Rod",
        "Media", "Media_P", "Perc_Media", "Ton_Cana", "Litros_Ton",
        "Ref1", "Ref2", "Unidade", "Safra", "Mes_Excel", "Semana_Excel",
        "Classe_Original", "Classe_Operacional", "Descricao_Proprietario_Original",
        "Potencia_CV_Abast"
    ]
    # renomeia por posição só se o número bater exatamente — evita ValueError
    if df_abast.shape[1] == len(expected):
        df_abast.columns = expected
    else:
        # tenta renomear por correspondência de nomes (case-insensitive)
        rename_map = {}
        for exp in expected:
            keywords = [exp.lower()]
            # acrescenta sinônimos simples
            if "data" in exp.lower():
                keywords += ["date"]
            if "cod_equip" in exp.lower():
                keywords += ["cod", "codigo", "equipamento", "cod_equipamento"]
            if "descricao" in exp.lower():
                keywords += ["descricao", "descrição", "descri"]
            if "qtde" in exp.lower() or "litros" in exp.lower():
                keywords += ["litros", "qtde", "quantidade"]
            if "km" in exp.lower() or "Km_Hs_Rod" in exp:
                keywords += ["km", "kms", "quilometro", "quilômetros", "quilometros", "km_hs_rod"]
            if "unidade" in exp.lower() or "unidade" in exp:
                keywords += ["unid", "unidade", "unid.", "un."]
            # procura coluna no df_abast parecido
            found = find_col_like(df_abast, keywords)
            if found:
                rename_map[found] = exp
        if rename_map:
            df_abast = df_abast.rename(columns=rename_map)
        # se ainda faltar colunas esperadas, não forçamos — usaremos get(...) mais adiante

    # Merge para enriquecer abast com dados de frota
    # Se Cod_Equip não existir em abast, tenta detectar coluna equivalente
    if "Cod_Equip" not in df_abast.columns:
        # tenta descobrir coluna de código por heurística
        candidate = find_col_like(df_abast, ["cod", "equip", "codigo"])
        if candidate:
            df_abast = df_abast.rename(columns={candidate: "Cod_Equip"})
    # Se ainda não tem Cod_Equip, o merge ficará vazio; preferimos parar e avisar
    if "Cod_Equip" not in df_abast.columns:
        st.error("Coluna de equipamento (Cod_Equip) não encontrada na aba 'BD'. Verifique a planilha.")
        st.stop()

    df = pd.merge(df_abast, df_frotas, on="Cod_Equip", how="left")
    # Data
    if "Data" in df.columns:
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
        df.dropna(subset=["Data"], inplace=True)
    else:
        # sem coluna data nada de temporal funciona corretamente; avisamos mas continuamos
        st.warning("Atenção: coluna 'Data' não encontrada em 'BD'. Algumas funcionalidades podem não funcionar corretamente.")

    # Campos derivados se Data existe
    if "Data" in df.columns:
        df["Mes"] = df["Data"].dt.month
        df["Semana"] = df["Data"].dt.isocalendar().week
        df["Ano"] = df["Data"].dt.year
        df["AnoMes"] = df["Data"].dt.to_period("M").astype(str)
        df["AnoSemana"] = df["Data"].dt.strftime("%Y-%U")
    else:
        df["Mes"] = np.nan
        df["Semana"] = np.nan
        df["Ano"] = np.nan
        df["AnoMes"] = np.nan
        df["AnoSemana"] = np.nan

    # numéricos seguros: tenta converter se col existir
    for col in ["Qtde_Litros", "Media", "Media_P", "Km_Hs_Rod", COL_KM_HR_ATUAL]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # marca / fazenda
    if "Ref2" in df.columns:
        df["DESCRICAOMARCA"] = df["Ref2"].astype(str)
    else:
        df["DESCRICAOMARCA"] = df.get("DESCRICAOMARCA", "").astype(str)
    df["Fazenda"] = df.get("Ref1", "").astype(str)

    # Detecta coluna de medição atual (KM_HR_Atual ou outro candidato):
    # 1) Se COL_KM_HR_ATUAL explicitamente presente (ex.: você criou esta coluna), usa-a.
    # 2) Senão, tenta encontrar colunas semelhantes ("QUILOMETROS","HORAS","KM","Km_Atual", etc.)
    valor_col = None
    if COL_KM_HR_ATUAL in df.columns:
        valor_col = COL_KM_HR_ATUAL
        df["Valor_Atual"] = pd.to_numeric(df[valor_col], errors="coerce")
    else:
        # candidatos comuns
        cand_cols = ["KM", "Km", "Km_Atual", "KmAtual", "KM_ATUAL", "QUILOMETROS", "Quilometros", "HORAS", "Horas", "Hodometro", "Hodômetro", "Horimetro", "Horímetro"]
        # procura coluna numérica entre candidatos
        found_num = find_first_numeric_col(df, [c for c in cand_cols if c in df.columns])
        if found_num:
            valor_col = found_num
            df["Valor_Atual"] = pd.to_numeric(df[found_num], errors="coerce")
        else:
            # fallback para Km_Hs_Rod do próprio histórico (último por equipamento)
            if "Km_Hs_Rod" in df.columns:
                last_km = df.sort_values(["Cod_Equip", "Data"]).groupby("Cod_Equip")["Km_Hs_Rod"].last()
                df = df.merge(last_km.rename("Km_Last_from_hist"), on="Cod_Equip", how="left")
                df["Valor_Atual"] = df["Km_Last_from_hist"]
            else:
                df["Valor_Atual"] = np.nan

    # detecta coluna Unid/Unidade (ex.: 'QUILÔMETROS'/'HORAS' no BD) preferindo colunas existentes
    unid_col = None
    for candidate in ["Unidade","Unid","UNID","UNIDADE","Unidade_med","Unidade_medido","Unid."]:
        if candidate in df.columns:
            unid_col = candidate
            break
    if unid_col:
        df["Unidade"] = df[unid_col].astype(str)
    else:
        # tenta pegar do histórico antes do merge (df_abast)
        found_un = find_col_like(df_abast, ["unid", "unidade", "quil", "hora", "hor"])
        if found_un:
            df["Unidade"] = df_abast[found_un].astype(str)
        else:
            # fallback vazio
            df["Unidade"] = ""

    return df, df_frotas

def save_maintenance_log(excel_path: str, entries_df: pd.DataFrame, sheet_name: str = MANUT_LOG_SHEET):
    """
    Salva (append) as entradas de manutenção em uma aba MANUTENCAO_LOG.
    Se a aba existir, carrega e concatena, senão cria.
    """
    # garante formato de data/hora
    if "Timestamp" in entries_df.columns:
        entries_df["Timestamp"] = pd.to_datetime(entries_df["Timestamp"])

    # se arquivo não existir, cria com o sheet
    if not Path(excel_path).exists():
        try:
            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                entries_df.to_excel(writer, sheet_name=sheet_name, index=False)
            return True
        except Exception as e:
            st.error(f"Erro ao criar arquivo Excel para log: {e}")
            return False

    # se existir, lê workbook e regrava com sheet atualizado
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

# --------- Funções de manutenção / cálculo ------------
def build_maintenance_dataframe(df_frotas: pd.DataFrame, df_abast: pd.DataFrame, class_intervals: dict,
                                km_default: int, hr_default: int):
    """
    Constrói um dataframe com próximas revisões por equipamento usando:
    - df_frotas (base cadastral)
    - df_abast (pode conter Valor_Atual e Unidade)
    - class_intervals: dict[classe] = {"rev_km":[..3], "rev_hr":[..3]}
    - km_default / hr_default: valores default se não houver por classe
    Returns mf (frota enriched)
    """
    mf = df_frotas.copy()

    # busca valor atual por equipamento (procura na aba BD se existir)
    # tenta ler a última medição (Valor_Atual) da aba BD
    try:
        last_vals = df_abast.sort_values(["Cod_Equip", "Data"]).groupby("Cod_Equip")["Valor_Atual"].last()
    except Exception:
        last_vals = pd.Series(dtype=float)

    mf = mf.set_index("Cod_Equip")
    mf["Km_Hr_Atual"] = last_vals.reindex(mf.index)
    mf = mf.reset_index()

    # Unidade: pega último 'Unidade' por equipamento do histórico (se existir)
    try:
        last_unid = df_abast.sort_values(["Cod_Equip", "Data"]).groupby("Cod_Equip")["Unidade"].last()
    except Exception:
        last_unid = pd.Series(dtype=object)
    mf = mf.set_index("Cod_Equip")
    mf["Unid_Last"] = last_unid.reindex(mf.index)
    mf = mf.reset_index()

    # Se existe 'Unid' em df_frotas, prioriza, senão usa Unid_Last
    if "Unid" in df_frotas.columns:
        mf["Unid"] = df_frotas.set_index("Cod_Equip").reindex(mf["Cod_Equip"])["Unid"].values
        mf["Unid"] = np.where(pd.isna(mf["Unid"]), mf["Unid_Last"], mf["Unid"])
    else:
        mf["Unid"] = mf["Unid_Last"]

    # agora calcular para cada revisão (3 revisões) - define intervalos por classe (listas de 3)
    rev_nums = [1, 2, 3]
    for r in rev_nums:
        def choose_km_interval(cls):
            if cls in class_intervals:
                v = class_intervals[cls].get("rev_km")
                if isinstance(v, list) and len(v) >= r:
                    return v[r-1]
                if isinstance(v, (int, float)):
                    return v
            # fallback padrão
            return km_default * r
        def choose_hr_interval(cls):
            if cls in class_intervals:
                v = class_intervals[cls].get("rev_hr")
                if isinstance(v, list) and len(v) >= r:
                    return v[r-1]
                if isinstance(v, (int, float)):
                    return v
            return hr_default * r

        mf[f"Rev{r}_Interval"] = mf["Classe_Operacional"].apply(lambda cls: choose_km_interval(cls))
        mf[f"Rev{r}_Interval_HR"] = mf["Classe_Operacional"].apply(lambda cls: choose_hr_interval(cls))

    # calc next due based on unit (Km_Hr_Atual)
    def calc_next(row, r):
        cur = row.get("Km_Hr_Atual", np.nan)
        unit = str(row.get("Unid", "")).strip().upper() if pd.notna(row.get("Unid")) else ""
        if pd.isna(cur):
            return (np.nan, np.nan)  # next, to_go
        if "QUIL" in unit or "KM" in unit or unit.startswith("K"):
            interval = row.get(f"Rev{r}_Interval", np.nan)
            next_due = cur + (interval if not pd.isna(interval) else np.nan)
            to_go = next_due - cur
            return (next_due, to_go)
        elif "HOR" in unit or "HR" in unit or unit.startswith("H"):
            interval = row.get(f"Rev{r}_Interval_HR", np.nan)
            next_due = cur + (interval if not pd.isna(interval) else np.nan)
            to_go = next_due - cur
            return (next_due, to_go)
        else:
            # sem unidade clara, usa km por padrão
            interval = row.get(f"Rev{r}_Interval", np.nan)
            next_due = cur + (interval if not pd.isna(interval) else np.nan)
            to_go = next_due - cur
            return (next_due, to_go)

    for r in rev_nums:
        mf[[f"Rev{r}_Next", f"Rev{r}_To_Go"]] = mf.apply(lambda row: pd.Series(calc_next(row, r)), axis=1)

    # flags default
    mf["Due_Rev"] = False
    mf["Due_Oil"] = False
    mf["Any_Due"] = False

    return mf

# ---------------- Layout / CSS leve ----------------
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

    # Carrega dados
    df, df_frotas = load_data(EXCEL_PATH)

    # Sidebar controles gerais + manutenção defaults
    with st.sidebar:
        st.header("Configurações")
        dark_mode = st.checkbox("🕶️ Dark Mode", value=False)
        st.markdown("---")
        st.header("Visual")
        top_n = st.slider("Top N classes antes de agrupar 'Outros'", min_value=3, max_value=30, value=10)
        hide_text_threshold = st.slider("Esconder valores nas barras quando categorias >", min_value=5, max_value=40, value=8)
        st.markdown("---")
        st.header("Manutenção - Defaults globais")
        km_interval_default = st.number_input("Intervalo padrão (km) revisão", min_value=100, max_value=200000, value=10000, step=100)
        hr_interval_default = st.number_input("Intervalo padrão (horas) lubrificação", min_value=1, max_value=5000, value=250, step=1)
        km_due_threshold = st.number_input("Alerta revisão se faltar <= (km)", min_value=10, max_value=5000, value=500, step=10)
        hr_due_threshold = st.number_input("Alerta lubrificação se faltar <= (horas)", min_value=1, max_value=500, value=20, step=1)

    # Aplica CSS
    apply_modern_css(dark_mode)
    palette = PALETTE_DARK if dark_mode else PALETTE_LIGHT
    plotly_template = "plotly_dark" if dark_mode else "plotly"

    # Configurações por classe (sessão) - inicializa se necessário
    if "manut_by_class" not in st.session_state:
        st.session_state.manut_by_class = {}
        classes = sorted(df["Classe_Operacional"].dropna().unique()) if "Classe_Operacional" in df.columns else []
        for cls in classes:
            st.session_state.manut_by_class[cls] = {
                "rev_km": [km_interval_default, km_interval_default*2, km_interval_default*3],
                "rev_hr": [hr_interval_default, hr_interval_default*2, hr_interval_default*3]
            }

    # Layout - abas (inclui Manutenção)
    tabs = st.tabs(["📊 Análise de Consumo", "🔎 Consulta de Frota", "📋 Tabela Detalhada", "⚙️ Configurações", "🛠️ Manutenção"])

    # ---------- Aba 1: Análise (simplificada) ----------
    with tabs[0]:
        st.header("Análise de Consumo")
        st.info("Visual principal — (melhorias aplicadas).")
        # Exemplo rápido: média por classe com agrupamento 'Outros'
        df_plot = df.copy()
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

    # ---------- Aba 2: Consulta de Frota ----------
    with tabs[1]:
        st.header("Ficha Individual do Equipamento")
        equip_label = st.selectbox("Selecione o Equipamento", options=df_frotas.sort_values("Cod_Equip")["label"])
        if equip_label:
            cod_sel = int(equip_label.split(" - ")[0])
            dados_eq = df_frotas.query("Cod_Equip == @cod_sel").iloc[0]
            # busca valor atual no histórico/enriquecido
            last_val = df.sort_values(["Cod_Equip", "Data"]).query("Cod_Equip == @cod_sel").groupby("Cod_Equip")["Valor_Atual"].last()
            val = last_val.iloc[0] if not last_val.empty else np.nan
            unidade = df.sort_values(["Cod_Equip", "Data"]).query("Cod_Equip == @cod_sel").groupby("Cod_Equip")["Unidade"].last().iloc[0] if not last_val.empty else ""
            st.subheader(f"{dados_eq.get('DESCRICAO_EQUIPAMENTO','–')} ({dados_eq.get('PLACA','–')})")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Status", dados_eq.get("ATIVO", "–"))
            c2.metric("Placa", dados_eq.get("PLACA", "–"))
            c3.metric("Medida Atual", f"{formatar_brasileiro(val)} {unidade}")
            # média geral do equipamento
            consumo_eq = df.query("Cod_Equip == @cod_sel")
            c4.metric("Média Geral", formatar_brasileiro(consumo_eq["Media"].mean()))

    # ---------- Aba 3: Tabela Detalhada ----------
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

    # ---------- Aba 4: Configurações ----------
    with tabs[3]:
        st.header("Padrões por Classe Operacional (Alertas & Intervalos)")
        st.markdown("Aqui você pode ajustar os intervalos por classe para as 3 revisões (KM e HORAS).")
        classes = sorted(df["Classe_Operacional"].dropna().unique()) if "Classe_Operacional" in df.columns else []
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

    # ---------- Aba 5: Manutenção ----------
    with tabs[4]:
        st.header("Controle de Revisões e Lubrificação")
        st.markdown("O sistema usa a coluna `KM_HR_Atual` (na aba BD) se existir, senão tenta detectar automaticamente. A coluna `Unid`/`Unidade` (ex.: 'QUILÔMETROS' ou 'HORAS') indica a unidade principal usada para cada equipamento.")

        # monta class_intervals a partir do session_state
        class_intervals = {}
        for k, v in st.session_state.manut_by_class.items():
            class_intervals[k] = {"rev_km": v.get("rev_km", []), "rev_hr": v.get("rev_hr", [])}

        mf = build_maintenance_dataframe(df_frotas, df, class_intervals, int(km_interval_default), int(hr_interval_default))

        # calcula flags de proximidade
        def set_due_flags(row):
            due_km = False
            due_hr = False
            unit = str(row.get("Unid","")).upper() if pd.notna(row.get("Unid")) else ""
            for r in [1,2,3]:
                to_go = row.get(f"Rev{r}_To_Go", np.nan)
                if pd.isna(to_go):
                    continue
                if "QUIL" in unit or unit.startswith("KM") or unit.startswith("K"):
                    if to_go <= km_due_threshold:
                        due_km = True
                elif "HOR" in unit or unit.startswith("HR") or unit.startswith("H"):
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
            st.markdown("Selecione o que foi feito e clique em **Salvar ações** — isso gravará um registro na aba `MANUTENCAO_LOG` do arquivo Excel.")

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
                cbd = cols[3].checkbox(f"Lubrificação (cod {cod})", key=f"lub_{cod}")
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
                            "Usuario": st.session_state.get("usuario","(anon)"),
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
