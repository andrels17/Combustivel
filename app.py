# app.py
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from st_aggrid import AgGrid, GridOptionsBuilder
from datetime import datetime
import textwrap
from pathlib import Path
from openpyxl import load_workbook

# ---------------- Configurações ----------------
EXCEL_PATH = "Acompto_Abast.xlsx"
MANUT_LOG_SHEET = "MANUTENCAO_LOG"
COL_KM_HR_PREFERRED = "KM_HR_Atual"  # se você criar essa coluna na aba BD, será usada como medidor atual
OUTROS_CLASSES = {"Motocicletas", "Mini Carregadeira", "Usina", "Veiculos Leves"}

PALETTE_LIGHT = px.colors.sequential.Blues_r
PALETTE_DARK = px.colors.sequential.Plasma_r

# possíveis nomes alternativos de colunas (serão detectados automaticamente)
COD_EQUIP_CANDIDATES = ["Cod_Equip", "COD_EQUIPAMENTO", "Codigo", "Código", "CODIGO", "cod_equip"]
CLASS_CANDIDATES = ["Classe_Operacional", "Classe Operacional", "Classe", "classe_operacional"]
UNID_CANDIDATES = ["Unid", "Unidade", "UNID", "UNIDADE"]
KMHR_CANDIDATES = [
    COL_KM_HR_PREFERRED, "KM_HR_ATUAL", "Medidor_Atual", "Km_Atual", "KmAtual",
    "Hodometro", "Hodômetro", "HORAS", "HORIMETRO", "Km", "KM"
]
# ---------------- Utilitários ----------------

def formatar_brasileiro(valor) -> str:
    if pd.isna(valor) or not np.isfinite(valor):
        return "–"
    return "{:,.2f}".format(float(valor)).replace(",", "X").replace(".", ",").replace("X", ".")

def wrap_labels(s: str, width: int = 18) -> str:
    if pd.isna(s):
        return ""
    parts = textwrap.wrap(str(s), width=width)
    return "<br>".join(parts) if parts else str(s)

def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Retorna primeiro candidato existente como nome exato de coluna."""
    for c in candidates:
        if c in df.columns:
            return c
    # tentar buscar sem acento e case-insensitive
    cols_norm = {col.lower().replace("ç","c").replace("ô","o").replace("ó","o"): col for col in df.columns}
    for cand in candidates:
        key = cand.lower().replace("ç","c").replace("ô","o").replace("ó","o")
        if key in cols_norm:
            return cols_norm[key]
    return None

@st.cache_data(show_spinner="Carregando e processando dados...")
def load_data(path: str) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Carrega BD e FROTAS. Normaliza colunas e retorna também um dicionário com mapeamentos detectados."""
    if not Path(path).exists():
        st.error(f"Arquivo `{path}` não encontrado. Coloque o arquivo na mesma pasta ou ajuste EXCEL_PATH.")
        st.stop()

    try:
        df_bd = pd.read_excel(path, sheet_name="BD", skiprows=2)
    except Exception as e:
        st.error(f"Erro ao ler a aba 'BD': {e}")
        st.stop()

    try:
        df_frotas = pd.read_excel(path, sheet_name="FROTAS", skiprows=1)
    except Exception as e:
        st.error(f"Erro ao ler a aba 'FROTAS': {e}")
        st.stop()

    # detecta colunas importantes com tolerância
    cod_col = find_column(df_bd, COD_EQUIP_CANDIDATES) or find_column(df_frotas, COD_EQUIP_CANDIDATES)
    class_col = find_column(df_bd, CLASS_CANDIDATES) or find_column(df_frotas, CLASS_CANDIDATES)
    unid_col = find_column(df_bd, UNID_CANDIDATES) or find_column(df_frotas, UNID_CANDIDATES)
    kmhr_col = find_column(df_bd, KMHR_CANDIDATES) or find_column(df_frotas, KMHR_CANDIDATES)

    # normalize column names expected downstream: vamos renomear as colunas do BD para nomes padronizados se possível
    # Mas só renomeamos se encontrarmos o cod_col correspondente.
    if cod_col and cod_col != "Cod_Equip":
        df_bd = df_bd.rename(columns={cod_col: "Cod_Equip"})
    if class_col and class_col != "Classe_Operacional":
        df_bd = df_bd.rename(columns={class_col: "Classe_Operacional"})
        df_frotas = df_frotas.rename(columns={class_col: "Classe_Operacional"}) if class_col in df_frotas.columns else df_frotas
    if unid_col and unid_col != "Unidade":
        df_bd = df_bd.rename(columns={unid_col: "Unidade"})
    if kmhr_col and kmhr_col != COL_KM_HR_PREFERRED:
        df_bd = df_bd.rename(columns={kmhr_col: COL_KM_HR_PREFERRED})
        df_frotas = df_frotas.rename(columns={kmhr_col: COL_KM_HR_PREFERRED}) if kmhr_col in df_frotas.columns else df_frotas

    # ensure FROTAS has Cod_Equip column - try to detect if FROTAS uses different name
    cod_frotas = find_column(df_frotas, COD_EQUIP_CANDIDATES)
    if cod_frotas and cod_frotas != "Cod_Equip":
        df_frotas = df_frotas.rename(columns={cod_frotas: "Cod_Equip"})

    # padroniza alguns valores
    df_frotas = df_frotas.rename(columns={"COD_EQUIPAMENTO":"Cod_Equip"}) if "COD_EQUIPAMENTO" in df_frotas.columns else df_frotas
    if "Cod_Equip" not in df_frotas.columns:
        st.error("Coluna do código do equipamento (ex.: 'Cod_Equip' / 'COD_EQUIPAMENTO') não encontrada na aba 'FROTAS'. Verifique o arquivo.")
        st.stop()

    # prepara label para select
    df_frotas = df_frotas.drop_duplicates(subset=["Cod_Equip"])
    df_frotas["label"] = df_frotas["Cod_Equip"].astype(str) + " - " + df_frotas.get("DESCRICAO_EQUIPAMENTO", "").fillna("").astype(str) + " (" + df_frotas.get("PLACA", "").fillna("Sem Placa").astype(str) + ")"

    # Normaliza BD: se não existir Cod_Equip em BD, erro controlado mas permitimos fallback (não todos os recursos funcionarão)
    if "Cod_Equip" not in df_bd.columns:
        st.warning("Coluna do código do equipamento não encontrada na aba 'BD'. Algumas funcionalidades (ex.: vincular histórico por equipamento) ficarão limitadas.")
    else:
        # garante tipo numérico quando for possível
        df_bd["Cod_Equip"] = pd.to_numeric(df_bd["Cod_Equip"], errors="coerce")

    # padroniza nomes de colunas do BD caso sejam diferentes dos esperados
    expected_bd_cols = [
        "Data", "Cod_Equip", "Descricao_Equip", "Qtde_Litros", "Km_Hs_Rod",
        "Media", "Media_P", "Perc_Media", "Ton_Cana", "Litros_Ton",
        "Ref1", "Ref2", "Unidade", "Safra", "Mes_Excel", "Semana_Excel",
        "Classe_Original", "Classe_Operacional", "Descricao_Proprietario_Original",
        "Potencia_CV_Abast"
    ]
    # Se o BD tiver número exato de colunas esperado podemos renomear diretamente com base na ordem,
    # mas isso é perigoso. Verificamos se já existem colunas parecidas.
    # Se BD já tem "Data" e "Qtde_Litros" e "Km_Hs_Rod", assumimos que está ok; caso contrário, não reescrevemos index-based.
    if all(c in df_bd.columns for c in ["Data", "Qtde_Litros", "Km_Hs_Rod"]):
        # já ok - nada a fazer aqui
        pass
    else:
        # se não tiver, tentamos não sobrescrever — apenas avisamos
        st.info("A aba 'BD' não contém as colunas esperadas padrão. O app tentará usar as colunas encontradas (Data, Qtde_Litros, Km_Hs_Rod, Unidade...).")

    # merge histórico com frotas quando houver Cod_Equip em BD
    if "Cod_Equip" in df_bd.columns:
        df = pd.merge(df_bd, df_frotas, on="Cod_Equip", how="left", suffixes=("", "_frota"))
    else:
        # não consegue relacionar por equipamento — mantemos histórico standalone
        df = df_bd.copy()
        # traz algumas colunas de frota sem vínculo (não possível)
    # garante datetime
    if "Data" in df.columns:
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
        df = df.dropna(subset=["Data"])
    # derivados
    if "Data" in df.columns:
        df["Mes"] = df["Data"].dt.month
        df["Semana"] = df["Data"].dt.isocalendar().week
        df["Ano"] = df["Data"].dt.year
        df["AnoMes"] = df["Data"].dt.to_period("M").astype(str)
        df["AnoSemana"] = df["Data"].dt.strftime("%Y-%U")

    # garante numérico em colunas comuns
    for col in ["Qtde_Litros", "Media", "Media_P", "Km_Hs_Rod", COL_KM_HR_PREFERRED]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # cria coluna de Valor_Atual para manutenção:
    # prioridade: COL_KM_HR_PREFERRED da BD; senão, última Km_Hs_Rod por equipamento (quando disponível)
    if COL_KM_HR_PREFERRED in df.columns and not df[COL_KM_HR_PREFERRED].isna().all():
        df["Valor_Atual"] = pd.to_numeric(df[COL_KM_HR_PREFERRED], errors="coerce")
    else:
        # se houver Cod_Equip, pega última Km_Hs_Rod por Cod_Equip
        if "Cod_Equip" in df.columns and "Km_Hs_Rod" in df.columns:
            last_km = df.sort_values(["Cod_Equip", "Data"]).groupby("Cod_Equip")["Km_Hs_Rod"].last()
            df = df.merge(last_km.rename("Km_Last_from_hist"), on="Cod_Equip", how="left")
            df["Valor_Atual"] = df["Km_Last_from_hist"]
        else:
            # fallback geral: tenta usar Km_Hs_Rod direto (sem agrupar)
            if "Km_Hs_Rod" in df.columns:
                df["Valor_Atual"] = pd.to_numeric(df["Km_Hs_Rod"], errors="coerce")
            else:
                df["Valor_Atual"] = np.nan

    # normalize Unidade col name
    unid_col_found = find_column(df, UNID_CANDIDATES)
    if unid_col_found and unid_col_found != "Unidade":
        df = df.rename(columns={unid_col_found: "Unidade"})
    # finalizam mapeamentos
    mappings = {
        "cod_col_bd": "Cod_Equip" if "Cod_Equip" in df.columns else None,
        "class_col": "Classe_Operacional" if "Classe_Operacional" in df.columns else None,
        "unid_col": "Unidade" if "Unidade" in df.columns else None,
        "kmhr_col": COL_KM_HR_PREFERRED if COL_KM_HR_PREFERRED in df.columns else None
    }
    return df, df_frotas, mappings

# grava log de manutenção
def save_maintenance_log(excel_path: str, df_entries: pd.DataFrame, sheet_name: str = MANUT_LOG_SHEET) -> bool:
    """Append (ou cria) sheet MANUTENCAO_LOG com as ações registradas."""
    try:
        df_entries = df_entries.copy()
        if "Timestamp" in df_entries.columns:
            df_entries["Timestamp"] = pd.to_datetime(df_entries["Timestamp"])
        # criar arquivo se não existe
        if not Path(excel_path).exists():
            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                df_entries.to_excel(writer, sheet_name=sheet_name, index=False)
            return True
        # se existe, le sheet existente se houver e concatena
        book = load_workbook(excel_path)
        if sheet_name in book.sheetnames:
            existing = pd.read_excel(excel_path, sheet_name=sheet_name)
            combined = pd.concat([existing, df_entries], ignore_index=True)
        else:
            combined = df_entries
        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            combined.to_excel(writer, sheet_name=sheet_name, index=False)
        return True
    except Exception as e:
        st.error(f"Falha ao escrever log de manutenção: {e}")
        return False

# construção da tabela de manutenção
def build_maintenance_dataframe(df_frotas: pd.DataFrame, df_bd: pd.DataFrame, class_intervals: dict,
                                km_default: int, hr_default: int) -> pd.DataFrame:
    """
    Retorna df com colunas:
     - Cod_Equip, DESCRICAO_EQUIPAMENTO, Km_Hr_Atual, Unidade, Rev1_Next, Rev1_To_Go, Rev2_..., Rev3_...
    class_intervals: {classe: {"rev_km":[r1,r2,r3], "rev_hr":[h1,h2,h3]}}
    """
    mf = df_frotas.copy()
    # index por Cod_Equip
    if "Cod_Equip" not in mf.columns:
        st.error("Coluna 'Cod_Equip' não encontrada em FROTAS — manutenção não pode ser montada.")
        return pd.DataFrame()

    mf = mf.set_index("Cod_Equip")

    # obtém último Valor_Atual por equipamento do histórico (df_bd)
    if "Cod_Equip" in df_bd.columns and "Valor_Atual" in df_bd.columns:
        last_vals = df_bd.sort_values(["Cod_Equip", "Data"]).groupby("Cod_Equip")["Valor_Atual"].last()
    else:
        last_vals = pd.Series(dtype=float)

    mf["Km_Hr_Atual"] = last_vals.reindex(mf.index).astype(float)

    # obtém Unidade (Unidade da última entrada)
    if "Cod_Equip" in df_bd.columns and "Unidade" in df_bd.columns:
        last_unid = df_bd.sort_values(["Cod_Equip", "Data"]).groupby("Cod_Equip")["Unidade"].last()
    else:
        last_unid = pd.Series(dtype=object)
    mf["Unidade"] = last_unid.reindex(mf.index).fillna("").astype(str)

    mf = mf.reset_index()

    # para cada revisão define intervalos por classe
    revs = [1, 2, 3]
    for r in revs:
        mf[f"Rev{r}_Interval_KM"] = mf["Classe_Operacional"].apply(
            lambda cls: class_intervals.get(cls, {}).get("rev_km", [km_default, km_default*2, km_default*3])[r-1]
            if isinstance(class_intervals.get(cls, {}).get("rev_km", None), list)
            else (class_intervals.get(cls, {}).get("rev_km") or km_default)
        )
        mf[f"Rev{r}_Interval_HR"] = mf["Classe_Operacional"].apply(
            lambda cls: class_intervals.get(cls, {}).get("rev_hr", [hr_default, hr_default*2, hr_default*3])[r-1]
            if isinstance(class_intervals.get(cls, {}).get("rev_hr", None), list)
            else (class_intervals.get(cls, {}).get("rev_hr") or hr_default)
        )

    # calcula próximo e to_go de acordo com unidade (QUILÔMETROS/HORAS) — se unidade estiver ambígua usa KM por padrão
    def calc_next_to_go(row, r):
        cur = row.get("Km_Hr_Atual", np.nan)
        unit = str(row.get("Unidade", "")).upper() if pd.notna(row.get("Unidade")) else ""
        if pd.isna(cur):
            return (np.nan, np.nan)
        if "QUIL" in unit or "KM" in unit:
            interval = row.get(f"Rev{r}_Interval_KM", np.nan)
            nxt = cur + (interval if not pd.isna(interval) else np.nan)
            return (nxt, nxt - cur)
        elif "HOR" in unit or "HR" in unit:
            interval = row.get(f"Rev{r}_Interval_HR", np.nan)
            nxt = cur + (interval if not pd.isna(interval) else np.nan)
            return (nxt, nxt - cur)
        else:
            # fallback km
            interval = row.get(f"Rev{r}_Interval_KM", np.nan)
            nxt = cur + (interval if not pd.isna(interval) else np.nan)
            return (nxt, nxt - cur)

    for r in revs:
        mf[[f"Rev{r}_Next", f"Rev{r}_To_Go"]] = mf.apply(lambda row: pd.Series(calc_next_to_go(row, r)), axis=1)

    mf["Due_Rev"] = False
    mf["Due_Oil"] = False
    mf["Any_Due"] = False
    return mf

# ---------------- App principal ----------------
def main():
    st.set_page_config(page_title="Dashboard Frotas & Manutenção", layout="wide")
    st.title("📊 Dashboard de Frotas & Manutenção")

    # Carrega dados
    df, df_frotas, maps = load_data(EXCEL_PATH)

    # Sidebar
    with st.sidebar:
        st.header("Configurações")
        dark = st.checkbox("Dark mode", value=False)
        st.markdown("---")
        st.header("Visual")
        top_n = st.slider("Top N classes antes de agrupar em 'Outros'", min_value=3, max_value=30, value=10)
        hide_text_threshold = st.slider("Esconder valores nas barras quando #categorias >", min_value=5, max_value=40, value=8)
        st.markdown("---")
        st.header("Manutenção - defaults")
        km_default = st.number_input("Intervalo padrão (km) revisão", min_value=100, max_value=200000, value=10000, step=100)
        hr_default = st.number_input("Intervalo padrão (horas) lubrificação", min_value=1, max_value=5000, value=250, step=1)
        km_due_threshold = st.number_input("Alerta revisão se faltar <= (km)", min_value=10, max_value=5000, value=500, step=10)
        hr_due_threshold = st.number_input("Alerta lubrificação se faltar <= (horas)", min_value=1, max_value=500, value=20, step=1)

    palette = PALETTE_DARK if dark else PALETTE_LIGHT
    plotly_template = "plotly_dark" if dark else "plotly"

    # estado: intervals por classe
    if "manut_by_class" not in st.session_state:
        st.session_state.manut_by_class = {}
        classes = sorted(df["Classe_Operacional"].dropna().unique()) if "Classe_Operacional" in df.columns else []
        for c in classes:
            st.session_state.manut_by_class[c] = {
                "rev_km": [km_default, km_default*2, km_default*3],
                "rev_hr": [hr_default, hr_default*2, hr_default*3]
            }

    # abas
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Análise", "🔎 Ficha", "📋 Tabela", "⚙️ Configurações", "🛠️ Manutenção"])

    # --- Aba 1: análise simplificada ---
    with tab1:
        st.header("Análise de Consumo")
        if "Classe_Operacional" not in df.columns:
            st.warning("Coluna 'Classe Operacional' não encontrada no BD. Alguns gráficos podem não estar disponíveis.")
        # agrupamento com "Outros"
        df_plot = df.copy()
        if "Classe_Operacional" in df_plot.columns:
            df_plot["Classe_Operacional"] = df_plot["Classe_Operacional"].fillna("Sem Classe")
            df_plot["Classe_Grouped"] = df_plot["Classe_Operacional"].apply(lambda s: "Outros" if s in OUTROS_CLASSES else s)
            media = df_plot.groupby("Classe_Grouped")["Media"].mean().reset_index().sort_values("Media", ascending=False)
            if media.shape[0] > top_n:
                top_keep = media.head(top_n)["Classe_Grouped"].tolist()
                df_plot["Classe_TopN"] = df_plot["Classe_Grouped"].apply(lambda s: s if s in top_keep else "Outros")
                media2 = df_plot.groupby("Classe_TopN")["Media"].mean().reset_index().rename(columns={"Classe_TopN":"Classe_Grouped"})
                media2["Classe_wrapped"] = media2["Classe_Grouped"].apply(lambda s: wrap_labels(s, width=16))
                fig = px.bar(media2, x="Classe_wrapped", y="Media", text="Media", color_discrete_sequence=palette)
            else:
                media["Classe_wrapped"] = media["Classe_Grouped"].apply(lambda s: wrap_labels(s, width=16))
                fig = px.bar(media, x="Classe_wrapped", y="Media", text="Media", color_discrete_sequence=palette)
            fig.update_layout(template=plotly_template, margin=dict(b=140))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sem 'Classe Operacional' — exiba a aba Tabela para ver dados brutos.")

    # --- Aba 2: ficha individual ---
    with tab2:
        st.header("Ficha Individual do Equipamento")
        equip_opts = df_frotas.sort_values("Cod_Equip")["label"].tolist()
        sel = st.selectbox("Selecione equipamento", options=[""] + equip_opts)
        if sel:
            cod = int(sel.split(" - ")[0])
            row = df_frotas.query("Cod_Equip == @cod").iloc[0]
            st.subheader(f"{row.get('DESCRICAO_EQUIPAMENTO','–')} ({row.get('PLACA','–')})")
            # valor atual do histórico
            if "Cod_Equip" in df.columns:
                last = df.sort_values(["Cod_Equip", "Data"]).query("Cod_Equip == @cod").groupby("Cod_Equip")["Valor_Atual"].last()
                val = last.iloc[0] if not last.empty else np.nan
                unidade = df.sort_values(["Cod_Equip", "Data"]).query("Cod_Equip == @cod").groupby("Cod_Equip")["Unidade"].last().iloc[0] if not last.empty else ""
            else:
                val = np.nan
                unidade = ""
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Status", row.get("ATIVO", "–"))
            c2.metric("Placa", row.get("PLACA", "–"))
            c3.metric("Medida Atual", f"{formatar_brasileiro(val)} {unidade}")
            if "Cod_Equip" in df.columns:
                consumo_eq = df.query("Cod_Equip == @cod")
                c4.metric("Média Geral (km/l)", formatar_brasileiro(consumo_eq["Media"].mean() if "Media" in consumo_eq.columns else np.nan))
            else:
                c4.metric("Média Geral (km/l)", "–")

    # --- Aba 3: tabela ---
    with tab3:
        st.header("Tabela detalhada")
        cols = ["Data", "Cod_Equip", "Descricao_Equip", "PLACA", "DESCRICAOMARCA", "ANOMODELO", "Qtde_Litros", "Media", "Media_P", "Classe_Operacional", "Valor_Atual", "Unidade"]
        present = [c for c in cols if c in df.columns or c in df.columns]
        df_tab = df[[c for c in cols if c in df.columns]]
        st.download_button("⬇️ Exportar CSV", df_tab.to_csv(index=False).encode("utf-8"), "abastecimentos.csv", "text/csv")
        gb = GridOptionsBuilder.from_dataframe(df_tab)
        gb.configure_default_column(filterable=True, sortable=True, resizable=True)
        if "Media" in df_tab.columns:
            gb.configure_column("Media", type=["numericColumn"], precision=1)
        if "Qtde_Litros" in df_tab.columns:
            gb.configure_column("Qtde_Litros", type=["numericColumn"], precision=1)
        gb.configure_pagination(paginationAutoPageSize=False, paginationPageSize=15)
        gb.configure_selection("single", use_checkbox=True)
        AgGrid(df_tab, gridOptions=gb.build(), height=520, allow_unsafe_jscode=True)

    # --- Aba 4: Configurações por classe ---
    with tab4:
        st.header("Configurações: intervalos por classe")
        classes = sorted(df["Classe_Operacional"].dropna().unique()) if "Classe_Operacional" in df.columns else []
        st.markdown("Defina intervalos (KM / HR) para as 3 revisões por classe. Valores salvos na sessão do app.")
        for cls in classes:
            st.subheader(str(cls))
            existing = st.session_state.manut_by_class.get(cls, {"rev_km":[km_default, km_default*2, km_default*3],"rev_hr":[hr_default, hr_default*2, hr_default*3]})
            k1 = st.number_input(f"{cls} → Rev1 (km)", min_value=0, max_value=10**7, value=int(existing["rev_km"][0]), key=f"{cls}_r1km")
            k2 = st.number_input(f"{cls} → Rev2 (km)", min_value=0, max_value=10**7, value=int(existing["rev_km"][1]), key=f"{cls}_r2km")
            k3 = st.number_input(f"{cls} → Rev3 (km)", min_value=0, max_value=10**7, value=int(existing["rev_km"][2]), key=f"{cls}_r3km")
            h1 = st.number_input(f"{cls} → Rev1 (hr)", min_value=0, max_value=10**7, value=int(existing["rev_hr"][0]), key=f"{cls}_r1hr")
            h2 = st.number_input(f"{cls} → Rev2 (hr)", min_value=0, max_value=10**7, value=int(existing["rev_hr"][1]), key=f"{cls}_r2hr")
            h3 = st.number_input(f"{cls} → Rev3 (hr)", min_value=0, max_value=10**7, value=int(existing["rev_hr"][2]), key=f"{cls}_r3hr")
            st.session_state.manut_by_class[cls] = {"rev_km":[int(k1), int(k2), int(k3)], "rev_hr":[int(h1), int(h2), int(h3)]}

    # --- Aba 5: Manutenção ---
    with tab5:
        st.header("Manutenção & Lubrificação")
        st.markdown("Lista de equipamentos com revisões ou lubrificações próximas/vencidas. Marque as ações e salve (irá gravar em aba MANUTENCAO_LOG).")

        # monta dict de intervals por classe
        class_intervals = {}
        for k, v in st.session_state.manut_by_class.items():
            class_intervals[k] = {"rev_km": v.get("rev_km", []), "rev_hr": v.get("rev_hr", [])}

        mf = build_maintenance_dataframe(df_frotas, df, class_intervals, int(km_default), int(hr_default))

        if mf.empty:
            st.info("Não foi possível montar o plano de manutenção (verifique colunas 'Cod_Equip' em FROTAS e BD).")
            return

        # avalia due flags conforme thresholds do sidebar
        def compute_flags(row):
            due_rev = False
            due_oil = False
            unit = str(row.get("Unidade","")).upper() if pd.notna(row.get("Unidade")) else ""
            for r in (1,2,3):
                togo = row.get(f"Rev{r}_To_Go", np.nan)
                if pd.isna(togo):
                    continue
                if "QUIL" in unit or "KM" in unit:
                    if togo <= km_due_threshold:
                        due_rev = True
                elif "HOR" in unit or "HR" in unit:
                    if togo <= hr_due_threshold:
                        due_oil = True
                else:
                    # fallback km
                    if togo <= km_due_threshold:
                        due_rev = True
            return pd.Series({"Due_Rev": due_rev, "Due_Oil": due_oil})

        flags = mf.apply(compute_flags, axis=1)
        mf["Due_Rev"] = flags["Due_Rev"]
        mf["Due_Oil"] = flags["Due_Oil"]
        mf["Any_Due"] = mf["Due_Rev"] | mf["Due_Oil"]

        df_due = mf[mf["Any_Due"]].copy().sort_values(["Due_Rev", "Due_Oil"], ascending=False)
        st.subheader("Equipamentos com ação recomendada")
        st.write(f"Total: {len(df_due)}")
        if df_due.empty:
            st.info("Nenhum equipamento com alerta dentro dos thresholds configurados.")
        else:
            display_cols = ["Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Unidade", "Km_Hr_Atual",
                            "Rev1_To_Go", "Rev2_To_Go", "Rev3_To_Go", "Due_Rev", "Due_Oil"]
            avail = [c for c in display_cols if c in df_due.columns]
            st.dataframe(df_due[avail].reset_index(drop=True), use_container_width=True)

            # ações com checkboxes
            st.markdown("---")
            st.markdown("**Marcar revisões/lubrificações como concluídas**")
            actions = []
            for i, row in df_due.reset_index(drop=True).iterrows():
                cod = row["Cod_Equip"]
                name = row.get("DESCRICAO_EQUIPAMENTO","")
                unit = row.get("Unidade","")
                cur = row.get("Km_Hr_Atual", np.nan)
                st.markdown(f"**{int(cod)} - {name}** — Atual: {formatar_brasileiro(cur)} {unit}")
                c1, c2, c3, c4 = st.columns([1,1,1,1])
                cb1 = c1.checkbox("Rev1", key=f"rev1_{cod}")
                cb2 = c2.checkbox("Rev2", key=f"rev2_{cod}")
                cb3 = c3.checkbox("Rev3", key=f"rev3_{cod}")
                c4b = c4.checkbox("Lubrificação", key=f"lub_{cod}")
                if cb1: actions.append({"Cod_Equip": cod, "Tipo": "Rev1", "Valor_Atual": cur, "Unidade": unit})
                if cb2: actions.append({"Cod_Equip": cod, "Tipo": "Rev2", "Valor_Atual": cur, "Unidade": unit})
                if cb3: actions.append({"Cod_Equip": cod, "Tipo": "Rev3", "Valor_Atual": cur, "Unidade": unit})
                if c4b: actions.append({"Cod_Equip": cod, "Tipo": "Lubrificacao", "Valor_Atual": cur, "Unidade": unit})

            if st.button("Salvar ações (grava MANUTENCAO_LOG)"):
                if not actions:
                    st.info("Nenhuma ação selecionada.")
                else:
                    rows = []
                    now = datetime.now()
                    for a in actions:
                        rows.append({
                            "Timestamp": now,
                            "Cod_Equip": a["Cod_Equip"],
                            "Tipo": a["Tipo"],
                            "Valor_Atual": a["Valor_Atual"],
                            "Unidade": a["Unidade"],
                            "Usuario": st.session_state.get("usuario","(anon)")
                        })
                    entries = pd.DataFrame(rows)
                    ok = save_maintenance_log(EXCEL_PATH, entries, MANUT_LOG_SHEET)
                    if ok:
                        st.success(f"{len(rows)} ação(ões) gravada(s) em `{MANUT_LOG_SHEET}`.")
                        # refresh to clear checkboxes
                        st.experimental_rerun()
                    else:
                        st.error("Falha ao gravar. Verifique permissões / caminho do arquivo.")

        st.markdown("---")
        st.subheader("Visão geral - plano de manutenção")
        overview_cols = ["Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Unidade", "Km_Hr_Atual",
                         "Rev1_Next", "Rev1_To_Go", "Rev2_Next", "Rev2_To_Go", "Rev3_Next", "Rev3_To_Go"]
        avail_over = [c for c in overview_cols if c in mf.columns]
        st.dataframe(mf[avail_over].sort_values("Cod_Equip").reset_index(drop=True), use_container_width=True)
        st.download_button("Exportar CSV - Plano de Manutenção", mf[avail_over].to_csv(index=False).encode("utf-8"), "plano_manutencao.csv", "text/csv")

if __name__ == "__main__":
    main()
