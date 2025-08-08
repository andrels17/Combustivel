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

# Paletas
PALETTE_LIGHT = px.colors.sequential.Blues_r
PALETTE_DARK = px.colors.sequential.Plasma_r

# Classes agrupadas em 'Outros' (conforme pedido)
OUTROS_CLASSES = {"Motocicletas", "Mini Carregadeira", "Usina", "Veiculos Leves"}

# Nome exato (opcional) da coluna única que conterá hodômetro / horímetro atual na aba BD
COL_KM_HR_ATUAL = "KM_HR_Atual"  # (se existir, será usado; senão o app faz fallback)

# Nome da aba de log de manutenção
MANUT_LOG_SHEET = "MANUTENCAO_LOG"

# ---------------- Utilitários ----------------
def formatar_brasileiro(valor: float) -> str:
    if pd.isna(valor) or not np.isfinite(valor):
        return "–"
    return "{:,.2f}".format(valor).replace(",", "X").replace(".", ",").replace("X", ".")

def wrap_labels(s: str, width: int = 16) -> str:
    if pd.isna(s):
        return ""
    parts = textwrap.wrap(str(s), width=width)
    return "<br>".join(parts) if parts else str(s)

def safe_rename_if_same_len(df: pd.DataFrame, new_cols: list[str]) -> pd.DataFrame:
    """
    Renomeia colunas apenas se o número de colunas bater com new_cols.
    Retorna df inalterado em caso contrário (evita ValueError).
    """
    if len(df.columns) == len(new_cols):
        df.columns = new_cols
    return df

# ---------------- Carregamento & preparação ----------------
@st.cache_data(show_spinner="Carregando e processando dados...")
def load_data(path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Carrega as sheets 'BD' e 'FROTAS' com tratamento defensivo:
     - se 'BD' não tem os nomes esperados (quantidade diferente), não força renomeação que quebra.
     - calcula/garante colunas derivadas (Data, Mes, Ano, AnoMes, etc).
     - calcula 'Media' automaticamente a partir de Unid/Qtde_Litros/Km_Hs_Rod (hodômetro ou horímetro).
     - cria coluna unificada 'Valor_Atual' priorizando KM_HR_Atual se presente.
    """
    expected_abast_cols = [
        "Data", "Cod_Equip", "Descricao_Equip", "Qtde_Litros", "Km_Hs_Rod",
        "Media", "Media_P", "Perc_Media", "Ton_Cana", "Litros_Ton",
        "Ref1", "Ref2", "Unidade", "Safra", "Mes_Excel", "Semana_Excel",
        "Classe_Original", "Classe_Operacional", "Descricao_Proprietario_Original",
        "Potencia_CV_Abast"
    ]

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

    # Normaliza frotas (coluna COD_EQUIPAMENTO => Cod_Equip)
    if "COD_EQUIPAMENTO" in df_frotas.columns and "Cod_Equip" not in df_frotas.columns:
        df_frotas = df_frotas.rename(columns={"COD_EQUIPAMENTO": "Cod_Equip"})
    # garantir Cod_Equip presente (se não, tentar outras candidatas)
    if "Cod_Equip" not in df_frotas.columns:
        # tenta heurística: colunas com 'COD' ou 'EQUIP' no nome
        for c in df_frotas.columns:
            cl = c.upper()
            if "COD" in cl or "EQUIP" in cl:
                df_frotas = df_frotas.rename(columns={c: "Cod_Equip"})
                break

    # deduplicate e transformar ano modelo
    if "Cod_Equip" in df_frotas.columns:
        df_frotas = df_frotas.drop_duplicates(subset=["Cod_Equip"])
        df_frotas["ANOMODELO"] = pd.to_numeric(df_frotas.get("ANOMODELO", pd.Series()), errors="coerce")
    else:
        # garante index como equipamento se nada encontrado
        df_frotas = df_frotas.reset_index().rename(columns={"index": "Cod_Equip"})
        df_frotas["ANOMODELO"] = pd.to_numeric(df_frotas.get("ANOMODELO", pd.Series()), errors="coerce")

    # cria label amigável na frota
    df_frotas["label"] = (
        df_frotas["Cod_Equip"].astype(str)
        + " - "
        + df_frotas.get("DESCRICAO_EQUIPAMENTO", "").fillna("")
        + " ("
        + df_frotas.get("PLACA", "").fillna("Sem Placa")
        + ")"
    )

    # Tentar renomear df_abast quando o número de colunas bate com o esperado.
    df_abast = safe_rename_if_same_len(df_abast, expected_abast_cols)

    # Se o DataFrame não tem 'Data' ou 'Cod_Equip' sob os nomes esperados, tentar heurísticas
    if "Data" not in df_abast.columns:
        # tenta encontrar primeira coluna com tipo datetime ou com nome similar
        for c in df_abast.columns:
            if "data" in c.lower() or "date" in c.lower():
                df_abast = df_abast.rename(columns={c: "Data"})
                break

    if "Cod_Equip" not in df_abast.columns:
        for c in df_abast.columns:
            cl = c.upper()
            if "COD" in cl and ("EQUIP" in cl or "EQUI" in cl or "EQP" in cl):
                df_abast = df_abast.rename(columns={c: "Cod_Equip"})
                break

    # Garantias mínimas: se não existe Cod_Equip criamos um índice incremental (perigoso para merge, o app tentará enriquecer o máximo)
    if "Cod_Equip" not in df_abast.columns:
        df_abast = df_abast.reset_index().rename(columns={"index":"Cod_Equip"})

    # Normalizações de tipos e nomes alternativos
    # coluna 'Unidade' pode aparecer como 'Unid' -> padroniza
    if "Unid" in df_abast.columns and "Unidade" not in df_abast.columns:
        df_abast = df_abast.rename(columns={"Unid": "Unidade"})

    # Garantir colunas numéricas quando presentes
    for col in ["Qtde_Litros", "Km_Hs_Rod", "Media", "Media_P"]:
        if col in df_abast.columns:
            df_abast[col] = pd.to_numeric(df_abast[col], errors="coerce")

    # Merge para enriquecer abast com dados de frota (Se houver Cod_Equip compatível)
    if "Cod_Equip" in df_abast.columns and "Cod_Equip" in df_frotas.columns:
        df = pd.merge(df_abast, df_frotas, on="Cod_Equip", how="left", suffixes=("", "_frota"))
    else:
        # sem chave comum, deixa df igual a df_abast e junta frotas via concat com chave ausente
        df = df_abast.copy()

    # Parse Data e filtros
    if "Data" in df.columns:
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
        df = df.dropna(subset=["Data"])
    else:
        # se não existe coluna Data, cria Data com hoje para preservar fluxo (evita erros)
        df["Data"] = pd.to_datetime(datetime.now())

    df["Mes"] = df["Data"].dt.month
    df["Semana"] = df["Data"].dt.isocalendar().week
    df["Ano"] = df["Data"].dt.year
    df["AnoMes"] = df["Data"].dt.to_period("M").astype(str)
    df["AnoSemana"] = df["Data"].dt.strftime("%Y-%U")

    # Marca / Fazenda
    if "Ref2" in df.columns:
        df["DESCRICAOMARCA"] = df["Ref2"].astype(str)
    if "Ref1" in df.columns:
        df["Fazenda"] = df["Ref1"].astype(str)

    # --- Calcular Media automaticamente a partir da Unid ---
    # Possíveis nomes: 'Unidade' (padronizado), 'Unid' etc.
    unidade_col = None
    for candidate in ["Unidade", "Unid", "UNIDADE", "UNID"]:
        if candidate in df.columns:
            unidade_col = candidate
            break

    # se Qtde_Litros e Km_Hs_Rod existem, tentamos criar Media coerente
    if "Qtde_Litros" in df.columns and "Km_Hs_Rod" in df.columns:
        # padronizar Unidade string e usar contains para decidir km/hr
        def compute_media(row):
            q = row.get("Qtde_Litros", np.nan)
            km_hr = row.get("Km_Hs_Rod", np.nan)
            unit = ""
            if unidade_col:
                try:
                    unit = str(row.get(unidade_col, "")).upper()
                except Exception:
                    unit = ""
            # se Qtde_Litros inválido ou zero -> NaN
            if pd.isna(q) or q == 0:
                return np.nan
            # se unidade contém 'QUIL' ou 'KM' => use km/hr as km
            if "QUIL" in unit or "KM" in unit:
                return km_hr / q if pd.notna(km_hr) else np.nan
            # se unidade contém 'HOR' or 'HR' => use km_hr as horas
            if "HOR" in unit or "HR" in unit:
                return km_hr / q if pd.notna(km_hr) else np.nan
            # fallback: if existing Media column is present, keep it
            if "Media" in df.columns and pd.notna(row.get("Media", np.nan)):
                return row.get("Media", np.nan)
            # otherwise try km/h logic
            return km_hr / q if pd.notna(km_hr) else np.nan

        df["Media"] = df.apply(compute_media, axis=1)
    else:
        # se não há as colunas, tenta não quebrar: cria coluna Media com NaNs ou usa existente
        if "Media" not in df.columns:
            df["Media"] = np.nan

    # --- Unifica um valor atual usado em manutenção (Valor_Atual)
    # Prioriza COL_KM_HR_ATUAL (se existir na BD), senão pega último Km_Hs_Rod por equipamento como fallback.
    if COL_KM_HR_ATUAL in df.columns:
        df[COL_KM_HR_ATUAL] = pd.to_numeric(df[COL_KM_HR_ATUAL], errors="coerce")
        df["Valor_Atual"] = df[COL_KM_HR_ATUAL]
    else:
        # último Km_Hs_Rod por Cod_Equip
        if "Cod_Equip" in df.columns and "Km_Hs_Rod" in df.columns:
            last_km = df.sort_values(["Cod_Equip", "Data"]).groupby("Cod_Equip")["Km_Hs_Rod"].last().rename("Km_Last")
            # merge back by Cod_Equip if possible
            if "Cod_Equip" in df.columns:
                df = df.merge(last_km, on="Cod_Equip", how="left")
                df["Valor_Atual"] = df["Km_Last"]
            else:
                df["Valor_Atual"] = df["Km_Hs_Rod"]
        else:
            df["Valor_Atual"] = np.nan

    return df, df_frotas

# ---------------- Excel logging ----------------
def save_maintenance_log(excel_path: str, entries_df: pd.DataFrame, sheet_name: str = MANUT_LOG_SHEET) -> bool:
    """
    Adiciona registros de manutenção à aba MANUTENCAO_LOG.
    Se a aba já existir, concatena; caso contrário cria-a.
    Retorna True se ok.
    """
    try:
        # normaliza Timestamp
        if "Timestamp" in entries_df.columns:
            entries_df["Timestamp"] = pd.to_datetime(entries_df["Timestamp"])
        else:
            entries_df["Timestamp"] = pd.to_datetime(datetime.now())

        # se arquivo não existe, cria com sheet novo
        if not Path(excel_path).exists():
            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                entries_df.to_excel(writer, sheet_name=sheet_name, index=False)
            return True

        # arquivo existe -> abrir workbook e verificar sheets
        book = load_workbook(excel_path)
        if sheet_name in book.sheetnames:
            existing = pd.read_excel(excel_path, sheet_name=sheet_name)
            combined = pd.concat([existing, entries_df], ignore_index=True)
        else:
            combined = entries_df

        # escrever substituindo a sheet
        with pd.ExcelWriter(excel_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            combined.to_excel(writer, sheet_name=sheet_name, index=False)
        return True
    except Exception as e:
        st.error(f"Falha ao salvar log de manutenção: {e}")
        return False

# ---------------- Funções manutenção ----------------
def build_maintenance_dataframe(df_frotas: pd.DataFrame, df_abast: pd.DataFrame,
                                class_intervals: dict, km_default: int, hr_default: int) -> pd.DataFrame:
    """
    Gera DataFrame com próximos serviços por equipamento.
    class_intervals: {classe: {"rev_km":[r1,r2,r3], "rev_hr":[r1,r2,r3] } }
    """
    mf = df_frotas.copy()
    # tenta obter último valor atual (Valor_Atual) por equipamento da aba BD
    last_values = pd.Series(dtype=float)
    if ("Cod_Equip" in df_abast.columns) and ("Valor_Atual" in df_abast.columns):
        last_values = df_abast.sort_values(["Cod_Equip", "Data"]).groupby("Cod_Equip")["Valor_Atual"].last()

    mf = mf.set_index("Cod_Equip", drop=False)
    mf["Km_Hr_Atual"] = last_values.reindex(mf.index)
    mf = mf.reset_index(drop=True)

    # tenta obter unidade por equipamento (última Unid/Unidade no histórico)
    last_unid = pd.Series(dtype=object)
    if "Unidade" in df_abast.columns and "Cod_Equip" in df_abast.columns:
        last_unid = df_abast.sort_values(["Cod_Equip", "Data"]).groupby("Cod_Equip")["Unidade"].last()
    mf = mf.set_index("Cod_Equip", drop=False)
    mf["Unid"] = last_unid.reindex(mf.index).values
    mf = mf.reset_index(drop=True)

    # garantir colunas de classe
    if "Classe_Operacional" not in mf.columns:
        mf["Classe_Operacional"] = np.nan

    # para cada revisão (1..3) obter intervalos por classe
    for r in [1, 2, 3]:
        def get_km_interval(cls):
            if pd.isna(cls):
                return km_default
            cfg = class_intervals.get(cls, {})
            rev_km = cfg.get("rev_km")
            if isinstance(rev_km, list) and len(rev_km) >= r:
                return rev_km[r - 1]
            if isinstance(rev_km, (int, float)):
                return rev_km
            # default progressive: r * km_default
            return km_default * r

        def get_hr_interval(cls):
            if pd.isna(cls):
                return hr_default
            cfg = class_intervals.get(cls, {})
            rev_hr = cfg.get("rev_hr")
            if isinstance(rev_hr, list) and len(rev_hr) >= r:
                return rev_hr[r - 1]
            if isinstance(rev_hr, (int, float)):
                return rev_hr
            return hr_default * r

        mf[f"Rev{r}_Interval_km"] = mf["Classe_Operacional"].apply(get_km_interval)
        mf[f"Rev{r}_Interval_hr"] = mf["Classe_Operacional"].apply(get_hr_interval)

    # calcula próximos e to_go conforme Unid (KM or HORAS)
    def calc_next_for_row(row, r):
        cur = row.get("Km_Hr_Atual", np.nan)
        unid = str(row.get("Unid", "")).upper() if pd.notna(row.get("Unid", "")) else ""
        if pd.isna(cur):
            return (np.nan, np.nan)
        if "QUIL" in unid or "KM" in unid:
            interval = row.get(f"Rev{r}_Interval_km", np.nan)
            nxt = cur + (interval if not pd.isna(interval) else np.nan)
            return (nxt, nxt - cur)
        if "HOR" in unid or "HR" in unid:
            interval = row.get(f"Rev{r}_Interval_hr", np.nan)
            nxt = cur + (interval if not pd.isna(interval) else np.nan)
            return (nxt, nxt - cur)
        # fallback: tratar como km
        interval = row.get(f"Rev{r}_Interval_km", np.nan)
        nxt = cur + (interval if not pd.isna(interval) else np.nan)
        return (nxt, nxt - cur)

    for r in [1, 2, 3]:
        mf[[f"Rev{r}_Next", f"Rev{r}_To_Go"]] = mf.apply(lambda row: pd.Series(calc_next_for_row(row, r)), axis=1)

    # flags iniciais (serão atualizadas no contexto do app com thresholds)
    mf["Due_Rev"] = False
    mf["Due_Oil"] = False
    mf["Any_Due"] = False

    return mf

# ---------------- Layout / CSS ----------------
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

# ---------------- App principal ----------------
def main():
    st.set_page_config(page_title="Dashboard de Frotas e Abastecimentos", layout="wide")
    st.title("📊 Dashboard de Frotas e Abastecimentos — Manutenção Integrada")

    # Carrega dados
    df, df_frotas = load_data(EXCEL_PATH)

    # Sidebar: controles
    with st.sidebar:
        st.header("Configurações")
        dark_mode = st.checkbox("🕶️ Dark Mode", value=False)
        st.markdown("---")
        st.header("Visual")
        top_n = st.slider("Top N classes antes de agrupar 'Outros'", min_value=3, max_value=30, value=10)
        hide_text_threshold = st.slider("Esconder valores nas barras quando categorias >", min_value=5, max_value=40, value=8)
        st.markdown("---")
        st.header("Manutenção - Defaults")
        km_interval_default = st.number_input("Intervalo padrão (km) revisão", min_value=100, max_value=200000, value=10000, step=100)
        hr_interval_default = st.number_input("Intervalo padrão (horas) lubrificação", min_value=1, max_value=5000, value=250, step=1)
        km_due_threshold = st.number_input("Alerta revisão se faltar <= (km)", min_value=10, max_value=5000, value=500, step=10)
        hr_due_threshold = st.number_input("Alerta lubrificação se faltar <= (horas)", min_value=1, max_value=500, value=20, step=1)
        st.markdown("---")
        if st.button("🔄 Limpar sessão"):
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.experimental_rerun()

    apply_modern_css(dark_mode)
    palette = PALETTE_DARK if dark_mode else PALETTE_LIGHT
    plotly_template = "plotly_dark" if dark_mode else "plotly"

    # Inicializa intervalos por classe em session_state (se inexistente)
    if "manut_by_class" not in st.session_state:
        st.session_state.manut_by_class = {}
        classes = sorted(df["Classe_Operacional"].dropna().unique()) if "Classe_Operacional" in df.columns else []
        for cls in classes:
            st.session_state.manut_by_class[cls] = {
                "rev_km": [km_interval_default, km_interval_default*2, km_interval_default*3],
                "rev_hr": [hr_interval_default, hr_interval_default*2, hr_interval_default*3],
            }

    # Abas
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📊 Análise de Consumo",
        "🔎 Consulta de Frota",
        "📋 Tabela Detalhada",
        "⚙️ Configurações",
        "🛠️ Manutenção"
    ])

    # ---------- Aba 1: Análise ----------
    with tab1:
        st.header("Análise de Consumo")
        # defensivo: se Media não existe, cria vazia
        if "Media" not in df.columns:
            df["Media"] = np.nan

        # agrupamento e Outras classes
        df_plot = df.copy()
        # coluna de classe: tentar variantes se não encontrar
        cls_col = "Classe_Operacional" if "Classe_Operacional" in df_plot.columns else None
        if cls_col is None:
            # tenta achar alguma coluna parecida
            for c in df_plot.columns:
                if "CLASSE" in c.upper():
                    cls_col = c
                    break
        if cls_col is None:
            df_plot["Classe_Operacional"] = "Sem Classe"
            cls_col = "Classe_Operacional"

        df_plot[cls_col] = df_plot[cls_col].fillna("Sem Classe")
        df_plot["Classe_Grouped"] = df_plot[cls_col].apply(lambda s: "Outros" if s in OUTROS_CLASSES else s)

        # média por classe e top_n lógica
        media_op_full = df_plot.groupby("Classe_Grouped", dropna=False)["Media"].mean().reset_index()
        media_op_full["Media"] = media_op_full["Media"].round(1).fillna(0)
        media_sorted = media_op_full.sort_values("Media", ascending=False).reset_index(drop=True)
        if media_sorted.shape[0] > top_n:
            top_keep = media_sorted.head(top_n)["Classe_Grouped"].tolist()
            df_plot["Classe_TopN"] = df_plot["Classe_Grouped"].apply(lambda s: s if s in top_keep else "Outros")
            media_op = df_plot.groupby("Classe_TopN", dropna=False)["Media"].mean().reset_index().rename(columns={"Classe_TopN": "Classe_Grouped"})
            media_op["Media"] = media_op["Media"].round(1)
            # força colocar "Outros" no fim
            outros = media_op[media_op["Classe_Grouped"] == "Outros"]
            media_op = media_op[media_op["Classe_Grouped"] != "Outros"].sort_values("Media", ascending=False)
            if not outros.empty:
                media_op = pd.concat([media_op, outros], ignore_index=True)
        else:
            media_op = media_sorted

        media_op["Classe_wrapped"] = media_op["Classe_Grouped"].astype(str).apply(lambda s: wrap_labels(s, width=16))
        fig = px.bar(media_op, x="Classe_wrapped", y="Media", text="Media", color_discrete_sequence=palette)
        fig.update_layout(template=plotly_template, margin=dict(b=160, t=50))
        fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

    # ---------- Aba 2: Consulta ----------
    with tab2:
        st.header("Ficha Individual do Equipamento")
        if "label" in df_frotas.columns:
            options = df_frotas.sort_values("Cod_Equip")["label"].tolist()
        else:
            options = df_frotas.index.astype(str).tolist()
        equip_label = st.selectbox("Selecione o Equipamento", options=options)
        if equip_label:
            cod_sel = None
            try:
                cod_sel = int(str(equip_label).split(" - ")[0])
            except Exception:
                # tenta pegar index
                cod_sel = equip_label
            # procura equipamento
            row = None
            if "Cod_Equip" in df_frotas.columns:
                row = df_frotas[df_frotas["Cod_Equip"].astype(str) == str(cod_sel)]
                if not row.empty:
                    row = row.iloc[0]
            else:
                # fallback: primeira linha
                row = df_frotas.iloc[0]

            st.subheader(f"{row.get('DESCRICAO_EQUIPAMENTO','–')} ({row.get('PLACA','–')})")
            # busca último valor atual na aba BD
            val = np.nan
            unit = ""
            if "Cod_Equip" in df.columns:
                try:
                    df_sel = df[df["Cod_Equip"].astype(str) == str(row.get("Cod_Equip"))].sort_values("Data")
                    if not df_sel.empty:
                        val = df_sel["Valor_Atual"].dropna().iloc[-1] if "Valor_Atual" in df_sel.columns else np.nan
                        unit = df_sel["Unidade"].dropna().iloc[-1] if "Unidade" in df_sel.columns else ""
                except Exception:
                    val = np.nan
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Status", row.get("ATIVO", "–"))
            c2.metric("Placa", row.get("PLACA", "–"))
            c3.metric("Valor Atual", f"{formatar_brasileiro(val)} {unit}")
            # média geral (do histórico)
            if "Cod_Equip" in df.columns:
                consumo_eq = df[df["Cod_Equip"].astype(str) == str(row.get("Cod_Equip"))]
                c4.metric("Média Histórica", formatar_brasileiro(consumo_eq["Media"].mean() if "Media" in consumo_eq.columns else np.nan))
            else:
                c4.metric("Média Histórica", "–")

    # ---------- Aba 3: Tabela ----------
    with tab3:
        st.header("Tabela Detalhada de Abastecimentos")
        cols_request = ["Data", "Cod_Equip", "Descricao_Equip", "PLACA", "DESCRICAOMARCA", "ANOMODELO",
                        "Qtde_Litros", "Media", "Media_P", "Classe_Operacional", "Valor_Atual", "Unidade"]
        df_tab = df[[c for c in cols_request if c in df.columns]]
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
    with tab4:
        st.header("Padrões por Classe Operacional (Alertas & Intervalos)")
        classes = sorted(df["Classe_Operacional"].dropna().unique()) if "Classe_Operacional" in df.columns else []
        st.markdown("Ajuste intervalos por classe para Rev1/Rev2/Rev3 (km e horas).")
        for cls in classes:
            rev = st.session_state.manut_by_class.get(cls, {})
            rev_km = rev.get("rev_km", [km_interval_default, km_interval_default*2, km_interval_default*3])
            rev_hr = rev.get("rev_hr", [hr_interval_default, hr_interval_default*2, hr_interval_default*3])
            st.subheader(str(cls))
            c1, c2 = st.columns(2)
            with c1:
                nk1 = st.number_input(f"{cls} → Rev1 (km)", min_value=0, max_value=10**7, value=int(rev_km[0]), key=f"{cls}_r1km")
                nk2 = st.number_input(f"{cls} → Rev2 (km)", min_value=0, max_value=10**7, value=int(rev_km[1]), key=f"{cls}_r2km")
                nk3 = st.number_input(f"{cls} → Rev3 (km)", min_value=0, max_value=10**7, value=int(rev_km[2]), key=f"{cls}_r3km")
            with c2:
                nh1 = st.number_input(f"{cls} → Rev1 (hr)", min_value=0, max_value=10**7, value=int(rev_hr[0]), key=f"{cls}_r1hr")
                nh2 = st.number_input(f"{cls} → Rev2 (hr)", min_value=0, max_value=10**7, value=int(rev_hr[1]), key=f"{cls}_r2hr")
                nh3 = st.number_input(f"{cls} → Rev3 (hr)", min_value=0, max_value=10**7, value=int(rev_hr[2]), key=f"{cls}_r3hr")
            st.session_state.manut_by_class[cls] = {"rev_km":[int(nk1), int(nk2), int(nk3)], "rev_hr":[int(nh1), int(nh2), int(nh3)]}

    # ---------- Aba 5: Manutenção ----------
    with tab5:
        st.header("Controle de Revisões e Lubrificação")
        st.markdown("Usa `KM_HR_Atual` (se presente) ou o último valor do histórico. A coluna `Unidade` determina se é KM ou HORAS (procure por 'QUIL'/'KM' ou 'HOR'/'HR').")

        # montar class_intervals a partir do session_state
        class_intervals = {}
        for k, v in st.session_state.manut_by_class.items():
            class_intervals[k] = {"rev_km": v.get("rev_km", []), "rev_hr": v.get("rev_hr", [])}

        mf = build_maintenance_dataframe(df_frotas, df, class_intervals, int(km_interval_default), int(hr_interval_default))

        # Aplica thresholds para marcar due
        def set_due_flags_row(row):
            due_km = False
            due_hr = False
            unit = str(row.get("Unid", "") or row.get("Unid", "") or row.get("Unid", "")).upper() if "Unid" in row else str(row.get("Unid","")).upper()
            # Em nosso mf usamos 'Unid' (se veio da frota). Também pegamos 'Unid' de last_unid se existir.
            unit = str(row.get("Unid", "")).upper() if pd.notna(row.get("Unid", "")) else str(row.get("Unid", "")).upper()
            for r in [1,2,3]:
                to_go = row.get(f"Rev{r}_To_Go", np.nan)
                if pd.isna(to_go):
                    continue
                # compara com threshold adequado
                if "QUIL" in unit or "KM" in unit:
                    if to_go <= km_due_threshold:
                        due_km = True
                elif "HOR" in unit or "HR" in unit:
                    if to_go <= hr_due_threshold:
                        due_hr = True
                else:
                    # sem unidade: avaliar com km threshold por padrão
                    if to_go <= km_due_threshold:
                        due_km = True
            return pd.Series({"Due_Rev": due_km, "Due_Oil": due_hr})

        if not mf.empty:
            flags = mf.apply(set_due_flags_row, axis=1)
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
            st.markdown("Marque a revisão/lubrificação realizada e clique em **Salvar ações** (registro será gravado em `MANUTENCAO_LOG`).")

            actions = []
            for idx, row in df_due.reset_index(drop=True).iterrows():
                cod = row.get("Cod_Equip", "")
                name = row.get("DESCRICAO_EQUIPAMENTO", "")
                unit = row.get("Unid", "") if "Unid" in row else row.get("Unidade", "")
                cur_val = row.get("Km_Hr_Atual", np.nan)
                st.markdown(f"**{cod} - {name}** — Atual: {formatar_brasileiro(cur_val)} {unit}")
                cols = st.columns([1,1,1,1])
                cb1 = cols[0].checkbox(f"Rev1 (cod {cod})", key=f"r1_{cod}")
                cb2 = cols[1].checkbox(f"Rev2 (cod {cod})", key=f"r2_{cod}")
                cb3 = cols[2].checkbox(f"Rev3 (cod {cod})", key=f"r3_{cod}")
                cbd = cols[3].checkbox(f"Lubrificação (cod {cod})", key=f"lub_{cod}")
                if cb1:
                    actions.append({"Cod_Equip": cod, "Tipo": "Rev1", "Valor_Atual": cur_val, "Unid": unit})
                if cb2:
                    actions.append({"Cod_Equip": cod, "Tipo": "Rev2", "Valor_Atual": cur_val, "Unid": unit})
                if cb3:
                    actions.append({"Cod_Equip": cod, "Tipo": "Rev3", "Valor_Atual": cur_val, "Unid": unit})
                if cbd:
                    actions.append({"Cod_Equip": cod, "Tipo": "Lubrificacao", "Valor_Atual": cur_val, "Unid": unit})

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
                            "Usuario": st.session_state.get("usuario", "(anon)"),
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
        st.subheader("Visão geral do plano de manutenção")
        overview_cols = ["Cod_Equip", "DESCRICAO_EQUIPAMENTO", "Km_Hr_Atual", "Unid",
                         "Rev1_Next", "Rev1_To_Go", "Rev2_Next", "Rev2_To_Go", "Rev3_Next", "Rev3_To_Go"]
        available_over = [c for c in overview_cols if c in mf.columns]
        st.dataframe(mf[available_over].sort_values("Cod_Equip").reset_index(drop=True), use_container_width=True)
        st.download_button("⬇️ Exportar CSV - Plano de Manutenção (Visão Geral)",
                           mf[available_over].to_csv(index=False).encode("utf-8"),
                           "manutencao_overview.csv", "text/csv")

if __name__ == "__main__":
    main()
