"""
Kion Dental — Alerta Diario de Faturamento
Executa todo dia as 10h (America/Sao_Paulo) via cron no container Docker.

Uso:
    python alerta_faturamento.py            → roda normalmente
    python alerta_faturamento.py --dry-run  → processa mas nao envia e-mails (so loga)

Variaveis de ambiente (obrigatorias em producao — defina no .env):
    GRAPH_TENANT_ID     → ID do diretorio (tenant) no Entra ID
    GRAPH_CLIENT_ID     → ID do aplicativo registrado no Entra ID
    GRAPH_CLIENT_SECRET → segredo do cliente (client secret)
    KION_BASE           → caminho base do projeto (padrao: /app em Docker, C:\\KionDental local)

Metricas calculadas por cliente:
    Faturamento: MRR real, projecao mes, mediana 12M, variacao %, meses consecutivos, impacto R$
    Casos novos: total mes, projecao mes, mediana 12M, variacao %, meses consecutivos
    Qualidade:   % ajustes, % repeticoes (alertas quando acima de 20%)
    Financeiro:  status (Ativo/Bloqueado/Inativo) e motivo do bloqueio
"""

import base64
import glob
import os
import sys
import logging
from datetime import datetime

import numpy as np
import pandas as pd
import requests
import yaml

# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────

def brl(v) -> str:
    """Formata numero no padrao brasileiro: 1.234.567 (ponto como separador de milhar)."""
    return f"{float(v or 0):,.0f}".replace(",", ".")


# ─────────────────────────────────────────────
#  CONFIGURACAO
# ─────────────────────────────────────────────

BASE_DIR    = os.environ.get("KION_BASE", r"C:\KionDental")
CONFIG_PATH = os.path.join(BASE_DIR, "config", "config.yaml")
DRY_RUN     = "--dry-run" in sys.argv
PREVIEW     = "--preview" in sys.argv

# --simdate=YYYY-MM-DD  → simula execucao em outra data (para testes/previews)
SIM_DATE: pd.Timestamp | None = None
for _arg in sys.argv:
    if _arg.startswith("--simdate="):
        SIM_DATE = pd.Timestamp(_arg.split("=", 1)[1])


def carregar_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    for chave in ["clientes", "producao_2025", "producao_2026", "logs"]:
        cfg["caminhos"][chave] = os.path.join(BASE_DIR, cfg["caminhos"][chave])
    cfg["caminhos"]["pedidos"] = os.path.join(BASE_DIR, cfg["caminhos"]["pedidos"])
    return cfg


# ─────────────────────────────────────────────
#  LOGGING
# ─────────────────────────────────────────────

def configurar_log(log_dir):
    os.makedirs(log_dir, exist_ok=True)
    hoje    = datetime.now().strftime("%Y-%m-%d")
    log_file = os.path.join(log_dir, f"alerta_{hoje}.log")

    console = logging.StreamHandler(sys.stdout)
    console.stream = open(sys.stdout.fileno(), mode="w",
                          encoding="utf-8", buffering=1, closefd=False)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler(log_file, encoding="utf-8"), console],
    )


# ─────────────────────────────────────────────
#  CONSTANTES DE PERIODO
# ─────────────────────────────────────────────

MESES_MAP = {
    "jan": 1, "fev": 2, "mar": 3, "abr": 4, "mai": 5, "jun": 6,
    "jul": 7, "ago": 8, "set": 9, "out": 10, "nov": 11, "dez": 12,
}

MESES_2025 = ["abr 2025", "mai 2025", "jun 2025", "jul 2025", "ago 2025",
               "set 2025", "out 2025", "nov 2025", "dez 2025"]

MESES_2026 = ["jan 2026", "fev 2026", "mar 2026", "abr 2026", "mai 2026",
               "jun 2026"]

JANELA_12M = ["jun 2025", "jul 2025", "ago 2025", "set 2025", "out 2025",
               "nov 2025", "dez 2025", "jan 2026", "fev 2026", "mar 2026",
               "abr 2026", "mai 2026"]  # atualizado dinamicamente em processar()

MESES_NOMES = {v: k for k, v in MESES_MAP.items()}


def calcular_janela_12m(mes_atual_str: str) -> list:
    """Janela de 12 meses terminando em mes_atual_str (rolling, nao hardcoded)."""
    periodo = mes_str_to_period(mes_atual_str)
    return [f"{MESES_NOMES[(periodo - i).month]} {(periodo - i).year}"
            for i in range(11, -1, -1)]


def mes_str_to_period(mes_str: str) -> pd.Period:
    """Converte 'jun 2025' → pd.Period('2025-06', 'M')."""
    p = mes_str.strip().split()
    return pd.Period(f"{p[1]}-{MESES_MAP[p[0]]:02d}", "M")


def dias_uteis_mes(ano: int, mes: int) -> int:
    """Total de dias uteis (seg–sex) no mes."""
    inicio = pd.Timestamp(ano, mes, 1)
    fim    = inicio + pd.offsets.MonthEnd(0)
    return len(pd.bdate_range(inicio, fim))


def dias_uteis_ate(ano: int, mes: int, ate: pd.Timestamp) -> int:
    """Dias uteis decorridos do inicio do mes ate `ate` (inclusive)."""
    inicio  = pd.Timestamp(ano, mes, 1)
    fim_mes = inicio + pd.offsets.MonthEnd(0)
    ate_clamped = min(ate, fim_mes)
    if ate_clamped < inicio:
        return 0
    return len(pd.bdate_range(inicio, ate_clamped))


# ─────────────────────────────────────────────
#  LEITURA DOS DADOS
# ─────────────────────────────────────────────

def ler_dados(cfg):
    logging.info("Lendo arquivos de faturamento...")
    df_cli  = pd.read_excel(cfg["caminhos"]["clientes"])
    df_2025 = pd.read_excel(cfg["caminhos"]["producao_2025"], sheet_name="Dados")
    df_2026 = pd.read_excel(cfg["caminhos"]["producao_2026"], sheet_name="Dados")
    logging.info(f"  clientes.xlsx       → {len(df_cli):,} linhas")
    logging.info(f"  producao_2025.xlsx  → {len(df_2025):,} linhas")
    logging.info(f"  producao_2026.xlsx  → {len(df_2026):,} linhas")
    return df_cli, df_2025, df_2026


def ler_pedidos(pasta: str) -> pd.DataFrame:
    """
    Carrega e consolida todos os arquivos xlsx da pasta de pedidos.
    Retorna DataFrame limpo com colunas padronizadas.
    """
    arquivos = sorted(glob.glob(os.path.join(pasta, "*.xlsx")))
    if not arquivos:
        logging.warning("  Pasta de pedidos sem arquivos xlsx — metricas de casos desativadas.")
        return pd.DataFrame()

    logging.info(f"Lendo pedidos ({len(arquivos)} arquivos)...")
    dfs = []
    for arq in arquivos:
        try:
            df = pd.read_excel(arq, dtype={"Nº pedido": str})
            dfs.append(df)
            logging.info(f"  {os.path.basename(arq):30s} → {len(df):,} linhas")
        except Exception as exc:
            logging.warning(f"  Erro em {os.path.basename(arq)}: {exc}")

    if not dfs:
        return pd.DataFrame()

    df_all = pd.concat(dfs, ignore_index=True)

    # Coluna Repeticao|Ajuste
    col_ra = next((c for c in df_all.columns if "Repeti" in c), None)
    df_all["rep_ajuste"] = df_all[col_ra].fillna("").astype(str).str.strip().str.upper() if col_ra else ""

    # Parse data de entrada (formato brasileiro DD/MM/YYYY)
    df_all["data_entrada"] = pd.to_datetime(
        df_all.get("Data de entrada"), dayfirst=True, errors="coerce"
    )
    df_all["mes_period"] = df_all["data_entrada"].dt.to_period("M")

    # Limpar Nº pedido — remove linhas sem pedido (itens filhos)
    df_all["Nº pedido"] = df_all["Nº pedido"].astype(str).str.strip()
    df_all = df_all[~df_all["Nº pedido"].isin(["", "nan", "NaN"])].copy()

    # Valor total numerico
    df_all["valor_total"] = pd.to_numeric(df_all.get("Valor total"), errors="coerce").fillna(0)

    # Status producao
    df_all["status_pedido"] = df_all.get("Status", pd.Series("")).astype(str).str.strip()

    # Cliente normalizado
    df_all["Cliente"] = df_all["Cliente"].astype(str).str.strip()

    # Excluir pedidos cancelados para contagem de casos
    df_all = df_all[df_all["status_pedido"] != "Cancelado"].copy()

    # Deduplica por Nº pedido (multiplas linhas de servico por pedido)
    df_all = df_all.drop_duplicates(subset=["Nº pedido"], keep="first")

    # Tipo do pedido
    df_all["tipo"] = "novo"
    df_all.loc[df_all["rep_ajuste"] == "R", "tipo"] = "repeticao"
    df_all.loc[df_all["rep_ajuste"] == "A", "tipo"] = "ajuste"

    logging.info(f"  Total pedidos consolidados: {len(df_all):,} (unicos, nao cancelados)")
    return df_all


# ─────────────────────────────────────────────
#  METRICAS DE PEDIDOS
# ─────────────────────────────────────────────

def calcular_metricas_pedidos(df_pedidos: pd.DataFrame, janela: list) -> pd.DataFrame:
    """
    Retorna DataFrame indexado por Cliente com colunas:
        casos_{mes}   → casos novos no mes
        rep_pct_{mes} → % repeticoes no mes
        adj_pct_{mes} → % ajustes no mes
    """
    if df_pedidos.empty:
        return pd.DataFrame()

    periodos = {m: mes_str_to_period(m) for m in janela}
    clientes = df_pedidos["Cliente"].unique()
    result   = pd.DataFrame(index=clientes)
    result.index.name = "Cliente"

    for mes_str, periodo in periodos.items():
        dm = df_pedidos[df_pedidos["mes_period"] == periodo]
        if dm.empty:
            result[f"casos_{mes_str}"]   = 0
            result[f"rep_pct_{mes_str}"] = 0.0
            result[f"adj_pct_{mes_str}"] = 0.0
            continue

        grp   = dm.groupby("Cliente")
        total = grp["Nº pedido"].count()
        novos = grp["tipo"].apply(lambda x: (x == "novo").sum())
        reps  = grp["tipo"].apply(lambda x: (x == "repeticao").sum())
        adjs  = grp["tipo"].apply(lambda x: (x == "ajuste").sum())

        result[f"casos_{mes_str}"]   = novos.reindex(result.index, fill_value=0)
        result[f"rep_pct_{mes_str}"] = (
            (reps / total.replace(0, np.nan) * 100)
            .reindex(result.index, fill_value=0).round(1)
        )
        result[f"adj_pct_{mes_str}"] = (
            (adjs / total.replace(0, np.nan) * 100)
            .reindex(result.index, fill_value=0).round(1)
        )

    return result.reset_index()


def calcular_projecao(df_pedidos: pd.DataFrame, mes_str: str, hoje: pd.Timestamp):
    """
    Projeta faturamento e casos novos para o mes em andamento.

    Formula:
        taxa_diaria  = acumulado ÷ dias_uteis_decorridos
        projecao     = acumulado + (taxa_diaria × dias_uteis_restantes)

    Retorna (serie_fat_projetado, serie_casos_projetados) indexadas por Cliente.
    """
    if df_pedidos.empty:
        return pd.Series(dtype=float), pd.Series(dtype=float)

    periodo = mes_str_to_period(mes_str)
    ano, mes = periodo.year, periodo.month

    du_total     = dias_uteis_mes(ano, mes)
    du_decorridos = dias_uteis_ate(ano, mes, hoje)
    du_restantes  = du_total - du_decorridos

    if du_decorridos == 0:
        return pd.Series(dtype=float), pd.Series(dtype=float)

    dm = df_pedidos[df_pedidos["mes_period"] == periodo]

    # ── Faturamento projetado ──────────────────────────────────────────────
    dm_fat   = dm[dm["status_pedido"] == "Finalizado"]
    fat_acum = dm_fat.groupby("Cliente")["valor_total"].sum()

    taxa_fat = fat_acum / du_decorridos                        # R$/dia util
    fat_proj = (fat_acum + taxa_fat * du_restantes).round(0)   # acumulado + restante estimado

    # ── Casos novos projetados ─────────────────────────────────────────────
    dm_novos   = dm[dm["tipo"] == "novo"]
    casos_acum = dm_novos.groupby("Cliente")["Nº pedido"].count()

    taxa_casos  = casos_acum / du_decorridos
    casos_proj  = (casos_acum + taxa_casos * du_restantes).round(0)

    return fat_proj, casos_proj


# ─────────────────────────────────────────────
#  PROCESSAMENTO FATURAMENTO
# ─────────────────────────────────────────────

def detectar_mes_atual(df_2026):
    """Retorna o mes mais recente com pelo menos 100 clientes faturando."""
    for mes in reversed(MESES_2026):
        if mes in df_2026.columns and (pd.to_numeric(df_2026[mes], errors="coerce") > 0).sum() >= 100:
            return mes
    return MESES_2026[-1]


def meses_consecutivos_queda(row, cols_janela: list, col_atual: str) -> int:
    """
    Conta meses FECHADOS consecutivos onde o faturamento/casos ficaram
    abaixo da mediana historica (nao compara com o mes atual parcial).

    - Para faturamento: compara cada mes fechado vs media_12m
    - Para casos:       compara cada mes fechado vs casos_mediana_12m
    """
    if col_atual not in cols_janela:
        return 0
    idx = cols_janela.index(col_atual)
    if idx == 0:
        return 0
    meses_anteriores = list(reversed(cols_janela[:idx]))

    # Seleciona a mediana correta conforme o tipo de coluna
    if col_atual.startswith("casos_"):
        media = float(row.get("casos_mediana_12m") or 0)
    else:
        media = float(row.get("media_12m") or 0)

    if media == 0:
        return 0

    count = 0
    for col in meses_anteriores:
        val = float(row.get(col) or 0)
        if val > 0 and val < media:
            count += 1
        else:
            break
    return count


def calcular_nivel_risco(row, cfg, mes_atual):
    th = cfg["thresholds"]
    if row["media_12m"] == 0 or row.get("fat_projetado", row[mes_atual]) == 0:
        return "SEM HISTORICO"
    if row["variacao_pct"] <= -th["alto_queda_pct"] and row["meses_queda"] >= th["alto_meses_min"]:
        return "ALTO"
    if row["variacao_pct"] <= -th["medio_queda_pct"]:
        return "MEDIO"
    if row["variacao_pct"] < 0:
        return "ATENCAO"
    return "ESTAVEL"


def processar(df_cli, df_2025, df_2026, df_pedidos, cfg):
    logging.info("Processando dados...")
    hoje = SIM_DATE if SIM_DATE is not None else pd.Timestamp.now()
    if SIM_DATE:
        logging.info(f"  ⚠️  Data simulada: {hoje.strftime('%d/%m/%Y')}")

    # ── Faturamento ────────────────────────────────────────────────────────
    meses_25_disp = [m for m in MESES_2025 if m in df_2025.columns]
    df_fat = pd.merge(
        df_2025[["Cliente"] + meses_25_disp],
        df_2026[["Cliente"] + [m for m in MESES_2026 if m in df_2026.columns]],
        on="Cliente", how="outer",
    ).fillna(0)

    # Status financeiro + dados cadastrais
    df_cli_limpo = df_cli[["Nome", "VENDAS", "Tabela de preço", "Status",
                            "Motivo de bloqueio"]].copy()
    df_cli_limpo.columns = ["Cliente", "vendas", "tabela", "status_financeiro",
                             "motivo_bloqueio"]

    df = pd.merge(df_fat, df_cli_limpo, on="Cliente", how="left")

    mes_atual = detectar_mes_atual(df_2026)
    logging.info(f"  Mes de referencia: {mes_atual}")

    # Janela rolling de 12 meses a partir do mes_atual
    janela = [m for m in calcular_janela_12m(mes_atual) if m in df.columns]
    logging.info(f"  Janela 12M: {janela[0]} → {janela[-1]} ({len(janela)} meses)")

    def mediana_sem_zeros(row):
        vals = [row[m] for m in janela if row.get(m, 0) > 0]
        return float(np.median(vals)) if vals else 0.0

    df["media_12m"]    = df.apply(mediana_sem_zeros, axis=1)
    df["meses_ativos"] = df[janela].apply(lambda r: (r > 0).sum(), axis=1)
    df["mes_atual"]    = df[mes_atual].astype(float)
    df["meses_queda"]  = df.apply(lambda r: meses_consecutivos_queda(r, janela, mes_atual), axis=1)
    df["mes_ref"]      = mes_atual

    # Projecao antecipada — necessaria para classificar risco corretamente
    _periodo = mes_str_to_period(mes_atual)
    _du_total = dias_uteis_mes(_periodo.year, _periodo.month)
    _du_dec   = dias_uteis_ate(_periodo.year, _periodo.month, hoje)
    _du_rest  = _du_total - _du_dec
    if _du_dec > 0:
        df["fat_projetado"] = (
            df["mes_atual"] + df["mes_atual"] / _du_dec * _du_rest
        ).round(0)
    else:
        df["fat_projetado"] = df["mes_atual"]

    # Variacao e risco baseados na PROJECAO (nao no parcial do mes)
    df["variacao_pct"] = (
        (df["fat_projetado"] - df["media_12m"]) / df["media_12m"].replace(0, np.nan) * 100
    )
    df["risco"] = df.apply(lambda r: calcular_nivel_risco(r, cfg, mes_atual), axis=1)

    # ── Clientes ativos: faturou em pelo menos 1 dos ultimos 6 meses ─────────
    ultimos_6 = [m for m in janela if m in df.columns][-6:]
    df["ativo_6m"] = df[ultimos_6].apply(lambda r: (r > 0).any(), axis=1)
    df_ativos = df[df["ativo_6m"]].copy()
    logging.info(f"  Clientes ativos (ultimos 6 meses): {len(df_ativos):,}")

    # ── Metricas de pedidos ────────────────────────────────────────────────
    if not df_pedidos.empty:
        logging.info("  Calculando metricas de pedidos (casos, ajustes, repeticoes)...")

        df_metricas = calcular_metricas_pedidos(df_pedidos, janela)

        if not df_metricas.empty:
            df_ativos = pd.merge(df_ativos, df_metricas, on="Cliente", how="left")

            # Colunas de casos novos por mes
            casos_cols = [f"casos_{m}" for m in janela if f"casos_{m}" in df_ativos.columns]
            for c in casos_cols:
                df_ativos[c] = df_ativos[c].fillna(0).astype(int)

            # Casos novos mes atual
            col_casos_atual = f"casos_{mes_atual}"
            df_ativos["casos_novos_atual"] = (
                df_ativos[col_casos_atual].fillna(0).astype(int)
                if col_casos_atual in df_ativos.columns else 0
            )

            # Mediana 12M de casos novos (exclui zeros)
            def mediana_casos(row):
                vals = [row[c] for c in casos_cols if row.get(c, 0) > 0]
                return float(np.median(vals)) if vals else 0.0

            df_ativos["casos_mediana_12m"] = df_ativos.apply(mediana_casos, axis=1)

            # Variacao e meses consecutivos de queda em casos
            df_ativos["casos_var_pct"] = (
                (df_ativos["casos_novos_atual"] - df_ativos["casos_mediana_12m"])
                / df_ativos["casos_mediana_12m"].replace(0, np.nan) * 100
            )
            df_ativos["casos_meses_queda"] = df_ativos.apply(
                lambda r: meses_consecutivos_queda(r, casos_cols, col_casos_atual), axis=1
            )

            # % ajustes e repeticoes do mes atual
            col_rep = f"rep_pct_{mes_atual}"
            col_adj = f"adj_pct_{mes_atual}"
            df_ativos["rep_pct_atual"] = (
                df_ativos[col_rep].fillna(0) if col_rep in df_ativos.columns else 0
            )
            df_ativos["adj_pct_atual"] = (
                df_ativos[col_adj].fillna(0) if col_adj in df_ativos.columns else 0
            )
        else:
            for col in ["casos_novos_atual", "casos_mediana_12m", "casos_var_pct",
                        "casos_meses_queda", "rep_pct_atual", "adj_pct_atual"]:
                df_ativos[col] = 0

        # ── Projecao do mes atual ──────────────────────────────────────────
        # Usa MRR atual do ERP como acumulado (mais preciso que pedidos nao finalizados)
        # Formula: acumulado + (acumulado ÷ dias_decorridos × dias_restantes)
        periodo_atual = mes_str_to_period(mes_atual)
        du_total  = dias_uteis_mes(periodo_atual.year, periodo_atual.month)
        du_dec    = dias_uteis_ate(periodo_atual.year, periodo_atual.month, hoje)
        du_rest   = du_total - du_dec

        if du_dec > 0:
            df_ativos["fat_projetado"] = (
                df_ativos["mes_atual"] +
                (df_ativos["mes_atual"] / du_dec * du_rest)
            ).round(0)
            df_ativos["casos_projetados"] = (
                df_ativos["casos_novos_atual"] +
                (df_ativos["casos_novos_atual"] / du_dec * du_rest)
            ).clip(lower=0).round(0).astype(int)
        else:
            df_ativos["fat_projetado"]    = df_ativos["mes_atual"]
            df_ativos["casos_projetados"] = df_ativos["casos_novos_atual"]

    else:
        for col in ["casos_novos_atual", "casos_mediana_12m", "casos_var_pct",
                    "casos_meses_queda", "rep_pct_atual", "adj_pct_atual"]:
            df_ativos[col] = 0
        df_ativos["fat_projetado"]    = df_ativos["mes_atual"]
        df_ativos["casos_projetados"] = 0

    # ── Impacto recalculado: Mediana 12M − Proj. Mes ──────────────────────
    df_ativos["impacto_rs"] = (
        df_ativos["media_12m"] - df_ativos["fat_projetado"]
    ).clip(lower=0).round(0)

    logging.info(f"  Clientes ativos em {mes_atual}: {len(df_ativos):,}")
    for nivel in ["ALTO", "MEDIO", "ATENCAO", "ESTAVEL"]:
        logging.info(f"    {nivel}: {(df_ativos['risco'] == nivel).sum()}")

    return df_ativos, mes_atual


# ─────────────────────────────────────────────
#  TEMPLATES HTML
# ─────────────────────────────────────────────

EMOJI = {"ALTO": "🔴", "MEDIO": "🟡", "ATENCAO": "🟢", "ESTAVEL": "✅"}


def _status_badge(r) -> str:
    status = str(r.get("status_financeiro", "Ativo") or "Ativo").strip()
    motivo = str(r.get("motivo_bloqueio", "") or "").strip()
    if status == "Bloqueado":
        mot = f" · {motivo}" if motivo and motivo != "nan" else ""
        return (f" <span style='background:#c0392b;color:#fff;font-size:9px;"
                f"padding:1px 5px;border-radius:3px;font-weight:700'>"
                f"&#128683; BLOQUEADO{mot}</span>")
    if status == "Inativo":
        return (" <span style='background:#bdc3c7;color:#fff;font-size:9px;"
                "padding:1px 5px;border-radius:3px'>INATIVO</span>")
    return ""


def _narrativa(r) -> str:
    """Linha de contexto compacta — sem prose longo."""
    queda = abs(r["variacao_pct"])
    meses = int(r["meses_queda"])
    return (
        f"<span style='font-size:10px;color:#9AA0A6'>"
        f"{queda:.0f}% abaixo da mediana"
        f"{f' · {meses}m em queda' if meses > 0 else ''}"
        f"</span>"
    )


def _bloco_risco(df_terr, nivel, label_acao, show_vendas=False, cfg=None) -> str:
    grupo = df_terr[df_terr["risco"] == nivel].sort_values("impacto_rs", ascending=False)
    if grupo.empty:
        return ""

    emoji       = EMOJI.get(nivel, "")
    tem_pedidos = "casos_novos_atual" in df_terr.columns
    n           = len(grupo)
    rows        = ""

    COR = {"ALTO": "#c0392b", "MEDIO": "#e67e22", "ATENCAO": "#27ae60"}
    cor_nivel = COR.get(nivel, "#00B1D2")

    for _, r in grupo.iterrows():
        tabela    = str(r.get("tabela") or "")
        if pd.isna(r.get("tabela")): tabela = ""
        fat_proj  = float(r.get("fat_projetado") or r["mes_atual"])
        casos     = int(r.get("casos_novos_atual") or 0)
        casos_prj = int(r.get("casos_projetados") or casos)
        casos_med = float(r.get("casos_mediana_12m") or 0)
        casos_var = r.get("casos_var_pct", None)
        casos_mq  = int(r.get("casos_meses_queda") or 0)
        adj_pct   = float(r.get("adj_pct_atual") or 0)
        rep_pct   = float(r.get("rep_pct_atual") or 0)

        # Variacao faturamento
        vf = r["variacao_pct"]
        cor_vf  = "#c0392b" if vf < 0 else "#27ae60"

        # Variacao casos
        if casos_var is not None and not pd.isna(casos_var):
            cor_vc = "#c0392b" if casos_var < -10 else "#27ae60" if casos_var > 5 else "#888"
            vc_str = f"<span style='color:{cor_vc};font-weight:700'>{casos_var:+.0f}%</span>"
        else:
            vc_str = "<span style='color:#ccc'>—</span>"

        # Qualidade
        adj_cor = "#c0392b" if adj_pct > 20 else "#888"
        rep_cor = "#c0392b" if rep_pct > 20 else "#888"

        # Sub-linha operacional
        if tem_pedidos:
            sub_row = f"""
        <tr class="op-row">
          <td style="color:#888;font-style:italic">{tabela}</td>
          <td><strong style="color:#00B1D2">{casos}</strong></td>
          <td>{casos_prj}</td>
          <td style="color:#aaa">{casos_med:.0f}</td>
          <td>{vc_str}</td>
          <td style="color:#aaa">{casos_mq}m &#8595;</td>
          <td><span style="color:{adj_cor}">Adj {adj_pct:.0f}%</span> 
              <span style="color:{rep_cor}">Rep {rep_pct:.0f}%</span></td>
        </tr>"""
        else:
            sub_row = ""

        # Coluna de responsavel (so no e-mail do gestor)
        resp_cell = ""
        if show_vendas and cfg:
            terr = str(r.get("vendas") or "—")
            nome_resp = cfg.get("territorios", {}).get(terr, {}).get("nome", terr) if terr != "—" else "—"
            resp_cell = (
                f"<td style='color:#555;font-size:11px'>"
                f"<strong style='color:#00B1D2'>{terr}</strong><br>{nome_resp}</td>"
            )

        rows += f"""
        <tr>
          <td style="padding:9px 10px 5px">
            <strong style="font-size:12px">{r['Cliente']}</strong>{_status_badge(r)}<br>
            {_narrativa(r)}
          </td>
          {resp_cell}
          <td style="font-weight:700">R$ {brl(r['mes_atual'])}</td>
          <td style="color:#555">R$ {brl(fat_proj)}</td>
          <td style="color:#888">R$ {brl(r['media_12m'])}</td>
          <td style="font-weight:700;color:{cor_vf}">{vf:+.0f}%</td>
          <td style="color:#555">{int(r['meses_queda'])}m</td>
          <td style="font-weight:700;color:#c0392b">-R$ {brl(r['impacto_rs'])}</td>
        </tr>{sub_row}"""

    col_resp_th = "<th>Responsavel</th>" if show_vendas else ""
    col_resp_op = "<th style='color:#00B1D2;font-size:10px;font-weight:600'> </th>" if show_vendas else ""
    cliente_width = "28%" if show_vendas else "34%"

    return f"""
    <p class="section-title" style="border-left:3px solid {cor_nivel};padding-left:8px">
      {emoji} {nivel} - {label_acao}
      <span style="font-weight:400;color:#9AA0A6;font-size:10px;margin-left:6px">{n} cliente(s)</span>
    </p>
    <table>
      <thead>
        <tr>
          <th style="width:{cliente_width}">Cliente</th>
          {col_resp_th}
          <th>MRR Atual</th>
          <th>Proj. Mes</th>
          <th>Mediana 12M</th>
          <th>Variacao</th>
          <th>Meses &#8595;</th>
          <th>Impacto/mes</th>
        </tr>
        {f'''<tr style="background:#e8f8fd">
          <th style="color:#00B1D2;font-size:10px;font-weight:600">Tabela de Preco</th>
          {col_resp_op}
          <th style="color:#00B1D2;font-size:10px;font-weight:600">Casos Novos</th>
          <th style="color:#00B1D2;font-size:10px;font-weight:600">Proj. Casos</th>
          <th style="color:#00B1D2;font-size:10px;font-weight:600">Med. Casos 12M</th>
          <th style="color:#00B1D2;font-size:10px;font-weight:600">Var. Casos</th>
          <th style="color:#00B1D2;font-size:10px;font-weight:600">Meses &#8595;</th>
          <th style="color:#00B1D2;font-size:10px;font-weight:600">Qualidade</th>
        </tr>''' if tem_pedidos else ''}
      </thead>
      <tbody>{rows}</tbody>
    </table>"""


# ── CSS moderno para o relatorio HTML standalone (anexo) ─────────────────────
CSS_RELATORIO = """
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
  :root{
    --azul:#00B1D2;--azul-vivo:#00F5FF;--amarelo:#FAEB1E;
    --cinza-dark:#282828;--vermelho:#c0392b;--laranja:#e67e22;--verde:#27ae60;
    --bg:#f5f7fa;--card:#fff;--borda:#e9ecef;--texto:#2d3748;--suave:#718096;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Inter',Arial,sans-serif;background:var(--bg);color:var(--texto);font-size:13px}
  .container{max-width:1100px;margin:0 auto;background:var(--card);box-shadow:0 2px 16px rgba(0,0,0,.08)}

  /* ── Header ── */
  .rpt-header{background:var(--cinza-dark);padding:16px 28px;display:flex;align-items:center;justify-content:space-between}
  .rpt-brand{display:flex;align-items:center;gap:12px}
  .rpt-nome{font-size:22px;font-weight:800;letter-spacing:3px;color:var(--azul-vivo)}
  .rpt-sub{font-size:10px;color:#8D8E8F;letter-spacing:1px;margin-top:1px}
  .rpt-tag{font-size:11px;color:#8D8E8F;text-align:right;line-height:1.6}
  .rpt-tag strong{color:var(--azul)}
  .rpt-bar{height:4px;background:linear-gradient(90deg,var(--azul-vivo) 0%,var(--azul) 50%,var(--amarelo) 100%)}

  /* ── Corpo ── */
  .rpt-body{padding:24px 28px}
  .rpt-title{font-size:16px;font-weight:700;color:var(--cinza-dark);padding-bottom:10px;border-bottom:2px solid var(--azul);margin-bottom:4px}
  .rpt-subtitle{font-size:11px;color:var(--suave);margin-bottom:18px;margin-top:4px}

  /* ── KPI Cards ── */
  .kpi-row{display:flex;gap:10px;margin-bottom:20px;flex-wrap:wrap}
  .kpi-card{flex:1;min-width:120px;background:var(--card);border:1px solid var(--borda);border-radius:8px;padding:14px 16px;box-shadow:0 1px 4px rgba(0,0,0,.05)}
  .kpi-card.destaque{border-left:3px solid var(--azul)}
  .kpi-card.risco{border-left:3px solid var(--vermelho)}
  .kpi-card.bloqueado{border-left:3px solid var(--vermelho);background:#fff5f5}
  .kpi-lbl{font-size:9px;font-weight:600;text-transform:uppercase;letter-spacing:.7px;color:var(--suave);margin-bottom:5px}
  .kpi-val{font-size:18px;font-weight:700;color:var(--cinza-dark);line-height:1.1}
  .kpi-sub{font-size:10px;color:var(--suave);margin-top:3px}
  .kpi-red{color:var(--vermelho)}

  /* ── Secoes de risco ── */
  .section{margin-bottom:28px}
  .section-hd{display:flex;align-items:center;gap:8px;padding:8px 12px;border-radius:6px;margin-bottom:8px}
  .section-hd.alto{background:#fff0f0;border-left:4px solid var(--vermelho)}
  .section-hd.medio{background:#fff8ee;border-left:4px solid var(--laranja)}
  .section-hd.atenc{background:#f0fff4;border-left:4px solid var(--verde)}
  .section-ttl{font-size:13px;font-weight:700}
  .section-cnt{font-size:11px;color:var(--suave);margin-left:auto}

  /* ── Tabela ── */
  .tbl-wrap{overflow-x:auto;border:1px solid var(--borda);border-radius:6px}
  table{border-collapse:collapse;width:100%;font-size:12px}
  thead{position:sticky;top:0;z-index:5}
  thead tr:first-child th{background:var(--azul);color:#fff;padding:8px 10px;text-align:left;font-size:10px;text-transform:uppercase;letter-spacing:.5px;white-space:nowrap}
  thead tr:last-child th{background:#e0f6fb;color:var(--azul);font-size:10px;font-weight:600;padding:5px 10px;text-transform:uppercase;letter-spacing:.4px}
  tbody tr:hover td{background:#f0fbff}
  tbody tr td{padding:9px 10px;border-bottom:1px solid #f0f3f6;vertical-align:middle}
  .op-row td{background:#f8fbfe;padding:4px 10px 7px;border-bottom:2px solid #e4edf3;font-size:11px;color:var(--suave)}
  .op-row:hover td{background:#f0f6fb}

  /* ── Badges ── */
  .bdg{display:inline-block;padding:2px 7px;border-radius:10px;font-size:10px;font-weight:700;margin-right:3px}
  .bdg-alto{background:#fde8e8;color:var(--vermelho)}
  .bdg-medio{background:#fff3e0;color:var(--laranja)}
  .bdg-atenc{background:#e8f5e9;color:var(--verde)}
  .neg{color:var(--vermelho);font-weight:700}
  .pos{color:var(--verde);font-weight:700}

  /* ── Status ── */
  .st-bloq{background:var(--vermelho);color:#fff;font-size:9px;padding:1px 5px;border-radius:3px;font-weight:700;vertical-align:middle}
  .st-inat{background:#bdc3c7;color:#fff;font-size:9px;padding:1px 5px;border-radius:3px}

  /* ── Legenda ── */
  .legenda{margin-top:24px;padding-top:14px;border-top:1px solid var(--borda);font-size:10px;color:#aaa;line-height:1.9}
  .legenda strong{color:#bbb}

  /* ── Rodape ── */
  .rpt-footer{background:var(--cinza-dark);padding:14px 28px;display:flex;align-items:center;justify-content:space-between;margin-top:0}
  .rpt-footer p{font-size:11px;color:#8D8E8F;line-height:1.8}
  .rpt-footer strong{color:var(--azul-vivo)}

  /* ── Botao imprimir ── */
  .btn-print{position:fixed;bottom:20px;right:20px;background:var(--azul);color:#fff;border:none;padding:10px 18px;border-radius:20px;font-size:12px;font-weight:600;cursor:pointer;box-shadow:0 3px 10px rgba(0,177,210,.35);z-index:99}
  .btn-print:hover{background:#009ab8}

  /* ── Banner teste ── */
  .banner{background:var(--azul);color:#fff;padding:8px 28px;font-size:11px;font-weight:600}

  @media print{
    .btn-print,.banner{display:none}
    .container{box-shadow:none}
    thead{position:relative}
    .rpt-body{padding:12px}
    body{background:#fff}
  }
  @media(max-width:700px){
    .kpi-row{gap:6px}
    .kpi-card{min-width:calc(50% - 6px)}
    .rpt-header{flex-direction:column;gap:8px}
  }
</style>
"""

# ── Kion logo arc ─────────────────────────────────────────────────────────────
KION_ARC = (
    "<!--[if !mso]><!-->"
    '<svg viewBox="0 0 110 100" width="36" height="32" '
    'style="vertical-align:middle;display:inline-block">'
    "<defs><linearGradient id=\"kg\" x1=\"0%\" y1=\"0%\" x2=\"100%\" y2=\"0%\">"
    '<stop offset="0%" stop-color="#00F5FF"/>'
    '<stop offset="100%" stop-color="#FAEB1E"/>'
    "</linearGradient></defs>"
    '<path d="M 12 92 A 46 46 0 1 1 98 92" stroke="url(#kg)" '
    'stroke-width="10" fill="none" stroke-linecap="round"/>'
    "</svg>"
    "<!--<![endif]-->"
)

# ── CSS ────────────────────────────────────────────────────────────────────────
CSS = """
<style>
  body{{margin:0;padding:0;background:#f0f2f5;
        font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#282828}}
  .wrapper{{max-width:860px;margin:0 auto;background:#fff;border-radius:0 0 8px 8px}}
  .body-wrap{{padding:20px 24px}}
  h2{{color:#282828;font-size:15px;font-weight:700;margin:0 0 2px;
      padding-bottom:8px;border-bottom:2px solid #00B1D2}}
  .subtitle{{color:#9AA0A6;font-size:11px;margin:0 0 14px}}
  .section-title{{color:#00B1D2;font-size:12px;font-weight:700;
                  margin:22px 0 6px;text-transform:uppercase;letter-spacing:.6px}}
  /* ── Tabela de clientes ── */
  .body-wrap table{{border-collapse:collapse;width:100%;margin:0 0 6px;font-size:12px}}
  .body-wrap th{{background:#00B1D2;color:#fff;padding:7px 10px;text-align:left;
                 font-size:10px;text-transform:uppercase;letter-spacing:.5px;font-weight:600}}
  .body-wrap td{{padding:8px 10px;border-bottom:1px solid #edf0f3;vertical-align:middle}}
  /* ── Badges ── */
  .badge{{display:inline-block;padding:2px 8px;border-radius:10px;
          font-size:11px;font-weight:700;margin-right:3px}}
  .badge-alto{{background:#fde8e8;color:#c0392b}}
  .badge-medio{{background:#fff8e1;color:#e67e22}}
  .badge-atenc{{background:#e8f5e9;color:#27ae60}}
  .neg{{color:#c0392b;font-weight:700}}
  .pos{{color:#27ae60;font-weight:700}}
  /* ── Resumo ── */
  .kpi-label{{font-size:9px;color:#9AA0A6;text-transform:uppercase;
              letter-spacing:.6px;margin-bottom:4px}}
  .kpi-value{{font-size:17px;font-weight:700;color:#282828;line-height:1.1}}
  .kpi-sub{{font-size:10px;color:#9AA0A6;margin-top:2px}}
  /* ── Banner teste ── */
  .banner{{background:#00B1D2;color:#fff;padding:8px 12px;font-weight:700;
           border-radius:4px;margin-bottom:14px;font-size:11px}}
  /* ── Sub-linha operacional ── */
  .op-row td{{background:#f7f9fb;font-size:11px;color:#666;
              padding:4px 10px 6px;border-bottom:2px solid #e4e9ef}}
</style>
"""


def _header(tag_titulo):
    return (
        '<table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#282828" '
        'style="border-collapse:collapse;background:#282828">'
        "<tr>"
        '<td align="left" bgcolor="#282828" '
        'style="padding:16px 24px;vertical-align:middle;background:#282828">'
        '<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse"><tr>'
        '<td style="vertical-align:middle;padding-right:10px;font-size:0;line-height:0">'
        f"{KION_ARC}</td>"
        '<td style="vertical-align:middle">'
        '<p style="margin:0;padding:0;font-size:21px;font-weight:900;letter-spacing:3px;'
        'color:#00F5FF;font-family:Arial,Helvetica,sans-serif;line-height:1.2">KION</p>'
        '<p style="margin:0;padding:0;font-size:10px;color:#8D8E8F;letter-spacing:1px;'
        'font-family:Arial,Helvetica,sans-serif">DENTAL TECHNOLOGY</p>'
        "</td></tr></table></td>"
        '<td align="right" bgcolor="#282828" '
        'style="padding:16px 24px;vertical-align:middle;text-align:right;background:#282828">'
        '<p style="margin:0;padding:0;font-size:11px;color:#8D8E8F;line-height:1.5;'
        'font-family:Arial,Helvetica,sans-serif">'
        f"Alerta Comercial<br><strong style=\"color:#00B1D2\">{tag_titulo}</strong>"
        "</p></td></tr></table>"
        '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse">'
        '<tr><td height="4" bgcolor="#00B1D2" style="height:4px;font-size:0;line-height:0;'
        "background:#00B1D2;background:linear-gradient(90deg,#00F5FF 0%,#00B1D2 50%,#FAEB1E 100%)\">"
        " </td></tr></table>"
    )


def _legenda() -> str:
    """Bloco de legenda sutil antes do rodape."""
    return (
        '<table width="100%" cellpadding="0" cellspacing="0" '
        'style="border-collapse:collapse;margin-top:24px">'
        '<tr><td style="border-top:1px solid #edf0f3;padding:12px 0 4px">'
        '<p style="margin:0 0 6px;font-size:9px;color:#bdc3c7;'
        'text-transform:uppercase;letter-spacing:.8px;font-family:Arial,sans-serif">'
        'Legenda</p>'
        '<p style="margin:0;font-size:10px;color:#bdc3c7;line-height:1.8;'
        'font-family:Arial,sans-serif">'
        '<strong style="color:#9AA0A6">MRR Atual</strong> Faturamento real do mes (ERP)  |  '
        '<strong style="color:#9AA0A6">Proj. Mes</strong> Estimativa de fechamento = acumulado / dias uteis decorridos x dias uteis totais  |  '
        '<strong style="color:#9AA0A6">Mediana 12M</strong> Baseline historico (mediana dos ultimos 12 meses com faturamento > 0)  |  '
        '<strong style="color:#9AA0A6">Variacao</strong> (MRR Atual - Mediana) / Mediana  |  '
        '<strong style="color:#9AA0A6">Meses &#8595;</strong> Meses consecutivos de queda  |  '
        '<strong style="color:#9AA0A6">Impacto</strong> Mediana - MRR Atual'
        '</p>'
        '<p style="margin:6px 0 0;font-size:10px;color:#bdc3c7;line-height:1.8;'
        'font-family:Arial,sans-serif">'
        '<strong style="color:#9AA0A6">Casos Novos</strong> Pedidos unicos sem Repeticao ou Ajuste no mes  |  '
        '<strong style="color:#9AA0A6">Proj. Casos</strong> Mesma formula de projecao aplicada a casos  |  '
        '<strong style="color:#9AA0A6">Med. Casos 12M</strong> Mediana historica de casos novos  |  '
        '<strong style="color:#9AA0A6">Adj %</strong> Pedidos de ajuste / total  |  '
        '<strong style="color:#9AA0A6">Rep %</strong> Pedidos de repeticao / total  |  '
        '<span style="color:#c0392b">Vermelho</span> quando Adj ou Rep > 20%'
        '</p>'
        '</td></tr></table>'
    )


def _footer(mes_ref):
    return (
        '<table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#282828" '
        'style="border-collapse:collapse;background:#282828">'
        "<tr>"
        '<td colspan="2" bgcolor="#282828" style="padding:0 24px;background:#282828;font-size:0;line-height:0">'
        '<table width="100%" cellpadding="0" cellspacing="0" border="0" '
        'style="border-collapse:collapse;border-top:1px solid #3a3a3a">'
        '<tr><td style="height:1px;font-size:0;line-height:0"> </td></tr>'
        "</table></td></tr><tr>"
        '<td bgcolor="#282828" style="padding:12px 24px 14px;vertical-align:middle;background:#282828">'
        '<p style="margin:0;padding:0;font-size:11px;color:#8D8E8F;line-height:1.9;'
        'font-family:Arial,Helvetica,sans-serif">'
        '<strong style="color:#00F5FF">Analytics Kion Dental</strong><br>'
        "analytics@kiondental.tech<br>"
        "Analise desenvolvida pelo "
        '<strong style="color:#00F5FF">Time de Tecnologia e Inovacao da Kion</strong>'
        "</p></td>"
        '<td align="right" bgcolor="#282828" '
        'style="padding:12px 24px 14px;text-align:right;vertical-align:middle;background:#282828">'
        f"{KION_ARC}"
        f'<p style="margin:4px 0 0;padding:0;font-size:10px;color:#5A5A5A;line-height:1.8;'
        f'font-family:Arial,Helvetica,sans-serif">'
        f"Gerado automaticamente | {mes_ref.upper()}"
        f"</p></td></tr></table>"
    )


def _kpi_row_relatorio(fat_total, fat_proj, total_ativos, n_alto, n_medio,
                       n_atenc, fat_risco, impacto_total, total_casos, casos_proj, n_bloqueados, mes_ref):
    """KPI cards para o relatorio HTML standalone (CSS moderno)."""
    return f"""
    <div class="kpi-row">
      <div class="kpi-card destaque">
        <div class="kpi-lbl">Faturamento {mes_ref}</div>
        <div class="kpi-val">R$ {brl(fat_total)}</div>
        <div class="kpi-sub">Proj: R$ {brl(fat_proj)}</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-lbl">Clientes Ativos</div>
        <div class="kpi-val">{total_ativos}</div>
      </div>
      <div class="kpi-card risco">
        <div class="kpi-lbl">Em Risco</div>
        <div class="kpi-val kpi-red">{n_alto + n_medio}</div>
        <div class="kpi-sub">
          <span class="bdg bdg-alto">&#128308; {n_alto}</span>
          <span class="bdg bdg-medio">&#128993; {n_medio}</span>
          {f'<span class="bdg bdg-atenc">&#128994; {n_atenc}</span>' if n_atenc > 0 else ''}
        </div>
      </div>
      <div class="kpi-card risco">
        <div class="kpi-lbl">Fat. em Risco</div>
        <div class="kpi-val kpi-red">R$ {brl(fat_risco)}</div>
        <div class="kpi-sub">{fat_risco/fat_total*100:.1f}% da carteira</div>
      </div>
      {f'''<div class="kpi-card destaque">
        <div class="kpi-lbl">Casos Novos</div>
        <div class="kpi-val">{total_casos}</div>
        <div class="kpi-sub">Proj: {casos_proj}</div>
      </div>''' if total_casos > 0 else ''}
      {f'''<div class="kpi-card bloqueado">
        <div class="kpi-lbl" style="color:var(--vermelho)">Bloqueados</div>
        <div class="kpi-val kpi-red">&#128683; {n_bloqueados}</div>
      </div>''' if n_bloqueados > 0 else ''}
    </div>"""


def _bloco_risco_relatorio(df_terr, nivel, label_acao, show_vendas=False, cfg=None):
    """Versao do bloco de risco para o relatorio HTML standalone (CSS moderno)."""
    grupo = df_terr[df_terr["risco"] == nivel].sort_values("impacto_rs", ascending=False)
    if grupo.empty:
        return ""

    n           = len(grupo)
    tem_pedidos = "casos_novos_atual" in df_terr.columns
    COR_HD      = {"ALTO": "alto", "MEDIO": "medio", "ATENCAO": "atenc"}
    cls_hd      = COR_HD.get(nivel, "atenc")
    EMOJI_MAP   = {"ALTO": "🔴", "MEDIO": "🟡", "ATENCAO": "🟢"}
    emoji       = EMOJI_MAP.get(nivel, "")
    rows        = ""

    for _, r in grupo.iterrows():
        tabela    = str(r.get("tabela") or ""); tabela = "" if pd.isna(r.get("tabela")) else tabela
        fat_proj  = float(r.get("fat_projetado") or r["mes_atual"])
        casos     = int(r.get("casos_novos_atual") or 0)
        casos_prj = int(r.get("casos_projetados") or casos)
        casos_med = float(r.get("casos_mediana_12m") or 0)
        casos_var = r.get("casos_var_pct")
        casos_mq  = int(r.get("casos_meses_queda") or 0)
        adj_pct   = float(r.get("adj_pct_atual") or 0)
        rep_pct   = float(r.get("rep_pct_atual") or 0)
        vf        = r["variacao_pct"]
        cor_vf    = "neg" if vf < 0 else "pos"

        status = str(r.get("status_financeiro") or "Ativo")
        st_badge = ""
        if status == "Bloqueado":
            mot = str(r.get("motivo_bloqueio") or "")
            mot_txt = f" · {mot}" if mot and mot != "nan" else ""
            st_badge = f' <span class="st-bloq">&#128683; BLOQUEADO{mot_txt}</span>'
        elif status == "Inativo":
            st_badge = ' <span class="st-inat">INATIVO</span>'

        resp_cell = ""
        if show_vendas and cfg:
            terr = str(r.get("vendas") or "—")
            nome_r = cfg.get("territorios", {}).get(terr, {}).get("nome", terr) if terr != "—" else "—"
            resp_cell = f"<td><strong style='color:var(--azul)'>{terr}</strong><br><span style='color:var(--suave);font-size:11px'>{nome_r}</span></td>"

        if tem_pedidos:
            vc_str = "—"
            if casos_var is not None and not pd.isna(casos_var):
                cl = "neg" if casos_var < -10 else "pos" if casos_var > 5 else ""
                vc_str = f"<span class='{cl}'>{casos_var:+.0f}%</span>"
            adj_cl = "neg" if adj_pct > 20 else ""
            rep_cl = "neg" if rep_pct > 20 else ""
            sub_row = f"""<tr class="op-row">
              <td style="font-style:italic;color:var(--suave)">{tabela}</td>
              {f"<td></td>" if show_vendas else ""}
              <td><strong style="color:var(--azul)">{casos}</strong></td>
              <td style="color:var(--suave)">{casos_prj}</td>
              <td style="color:var(--suave)">{casos_med:.0f}</td>
              <td>{vc_str}</td>
              <td style="color:var(--suave)">{casos_mq}m &#8595;</td>
              <td><span class="{adj_cl}">Adj {adj_pct:.0f}%</span> <span class="{rep_cl}">Rep {rep_pct:.0f}%</span></td>
            </tr>"""
        else:
            sub_row = ""

        rows += f"""<tr>
          <td style="padding:9px 10px 5px">
            <strong>{r['Cliente']}</strong>{st_badge}<br>
            <span style="font-size:10px;color:var(--suave)">{abs(vf):.0f}% abaixo da mediana{f" · {int(r['meses_queda'])}m em queda" if r['meses_queda'] > 0 else ""}</span>
          </td>
          {resp_cell}
          <td style="font-weight:700">R$ {brl(r['mes_atual'])}</td>
          <td style="color:var(--suave)">R$ {brl(fat_proj)}</td>
          <td style="color:var(--suave)">R$ {brl(r['media_12m'])}</td>
          <td class="{cor_vf}" style="font-weight:700">{vf:+.0f}%</td>
          <td style="color:var(--suave)">{int(r['meses_queda'])}m</td>
          <td class="neg" style="font-weight:700">-R$ {brl(r['impacto_rs'])}</td>
        </tr>{sub_row}"""

    col_resp_th = "<th>Responsavel</th>" if show_vendas else ""
    col_resp_op = "<th></th>" if show_vendas else ""
    cli_w = "26%" if show_vendas else "32%"

    return f"""
    <div class="section">
      <div class="section-hd {cls_hd}">
        <span style="font-size:16px">{emoji}</span>
        <span class="section-ttl">{nivel} - {label_acao}</span>
        <span class="section-cnt">{n} cliente(s)</span>
      </div>
      <div class="tbl-wrap">
        <table>
          <thead>
            <tr>
              <th style="width:{cli_w}">Cliente</th>{col_resp_th}
              <th>MRR Atual</th><th>Proj. Mes</th>
              <th>Mediana 12M</th><th>Variacao</th>
              <th>Meses &#8595;</th><th>Impacto/mes</th>
            </tr>
            {f'''<tr>
              <th>Tabela de Preco</th>{col_resp_op}
              <th>Casos Novos</th><th>Proj. Casos</th>
              <th>Med. Casos 12M</th><th>Var. Casos</th>
              <th>Meses &#8595;</th><th>Qualidade</th>
            </tr>''' if tem_pedidos else ''}
          </thead>
          <tbody>{rows}</tbody>
        </table>
      </div>
    </div>"""


def _legenda_relatorio():
    return """
    <div class="legenda">
      <strong>MRR Atual</strong> Faturamento real do mes (ERP)  | 
      <strong>Proj. Mes</strong> Acumulado ÷ dias uteis decorridos × dias uteis totais  | 
      <strong>Mediana 12M</strong> Mediana dos ultimos 12 meses com faturamento > 0  | 
      <strong>Variacao</strong> (MRR Atual − Mediana) ÷ Mediana  | 
      <strong>Meses ↓</strong> Meses consecutivos de queda  | 
      <strong>Impacto</strong> Mediana − MRR Atual<br>
      <strong>Casos Novos</strong> Pedidos unicos sem R ou A  | 
      <strong>Adj %</strong> Ajustes ÷ total pedidos  | 
      <strong>Rep %</strong> Repeticoes ÷ total pedidos  | 
      <span style="color:var(--vermelho)">Vermelho</span> quando Adj ou Rep > 20%
    </div>"""


def _header_relatorio(tag_titulo):
    return f"""
    <div class="rpt-header">
      <div class="rpt-brand">
        {KION_ARC}
        <div>
          <div class="rpt-nome">KION</div>
          <div class="rpt-sub">DENTAL TECHNOLOGY</div>
        </div>
      </div>
      <div class="rpt-tag">Alerta Comercial<br><strong>{tag_titulo}</strong></div>
    </div>
    <div class="rpt-bar"></div>"""


def _footer_relatorio(mes_ref):
    return f"""
    <div class="rpt-footer">
      <p>
        <strong>Analytics Kion Dental</strong><br>
        analytics@kiondental.tech<br>
        Analise desenvolvida pelo <strong>Time de Tecnologia e Inovacao da Kion</strong>
      </p>
      <p style="text-align:right">
        {KION_ARC}<br>
        Gerado automaticamente | {mes_ref.upper()}
      </p>
    </div>"""


def gerar_relatorio_territorio(nome_resp, codigo, df_terr, mes_ref, cfg, modo_teste):
    """HTML completo e moderno para o arquivo ANEXO (abre no browser)."""
    total_ativos  = len(df_terr)
    em_risco      = df_terr[df_terr["risco"].isin(["ALTO", "MEDIO"])]
    fat_risco     = em_risco["mes_atual"].sum()
    impacto_total = em_risco["impacto_rs"].sum()
    fat_total     = df_terr["mes_atual"].sum()
    fat_proj      = float(df_terr.get("fat_projetado", df_terr["mes_atual"]).sum())
    n_alto        = (df_terr["risco"] == "ALTO").sum()
    n_medio       = (df_terr["risco"] == "MEDIO").sum()
    n_atenc       = (df_terr["risco"] == "ATENCAO").sum()
    n_bloqueados  = (df_terr.get("status_financeiro", pd.Series([])).astype(str) == "Bloqueado").sum()
    total_casos   = int(df_terr.get("casos_novos_atual", pd.Series([0])).sum())
    casos_proj    = int(df_terr.get("casos_projetados", pd.Series([0])).sum())

    banner = ""  # banner de teste removido

    blocos = (
        _bloco_risco_relatorio(df_terr, "ALTO",    "LIGAR HOJE") +
        _bloco_risco_relatorio(df_terr, "MEDIO",   "CHECK-IN ESTA SEMANA") +
        _bloco_risco_relatorio(df_terr, "ATENCAO", "MONITORAR")
    ) or "<p style='color:var(--verde);font-size:13px'>✅ Nenhum cliente em risco no momento.</p>"

    kpis = _kpi_row_relatorio(fat_total, fat_proj, total_ativos, n_alto, n_medio,
                               n_atenc, fat_risco, impacto_total, total_casos, casos_proj, n_bloqueados, mes_ref)

    return f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Alerta {mes_ref.upper()} — {nome_resp} ({codigo})</title>
{CSS_RELATORIO}</head>
<body>
<button class="btn-print" onclick="window.print()">🖨 Imprimir</button>
<div class="container">
  {_header_relatorio('Monitoramento de Faturamento')}
  <div class="rpt-body">
    {banner}
    <p class="rpt-title">Alerta de Faturamento — {nome_resp} ({codigo})</p>
    <p class="rpt-subtitle">{mes_ref.upper()} | Carteira do responsavel | Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
    {kpis}
    {blocos}
    {_legenda_relatorio()}
  </div>
  {_footer_relatorio(mes_ref)}
</div>
</body></html>"""


def gerar_relatorio_gestor(df_ativos, mes_ref, cfg, modo_teste):
    """HTML completo e moderno para o arquivo ANEXO do gestor."""
    em_risco      = df_ativos[df_ativos["risco"].isin(["ALTO", "MEDIO"])]
    fat_total     = df_ativos["mes_atual"].sum()
    fat_proj      = float(df_ativos.get("fat_projetado", df_ativos["mes_atual"]).sum())
    fat_risco     = em_risco["mes_atual"].sum()
    impacto_tot   = em_risco["impacto_rs"].sum()
    n_bloqueados  = (df_ativos.get("status_financeiro", pd.Series([])).astype(str) == "Bloqueado").sum()
    casos_total   = int(df_ativos.get("casos_novos_atual", pd.Series([0])).sum())
    casos_proj_t  = int(df_ativos.get("casos_projetados", pd.Series([0])).sum())
    n_alto        = (df_ativos["risco"] == "ALTO").sum()
    n_medio       = (df_ativos["risco"] == "MEDIO").sum()
    n_atenc       = (df_ativos["risco"] == "ATENCAO").sum()

    banner = ""  # banner de teste removido

    # Tabela de resumo por territorio
    linhas_terr = ""
    for cod, info in cfg["territorios"].items():
        t = df_ativos[df_ativos["vendas"] == cod]
        t_risco = t[t["risco"].isin(["ALTO", "MEDIO"])]
        if t.empty: continue
        t_fp   = float(t.get("fat_projetado", t["mes_atual"]).sum())
        t_cas  = int(t.get("casos_novos_atual", pd.Series([0])).sum())
        t_bloq = (t.get("status_financeiro", pd.Series([])).astype(str) == "Bloqueado").sum()
        bloq_s = f' <span class="st-bloq" style="font-size:9px">&#128683; {t_bloq}</span>' if t_bloq > 0 else ""
        linhas_terr += (
            f"<tr><td><strong style='color:var(--azul)'>{cod}</strong> {info['nome']}</td>"
            f"<td style='font-weight:600'>R$ {brl(t['mes_atual'].sum())}"
            f"<br><span style='font-size:10px;color:var(--suave)'>Proj: R$ {brl(t_fp)}</span></td>"
            f"<td>{len(t)}{bloq_s}</td>"
            f"<td><span class='bdg bdg-alto'>{(t['risco']=='ALTO').sum()}</span> "
            f"<span class='bdg bdg-medio'>{(t['risco']=='MEDIO').sum()}</span></td>"
            f"<td class='neg' style='font-weight:700'>-R$ {brl(t_risco['impacto_rs'].sum())}</td>"
            f"<td style='color:var(--suave)'>{t_cas if t_cas > 0 else '—'}</td></tr>"
        )

    kpis = _kpi_row_relatorio(fat_total, fat_proj, len(df_ativos), n_alto, n_medio,
                               n_atenc, fat_risco, impacto_tot, casos_total, casos_proj_t, n_bloqueados, mes_ref)

    blocos = (
        _bloco_risco_relatorio(df_ativos, "ALTO",    "LIGAR HOJE",           show_vendas=True, cfg=cfg) +
        _bloco_risco_relatorio(df_ativos, "MEDIO",   "CHECK-IN ESTA SEMANA", show_vendas=True, cfg=cfg) +
        _bloco_risco_relatorio(df_ativos, "ATENCAO", "MONITORAR",            show_vendas=True, cfg=cfg)
    )

    return f"""<!DOCTYPE html>
<html lang="pt-BR"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Alerta Consolidado {mes_ref.upper()}</title>
{CSS_RELATORIO}</head>
<body>
<button class="btn-print" onclick="window.print()">🖨 Imprimir</button>
<div class="container">
  {_header_relatorio('Visao Consolidada')}
  <div class="rpt-body">
    {banner}
    <p class="rpt-title">Alerta de Faturamento — Visao Consolidada</p>
    <p class="rpt-subtitle">{mes_ref.upper()} | Todos os territorios | Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
    {kpis}
    <div class="section">
      <div class="section-hd atenc"><span style="font-size:16px">📊</span>
        <span class="section-ttl">Resumo por Territorio</span>
      </div>
      <div class="tbl-wrap"><table>
        <thead><tr>
          <th>Territorio</th><th>Faturamento / Proj.</th><th>Ativos</th>
          <th>Em Risco</th><th>Impacto</th><th>Casos Novos</th>
        </tr></thead>
        <tbody>{linhas_terr}</tbody>
      </table></div>
    </div>
    {blocos}
    {_legenda_relatorio()}
  </div>
  {_footer_relatorio(mes_ref)}
</div>
</body></html>"""


def _corpo_simples(titulo, subtitulo, mes_ref, fat_total, fat_proj,
                   n_alto, n_medio, fat_risco, impacto, top5_linhas,
                   modo_teste, banner_txt, n_bloqueados=0, nome_arquivo=""):
    """E-mail de notificacao simples — corpo enxuto, detalhes no anexo."""
    bloq = (f"<tr><td colspan='3' style='padding:6px 0;color:#c0392b;font-size:11px'>"
            f"&#128683; {n_bloqueados} cliente(s) bloqueado(s)</td></tr>"
            if n_bloqueados > 0 else "")
    anexo_msg = (
        f"<p style='font-size:12px;color:#555;background:#f0fbff;"
        f"border-left:3px solid #00B1D2;padding:10px 14px;margin:0 0 16px;border-radius:0 4px 4px 0'>"
        f"&#128206; O relatorio completo esta no arquivo <strong>{nome_arquivo}</strong> anexado a este e-mail."
        f"</p>"
        if nome_arquivo else ""
    )
    banner = ""  # banner de teste removido
    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">{CSS}</head><body>
<div class='wrapper'>
  {_header('Monitoramento de Faturamento')}
  <div class='body-wrap'>
    {banner}
    <h2>{titulo}</h2>
    <p class='subtitle'>{subtitulo}</p>
    {anexo_msg}
    <table width="100%" cellpadding="0" cellspacing="0"
           style="border-collapse:collapse;background:#f8fbff;border:1px solid #d6eaf8;
                  border-radius:6px;margin-bottom:18px">
      <tr>
        <td style="padding:12px 14px;border-right:1px solid #d6eaf8;vertical-align:top">
          <div class="kpi-label">Faturamento {mes_ref}</div>
          <div class="kpi-value">R$ {brl(fat_total)}</div>
          <div class="kpi-sub">Proj: R$ {brl(fat_proj)}</div>
        </td>
        <td style="padding:12px 14px;border-right:1px solid #d6eaf8;vertical-align:top">
          <div class="kpi-label">Em Risco</div>
          <div class="kpi-value" style="color:#c0392b">{n_alto + n_medio}</div>
          <div class="kpi-sub">
            <span class="badge badge-alto">&#128308; {n_alto}</span> 
            <span class="badge badge-medio">&#128993; {n_medio}</span>
          </div>
        </td>
        <td style="padding:12px 14px;border-right:1px solid #d6eaf8;vertical-align:top">
          <div class="kpi-label">Fat. em Risco</div>
          <div class="kpi-value" style="color:#c0392b">R$ {brl(fat_risco)}</div>
          <div class="kpi-sub">{fat_risco/fat_total*100:.1f}% da carteira</div>
        </td>
      </tr>
    </table>
    <p class="section-title" style="border-left:3px solid #c0392b;padding-left:8px">
      &#128308; Top clientes ALTO risco
    </p>
    <table>
      <thead><tr>
        <th>Cliente</th><th>MRR Atual</th><th>Impacto/mes</th>
      </tr></thead>
      <tbody>{bloq}{top5_linhas}</tbody>
    </table>
  </div>
  {_footer(mes_ref)}
</div>
</body></html>"""


def gerar_email_territorio(nome_resp, codigo, df_terr, mes_ref, cfg, modo_teste,
                            nome_arquivo=""):
    """Corpo simples do e-mail — detalhes completos no anexo HTML."""
    total_ativos = len(df_terr)
    em_risco     = df_terr[df_terr["risco"].isin(["ALTO", "MEDIO"])]
    fat_total    = df_terr["mes_atual"].sum()          # acumulado do setor em junho
    n_alto       = (df_terr["risco"] == "ALTO").sum()
    n_medio      = (df_terr["risco"] == "MEDIO").sum()
    n_atenc      = (df_terr["risco"] == "ATENCAO").sum()

    # ── Projecao no nivel do setor ─────────────────────────────────────────
    # acumulado do setor ÷ dias uteis decorridos × dias uteis restantes
    _periodo  = mes_str_to_period(mes_ref)
    _hoje     = SIM_DATE if SIM_DATE is not None else pd.Timestamp.now()
    _du_total = dias_uteis_mes(_periodo.year, _periodo.month)
    _du_dec   = dias_uteis_ate(_periodo.year, _periodo.month, _hoje)
    _du_rest  = _du_total - _du_dec
    if _du_dec > 0:
        fat_proj = float(fat_total + fat_total / _du_dec * _du_rest)
    else:
        fat_proj = float(fat_total)

    # ── Faturamento e impacto em risco (baseado na projecao) ──────────────
    fat_risco     = em_risco["fat_projetado"].sum()
    impacto_total = em_risco["impacto_rs"].sum()

    n_bloqueados = (df_terr.get("status_financeiro", pd.Series([])).astype(str) == "Bloqueado").sum()

    # Top 5 ALTO por impacto
    top5 = df_terr[df_terr["risco"] == "ALTO"].sort_values("impacto_rs", ascending=False).head(5)
    top5_linhas = ""
    for _, r in top5.iterrows():
        st = _status_badge(r)
        top5_linhas += (
            f"<tr><td><strong>{r['Cliente']}</strong>{st}</td>"
            f"<td>R$ {brl(r['mes_atual'])}</td>"
            f"<td style='color:#c0392b;font-weight:700'>-R$ {brl(r['impacto_rs'])}</td></tr>"
        )

    return _corpo_simples(
        titulo=f"Alerta de Faturamento - {nome_resp} ({codigo})",
        subtitulo=f"{mes_ref.upper()} | {total_ativos} clientes ativos | Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        mes_ref=mes_ref, fat_total=fat_total, fat_proj=fat_proj,
        n_alto=n_alto, n_medio=n_medio, fat_risco=fat_risco,
        impacto=impacto_total, top5_linhas=top5_linhas,
        modo_teste=modo_teste,
        banner_txt=f"Em producao este e-mail vai para {nome_resp}",
        n_bloqueados=n_bloqueados, nome_arquivo=nome_arquivo,
    )


def gerar_email_gestor(df_ativos, mes_ref, cfg, modo_teste, nome_arquivo=""):
    """Corpo simples do e-mail consolidado — detalhes no anexo HTML."""
    em_risco  = df_ativos[df_ativos["risco"].isin(["ALTO", "MEDIO"])]
    fat_total = df_ativos["mes_atual"].sum()

    # Projecao consolidada no nivel da base
    _periodo  = mes_str_to_period(mes_ref)
    _hoje     = SIM_DATE if SIM_DATE is not None else pd.Timestamp.now()
    _du_total = dias_uteis_mes(_periodo.year, _periodo.month)
    _du_dec   = dias_uteis_ate(_periodo.year, _periodo.month, _hoje)
    _du_rest  = _du_total - _du_dec
    if _du_dec > 0:
        fat_proj_tot = float(fat_total + fat_total / _du_dec * _du_rest)
    else:
        fat_proj_tot = float(fat_total)

    fat_risco   = em_risco["fat_projetado"].sum()
    impacto_tot = em_risco["impacto_rs"].sum()
    n_bloqueados  = (df_ativos.get("status_financeiro", pd.Series([])).astype(str) == "Bloqueado").sum()
    n_alto_total  = (df_ativos["risco"] == "ALTO").sum()
    n_medio_total = (df_ativos["risco"] == "MEDIO").sum()

    # Top 5 ALTO por impacto (todos os territorios)
    top5 = df_ativos[df_ativos["risco"] == "ALTO"].sort_values("impacto_rs", ascending=False).head(5)
    top5_linhas = ""
    for _, r in top5.iterrows():
        terr = str(r.get("vendas") or "—")
        nome_r = cfg.get("territorios", {}).get(terr, {}).get("nome", terr)
        st = _status_badge(r)
        top5_linhas += (
            f"<tr><td><strong>{r['Cliente']}</strong>{st}"
            f"<br><small style='color:#9AA0A6'>{terr} {nome_r}</small></td>"
            f"<td>R$ {brl(r['mes_atual'])}</td>"
            f"<td style='color:#c0392b;font-weight:700'>-R$ {brl(r['impacto_rs'])}</td></tr>"
        )

    # Resumo de territorios como linhas extras
    linhas_terr = ""
    for cod, info in cfg["territorios"].items():
        t       = df_ativos[df_ativos["vendas"] == cod]
        t_risco = t[t["risco"].isin(["ALTO", "MEDIO"])]
        if t.empty:
            continue
        t_fat_proj = float(t.get("fat_projetado", t["mes_atual"]).sum())
        t_bloq     = (t.get("status_financeiro", pd.Series([])).astype(str) == "Bloqueado").sum()
        bloq_txt   = f" <span style='color:#c0392b;font-size:10px'>&#128683; {t_bloq}</span>" if t_bloq > 0 else ""
        linhas_terr += (
            f"<tr>"
            f"<td><strong style='color:#00B1D2'>{cod}</strong> {info['nome']}</td>"
            f"<td style='font-weight:700'>R$ {brl(t['mes_atual'].sum())}"
            f"<br><small style='color:#9AA0A6'>Proj: R$ {brl(t_fat_proj)}</small></td>"
            f"<td>{len(t)}{bloq_txt}</td>"
            f"<td><span class='badge badge-alto'>{(t['risco']=='ALTO').sum()}</span>"
            f" <span class='badge badge-medio'>{(t['risco']=='MEDIO').sum()}</span></td>"
            f"<td style='font-weight:700;color:#c0392b'>-R$ {brl(t_risco['impacto_rs'].sum())}</td>"
            f"</tr>"
        )

    # Corpo simples: top 5 + resumo por territorio
    full_top5 = top5_linhas + linhas_terr

    return _corpo_simples(
        titulo="Alerta de Faturamento - Visao Consolidada",
        subtitulo=f"{mes_ref.upper()} | Todos os territorios | Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}",
        mes_ref=mes_ref, fat_total=fat_total, fat_proj=fat_proj_tot,
        n_alto=n_alto_total, n_medio=n_medio_total, fat_risco=fat_risco,
        impacto=impacto_tot, top5_linhas=full_top5,
        modo_teste=modo_teste,
        banner_txt="Em producao este e-mail vai para Bruno Garcia",
        n_bloqueados=n_bloqueados, nome_arquivo=nome_arquivo,
    )


# ─────────────────────────────────────────────
#  ENVIO — Microsoft Graph API
# ─────────────────────────────────────────────

def _get_graph_token():
    tenant_id     = os.environ.get("GRAPH_TENANT_ID", "")
    client_id     = os.environ.get("GRAPH_CLIENT_ID", "")
    client_secret = os.environ.get("GRAPH_CLIENT_SECRET", "")
    if not all([tenant_id, client_id, client_secret]):
        raise ValueError("GRAPH_TENANT_ID, GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET nao configurados.")
    resp = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data={"grant_type": "client_credentials", "client_id": client_id,
              "client_secret": client_secret, "scope": "https://graph.microsoft.com/.default"},
        timeout=15,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def enviar_email(assunto, html, destinatarios, cfg, dry_run=False,
                 nome_anexo=None, html_anexo=None):
    """
    Envia e-mail via Microsoft Graph API.
    html       → corpo do e-mail (notificacao simples)
    html_anexo → conteudo do arquivo anexo (relatorio completo)
    nome_anexo → nome do arquivo .html anexado
    """
    if not isinstance(destinatarios, list):
        destinatarios = [destinatarios]
    if dry_run:
        logging.info(f"  [DRY-RUN] {assunto} → {destinatarios}")
        return True
    try:
        token = _get_graph_token()
        rem   = cfg["remetente"]

        # Anexa o relatorio completo (html_anexo), nao o corpo simples
        attachments = []
        conteudo_anexo = html_anexo or html
        if nome_anexo and cfg.get("envio", {}).get("anexar_html", False):
            html_b64 = base64.b64encode(conteudo_anexo.encode("utf-8")).decode("utf-8")
            attachments.append({
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": nome_anexo,
                "contentType": "text/html",
                "contentBytes": html_b64,
            })

        payload = {
            "message": {
                "subject": assunto,
                "body": {"contentType": "HTML", "content": html},
                "from": {"emailAddress": {"address": rem["email"], "name": rem["nome"]}},
                "toRecipients": [{"emailAddress": {"address": d}} for d in destinatarios],
                **({"attachments": attachments} if attachments else {}),
            },
            "saveToSentItems": "false",
        }
        resp = requests.post(
            f"https://graph.microsoft.com/v1.0/users/{rem['email']}/sendMail",
            json=payload,
            headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
            timeout=30,
        )
        resp.raise_for_status()
        logging.info(f"  ✅ Enviado: {assunto} → {destinatarios}")
        return True
    except Exception as exc:
        logging.error(f"  ❌ Falha ao enviar '{assunto}': {exc}")
        return False


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────

def main():
    cfg = carregar_config()
    configurar_log(cfg["caminhos"]["logs"])

    modo_teste = cfg.get("modo_teste", True)
    prefixo    = "[TESTE] " if modo_teste else ""

    logging.info("=" * 55)
    logging.info(f"  Kion Dental — Alerta Faturamento  {'(MODO TESTE)' if modo_teste else ''}")
    logging.info(f"  {datetime.now().strftime('%d/%m/%Y %H:%M')}  |  dry-run={DRY_RUN}  |  preview={PREVIEW}")
    logging.info("=" * 55)

    # Verificar se o disparo esta ativo
    if not cfg.get("ativo", True) and not PREVIEW and not DRY_RUN:
        logging.info("  ⏸  Sistema PAUSADO (ativo: false no config.yaml) — nenhum e-mail enviado.")
        logging.info("  Para ativar: altere 'ativo: true' no config/config.yaml")
        logging.info("=" * 55)
        return

    # 1. Ler dados
    df_cli, df_2025, df_2026 = ler_dados(cfg)
    df_pedidos = ler_pedidos(cfg["caminhos"]["pedidos"])

    # 2. Processar
    df_ativos, mes_ref = processar(df_cli, df_2025, df_2026, df_pedidos, cfg)

    # ── Modo preview: salva HTMLs em disco sem enviar ─────────────────────
    if PREVIEW:
        preview_dir = os.path.join(BASE_DIR, "saidas", "preview_emails")
        os.makedirs(preview_dir, exist_ok=True)
        logging.info(f"Salvando previews em {preview_dir} ...")

        for cod, info in cfg["territorios"].items():
            df_terr = df_ativos[df_ativos["vendas"] == cod]
            if df_terr.empty:
                continue
            # Salva o relatorio completo (anexo) como preview
            html = gerar_relatorio_territorio(info["nome"], cod, df_terr, mes_ref, cfg, True)
            path = os.path.join(preview_dir, f"{cod}_{info['nome']}.html")
            with open(path, "w", encoding="utf-8") as f:
                f.write(html)
            logging.info(f"  ✅ {os.path.basename(path)}")

        html_g = gerar_relatorio_gestor(df_ativos, mes_ref, cfg, True)
        path_g = os.path.join(preview_dir, "GESTOR_Bruno.html")
        with open(path_g, "w", encoding="utf-8") as f:
            f.write(html_g)
        logging.info(f"  ✅ {os.path.basename(path_g)}")
        logging.info("=" * 55)
        logging.info(f"  Preview completo — {len(cfg['territorios']) + 1} arquivos salvos")
        logging.info("=" * 55)
        return

    # ── Envio normal ───────────────────────────────────────────────────────
    enviados = falhas = 0

    logging.info("Gerando e-mails por territorio...")
    for cod, info in cfg["territorios"].items():
        df_terr = df_ativos[df_ativos["vendas"] == cod]
        if df_terr.empty:
            logging.info(f"  {cod} ({info['nome']}): sem clientes ativos, pulando")
            continue

        n_alto  = (df_terr["risco"] == "ALTO").sum()
        n_medio = (df_terr["risco"] == "MEDIO").sum()
        impacto = df_terr[df_terr["risco"].isin(["ALTO", "MEDIO"])]["impacto_rs"].sum()
        assunto = (
            f"{prefixo}[Churn Alert] {mes_ref.upper()} | "
            f"{cod} {info['nome']} | "
            f"🔴 {n_alto} | 🟡 {n_medio} | R$ {brl(impacto)} em risco"
        )
        nome_anexo  = f"Alerta_{cod}_{info['nome']}_{mes_ref.replace(' ','').upper()}.html"
        html_corpo  = gerar_email_territorio(info["nome"], cod, df_terr, mes_ref, cfg, modo_teste,
                                              nome_arquivo=nome_anexo)
        html_relat  = gerar_relatorio_territorio(info["nome"], cod, df_terr, mes_ref, cfg, modo_teste)
        dests       = cfg["emails_teste"] if modo_teste else [info["email"]]

        ok = enviar_email(assunto, html_corpo, dests, cfg, DRY_RUN,
                          nome_anexo=nome_anexo, html_anexo=html_relat)
        if ok: enviados += 1
        else:  falhas   += 1

    logging.info("Gerando e-mail consolidado para o gestor...")
    em_risco_total = df_ativos[df_ativos["risco"].isin(["ALTO", "MEDIO"])]
    assunto_gestor = (
        f"{prefixo}[Churn CONSOLIDADO] {mes_ref.upper()} | "
        f"R$ {brl(em_risco_total['impacto_rs'].sum())} em risco | "
        f"🔴 {(df_ativos['risco'] == 'ALTO').sum()} | "
        f"🟡 {(df_ativos['risco'] == 'MEDIO').sum()}"
    )
    nome_anexo_g   = f"Alerta_CONSOLIDADO_{mes_ref.replace(' ','').upper()}.html"
    html_gestor    = gerar_email_gestor(df_ativos, mes_ref, cfg, modo_teste, nome_arquivo=nome_anexo_g)
    html_relat_g   = gerar_relatorio_gestor(df_ativos, mes_ref, cfg, modo_teste)
    dests_gestor   = cfg["emails_teste"] if modo_teste else [cfg["gestor"]["email"]]

    ok = enviar_email(assunto_gestor, html_gestor, dests_gestor, cfg, DRY_RUN,
                      nome_anexo=nome_anexo_g, html_anexo=html_relat_g)
    if ok: enviados += 1
    else:  falhas   += 1

    logging.info("=" * 55)
    logging.info(f"  Concluido — {enviados} enviados | {falhas} falhas")
    logging.info("=" * 55)


if __name__ == "__main__":
    main()
