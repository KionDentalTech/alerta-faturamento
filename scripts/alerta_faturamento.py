"""
Kion Dental — Alerta Diário de Faturamento
Executa todo dia às 10h (America/Sao_Paulo) via cron no container Docker.

Uso:
    python alerta_faturamento.py            → roda normalmente
    python alerta_faturamento.py --dry-run  → processa mas não envia e-mails (só loga)

Variáveis de ambiente:
    SMTP_SENHA   → senha do analytics@kiondental.tech (obrigatório em produção)
    KION_BASE    → caminho base do projeto (padrão: /app em Docker, C:\\KionDental local)
"""

import pandas as pd
import numpy as np
import yaml
import smtplib
import logging
import sys
import os
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# ─────────────────────────────────────────────
#  CONFIGURAÇÃO
# ─────────────────────────────────────────────

# Em Docker: KION_BASE=/app  |  Local Windows: C:\KionDental
BASE_DIR    = os.environ.get("KION_BASE", r"C:\KionDental")
CONFIG_PATH = os.path.join(BASE_DIR, "config", "config.yaml")
DRY_RUN     = "--dry-run" in sys.argv

def carregar_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    # Resolve caminhos relativos ao BASE_DIR
    for chave in ["clientes", "producao_2025", "producao_2026", "logs"]:
        cfg["caminhos"][chave] = os.path.join(BASE_DIR, cfg["caminhos"][chave])
    return cfg

# ─────────────────────────────────────────────
#  LOGGING
# ─────────────────────────────────────────────

def configurar_log(log_dir):
    os.makedirs(log_dir, exist_ok=True)
    hoje = datetime.now().strftime("%Y-%m-%d")
    log_file = os.path.join(log_dir, f"alerta_{hoje}.log")

    console = logging.StreamHandler(sys.stdout)
    console.stream = open(sys.stdout.fileno(), mode="w",
                          encoding="utf-8", buffering=1, closefd=False)

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            console,
        ],
    )

# ─────────────────────────────────────────────
#  LEITURA DOS DADOS
# ─────────────────────────────────────────────

def ler_dados(cfg):
    logging.info("Lendo arquivos Excel...")

    df_cli  = pd.read_excel(cfg["caminhos"]["clientes"])
    df_2025 = pd.read_excel(cfg["caminhos"]["producao_2025"], sheet_name="Dados")
    df_2026 = pd.read_excel(cfg["caminhos"]["producao_2026"], sheet_name="Dados")

    logging.info(f"  clientes.xlsx       → {len(df_cli):,} linhas")
    logging.info(f"  producao_2025.xlsx  → {len(df_2025):,} linhas")
    logging.info(f"  producao_2026.xlsx  → {len(df_2026):,} linhas")

    return df_cli, df_2025, df_2026

# ─────────────────────────────────────────────
#  PROCESSAMENTO
# ─────────────────────────────────────────────

MESES_2025 = ["abr 2025","mai 2025","jun 2025","jul 2025","ago 2025",
               "set 2025","out 2025","nov 2025","dez 2025"]

MESES_2026 = ["jan 2026","fev 2026","mar 2026","abr 2026","mai 2026"]

JANELA_12M = ["jun 2025","jul 2025","ago 2025","set 2025","out 2025",
               "nov 2025","dez 2025","jan 2026","fev 2026","mar 2026",
               "abr 2026","mai 2026"]


def detectar_mes_atual(df_2026):
    """Retorna o mês mais recente com pelo menos 100 clientes faturando."""
    for mes in reversed(MESES_2026):
        if mes in df_2026.columns and (pd.to_numeric(df_2026[mes], errors="coerce") > 0).sum() >= 100:
            return mes
    return MESES_2026[-1]


def meses_consecutivos_queda(row, janela, mes_atual):
    """Conta meses seguidos de queda antes do mês atual."""
    idx = janela.index(mes_atual)
    meses_anteriores = list(reversed(janela[:idx]))
    count = 0
    for m in meses_anteriores:
        if row.get(m, 0) > 0 and row[mes_atual] < row[m]:
            count += 1
        else:
            break
    return count


def calcular_nivel_risco(row, cfg, mes_atual):
    th = cfg["thresholds"]
    if row["media_12m"] == 0 or row[mes_atual] == 0:
        return "SEM HISTÓRICO"
    if (row["variacao_pct"] <= -th["alto_queda_pct"]
            and row["meses_queda"] >= th["alto_meses_min"]):
        return "ALTO"
    if row["variacao_pct"] <= -th["medio_queda_pct"]:
        return "MÉDIO"
    if row["variacao_pct"] < 0:
        return "ATENÇÃO"
    return "ESTÁVEL"


def processar(df_cli, df_2025, df_2026, cfg):
    logging.info("Processando dados...")

    # Meses disponíveis em 2025
    meses_25_disp = [m for m in MESES_2025 if m in df_2025.columns]

    # Consolidar faturamento
    df_fat = pd.merge(
        df_2025[["Cliente"] + meses_25_disp],
        df_2026[["Cliente"] + MESES_2026],
        on="Cliente", how="outer"
    ).fillna(0)

    # Join com cadastro de clientes
    df_cli_limpo = df_cli[["Nome","VENDAS","Tabela de preço","Status"]].copy()
    df_cli_limpo.columns = ["Cliente","vendas","tabela","status"]

    df = pd.merge(df_fat, df_cli_limpo, on="Cliente", how="left")

    # Mês atual (mais recente com dados)
    mes_atual = detectar_mes_atual(df_2026)
    logging.info(f"  Mês de referência: {mes_atual}")

    # Janela de 12 meses disponível
    janela = [m for m in JANELA_12M if m in df.columns]

    # Métricas
    df["media_12m"]    = df[janela].mean(axis=1)
    df["mes_atual"]    = df[mes_atual].astype(float)
    df["variacao_pct"] = ((df["mes_atual"] - df["media_12m"])
                          / df["media_12m"].replace(0, np.nan) * 100)
    df["meses_queda"]  = df.apply(
        lambda r: meses_consecutivos_queda(r, janela, mes_atual), axis=1
    )
    df["risco"]        = df.apply(
        lambda r: calcular_nivel_risco(r, cfg, mes_atual), axis=1
    )
    df["impacto_rs"]   = (df["media_12m"] - df["mes_atual"]).clip(lower=0)
    df["mes_ref"]      = mes_atual

    # Apenas clientes com faturamento no mês atual
    df_ativos = df[df["mes_atual"] > 0].copy()

    logging.info(f"  Clientes ativos em {mes_atual}: {len(df_ativos):,}")
    for nivel in ["ALTO","MÉDIO","ATENÇÃO","ESTÁVEL"]:
        n = (df_ativos["risco"] == nivel).sum()
        logging.info(f"    {nivel}: {n}")

    return df_ativos, mes_atual


# ─────────────────────────────────────────────
#  TEMPLATES HTML
# ─────────────────────────────────────────────

EMOJI = {"ALTO": "🔴", "MÉDIO": "🟡", "ATENÇÃO": "🟢", "ESTÁVEL": "✅"}


def _tabela_html(df_grupo, colunas, headers):
    linhas = ""
    for _, r in df_grupo.iterrows():
        linhas += "<tr>" + "".join(f"<td>{r[c]}</td>" for c in colunas) + "</tr>"
    cabecalho = "".join(f"<th>{h}</th>" for h in headers)
    return f"""
    <table>
      <thead><tr>{cabecalho}</tr></thead>
      <tbody>{linhas}</tbody>
    </table>"""


def _narrativa(r):
    meses_txt = f"{int(r['meses_queda'])} {'mês' if r['meses_queda'] == 1 else 'meses'} seguidos"
    return (
        f"<em style='color:#5A5A5A;font-size:11px;line-height:1.6'>"
        f"Costumava gerar <strong>R$ {r['media_12m']:,.0f}/m&ecirc;s</strong>. "
        f"Hoje est&aacute; em <strong>R$ {r['mes_atual']:,.0f}</strong>. "
        f"Queda de <strong style='color:#c0392b'>{abs(r['variacao_pct']):.0f}%</strong>, "
        f"h&aacute; <strong>{meses_txt}</strong>. "
        f"Isso representa <strong style='color:#c0392b'>R$ {r['impacto_rs']:,.0f}/m&ecirc;s</strong> "
        f"que a Kion parou de receber desse cliente."
        f"</em>"
    )


def _bloco_risco(df_terr, nivel, label_acao):
    grupo = df_terr[df_terr["risco"] == nivel].sort_values("impacto_rs", ascending=False)
    if grupo.empty:
        return ""
    emoji = EMOJI.get(nivel, "")
    rows = ""
    for _, r in grupo.iterrows():
        tabela = r['tabela'] if pd.notna(r['tabela']) else ''
        rows += f"""
        <tr>
          <td>
            <strong>{r['Cliente']}</strong><br>
            <small style='color:#8D8E8F'>{tabela}</small><br>
            {_narrativa(r)}
          </td>
          <td>R$ {r['mes_atual']:,.0f}</td>
          <td>R$ {r['media_12m']:,.0f}</td>
          <td class='neg'><strong>{r['variacao_pct']:+.0f}%</strong></td>
          <td>{int(r['meses_queda'])}m</td>
          <td class='neg'>-R$ {r['impacto_rs']:,.0f}</td>
        </tr>"""
    return f"""
    <h3>{emoji} {nivel} &mdash; {label_acao}</h3>
    <table>
      <thead>
        <tr>
          <th>Cliente</th><th>MRR Atual</th><th>M&eacute;dia 12M</th>
          <th>Varia&ccedil;&atilde;o</th><th>Meses queda</th><th>Impacto/m&ecirc;s</th>
        </tr>
      </thead>
      <tbody>{rows}</tbody>
    </table>"""


# ── Kion logo arc (SVG inline, compatível com clientes de e-mail modernos) ──
KION_ARC = (
    '<svg viewBox="0 0 110 100" width="36" height="32" '
    'style="vertical-align:middle;margin-right:10px">'
    '<defs><linearGradient id="kg" x1="0%" y1="0%" x2="100%" y2="0%">'
    '<stop offset="0%" stop-color="#00F5FF"/>'
    '<stop offset="100%" stop-color="#FAEB1E"/>'
    '</linearGradient></defs>'
    '<path d="M 12 92 A 46 46 0 1 1 98 92" stroke="url(#kg)" '
    'stroke-width="10" fill="none" stroke-linecap="round"/>'
    '</svg>'
)

# ── CSS Kion Brand ──────────────────────────────────────────────────────────
CSS = """
<style>
  body{{margin:0;padding:0;background:#f4f4f4;
        font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#282828}}
  .wrapper{{max-width:920px;margin:0 auto;background:#fff}}
  .kion-header{{background:#282828;padding:16px 24px;
                display:flex;align-items:center;justify-content:space-between}}
  .kion-name{{font-size:21px;font-weight:900;letter-spacing:3px;color:#00F5FF}}
  .kion-sub{{font-size:10px;color:#8D8E8F;letter-spacing:1px}}
  .kion-tag{{font-size:11px;color:#8D8E8F;text-align:right;line-height:1.5}}
  .kion-tag strong{{color:#00B1D2}}
  .kion-bar{{height:4px;background:linear-gradient(90deg,#00F5FF 0%,#00B1D2 50%,#FAEB1E 100%)}}
  .body-wrap{{padding:22px 26px}}
  h2{{color:#282828;font-size:16px;margin:0 0 3px;padding-bottom:8px;border-bottom:2px solid #00B1D2}}
  .subtitle{{color:#8D8E8F;font-size:12px;margin:0 0 16px}}
  h3{{color:#00B1D2;font-size:13px;margin:26px 0 5px}}
  table{{border-collapse:collapse;width:100%;margin:4px 0 18px;font-size:12px}}
  th{{background:#00B1D2;color:#fff;padding:8px 10px;text-align:left;
      font-size:11px;text-transform:uppercase;letter-spacing:.4px}}
  td{{padding:7px 10px;border-bottom:1px solid #e8e8e8;vertical-align:top}}
  tr:nth-child(even) td{{background:#f8fdff}}
  .resumo{{background:#f0fbff;border-left:4px solid #00F5FF;
           padding:13px 17px;margin-bottom:20px;border-radius:0 6px 6px 0;line-height:2}}
  .badge{{display:inline-block;padding:2px 8px;border-radius:12px;
          font-size:11px;font-weight:bold;margin-right:4px}}
  .alto{{background:#fde8e8;color:#c0392b}}
  .medio{{background:#fff8cc;color:#8a6500}}
  .neg{{color:#c0392b;font-weight:bold}}
  .banner{{background:#00B1D2;color:#fff;padding:9px 14px;
           font-weight:bold;border-radius:4px;margin-bottom:16px;font-size:12px}}
  .kion-footer{{background:#282828;padding:14px 24px}}
  .kion-footer-inner{{border-top:1px solid #3a3a3a;padding-top:12px;
                      display:flex;align-items:center;justify-content:space-between}}
  .footer-left{{font-size:11px;color:#8D8E8F;line-height:1.9}}
  .footer-left strong{{color:#00F5FF}}
  .footer-right{{font-size:10px;color:#5A5A5A;text-align:right;line-height:1.8}}
</style>
"""

def _header(tag_titulo):
    return (
        f"<div class='kion-header'>"
        f"<div style='display:flex;align-items:center'>{KION_ARC}"
        f"<div><div class='kion-name'>KION</div>"
        f"<div class='kion-sub'>DENTAL TECHNOLOGY</div></div></div>"
        f"<div class='kion-tag'>Alerta Comercial<br><strong>{tag_titulo}</strong></div>"
        f"</div><div class='kion-bar'></div>"
    )

def _footer(mes_ref):
    return (
        f"<div class='kion-footer'><div class='kion-footer-inner'>"
        f"<div class='footer-left'>"
        f"<strong>Analytics Kion Dental</strong><br>"
        f"analytics@kiondental.tech<br>"
        f"An&aacute;lise desenvolvida pelo "
        f"<strong>Time de Tecnologia e Inova&ccedil;&atilde;o da Kion</strong>"
        f"</div>"
        f"<div class='footer-right'>{KION_ARC}<br>"
        f"Gerado automaticamente &bull; {mes_ref.upper()}</div>"
        f"</div></div>"
    )


def gerar_email_territorio(nome_resp, codigo, df_terr, mes_ref, cfg, modo_teste):
    total_ativos  = len(df_terr)
    em_risco      = df_terr[df_terr["risco"].isin(["ALTO","MÉDIO"])]
    fat_risco     = em_risco["mes_atual"].sum()
    impacto_total = em_risco["impacto_rs"].sum()
    fat_total     = df_terr["mes_atual"].sum()
    n_alto        = (df_terr["risco"] == "ALTO").sum()
    n_medio       = (df_terr["risco"] == "MÉDIO").sum()
    n_atenc       = (df_terr["risco"] == "ATENÇÃO").sum()

    banner = (f"<div class='banner'>&#9888;&#65039; MODO TESTE &mdash; "
              f"Em produ&ccedil;&atilde;o este e-mail vai para {nome_resp}</div>"
              if modo_teste else "")

    blocos = (
        _bloco_risco(df_terr, "ALTO",   "LIGAR HOJE") +
        _bloco_risco(df_terr, "MÉDIO",  "CHECK-IN ESTA SEMANA") +
        _bloco_risco(df_terr, "ATENÇÃO","MONITORAR")
    )
    if not blocos:
        blocos = "<p style='color:#00B1D2'>&#9989; Nenhum cliente em risco no momento.</p>"

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">{CSS}</head><body>
<div class='wrapper'>
  {_header('Monitoramento de Faturamento')}
  <div class='body-wrap'>
    {banner}
    <h2>Alerta de Faturamento &mdash; {nome_resp} ({codigo})</h2>
    <p class='subtitle'>{mes_ref.upper()} &bull; Carteira do respons&aacute;vel
       &bull; Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
    <div class='resumo'>
      <strong>Resumo da carteira</strong><br>
      Faturamento {mes_ref}: <strong>R$ {fat_total:,.0f}</strong>
      &nbsp;|&nbsp; Clientes ativos: <strong>{total_ativos}</strong><br>
      Em risco:
        <span class='badge alto'>&#128308; ALTO: {n_alto}</span>
        <span class='badge medio'>&#129473; M&Eacute;DIO: {n_medio}</span>
        &nbsp;Aten&ccedil;&atilde;o: {n_atenc}<br>
      Faturamento em risco:
        <strong class='neg'>R$ {fat_risco:,.0f}</strong>
        ({fat_risco/fat_total*100:.1f}% da carteira)<br>
      Impacto vs. m&eacute;dia 12M:
        <strong class='neg'>-R$ {impacto_total:,.0f}/m&ecirc;s</strong>
    </div>
    {blocos}
  </div>
  {_footer(mes_ref)}
</div>
</body></html>"""


def gerar_email_gestor(df_ativos, mes_ref, cfg, modo_teste):
    em_risco    = df_ativos[df_ativos["risco"].isin(["ALTO","MÉDIO"])]
    fat_total   = df_ativos["mes_atual"].sum()
    fat_risco   = em_risco["mes_atual"].sum()
    impacto_tot = em_risco["impacto_rs"].sum()

    banner = ("<div class='banner'>&#9888;&#65039; MODO TESTE &mdash; "
              "Em produ&ccedil;&atilde;o este e-mail vai para Bruno Garcia</div>"
              if modo_teste else "")

    top = em_risco.sort_values("impacto_rs", ascending=False).head(15)
    linhas_top = ""
    for _, r in top.iterrows():
        terr  = r["vendas"] if pd.notna(r["vendas"]) else "—"
        resp  = cfg["territorios"].get(terr, {}).get("nome", terr)
        emoji = EMOJI.get(r["risco"], "")
        tab   = r["tabela"] if pd.notna(r["tabela"]) else ""
        linhas_top += (
            f"<tr><td>{emoji} <strong>{r['risco']}</strong></td>"
            f"<td><strong>{r['Cliente']}</strong><br>"
            f"<small style='color:#8D8E8F'>{tab}</small></td>"
            f"<td>{terr} {resp}</td>"
            f"<td>R$ {r['mes_atual']:,.0f}</td>"
            f"<td class='neg'>{r['variacao_pct']:+.0f}%</td>"
            f"<td>{int(r['meses_queda'])}m</td>"
            f"<td class='neg'><strong>-R$ {r['impacto_rs']:,.0f}</strong></td></tr>"
        )

    linhas_terr = ""
    for cod, info in cfg["territorios"].items():
        t       = df_ativos[df_ativos["vendas"] == cod]
        t_risco = t[t["risco"].isin(["ALTO","MÉDIO"])]
        if t.empty:
            continue
        linhas_terr += (
            f"<tr>"
            f"<td><strong style='color:#00B1D2'>{cod}</strong> {info['nome']}</td>"
            f"<td>{len(t)}</td>"
            f"<td><span class='badge alto'>{(t['risco']=='ALTO').sum()}</span>"
            f"<span class='badge medio'>{(t['risco']=='MÉDIO').sum()}</span></td>"
            f"<td class='neg'><strong>R$ {t_risco['impacto_rs'].sum():,.0f}</strong></td>"
            f"</tr>"
        )

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8">{CSS}</head><body>
<div class='wrapper'>
  {_header('Vis&atilde;o Consolidada')}
  <div class='body-wrap'>
    {banner}
    <h2>Alerta de Faturamento &mdash; Vis&atilde;o Consolidada</h2>
    <p class='subtitle'>{mes_ref.upper()} &bull; Resumo executivo de todos os territ&oacute;rios
       &bull; Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
    <div class='resumo'>
      <strong>Vis&atilde;o Executiva</strong><br>
      Faturamento total {mes_ref}: <strong>R$ {fat_total:,.0f}</strong>
      &nbsp;|&nbsp; Clientes ativos: <strong>{len(df_ativos):,}</strong><br>
      Em risco:
        <span class='badge alto'>&#128308; ALTO: {(df_ativos['risco']=='ALTO').sum()}</span>
        <span class='badge medio'>&#129473; M&Eacute;DIO: {(df_ativos['risco']=='MÉDIO').sum()}</span><br>
      Faturamento em risco:
        <strong class='neg'>R$ {fat_risco:,.0f}</strong>
        ({fat_risco/fat_total*100:.1f}% do total)<br>
      Impacto potencial vs. m&eacute;dia 12M:
        <strong class='neg'>-R$ {impacto_tot:,.0f}/m&ecirc;s</strong>
    </div>
    <h3>&#127942; Top 15 &mdash; Maior Impacto Financeiro</h3>
    <table>
      <thead><tr>
        <th>Risco</th><th>Cliente</th><th>Respons&aacute;vel</th>
        <th>MRR Atual</th><th>Varia&ccedil;&atilde;o</th>
        <th>Meses</th><th>Impacto</th>
      </tr></thead>
      <tbody>{linhas_top}</tbody>
    </table>
    <h3>&#128202; Resumo por Territ&oacute;rio</h3>
    <table>
      <thead><tr>
        <th>Territ&oacute;rio</th><th>Ativos</th>
        <th>Em Risco</th><th>Impacto R$</th>
      </tr></thead>
      <tbody>{linhas_terr}</tbody>
    </table>
  </div>
  {_footer(mes_ref)}
</div>
</body></html>"""


# ─────────────────────────────────────────────
#  ENVIO DE E-MAIL
# ─────────────────────────────────────────────

def enviar_email(assunto, html, destinatarios, cfg, dry_run=False):
    if not isinstance(destinatarios, list):
        destinatarios = [destinatarios]

    msg = MIMEMultipart("alternative")
    msg["Subject"] = assunto
    msg["From"]    = f"{cfg['remetente']['nome']} <{cfg['remetente']['email']}>"
    msg["To"]      = ", ".join(destinatarios)
    msg.attach(MIMEText(html, "html", "utf-8"))

    if dry_run:
        logging.info(f"  [DRY-RUN] {assunto} → {destinatarios}")
        return True

    try:
        # Senha via env var (produção) ou config.yaml (local)
        senha = os.environ.get("SMTP_SENHA") or cfg["smtp"]["senha"]
        with smtplib.SMTP(cfg["smtp"]["servidor"], cfg["smtp"]["porta"]) as server:
            server.ehlo()
            server.starttls()
            server.login(cfg["smtp"]["usuario"], senha)
            server.sendmail(cfg["remetente"]["email"], destinatarios, msg.as_string())
        logging.info(f"  ✅ Enviado: {assunto} → {destinatarios}")
        return True
    except Exception as e:
        logging.error(f"  ❌ Falha ao enviar '{assunto}': {e}")
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
    logging.info(f"  Kion Dental — Alerta Churn  {'(MODO TESTE)' if modo_teste else ''}")
    logging.info(f"  {datetime.now().strftime('%d/%m/%Y %H:%M')}  |  dry-run={DRY_RUN}")
    logging.info("=" * 55)

    # 1. Ler dados
    df_cli, df_2025, df_2026 = ler_dados(cfg)

    # 2. Processar
    df_ativos, mes_ref = processar(df_cli, df_2025, df_2026, cfg)

    enviados = 0
    falhas   = 0

    # 3. E-mails por território
    logging.info("Gerando e-mails por território...")
    for cod, info in cfg["territorios"].items():
        df_terr = df_ativos[df_ativos["vendas"] == cod]

        if df_terr.empty:
            logging.info(f"  {cod} ({info['nome']}): sem clientes ativos, pulando")
            continue

        em_risco = (df_terr["risco"].isin(["ALTO","MÉDIO"])).sum()
        assunto  = (
            f"{prefixo}[Churn Alert] {mes_ref.upper()} | "
            f"{cod} {info['nome']} | "
            f"🔴 {(df_terr['risco']=='ALTO').sum()} | "
            f"🟡 {(df_terr['risco']=='MÉDIO').sum()} | "
            f"R$ {df_terr[df_terr['risco'].isin(['ALTO','MÉDIO'])]['impacto_rs'].sum():,.0f} em risco"
        )
        html  = gerar_email_territorio(info["nome"], cod, df_terr, mes_ref, cfg, modo_teste)
        dests = cfg["emails_teste"] if modo_teste else [info["email"]]

        ok = enviar_email(assunto, html, dests, cfg, DRY_RUN)
        if ok: enviados += 1
        else:  falhas   += 1

    # 4. E-mail consolidado gestor
    logging.info("Gerando e-mail consolidado para o gestor...")
    em_risco_total = df_ativos[df_ativos["risco"].isin(["ALTO","MÉDIO"])]
    assunto_gestor = (
        f"{prefixo}[Churn CONSOLIDADO] {mes_ref.upper()} | "
        f"R$ {em_risco_total['impacto_rs'].sum():,.0f} em risco | "
        f"🔴 {(df_ativos['risco']=='ALTO').sum()} | "
        f"🟡 {(df_ativos['risco']=='MÉDIO').sum()}"
    )
    html_gestor = gerar_email_gestor(df_ativos, mes_ref, cfg, modo_teste)
    dests_gestor = cfg["emails_teste"] if modo_teste else [cfg["gestor"]["email"]]

    ok = enviar_email(assunto_gestor, html_gestor, dests_gestor, cfg, DRY_RUN)
    if ok: enviados += 1
    else:  falhas   += 1

    # 5. Resumo
    logging.info("=" * 55)
    logging.info(f"  Concluído — {enviados} enviados | {falhas} falhas")
    logging.info("=" * 55)


if __name__ == "__main__":
    main()
