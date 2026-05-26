"""Gera arquivos HTML de preview de cada e-mail por território."""

import pandas as pd
import numpy as np
import yaml
import os

with open(r"C:\KionDental\config\config.yaml", encoding="utf-8") as f:
    cfg = yaml.safe_load(f)

MESES_2025 = ["abr 2025","mai 2025","jun 2025","jul 2025","ago 2025",
               "set 2025","out 2025","nov 2025","dez 2025"]
MESES_2026 = ["jan 2026","fev 2026","mar 2026","abr 2026","mai 2026"]
JANELA_12M = ["jun 2025","jul 2025","ago 2025","set 2025","out 2025",
               "nov 2025","dez 2025","jan 2026","fev 2026","mar 2026",
               "abr 2026","mai 2026"]

# ── leitura ──
df_cli  = pd.read_excel(r"C:\KionDental\dados\clientes.xlsx")
df_2025 = pd.read_excel(r"C:\KionDental\dados\producao_2025.xlsx", sheet_name="Dados")
df_2026 = pd.read_excel(r"C:\KionDental\dados\producao_2026.xlsx", sheet_name="Dados")

meses_25_disp = [m for m in MESES_2025 if m in df_2025.columns]
df_fat = pd.merge(
    df_2025[["Cliente"] + meses_25_disp],
    df_2026[["Cliente"] + MESES_2026],
    on="Cliente", how="outer"
).fillna(0)

df_c = df_cli[["Nome","VENDAS","Tabela de preço","Status"]].copy()
df_c.columns = ["Cliente","vendas","tabela","status"]
df = pd.merge(df_fat, df_c, on="Cliente", how="left")

mes_atual = "mai 2026"
janela    = [m for m in JANELA_12M if m in df.columns]

df["media_12m"]    = df[janela].mean(axis=1)
df["mes_atual"]    = df[mes_atual].astype(float)
df["variacao_pct"] = ((df["mes_atual"] - df["media_12m"])
                       / df["media_12m"].replace(0, np.nan) * 100)

def meses_queda_fn(row):
    idx  = janela.index(mes_atual)
    prev = list(reversed(janela[:idx]))
    cnt  = 0
    for m in prev:
        if row.get(m, 0) > 0 and row[mes_atual] < row[m]:
            cnt += 1
        else:
            break
    return cnt

df["meses_queda"] = df.apply(meses_queda_fn, axis=1)

def risco_fn(row):
    if row["media_12m"] == 0 or row["mes_atual"] == 0:
        return "SEM_HISTORICO"
    if row["variacao_pct"] <= -20 and row["meses_queda"] >= 2:
        return "ALTO"
    if row["variacao_pct"] <= -10:
        return "MEDIO"
    if row["variacao_pct"] < 0:
        return "ATENCAO"
    return "ESTAVEL"

df["risco"]      = df.apply(risco_fn, axis=1)
df["impacto_rs"] = (df["media_12m"] - df["mes_atual"]).clip(lower=0)
df_ativos        = df[df["mes_atual"] > 0].copy()

# ── Kion logo arc (SVG inline) ──
KION_ARC = """<svg viewBox="0 0 110 100" width="38" height="34" style="vertical-align:middle;margin-right:10px">
  <defs>
    <linearGradient id="kg" x1="0%" y1="0%" x2="100%" y2="0%">
      <stop offset="0%"   stop-color="#00F5FF"/>
      <stop offset="100%" stop-color="#FAEB1E"/>
    </linearGradient>
  </defs>
  <path d="M 12 92 A 46 46 0 1 1 98 92"
        stroke="url(#kg)" stroke-width="10"
        fill="none" stroke-linecap="round"/>
</svg>"""

# ── CSS Kion Brand ──
CSS = """
<style>
  /* --- reset --- */
  body{margin:0;padding:0;background:#f4f4f4;font-family:Arial,Helvetica,sans-serif;
       font-size:13px;color:#282828}
  .wrapper{max-width:920px;margin:0 auto;background:#fff}

  /* --- header Kion --- */
  .kion-header{background:#282828;padding:18px 28px;display:flex;
               align-items:center;justify-content:space-between}
  .kion-brand{display:flex;align-items:center}
  .kion-name{font-size:22px;font-weight:900;letter-spacing:3px;color:#00F5FF}
  .kion-sub{font-size:10px;color:#8D8E8F;letter-spacing:1px;margin-top:1px}
  .kion-tag{font-size:11px;color:#8D8E8F;text-align:right;line-height:1.5}
  .kion-tag strong{color:#00B1D2}

  /* --- gradient bar --- */
  .kion-bar{height:4px;background:linear-gradient(90deg,#00F5FF 0%,#00B1D2 50%,#FAEB1E 100%)}

  /* --- body padding --- */
  .body-wrap{padding:24px 28px}

  /* --- título --- */
  h2{color:#282828;font-size:17px;margin:0 0 4px;padding-bottom:8px;
     border-bottom:2px solid #00B1D2}
  .subtitle{color:#8D8E8F;font-size:12px;margin:0 0 18px}
  h3{color:#00B1D2;font-size:14px;margin:28px 0 6px;display:flex;align-items:center;gap:6px}

  /* --- resumo box --- */
  .resumo{background:#f0fbff;border-left:4px solid #00F5FF;padding:14px 18px;
          margin-bottom:22px;border-radius:0 6px 6px 0;line-height:2.1}

  /* --- tabelas --- */
  table{border-collapse:collapse;width:100%;margin:4px 0 20px;font-size:12px}
  th{background:#00B1D2;color:#fff;padding:9px 10px;text-align:left;font-size:11px;
     text-transform:uppercase;letter-spacing:.4px}
  td{padding:8px 10px;border-bottom:1px solid #e8e8e8;vertical-align:top}
  tr:nth-child(even) td{background:#f8fdff}
  tr:hover td{background:#edf8fb}

  /* --- badges --- */
  .badge{display:inline-block;padding:2px 9px;border-radius:12px;
         font-size:11px;font-weight:bold;margin-right:4px}
  .alto{background:#fde8e8;color:#c0392b}
  .medio{background:#fff8cc;color:#8a6500}

  /* --- valores negativos --- */
  .neg{color:#c0392b;font-weight:bold}

  /* --- banner de teste --- */
  .banner{background:#00B1D2;color:#fff;padding:10px 16px;font-weight:bold;
          border-radius:4px;margin-bottom:18px;font-size:12px;letter-spacing:.3px}

  /* --- footer Kion --- */
  .kion-footer{background:#282828;padding:16px 28px;margin-top:10px}
  .kion-footer-inner{border-top:1px solid #3a3a3a;padding-top:14px;
                     display:flex;align-items:center;justify-content:space-between}
  .footer-left{font-size:11px;color:#8D8E8F;line-height:1.8}
  .footer-left strong{color:#00F5FF}
  .footer-right{font-size:10px;color:#5A5A5A;text-align:right;line-height:1.8}
</style>
"""

# ── helpers ──
def bloco(df_g, nivel, acao, emoji):
    sub = df_g[df_g["risco"] == nivel].sort_values("impacto_rs", ascending=False)
    if sub.empty:
        return ""
    rows = ""
    for _, r in sub.iterrows():
        tabela = r["tabela"] if pd.notna(r["tabela"]) else ""
        rows += f"""
        <tr>
          <td><strong>{r['Cliente']}</strong><br>
              <small style='color:#888'>{tabela}</small></td>
          <td>R$ {r['mes_atual']:,.0f}</td>
          <td>R$ {r['media_12m']:,.0f}</td>
          <td class='neg'>{r['variacao_pct']:+.0f}%</td>
          <td>{int(r['meses_queda'])} meses</td>
          <td class='neg'>-R$ {r['impacto_rs']:,.0f}</td>
        </tr>"""
    return f"""
    <h3>{emoji} {nivel} &mdash; {acao}</h3>
    <table>
      <thead>
        <tr><th>Cliente</th><th>MRR Atual</th><th>M&eacute;dia 12M</th>
            <th>Varia&ccedil;&atilde;o</th><th>Meses queda</th><th>Impacto/m&ecirc;s</th></tr>
      </thead>
      <tbody>{rows}</tbody>
    </table>"""


HEADER_HTML = """
<div class='kion-header'>
  <div class='kion-brand'>
    {arc}
    <div>
      <div class='kion-name'>KION</div>
      <div class='kion-sub'>DENTAL TECHNOLOGY</div>
    </div>
  </div>
  <div class='kion-tag'>
    Alerta Comercial<br>
    <strong>Monitoramento de Churn</strong>
  </div>
</div>
<div class='kion-bar'></div>
""".format(arc=KION_ARC)

FOOTER_HTML = """
<div class='kion-footer'>
  <div class='kion-footer-inner'>
    <div class='footer-left'>
      <strong>Analytics Kion Dental</strong><br>
      analytics@kiondental.tech<br>
      An&aacute;lise desenvolvida pelo <strong>Time de Tecnologia e Inova&ccedil;&atilde;o da Kion</strong>
    </div>
    <div class='footer-right'>
      {arc}<br>
      Gerado automaticamente &bull; {date}
    </div>
  </div>
</div>
""".format(arc=KION_ARC, date="MAI 2026")


def email_territorio(nome, cod, df_t):
    fat      = df_t["mes_atual"].sum()
    risco_df = df_t[df_t["risco"].isin(["ALTO","MEDIO"])]
    imp      = risco_df["impacto_rs"].sum()
    fat_r    = risco_df["mes_atual"].sum()
    pct      = fat_r / fat * 100 if fat else 0
    n_alto   = (df_t["risco"] == "ALTO").sum()
    n_medio  = (df_t["risco"] == "MEDIO").sum()
    n_atenc  = (df_t["risco"] == "ATENCAO").sum()

    cont = (bloco(df_t, "ALTO",    "LIGAR HOJE",            "&#128308;") +
            bloco(df_t, "MEDIO",   "CHECK-IN ESTA SEMANA",  "&#129473;") +
            bloco(df_t, "ATENCAO", "MONITORAR",             "&#128994;"))
    if not cont.strip():
        cont = "<p style='color:#00B1D2'>&#9989; Nenhum cliente em risco no momento.</p>"

    return f"""<!DOCTYPE html>
<html><head><meta charset='utf-8'>{CSS}</head><body>
<div class='wrapper'>
  {HEADER_HTML}
  <div class='body-wrap'>
    <div class='banner'>&#9888;&#65039; MODO TESTE &mdash; Em produ&ccedil;&atilde;o este e-mail vai para {nome}</div>
    <h2>Alerta de Churn &mdash; {nome} ({cod})</h2>
    <p class='subtitle'>MAI 2026 &nbsp;&bull;&nbsp; Carteira do respons&aacute;vel</p>

    <div class='resumo'>
      <strong>Resumo da carteira</strong><br>
      Faturamento mai/2026: <strong>R$ {fat:,.0f}</strong>
      &nbsp;&nbsp;|&nbsp;&nbsp; Clientes ativos: <strong>{len(df_t)}</strong><br>
      Em risco:
        <span class='badge alto'>&#128308; ALTO: {n_alto}</span>
        <span class='badge medio'>&#129473; M&Eacute;DIO: {n_medio}</span>
        &nbsp; Aten&ccedil;&atilde;o: {n_atenc}<br>
      Faturamento em risco: <strong class='neg'>R$ {fat_r:,.0f}</strong>
        ({pct:.1f}% da carteira)<br>
      Impacto vs. m&eacute;dia 12M:
        <strong class='neg'>-R$ {imp:,.0f}/m&ecirc;s</strong>
    </div>

    {cont}
  </div>
  {FOOTER_HTML}
</div>
</body></html>"""


def email_gestor(df_a):
    fat   = df_a["mes_atual"].sum()
    er    = df_a[df_a["risco"].isin(["ALTO","MEDIO"])]
    imp   = er["impacto_rs"].sum()
    fat_r = er["mes_atual"].sum()
    top   = er.sort_values("impacto_rs", ascending=False).head(15)

    rows_top = ""
    for _, r in top.iterrows():
        cod_t  = r["vendas"] if pd.notna(r["vendas"]) else "—"
        resp   = cfg["territorios"].get(cod_t, {}).get("nome", cod_t)
        emoji  = "&#128308;" if r["risco"] == "ALTO" else "&#129473;"
        tabela = r["tabela"] if pd.notna(r["tabela"]) else ""
        rows_top += f"""
        <tr>
          <td>{emoji} <strong>{r['risco']}</strong></td>
          <td><strong>{r['Cliente']}</strong><br><small style='color:#8D8E8F'>{tabela}</small></td>
          <td>{cod_t} {resp}</td>
          <td>R$ {r['mes_atual']:,.0f}</td>
          <td class='neg'>{r['variacao_pct']:+.0f}%</td>
          <td>{int(r['meses_queda'])}m</td>
          <td class='neg'><strong>-R$ {r['impacto_rs']:,.0f}</strong></td>
        </tr>"""

    rows_terr = ""
    for cod_t, info in cfg["territorios"].items():
        t  = df_a[df_a["vendas"] == cod_t]
        tr = t[t["risco"].isin(["ALTO","MEDIO"])]
        if t.empty:
            continue
        rows_terr += f"""
        <tr>
          <td><strong style='color:#00B1D2'>{cod_t}</strong> &nbsp;{info['nome']}</td>
          <td>{len(t)}</td>
          <td>
            <span class='badge alto'>{(t['risco']=='ALTO').sum()}</span>
            <span class='badge medio'>{(t['risco']=='MEDIO').sum()}</span>
          </td>
          <td class='neg'><strong>R$ {tr['impacto_rs'].sum():,.0f}</strong></td>
        </tr>"""

    return f"""<!DOCTYPE html>
<html><head><meta charset='utf-8'>{CSS}</head><body>
<div class='wrapper'>
  {HEADER_HTML}
  <div class='body-wrap'>
    <div class='banner'>&#9888;&#65039; MODO TESTE &mdash; Em produ&ccedil;&atilde;o este e-mail vai para Bruno Garcia</div>
    <h2>Alerta de Churn &mdash; Vis&atilde;o Consolidada</h2>
    <p class='subtitle'>MAI 2026 &nbsp;&bull;&nbsp; Resumo executivo de todos os territ&oacute;rios</p>

    <div class='resumo'>
      <strong>Vis&atilde;o Executiva</strong><br>
      Faturamento total mai/2026: <strong>R$ {fat:,.0f}</strong>
      &nbsp;&nbsp;|&nbsp;&nbsp; Clientes ativos: <strong>{len(df_a):,}</strong><br>
      Em risco:
        <span class='badge alto'>&#128308; ALTO: {(df_a['risco']=='ALTO').sum()}</span>
        <span class='badge medio'>&#129473; M&Eacute;DIO: {(df_a['risco']=='MEDIO').sum()}</span><br>
      Faturamento em risco:
        <strong class='neg'>R$ {fat_r:,.0f}</strong>
        ({fat_r/fat*100:.1f}% do total)<br>
      Impacto potencial vs. m&eacute;dia 12M:
        <strong class='neg'>-R$ {imp:,.0f}/m&ecirc;s</strong>
    </div>

    <h3>&#127942; Top 15 &mdash; Maior Impacto Financeiro</h3>
    <table>
      <thead>
        <tr><th>Risco</th><th>Cliente</th><th>Respons&aacute;vel</th>
            <th>MRR Atual</th><th>Varia&ccedil;&atilde;o</th><th>Meses</th><th>Impacto</th></tr>
      </thead>
      <tbody>{rows_top}</tbody>
    </table>

    <h3>&#128202; Resumo por Territ&oacute;rio</h3>
    <table>
      <thead>
        <tr><th>Territ&oacute;rio</th><th>Ativos</th>
            <th>Em Risco</th><th>Impacto R$</th></tr>
      </thead>
      <tbody>{rows_terr}</tbody>
    </table>
  </div>
  {FOOTER_HTML}
</div>
</body></html>"""


# ── salvar ──
out_dir = r"C:\KionDental\saidas\preview_emails"
os.makedirs(out_dir, exist_ok=True)

for cod in ["IS1","IS2","IS3","T1","T2","T3","T4","T5","T6"]:
    info = cfg["territorios"][cod]
    df_t = df_ativos[df_ativos["vendas"] == cod]
    if df_t.empty:
        continue
    path = os.path.join(out_dir, f"{cod}_{info['nome']}.html")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(email_territorio(info["nome"], cod, df_t))
    n_risco = (df_t["risco"].isin(["ALTO","MEDIO"])).sum()
    print(f"OK  {cod}_{info['nome']}.html  "
          f"({len(df_t)} clientes | {n_risco} em risco)")

path_g = os.path.join(out_dir, "GESTOR_Bruno.html")
with open(path_g, "w", encoding="utf-8") as fh:
    fh.write(email_gestor(df_ativos))
print(f"OK  GESTOR_Bruno.html  (consolidado)")

print(f"\nArquivos salvos em: {out_dir}")
