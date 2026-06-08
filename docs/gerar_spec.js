const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  TableOfContents
} = require('docx');
const fs = require('fs');

// ── Cores da marca ──────────────────────────────────────────────────────────
const AZUL   = '00B1D2';
const CINZA  = '282828';
const BRANCO = 'FFFFFF';
const CINZA2 = '5A5A5A';
const CINZA3 = 'F2F2F2';

// ── Bordas padrão de tabela ─────────────────────────────────────────────────
const borda = (cor = 'CCCCCC') => ({ style: BorderStyle.SINGLE, size: 1, color: cor });
const bordas = (cor = 'CCCCCC') => ({ top: borda(cor), bottom: borda(cor), left: borda(cor), right: borda(cor) });

// ── Helpers ─────────────────────────────────────────────────────────────────
const p = (text, opts = {}) => new Paragraph({
  children: [new TextRun({ text, font: 'Arial', size: opts.size || 22, bold: opts.bold || false,
    color: opts.color || '000000', italics: opts.italic || false })],
  spacing: { before: opts.before || 0, after: opts.after || 120 },
  alignment: opts.align || AlignmentType.LEFT,
});

const pBullet = (text, ref = 'bullets') => new Paragraph({
  numbering: { reference: ref, level: 0 },
  children: [new TextRun({ text, font: 'Arial', size: 20 })],
  spacing: { before: 40, after: 40 },
});

const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun({ text, font: 'Arial', size: 30, bold: true, color: CINZA })],
  spacing: { before: 360, after: 180 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: AZUL, space: 4 } },
});

const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  children: [new TextRun({ text, font: 'Arial', size: 24, bold: true, color: AZUL })],
  spacing: { before: 240, after: 120 },
});

const espaco = () => new Paragraph({ children: [new TextRun('')], spacing: { before: 0, after: 80 } });

// ── Células de tabela ────────────────────────────────────────────────────────
const celulaHeader = (text, width) => new TableCell({
  width: { size: width, type: WidthType.DXA },
  borders: bordas(AZUL),
  shading: { fill: AZUL, type: ShadingType.CLEAR },
  margins: { top: 100, bottom: 100, left: 120, right: 120 },
  verticalAlign: VerticalAlign.CENTER,
  children: [new Paragraph({
    children: [new TextRun({ text, font: 'Arial', size: 18, bold: true, color: BRANCO })],
    alignment: AlignmentType.LEFT,
    spacing: { before: 0, after: 0 },
  })],
});

const celula = (text, width, shade = false, bold = false) => new TableCell({
  width: { size: width, type: WidthType.DXA },
  borders: bordas('DDDDDD'),
  shading: { fill: shade ? 'F5FCFE' : BRANCO, type: ShadingType.CLEAR },
  margins: { top: 80, bottom: 80, left: 120, right: 120 },
  verticalAlign: VerticalAlign.CENTER,
  children: [new Paragraph({
    children: [new TextRun({ text, font: 'Arial', size: 18, bold, color: bold ? CINZA : '333333' })],
    spacing: { before: 0, after: 0 },
  })],
});

const linhaTabela = (colunas, widths, shade = false) => new TableRow({
  children: colunas.map((c, i) => celula(c, widths[i], shade)),
});

const headerTabela = (colunas, widths) => new TableRow({
  tableHeader: true,
  children: colunas.map((c, i) => celulaHeader(c, widths[i])),
});

// ═══════════════════════════════════════════════════════════════════════════
//  CAPA
// ═══════════════════════════════════════════════════════════════════════════
const capa = [
  espaco(), espaco(), espaco(), espaco(),
  new Paragraph({
    children: [new TextRun({ text: 'KION CS HUB', font: 'Arial', size: 64, bold: true, color: AZUL })],
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 200 },
  }),
  new Paragraph({
    children: [new TextRun({ text: 'Especificação de Produto — v1.0', font: 'Arial', size: 30, color: CINZA2 })],
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 600 },
  }),
  new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: AZUL, space: 1 } },
    children: [new TextRun('')], spacing: { before: 0, after: 600 },
  }),
  new Paragraph({
    children: [new TextRun({ text: 'Kion Dental Technology', font: 'Arial', size: 24, bold: true, color: CINZA })],
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 120 },
  }),
  new Paragraph({
    children: [new TextRun({ text: 'Time de Tecnologia e Inovação', font: 'Arial', size: 22, color: CINZA2 })],
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 120 },
  }),
  new Paragraph({
    children: [new TextRun({ text: 'Maio 2026', font: 'Arial', size: 22, color: CINZA2 })],
    alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 },
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ═══════════════════════════════════════════════════════════════════════════
//  SUMÁRIO
// ═══════════════════════════════════════════════════════════════════════════
const sumario = [
  new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text: 'Sumário', font: 'Arial', size: 30, bold: true, color: CINZA })],
    spacing: { before: 0, after: 180 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: AZUL, space: 4 } },
  }),
  new TableOfContents('Sumário', { hyperlink: true, headingStyleRange: '1-2' }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ═══════════════════════════════════════════════════════════════════════════
//  SEÇÃO 1 — VISÃO GERAL
// ═══════════════════════════════════════════════════════════════════════════
const sec1 = [
  h1('1. Visão Geral'),
  h2('1.1 Objetivo'),
  p('O Kion CS Hub é um sistema web de Customer Success desenvolvido para fechar o ciclo do sistema de Alerta de Faturamento já existente. Enquanto o alerta identifica e notifica clientes em risco, o CS Hub permite que os consultores registrem suas atuações, que o gestor acompanhe se os clientes estão sendo trabalhados dentro do prazo, e que o diretor visualize KPIs consolidados de recuperação e performance do time.', { after: 160 }),

  h2('1.2 Problema que resolve'),
  p('Hoje as atuações dos consultores ficam na memória ou em planilhas isoladas. Não existe histórico centralizado, não é possível saber se um cliente em ALTO risco foi contatado, e o gestor não tem visibilidade sobre a atividade do time em tempo real.', { after: 160 }),

  h2('1.3 Proposta de valor'),
  pBullet('Consultor: "Sei exatamente quem ligar hoje e tenho contexto completo antes de cada ligação"'),
  pBullet('Gestor: "Vejo em tempo real quem está sendo negligenciado e quem está performando"'),
  pBullet('Diretor: "Acompanho a taxa de recuperação e a saúde da carteira com dados reais"'),
  espaco(),
];

// ═══════════════════════════════════════════════════════════════════════════
//  SEÇÃO 2 — ARQUITETURA
// ═══════════════════════════════════════════════════════════════════════════
const sec2 = [
  h1('2. Arquitetura Técnica'),
  h2('2.1 Stack Tecnológico'),
  pBullet('Frontend: Aplicação web responsiva (mobile-first)'),
  pBullet('Backend: Python'),
  pBullet('Banco de dados: MySQL'),
  pBullet('Autenticação: Microsoft Entra ID (OAuth2 — SSO com conta corporativa)'),
  pBullet('Hospedagem: Servidor Linux próprio da Kion (mesmo servidor do alerta)'),
  pBullet('Containerização: Docker'),
  espaco(),

  h2('2.2 Fluxo de Dados'),
  new Paragraph({
    children: [new TextRun({ text: 'Excel (ERP)  →  Script de Alerta (Python, 10h diariamente)  →  MySQL  ←  Web App CS Hub', font: 'Courier New', size: 18, color: CINZA2 })],
    shading: { fill: CINZA3, type: ShadingType.CLEAR },
    border: { left: { style: BorderStyle.SINGLE, size: 12, color: AZUL, space: 4 } },
    indent: { left: 360 }, spacing: { before: 120, after: 120 },
  }),
  p('Os consultores registram interações diretamente no Web App, que também armazena no MySQL.', { after: 160 }),

  h2('2.3 Integração com o Sistema de Alerta Existente'),
  p('O script alerta_faturamento.py (já em produção) passa a ter uma segunda função além de enviar e-mails: gravar no MySQL um snapshot diário de risco por cliente. Isso alimenta o CS Hub com dados sempre atualizados sem necessidade de o CRM ler os arquivos Excel diretamente.', { after: 80 }),
  espaco(),
];

// ═══════════════════════════════════════════════════════════════════════════
//  SEÇÃO 3 — PERFIS
// ═══════════════════════════════════════════════════════════════════════════
const w3 = [2200, 2800, 2500, 1860];
const sec3 = [
  h1('3. Perfis de Usuário e Permissões'),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: w3,
    rows: [
      headerTabela(['Perfil', 'Acesso', 'Ação principal', 'Frequência'], w3),
      linhaTabela(['Consultor', 'Somente sua carteira (território)', 'Registrar contatos, atualizar status do cliente', 'Diária'], w3, false),
      linhaTabela(['Gestor', 'Todos os territórios', 'Monitorar SLAs, acompanhar time', 'Semanal'], w3, true),
      linhaTabela(['Diretor', 'Visão consolidada (somente leitura)', 'Acompanhar KPIs e dashboards', 'Semanal/Mensal'], w3, false),
      linhaTabela(['Admin', 'Configurações do sistema', 'Cadastrar usuários, ajustar territórios e SLAs', 'Eventual'], w3, true),
    ],
  }),
  espaco(),
];

// ═══════════════════════════════════════════════════════════════════════════
//  SEÇÃO 4 — ESTADOS
// ═══════════════════════════════════════════════════════════════════════════
const sec4 = [
  h1('4. Estados de Saúde do Cliente'),
  h2('4.1 Estados Automáticos (baseados em dados do ERP)'),
  pBullet('ESTÁVEL — Faturamento dentro da normalidade histórica'),
  pBullet('ATENÇÃO — Queda leve vs. mediana 12M'),
  pBullet('MÉDIO — Queda moderada vs. mediana 12M'),
  pBullet('ALTO — Queda significativa + meses consecutivos em declínio'),
  pBullet('POSSÍVEL CHURN — Faturamento = R$0 por 1 mês'),
  pBullet('CHURN PROVÁVEL — Faturamento = R$0 por 2 ou mais meses consecutivos'),
  pBullet('RECUPERADO — Faturamento voltou acima de 70% da mediana histórica'),
  espaco(),

  h2('4.2 Estados Manuais (definidos pelo consultor)'),
  pBullet('EM TRATATIVA — Consultor está ativamente trabalhando a recuperação'),
  pBullet('PAUSADO — Ausência temporária esperada (ex: clínica em reforma, sazonalidade)'),
  pBullet('CHURN CONFIRMADO — Cliente comunicou formalmente o cancelamento'),
  espaco(),

  h2('4.3 Regras de Transição'),
  p('Automáticas: calculadas diariamente pelo script de alerta com base nos dados do ERP.', { after: 80 }),
  p('Manuais: alteradas pelo consultor responsável pelo território, com registro de data e justificativa.', { after: 80 }),
  p('Um cliente marcado como PAUSADO ou EM TRATATIVA não recebe alertas de SLA enquanto o status estiver ativo.', { after: 80 }),
  espaco(),
];

// ═══════════════════════════════════════════════════════════════════════════
//  SEÇÃO 5 — SCORE
// ═══════════════════════════════════════════════════════════════════════════
const w5a = [2400, 2100, 1100, 3760];
const w5b = [2000, 3000, 2160];
const sec5 = [
  h1('5. Score de Risco Kion (0–100)'),

  h2('5.1 Composição do Score'),
  p('O Score de Risco Kion é um número de 0 a 100 que representa o risco composto de churn de cada cliente. Quanto maior o score, maior o risco.', { after: 120 }),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: w5a,
    rows: [
      headerTabela(['Fator', 'Fonte', 'Peso', 'Lógica de cálculo'], w5a),
      linhaTabela(['Queda vs. mediana 12M', 'ERP — automático', '35%', '0% de queda = 0 pts; 100% de queda = 35 pts'], w5a, false),
      linhaTabela(['Meses consecutivos em queda', 'ERP — automático', '25%', '1 mês = baixo; 6+ meses = máximo (25 pts)'], w5a, true),
      linhaTabela(['Meses sem faturamento (zeros)', 'ERP — automático', '20%', '0 meses = 0 pts; 3+ meses = 20 pts'], w5a, false),
      linhaTabela(['Dias desde última compra', 'ERP — automático', '10%', '0–30 dias = baixo; 90+ dias = máximo (10 pts)'], w5a, true),
      linhaTabela(['Satisfação declarada', 'Consultor — manual', '10%', 'Nota 3 = 0 pts; Nota 1 = 10 pts'], w5a, false),
    ],
  }),
  espaco(),

  h2('5.2 Faixas de Classificação'),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: w5b,
    rows: [
      headerTabela(['Score', 'Classificação', 'Indicação visual'], w5b),
      linhaTabela(['0–25', 'Saudável', '🟢 Verde'], w5b, false),
      linhaTabela(['26–50', 'Em observação', '🟡 Amarelo'], w5b, true),
      linhaTabela(['51–75', 'Em risco', '🟠 Laranja'], w5b, false),
      linhaTabela(['76–100', 'Risco crítico', '🔴 Vermelho'], w5b, true),
    ],
  }),
  espaco(),

  h2('5.3 Auditabilidade do Score'),
  p('O consultor e o gestor podem expandir o score de qualquer cliente para ver a contribuição individual de cada fator. Isso torna o score explicável e confiável — não é uma caixa preta.', { after: 160 }),

  h2('5.4 Detector de Inconsistência'),
  p('O sistema exibe um alerta quando a satisfação declarada contradiz sistematicamente o comportamento financeiro. Critério: cliente com satisfação = 3 (feliz) E score de dados (excluindo satisfação) acima de 65. O alerta é visível apenas para o gestor e o admin — não para o consultor responsável.', { after: 80 }),
  espaco(),
];

// ═══════════════════════════════════════════════════════════════════════════
//  SEÇÃO 6 — SLA
// ═══════════════════════════════════════════════════════════════════════════
const w6 = [2200, 1800, 5360];
const sec6 = [
  h1('6. SLA de Atendimento'),
  h2('6.1 Prazos por Nível de Risco'),
  p('Os prazos abaixo são os valores padrão. Todos são configuráveis pelo Admin na tela de Configurações sem necessidade de alteração de código.', { after: 120 }),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: w6,
    rows: [
      headerTabela(['Nível de risco', 'Prazo padrão', 'O que acontece ao estourar'], w6),
      linhaTabela(['ALTO', '3 dias', 'Cliente entra em destaque vermelho no painel do gestor'], w6, false),
      linhaTabela(['MÉDIO', '7 dias', 'Cliente entra em destaque laranja no painel do gestor'], w6, true),
      linhaTabela(['ATENÇÃO', '15 dias', 'Cliente entra em destaque amarelo no painel do gestor'], w6, false),
      linhaTabela(['POSSÍVEL CHURN', '2 dias', 'Alerta prioritário — aparece no topo da lista do consultor'], w6, true),
      linhaTabela(['CHURN PROVÁVEL', '1 dia', 'Notificação automática para gestor e diretor'], w6, false),
    ],
  }),
  espaco(),

  h2('6.2 Contagem de SLA'),
  p('O contador de SLA é reiniciado a cada nova interação registrada pelo consultor. Clientes com status PAUSADO ou EM TRATATIVA não contam SLA.', { after: 80 }),
  espaco(),
];

// ═══════════════════════════════════════════════════════════════════════════
//  SEÇÃO 7 — REGISTRO DE INTERAÇÃO
// ═══════════════════════════════════════════════════════════════════════════
const w7 = [3200, 4200, 1960];
const sec7 = [
  h1('7. Registro de Interação'),
  h2('7.1 Campos do Formulário'),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: w7,
    rows: [
      headerTabela(['Campo', 'Tipo / Opções', 'Obrigatório'], w7),
      linhaTabela(['Tipo de contato', 'Ligação / WhatsApp / E-mail / Visita presencial / Proposta enviada / Outro', 'Sim'], w7, false),
      linhaTabela(['Data e hora', 'Automático (editável)', 'Sim'], w7, true),
      linhaTabela(['Resumo da conversa', 'Texto livre (máx. 1.000 caracteres)', 'Sim'], w7, false),
      linhaTabela(['Felicidade do cliente', '1 = Insatisfeito / 2 = Neutro / 3 = Satisfeito', 'Sim'], w7, true),
      linhaTabela(['Próximo passo', 'Texto livre', 'Não'], w7, false),
      linhaTabela(['Data do próximo contato', 'Data (obrigatório se próximo passo preenchido)', 'Condicional'], w7, true),
      linhaTabela(['Alteração de status', 'Em tratativa / Pausado / Churn Confirmado', 'Não'], w7, false),
      linhaTabela(['Duração da conversa', 'Minutos (numérico)', 'Não'], w7, true),
    ],
  }),
  espaco(),

  h2('7.2 Lembretes de Follow-up'),
  p('Quando o consultor preenche a data do próximo contato, o sistema envia um e-mail de lembrete na manhã daquele dia via o mesmo serviço de e-mail do alerta (Microsoft Graph API).', { after: 80 }),
  espaco(),
];

// ═══════════════════════════════════════════════════════════════════════════
//  SEÇÃO 8 — TELAS PRINCIPAIS
// ═══════════════════════════════════════════════════════════════════════════
const sec8 = [
  h1('8. Telas Principais'),

  h2('8.1 Consultor — Minha Carteira'),
  p('Lista completa dos clientes do território, ranqueada por Score de Risco (maior primeiro). Para cada cliente exibe:', { after: 80 }),
  pBullet('Nome e tabela de preço (tier)'),
  pBullet('Badge de status atual (ALTO, MÉDIO, etc.)'),
  pBullet('Score de Risco (0–100) com cor correspondente'),
  pBullet('Dias desde o último contato (vermelho se SLA estourado)'),
  pBullet('Última nota de satisfação registrada'),
  pBullet('Botão de ação rápida: "Registrar contato"'),
  p('Filtros disponíveis: por status, por SLA estourado, por satisfação.', { before: 80, after: 160 }),

  h2('8.2 Consultor — Perfil do Cliente'),
  p('A tela mais completa do sistema, dividida em três blocos:', { after: 80 }),
  pBullet('Bloco 1 — Saúde financeira: gráfico de barras do faturamento mês a mês (12 meses), linha de referência da mediana histórica, Score de Risco com detalhamento por fator'),
  pBullet('Bloco 2 — Histórico de satisfação: gráfico de linha com a evolução da nota ao longo das interações, indicador de tendência (subindo/estável/caindo)'),
  pBullet('Bloco 3 — Timeline de interações: todas as interações em ordem cronológica reversa, mostrando data, consultor, tipo de contato, resumo e nota'),
  espaco(),

  h2('8.3 Gestor — Painel do Time'),
  pBullet('Cards por território: total de clientes, quantos em ALTO/MÉDIO, quantos com SLA estourado'),
  pBullet('Lista de clientes críticos: SLA estourado há mais tempo, independente de território'),
  pBullet('Tabela de atividade do time: contatos registrados por consultor na semana atual'),
  espaco(),

  h2('8.4 Diretor — Dashboard Executivo'),
  p('KPIs consolidados em cards:', { after: 80 }),
  pBullet('Total de clientes ativos'),
  pBullet('Clientes em risco (ALTO + MÉDIO) — variação vs. mês anterior'),
  pBullet('Taxa de recuperação do mês (% de clientes que saíram do ALTO)'),
  pBullet('Faturamento em risco (R$) — variação vs. mês anterior'),
  p('Gráficos:', { before: 80, after: 80 }),
  pBullet('Evolução da carteira de risco nos últimos 6 meses (área empilhada por nível)'),
  pBullet('Ranking de consultores por recuperações realizadas no mês'),
  pBullet('Distribuição dos churns confirmados por motivo'),
  p('Lista: Top 10 clientes de maior impacto sem contato nos últimos 7 dias.', { before: 80, after: 160 }),

  h2('8.5 Admin — Configurações'),
  pBullet('Cadastro e edição de usuários (nome, e-mail corporativo, perfil, território)'),
  pBullet('Ajuste dos SLAs por nível de risco (campo numérico editável)'),
  pBullet('Ajuste dos thresholds de risco (queda % para ALTO, MÉDIO, ATENÇÃO)'),
  pBullet('Pesos do Score de Risco (sliders de 0–100%, soma deve ser 100%)'),
  espaco(),
];

// ═══════════════════════════════════════════════════════════════════════════
//  SEÇÃO 9 — MODELO DE DADOS
// ═══════════════════════════════════════════════════════════════════════════
const entidade = (nome, campos) => [
  new Paragraph({
    children: [new TextRun({ text: nome, font: 'Arial', size: 20, bold: true, color: AZUL })],
    spacing: { before: 160, after: 60 },
  }),
  new Paragraph({
    children: [new TextRun({ text: campos, font: 'Courier New', size: 17, color: CINZA2 })],
    shading: { fill: CINZA3, type: ShadingType.CLEAR },
    indent: { left: 360 },
    spacing: { before: 60, after: 120 },
  }),
];

const sec9 = [
  h1('9. Modelo de Dados (Conceitual)'),
  h2('9.1 Entidades Principais'),
  ...entidade('CLIENTE', 'id, nome, territorio_id, tabela_preco, status_atual, score_atual,\ndata_ultima_compra, data_ultimo_contato, criado_em, atualizado_em'),
  ...entidade('TERRITORIO', 'id, codigo (IS1, T1...), nome_responsavel, usuario_id'),
  ...entidade('USUARIO', 'id, nome, email_corporativo, perfil (consultor/gestor/diretor/admin),\nterritorio_id, ativo, criado_em'),
  ...entidade('INTERACAO', 'id, cliente_id, usuario_id, tipo_contato, resumo,\nsatisfacao (1/2/3), duracao_minutos, proximo_passo,\ndata_proximo_contato, criado_em'),
  ...entidade('SNAPSHOT_RISCO', 'id, cliente_id, data, faturamento_mes, mediana_12m,\nvariacao_pct, meses_queda, score_risco, status_risco, criado_em'),
  ...entidade('CONFIGURACAO', 'chave (ex: sla_alto, peso_queda_pct), valor, descricao,\natualizado_em, atualizado_por'),
  espaco(),
];

// ═══════════════════════════════════════════════════════════════════════════
//  SEÇÃO 10 — ROADMAP
// ═══════════════════════════════════════════════════════════════════════════
const w10 = [2000, 4560, 2800];
const sec10 = [
  h1('10. Roadmap Sugerido'),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: w10,
    rows: [
      headerTabela(['Fase', 'Escopo', 'Objetivo'], w10),
      linhaTabela(['Fase 1 — MVP', 'Autenticação Entra ID, Minha Carteira, Perfil do Cliente, registro de interação, snapshot diário no MySQL', 'Consultor registra contatos e vê histórico'], w10, false),
      linhaTabela(['Fase 2 — Gestão', 'Painel do Gestor, SLA com alertas, detector de inconsistência, lembretes de follow-up por e-mail', 'Gestor tem visibilidade total'], w10, true),
      linhaTabela(['Fase 3 — Inteligência', 'Score de Risco composto, Dashboard do Diretor, histórico de satisfação, estados de churn automáticos', 'Diretor acompanha KPIs e time tem score auditável'], w10, false),
      linhaTabela(['Fase 4 — Admin', 'Tela de configurações, ajuste de pesos e SLAs pelo front, relatórios exportáveis em Excel', 'Sistema autossuficiente sem dependência de TI'], w10, true),
    ],
  }),
  espaco(),
];

// ═══════════════════════════════════════════════════════════════════════════
//  SEÇÃO 11 — GLOSSÁRIO
// ═══════════════════════════════════════════════════════════════════════════
const w11 = [3000, 6360];
const sec11 = [
  h1('11. Glossário'),
  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: w11,
    rows: [
      headerTabela(['Termo', 'Definição'], w11),
      linhaTabela(['Score de Risco Kion', 'Índice composto de 0–100 que representa o risco de churn, calculado com quatro fatores de dados do ERP e um fator relacional'], w11, false),
      linhaTabela(['Mediana 12M', 'Mediana dos meses com faturamento > zero nos últimos 12 meses. Usada como baseline histórico do cliente'], w11, true),
      linhaTabela(['SLA de Atendimento', 'Prazo máximo (em dias) para o consultor registrar um contato após o cliente entrar em determinado nível de risco'], w11, false),
      linhaTabela(['Snapshot de Risco', 'Registro diário do estado de risco de cada cliente, gravado pelo script de alerta no banco de dados MySQL'], w11, true),
      linhaTabela(['Churn Confirmado', 'Status manual atribuído pelo consultor quando o cliente comunica formalmente o encerramento da parceria'], w11, false),
      linhaTabela(['Detector de Inconsistência', 'Regra automática que sinaliza quando a satisfação declarada contradiz os dados financeiros do cliente'], w11, true),
      linhaTabela(['EM TRATATIVA', 'Status manual que indica que o consultor está trabalhando a recuperação — suspende a contagem de SLA'], w11, false),
      linhaTabela(['RECUPERADO', 'Status automático atingido quando o faturamento volta acima de 70% da mediana histórica após um período de risco'], w11, true),
      linhaTabela(['Taxa de Recuperação', 'Percentual de clientes que estavam em ALTO ou MÉDIO e migraram para ESTÁVEL ou RECUPERADO no período'], w11, false),
      linhaTabela(['Entra ID', 'Plataforma de identidade da Microsoft (antigo Azure AD) usada para autenticação dos usuários no sistema'], w11, true),
    ],
  }),
  espaco(),
];

// ═══════════════════════════════════════════════════════════════════════════
//  DOCUMENTO FINAL
// ═══════════════════════════════════════════════════════════════════════════
const doc = new Document({
  numbering: {
    config: [{
      reference: 'bullets',
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: '•',
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      }],
    }],
  },
  styles: {
    default: {
      document: { run: { font: 'Arial', size: 22 } },
    },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 30, bold: true, font: 'Arial', color: CINZA },
        paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 },
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: 'Arial', color: AZUL },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 },
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          children: [
            new TextRun({ text: 'Kion CS Hub — Especificação de Produto v1.0', font: 'Arial', size: 18, color: CINZA2 }),
          ],
          border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: AZUL, space: 4 } },
          alignment: AlignmentType.RIGHT,
          spacing: { before: 0, after: 80 },
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          children: [
            new TextRun({ text: 'Kion Dental Technology  |  Confidencial  |  Página ', font: 'Arial', size: 16, color: CINZA2 }),
            new TextRun({ children: [PageNumber.CURRENT], font: 'Arial', size: 16, color: CINZA2 }),
          ],
          border: { top: { style: BorderStyle.SINGLE, size: 2, color: AZUL, space: 4 } },
          alignment: AlignmentType.CENTER,
          spacing: { before: 80, after: 0 },
        })],
      }),
    },
    children: [
      ...capa,
      ...sumario,
      ...sec1,
      ...sec2,
      ...sec3,
      ...sec4,
      ...sec5,
      ...sec6,
      ...sec7,
      ...sec8,
      ...sec9,
      ...sec10,
      ...sec11,
    ],
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('C:\\KionDental\\docs\\KionCSHub_Especificacao_v1.0.docx', buffer);
  console.log('✅ Documento gerado com sucesso!');
}).catch(err => {
  console.error('❌ Erro:', err.message);
});
