/**
 * RITZ · Painel — ponte entre a planilha "controle-gerencial" e o painel no GitHub Pages.
 *
 * Instalação (uma vez):
 *   1. Na planilha: Extensões → Apps Script → colar este arquivo inteiro → Salvar.
 *   2. Configurações do projeto → Propriedades do script → adicionar SENHA = <sua senha>.
 *   3. Rodar a função  configurarPlanilha  (menu ▶ Executar) e autorizar.
 *      Ela corrige as fórmulas que usavam macro (Compra/Venda) e cria a aba METAS.
 *   4. Implantar → Nova implantação → App da Web
 *      Executar como: Eu · Quem pode acessar: Qualquer pessoa → copiar a URL /exec.
 *
 *   5. Rodar a função  instalarRotinaDiaria  e autorizar (e-mail + gatilho).
 *      Todo dia às 8h: grava uma linha na aba HISTÓRICO e envia o resumo por e-mail.
 *      Destinatário: a propriedade EMAIL_RESUMO (vários separados por vírgula) ou, sem ela, o seu e-mail.
 *
 * Ao mudar este código: Implantar → Gerenciar implantações → editar → Nova versão.
 *
 * GET <url>/exec?token=SENHA  →  { ok, gerado, produtos, vendas, compras, metas, evolucao, snapshots, ... }
 */

var ABA_CADASTRO = 'CADASTRO E CONTROLE';
var ABA_SAIDAS   = 'SAÍDAS';
var ABA_ENTRADAS = 'ENTRADAS';
var ABA_PRECOS   = 'PREÇOS';
var ABA_METAS    = 'METAS';
var ABA_HIST     = 'HISTÓRICO';
var FUSO         = 'America/Sao_Paulo';
var URL_PAINEL   = 'https://anavrodovalho.github.io/ritz-painel/';
var DIAS_PARADO  = 45;   // sem venda há X dias (ou, se nunca vendeu, desde a 1ª entrada)

var HIST_COLUNAS = ['data', 'valor_estoque', 'pecas_estoque', 'itens_abaixo_minimo', 'itens_parados',
  'receita_mes', 'lucro_mes', 'pedidos_mes', 'caixa_atual'];

// brindes (preço 0) das sócias ficam fora dos números de cliente/venda
var SOCIAS = ['ana laura rodovalho', 'rafaella rempel', 'rafaela rempel'];

var CORES = {MAR:'Marrom',PRE:'Preto',AZM:'Azul Marinho',BOR:'Bordô',VMI:'Verde Militar',BRA:'Branco',
  BEG:'Bege',VER:'Verde',AZU:'Azul',OFF:'Off White',CCH:'Chumbo',ROS:'Rosa',NUD:'Nude',CAR:'Caramelo',
  VIN:'Vinho',CIN:'Cinza',AMA:'Amarelo',LAR:'Laranja',LIL:'Lilás',OFW:'Off White'};

var METAS_PADRAO = [
  ['meta_receita_mensal', 10000, 'Receita do mês (R$)'],
  ['margem_alvo_pct',     55,    'Margem alvo (%)'],
  ['meta_pedidos_mensal', 50,    'Pedidos no mês'],
  ['meta_ticket',         200,   'Ticket médio (R$)'],
  ['caixa_minimo',        3000,  'Caixa mínimo (R$)'],
  ['caixa_atual',         4283.6,'Caixa atual (R$) — atualize quando quiser']
];

/* ======================= Web App ======================= */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var senha = PropertiesService.getScriptProperties().getProperty('SENHA');
  if (!senha) return json_({ ok: false, erro: 'senha_nao_configurada' });
  if (p.token !== senha) return json_({ ok: false, erro: 'senha_invalida' });
  try {
    return json_(montarDados_());
  } catch (err) {
    return json_({ ok: false, erro: 'falha_leitura', detalhe: String(err && err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ======================= Leitura ======================= */

function montarDados_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cad = tabela_(ss, ABA_CADASTRO), ent = tabela_(ss, ABA_ENTRADAS), sai = tabela_(ss, ABA_SAIDAS);
  var pre = null; try { pre = tabela_(ss, ABA_PRECOS); } catch (x) {}

  // compras (ENTRADAS) — também dão o custo: último preço de compra de cada SKU
  var compras = [], ultimoCusto = {};
  ent.linhas.forEach(function (r) {
    var sku = txt_(r[ent.col['Cod. Item']]); if (!sku) return;
    var qtd = num_(r[ent.col['Quantidade']]), custo = num_(r[ent.col['Preço Compra']]);
    if (custo > 0) ultimoCusto[sku] = custo;
    compras.push({ data: data_(r[ent.col['Data']]), sku: sku, nome: txt_(r[ent.col['Descrição do Item']]),
      fornecedor: txt_(r[ent.col['Nome do Fornecedor']]), qtd: qtd, custo_unitario: custo, custo_total: custo * qtd });
  });

  // vendas (SAÍDAS) — a coluna "Pedido" guarda o nome da cliente
  var vendas = [], brindes = 0, precosVendidos = {}, saidas = [];
  sai.linhas.forEach(function (r) {
    var sku = txt_(r[sai.col['Cod. Item']]), data = data_(r[sai.col['Data']]);
    if (!sku || !data) return;
    var qtd = num_(r[sai.col['Quantidade']]) || 1, preco = num_(r[sai.col['Preço Venda']]);
    saidas.push({ data: data, sku: sku, qtd: qtd });   // inclui brindes: também tiram peça do estoque
    var cliente = txt_(r[sai.col['Pedido']]);
    var receita = num_(r[sai.col['Total Vendido']]) || preco * qtd;
    if (receita === 0 && SOCIAS.indexOf(cliente.toLowerCase()) >= 0) { brindes++; return; }
    if (preco > 0) (precosVendidos[sku] = precosVendidos[sku] || []).push(preco);
    vendas.push({ data: data, num_pedido: data + '·' + (cliente || '—'), cliente: cliente, sku: sku,
      nome: txt_(r[sai.col['Descrição do Item']]), qtd: qtd, preco_vendido: preco, receita: receita,
      status: 'entregue', canal: '' });
  });

  // tabela de preços (aba PREÇOS)
  var precoTabela = {};
  if (pre) {
    var pc = pre.col['Preço Venda (R$/peça)']; if (pc == null) pc = pre.col['Preço Venda']; if (pc == null) pc = pre.col['Preço'];
    pre.linhas.forEach(function (r) { var sku = txt_(r[pre.col['Cod. Item']]), v = num_(r[pc]); if (sku && v > 0) precoTabela[sku] = v; });
  }

  // produtos (CADASTRO E CONTROLE)
  // SKU cadastrado em duas linhas: as duas mostram o mesmo saldo (SUMIF), então conta só a primeira
  var produtos = [], linhaDoSku = {}, duplicados = [];
  cad.linhas.forEach(function (r, i) {
    var sku = txt_(r[cad.col['Cod. Item']]); if (!sku) return;
    var linha = cad.inicio + i;
    if (linhaDoSku[sku]) { duplicados.push({ sku: sku, linhas: [linhaDoSku[sku], linha] }); return; }
    linhaDoSku[sku] = linha;
    var partes = sku.split('-'), tam = '', cor = '';
    var nome = txt_(r[cad.col['Descrição do Item']]);
    // código conhecido manda (agrupa variações de escrita); senão vale a cor escrita na descrição "… (Cor) (Tam)"
    var corDesc = (nome.match(/\(([^()]+)\)\s*\([^()]+\)\s*$/) || [])[1];
    if (partes.length >= 4) { cor = CORES[partes[2]] || corDesc || partes[2]; tam = partes[partes.length - 1]; }
    var custo = ultimoCusto[sku] || num_(r[cad.col['Compra']]);
    var preco = precoTabela[sku] || num_(r[cad.col['Venda']]);
    if (!preco && precosVendidos[sku]) {
      var a = precosVendidos[sku]; preco = Math.round(a.reduce(function (s, x) { return s + x; }, 0) / a.length * 100) / 100;
    }
    produtos.push({ sku: sku, nome: nome, categoria: txt_(r[cad.col['Grupo']]) || '—',
      fornecedor: txt_(r[cad.col['Nome do Fornecedor']]) || '—', tamanho: tam, cor: cor,
      custo_unitario: custo, preco_venda: preco, estoque_atual: num_(r[cad.col['Saldo']]),
      estoque_minimo: num_(r[cad.col['Min.']]), pronto_entrega: true });
  });

  return { ok: true, gerado: new Date().toISOString(), produtos: produtos, vendas: vendas, compras: compras,
    pedidos: [], leads: [], campanhas: [], metas: lerMetas_(ss), snapshots: lerHistorico_(ss),
    evolucao: evolucaoEstoque_(produtos, compras, saidas), brindes_excluidos: brindes, duplicados: duplicados };
}

/**
 * Estoque dia a dia desde o 1º movimento, reconstruído de trás pra frente a partir do saldo de hoje
 * (saldo de ontem = hoje − entradas de hoje + saídas de hoje). Valor a custo atual de cada SKU.
 */
function evolucaoEstoque_(produtos, compras, saidas) {
  var custo = {}, pecas = 0, valor = 0;
  produtos.forEach(function (p) { custo[p.sku] = p.custo_unitario; pecas += p.estoque_atual; valor += p.estoque_atual * p.custo_unitario; });
  var mov = {};  // data → {p: Δpeças, v: Δvalor}
  function add(d, sku, q) { if (!d) return; var m = mov[d] = mov[d] || { p: 0, v: 0 }; m.p += q; m.v += q * (custo[sku] || 0); }
  compras.forEach(function (c) { add(c.data, c.sku, c.qtd); });
  saidas.forEach(function (s) { add(s.data, s.sku, -s.qtd); });
  var datas = Object.keys(mov).sort(); if (!datas.length) return [];
  var hoje = Utilities.formatDate(new Date(), FUSO, 'yyyy-MM-dd'), out = [];
  var d = new Date(hoje + 'T12:00:00Z'), ini = new Date(datas[0] + 'T12:00:00Z');
  while (d >= ini) {
    var k = d.toISOString().slice(0, 10);
    out.push({ data: k, pecas: Math.round(pecas), valor: Math.round(valor * 100) / 100 });
    if (mov[k]) { pecas -= mov[k].p; valor -= mov[k].v; }
    d.setUTCDate(d.getUTCDate() - 1);
  }
  return out.reverse();
}

function lerHistorico_(ss) {
  var sh = ss.getSheetByName(ABA_HIST); if (!sh || sh.getLastRow() < 2) return [];
  var v = sh.getDataRange().getValues(), h = v[0].map(function (c) { return txt_(c); });
  return v.slice(1).filter(function (r) { return r[0] !== ''; }).map(function (r) {
    var o = {}; h.forEach(function (k, j) { o[k] = k === 'data' ? data_(r[j]) : num_(r[j]); }); return o;
  });
}

/** Lê uma aba achando a linha de cabeçalho pela célula "Cod. Item" (nunca pela posição). */
function tabela_(ss, nome) {
  var sh = ss.getSheetByName(nome);
  if (!sh) throw new Error('Aba não encontrada: ' + nome);
  var v = sh.getDataRange().getValues(), h = -1;
  for (var i = 0; i < Math.min(v.length, 15) && h < 0; i++) {
    if (v[i].some(function (c) { return String(c).trim() === 'Cod. Item'; })) h = i;
  }
  if (h < 0) throw new Error('Cabeçalho "Cod. Item" não encontrado na aba ' + nome);
  var col = {};
  v[h].forEach(function (c, j) { var k = String(c).trim(); if (k && !(k in col)) col[k] = j; });
  return { col: col, linhas: v.slice(h + 1), inicio: h + 2 };   // inicio = nº da linha na planilha do 1º dado
}

function lerMetas_(ss) {
  var sh = ss.getSheetByName(ABA_METAS); if (!sh) return [];
  var m = {};
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    var k = txt_(r[0]); if (k && r[1] !== '' && r[1] != null) m[k] = num_(r[1]);
  });
  return Object.keys(m).length ? [m] : [];
}

function txt_(v) { return v == null ? '' : String(v).trim(); }
function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  if (v == null || v === '') return 0;
  var s = String(v).replace(/R\$\s?/, '').trim();
  if (/,\d{1,2}$/.test(s)) s = s.replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s); return isNaN(n) ? 0 : n;
}
/** Data → 'yyyy-mm-dd'. Aceita Date, número de série do Excel ou texto dd/mm/aaaa. */
function data_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, FUSO, 'yyyy-MM-dd');
  if (typeof v === 'number' && v > 20000) return Utilities.formatDate(new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5), 'UTC', 'yyyy-MM-dd');
  var s = String(v || '').trim(), iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[0];
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) { var y = m[3].length === 2 ? '20' + m[3] : m[3]; return y + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2); }
  return null;
}

/* ======================= Preparação da planilha ======================= */

/**
 * Rodar uma vez depois de converter o .xlsm.
 * - Compra/Venda/Lucro/Valor do Saldo usavam a macro VBA_SELECT (não existe no Google) → #NAME? e #DIV/0!.
 *   Troca por fórmulas nativas: Compra = último preço de compra (ENTRADAS);
 *   Venda = preço da aba PREÇOS, ou o último preço vendido (SAÍDAS) se o item não estiver lá.
 * - Limpa a sobra de VBA_SELECT fora da tabela em ENTRADAS (coluna sem cabeçalho).
 * - Cria a aba METAS, se não existir.
 */
function configurarPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ABA_CADASTRO), v = sh.getDataRange().getValues(), h = -1;
  for (var i = 0; i < 15 && h < 0; i++) if (v[i].some(function (c) { return String(c).trim() === 'Cod. Item'; })) h = i;
  if (h < 0) throw new Error('Cabeçalho não encontrado em ' + ABA_CADASTRO);
  var col = {}; v[h].forEach(function (c, j) { col[String(c).trim()] = j + 1; });
  var L = function (nome) { return colLetra_(col[nome]); };
  var B = L('Cod. Item'), N = L('Compra'), O = L('Venda'), P = L('Lucro'), Q = L('Valor do Saldo'), S = L('Saldo');

  var ent = ss.getSheetByName(ABA_ENTRADAS), sai = ss.getSheetByName(ABA_SAIDAS), pre = ss.getSheetByName(ABA_PRECOS);
  var e = cols_(ent), s = cols_(sai), p = pre ? cols_(pre) : null;
  var rE = function (c) { return "ENTRADAS!$" + e.l[c] + "$" + (e.h + 2) + ":$" + e.l[c]; };
  var rS = function (c) { return "'SAÍDAS'!$" + s.l[c] + "$" + (s.h + 2) + ":$" + s.l[c]; };
  var rP = function (c) { return "'PREÇOS'!$" + p.l[c] + "$" + (p.h + 2) + ":$" + p.l[c]; };
  var precoCol = p && (p.l['Preço Venda (R$/peça)'] ? 'Preço Venda (R$/peça)' : 'Preço Venda');

  var ini = h + 2, fim = sh.getMaxRows(), f = [];
  for (var r = ini; r <= fim; r++) {
    var k = '$' + B + r;
    var ultVenda = 'XLOOKUP(' + k + ',' + rS('Cod. Item') + ',' + rS('Preço Venda') + ',0,0,-1)';
    f.push([
      '=IF(' + k + '="","",IFERROR(XLOOKUP(' + k + ',' + rE('Cod. Item') + ',' + rE('Preço Compra') + ',"",0,-1),""))',
      '=IF(' + k + '="","",IFERROR(' + (p ? 'XLOOKUP(' + k + ',' + rP('Cod. Item') + ',' + rP(precoCol) + ',' + ultVenda + ')' : ultVenda) + ',0))',
      '=IF(AND(ISNUMBER($' + N + r + '),ISNUMBER($' + O + r + '),N($' + N + r + ')>0,N($' + O + r + ')>0),$' + O + r + '/$' + N + r + '-1,"")',
      '=IF(AND(ISNUMBER($' + O + r + '),ISNUMBER($' + S + r + ')),$' + S + r + '*$' + O + r + ',"")'
    ]);
  }
  if (col['Venda'] !== col['Compra'] + 1 || col['Lucro'] !== col['Compra'] + 2 || col['Valor do Saldo'] !== col['Compra'] + 3)
    throw new Error('Colunas Compra/Venda/Lucro/Valor do Saldo não estão lado a lado — nada foi alterado.');
  sh.getRange(ini, col['Compra'], f.length, 4).setFormulas(f);

  // sobra da macro em ENTRADAS: células com VBA_SELECT fora das colunas com cabeçalho
  var ef = ent.getDataRange().getFormulas(), limpas = 0;
  ef.forEach(function (row, i) { row.forEach(function (fx, j) {
    if (fx.indexOf('VBA_SELECT') >= 0) { ent.getRange(i + 1, j + 1).clearContent(); limpas++; }
  }); });

  if (!ss.getSheetByName(ABA_METAS)) {
    var m = ss.insertSheet(ABA_METAS);
    m.getRange(1, 1, 1, 3).setValues([['chave', 'valor', 'o que é']]).setFontWeight('bold');
    m.getRange(2, 1, METAS_PADRAO.length, 3).setValues(METAS_PADRAO);
    m.autoResizeColumns(1, 3);
  }
  SpreadsheetApp.flush();
  Logger.log('Fórmulas corrigidas em %s linhas; %s células de macro limpas em ENTRADAS.', f.length, limpas);
}

function cols_(sh) {
  var v = sh.getRange(1, 1, 15, sh.getLastColumn()).getValues(), h = -1;
  for (var i = 0; i < v.length && h < 0; i++) if (v[i].some(function (c) { return String(c).trim() === 'Cod. Item'; })) h = i;
  if (h < 0) throw new Error('Cabeçalho não encontrado em ' + sh.getName());
  var l = {}; v[h].forEach(function (c, j) { var k = String(c).trim(); if (k && !(k in l)) l[k] = colLetra_(j + 1); });
  return { h: h, l: l };
}
function colLetra_(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; }

/** Teste rápido no editor: mostra no log quantos registros o painel vai receber. */
function testarLeitura() {
  var d = montarDados_();
  Logger.log('produtos %s · vendas %s · compras %s · metas %s · histórico %s · evolução %s dias · brindes excluídos %s',
    d.produtos.length, d.vendas.length, d.compras.length, d.metas.length, d.snapshots.length, d.evolucao.length, d.brindes_excluidos);
}

/* ======================= Rotina diária: histórico + e-mail ======================= */

/** Rodar uma vez: cria o gatilho diário das 8h e já grava a primeira linha do histórico. */
function instalarRotinaDiaria() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'rotinaDiaria') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('rotinaDiaria').timeBased().everyDays(1).atHour(8).inTimezone(FUSO).create();
  registrarHistorico();
  Logger.log('Rotina diária instalada (8h). E-mail vai para: %s', destinatarios_());
}

/** Para desligar o e-mail e o histórico automáticos. */
function desinstalarRotinaDiaria() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'rotinaDiaria') ScriptApp.deleteTrigger(t);
  });
}

function rotinaDiaria() {
  var d = montarDados_();
  registrarHistorico(d);
  enviarResumoDiario(d);
}

/** Grava (ou atualiza) a linha de hoje na aba HISTÓRICO. */
function registrarHistorico(d) {
  d = d || montarDados_();
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(ABA_HIST);
  if (!sh) {
    sh = ss.insertSheet(ABA_HIST);
    sh.getRange(1, 1, 1, HIST_COLUNAS.length).setValues([HIST_COLUNAS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  var r = resumo_(d), hoje = r.hoje;
  var linha = [hoje, r.valorEstoque, r.pecas, r.abaixo.length, r.parados.length,
    r.mes.receita, r.mes.lucro, r.mes.pedidos, r.metas.caixa_atual != null ? r.metas.caixa_atual : ''];
  var datas = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (x) { return data_(x[0]); }) : [];
  var i = datas.indexOf(hoje);
  var alvo = i >= 0 ? i + 2 : sh.getLastRow() + 1;
  sh.getRange(alvo, 1, 1, linha.length).setValues([linha]);
  sh.getRange(alvo, 1).setNumberFormat('@');   // data como texto aaaa-mm-dd, sem conversão de fuso
}

/** Resumo por e-mail: vendas de ontem, mês x metas, reposição e parados. */
function enviarResumoDiario(d) {
  d = d || montarDados_();
  var r = resumo_(d), m = r.metas;
  var brl = function (v) { return 'R$ ' + Number(v || 0).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.'); };
  var pct = function (v) { return Math.round(v) + '%'; };
  var linha = function (rot, val, extra) {
    return '<tr><td style="padding:6px 0;color:#555">' + rot + '</td><td style="padding:6px 0;text-align:right;font-weight:600">' + val +
      (extra ? '<div style="font-weight:400;color:#888;font-size:12px">' + extra + '</div>' : '') + '</td></tr>';
  };
  var meta = function (atual, alvo, fmt) { return alvo > 0 ? 'meta ' + fmt(alvo) + ' · ' + pct(atual / alvo * 100) : ''; };
  var esc = function (s) { return String(s).replace(/[&<>]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]; }); };

  var ontemTxt = r.ontem.itens.length
    ? r.ontem.itens.map(function (v) { return '• ' + esc(v.nome) + ' — ' + esc(v.cliente || '—') + ' — ' + brl(v.receita); }).join('<br>')
    : 'Nenhuma venda lançada ontem.';
  var lista = function (arr, fmt) {
    if (!arr.length) return '<p style="color:#888;margin:4px 0;font-size:13px">Nada por aqui ✓</p>';
    var mais = arr.length > 12 ? '<br>… e mais ' + (arr.length - 12) : '';
    return '<p style="margin:4px 0;line-height:1.6;font-size:13px">' + arr.slice(0, 12).map(fmt).join('<br>') + mais + '</p>';
  };

  var html = '<div style="font-family:Arial,Helvetica,sans-serif;max-width:560px;color:#111">' +
    '<div style="letter-spacing:.4em;font-size:22px;font-weight:300;margin:8px 0 2px">RITZ</div>' +
    '<div style="color:#888;font-size:12px;margin-bottom:18px">Resumo de ' + r.hojeBR + '</div>' +
    '<h3 style="font-size:15px;margin:16px 0 6px">Ontem</h3>' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
      linha('Receita', brl(r.ontem.receita)) + linha('Pedidos', r.ontem.pedidos) + '</table>' +
    '<p style="font-size:13px;line-height:1.6;color:#333">' + ontemTxt + '</p>' +
    '<h3 style="font-size:15px;margin:18px 0 6px">Mês até agora</h3>' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px">' +
      linha('Receita', brl(r.mes.receita), meta(r.mes.receita, m.meta_receita_mensal, brl)) +
      linha('Lucro', brl(r.mes.lucro)) +
      linha('Margem', pct(r.mes.margem), m.margem_alvo_pct > 0 ? 'meta ' + pct(m.margem_alvo_pct) : '') +
      linha('Pedidos', r.mes.pedidos, meta(r.mes.pedidos, m.meta_pedidos_mensal, String)) +
      (m.caixa_atual != null ? linha('Caixa informado', brl(m.caixa_atual), m.caixa_minimo > 0 ? 'mínimo ' + brl(m.caixa_minimo) : '') : '') +
    '</table>' +
    (d.duplicados.length ? '<p style="background:#fdecea;border-radius:8px;padding:10px 12px;font-size:13px;color:#8a1c14;line-height:1.6">' +
      '<b>Cadastro duplicado (' + d.duplicados.length + ')</b> — cada SKU deve aparecer uma vez em ' + ABA_CADASTRO + ':<br>' +
      d.duplicados.map(function (x) { return '• ' + x.sku + ' nas linhas ' + x.linhas.join(' e '); }).join('<br>') + '</p>' : '') +
    '<h3 style="font-size:15px;margin:18px 0 6px">Repor (' + r.abaixo.length + ' abaixo do mínimo)</h3>' +
    lista(r.abaixo, function (p) { return '• ' + esc(p.nome) + ' — tem ' + p.estoque_atual + ', mínimo ' + p.estoque_minimo; }) +
    '<h3 style="font-size:15px;margin:18px 0 6px">Parados há ' + DIAS_PARADO + '+ dias (' + r.parados.length + ' itens · ' + brl(r.valorParado) + ' a custo)</h3>' +
    lista(r.parados, function (p) { return '• ' + esc(p.nome) + ' — ' + p.dias + ' dias, ' + p.estoque_atual + ' peça(s)'; }) +
    '<p style="margin:22px 0"><a href="' + URL_PAINEL + '" style="background:#111;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none;font-size:14px">Abrir o painel</a></p>' +
    '<p style="color:#aaa;font-size:11px">Estoque: ' + r.pecas + ' peças · ' + brl(r.valorEstoque) + ' a custo. Enviado automaticamente pelo Apps Script da planilha.</p></div>';

  MailApp.sendEmail({ to: destinatarios_(), subject: 'RITZ · resumo de ' + r.hojeBR + ' — ontem ' + brl(r.ontem.receita), htmlBody: html });
}

function destinatarios_() {
  return PropertiesService.getScriptProperties().getProperty('EMAIL_RESUMO') || Session.getEffectiveUser().getEmail();
}

/** Números do dia, compartilhados entre o e-mail e o histórico. */
function resumo_(d) {
  var agora = new Date(), hoje = Utilities.formatDate(agora, FUSO, 'yyyy-MM-dd');
  var ontem = Utilities.formatDate(new Date(agora.getTime() - 864e5), FUSO, 'yyyy-MM-dd');
  var ini = hoje.slice(0, 8) + '01';
  var custo = {}; d.produtos.forEach(function (p) { custo[p.sku] = p.custo_unitario; });
  var soma = function (vs) {
    var rec = 0, cst = 0, ped = {};
    vs.forEach(function (v) { rec += v.receita; cst += v.qtd * (custo[v.sku] || 0); ped[v.num_pedido] = 1; });
    return { receita: rec, lucro: rec - cst, margem: rec > 0 ? (rec - cst) / rec * 100 : 0, pedidos: Object.keys(ped).length };
  };
  var vOntem = d.vendas.filter(function (v) { return v.data === ontem; });
  var o = soma(vOntem); o.itens = vOntem;

  // parado: dias desde a última venda; se nunca vendeu, desde a 1ª entrada
  var ultVenda = {}, primEntrada = {};
  d.vendas.forEach(function (v) { if (!ultVenda[v.sku] || v.data > ultVenda[v.sku]) ultVenda[v.sku] = v.data; });
  d.compras.forEach(function (c) { if (c.data && (!primEntrada[c.sku] || c.data < primEntrada[c.sku])) primEntrada[c.sku] = c.data; });
  var dias = function (iso) { return iso ? Math.floor((new Date(hoje) - new Date(iso)) / 864e5) : null; };
  var parados = [], valorParado = 0;
  d.produtos.forEach(function (p) {
    if (p.estoque_atual <= 0) return;
    var n = dias(ultVenda[p.sku] || primEntrada[p.sku]);
    if (n != null && n >= DIAS_PARADO) { parados.push({ nome: p.nome, dias: n, estoque_atual: p.estoque_atual }); valorParado += p.estoque_atual * p.custo_unitario; }
  });
  parados.sort(function (a, b) { return b.dias - a.dias; });

  return {
    hoje: hoje, hojeBR: hoje.split('-').reverse().join('/'),
    ontem: o, mes: soma(d.vendas.filter(function (v) { return v.data >= ini && v.data <= hoje; })),
    metas: d.metas[0] || {},
    abaixo: d.produtos.filter(function (p) { return p.estoque_minimo > 0 && p.estoque_atual < p.estoque_minimo; }),
    parados: parados, valorParado: valorParado,
    pecas: d.produtos.reduce(function (s, p) { return s + p.estoque_atual; }, 0),
    valorEstoque: Math.round(d.produtos.reduce(function (s, p) { return s + p.estoque_atual * p.custo_unitario; }, 0) * 100) / 100
  };
}
