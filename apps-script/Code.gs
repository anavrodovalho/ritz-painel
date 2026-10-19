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
 * Ao mudar este código: Implantar → Gerenciar implantações → editar → Nova versão.
 *
 * GET <url>/exec?token=SENHA  →  { ok, gerado, produtos, vendas, compras, metas, ... }
 */

var ABA_CADASTRO = 'CADASTRO E CONTROLE';
var ABA_SAIDAS   = 'SAÍDAS';
var ABA_ENTRADAS = 'ENTRADAS';
var ABA_PRECOS   = 'PREÇOS';
var ABA_METAS    = 'METAS';
var FUSO         = 'America/Sao_Paulo';

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
  var vendas = [], brindes = 0, precosVendidos = {};
  sai.linhas.forEach(function (r) {
    var sku = txt_(r[sai.col['Cod. Item']]), data = data_(r[sai.col['Data']]);
    if (!sku || !data) return;
    var qtd = num_(r[sai.col['Quantidade']]) || 1, preco = num_(r[sai.col['Preço Venda']]);
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
  var produtos = [];
  cad.linhas.forEach(function (r) {
    var sku = txt_(r[cad.col['Cod. Item']]); if (!sku) return;
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
    pedidos: [], leads: [], campanhas: [], metas: lerMetas_(ss), snapshots: [], brindes_excluidos: brindes };
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
  return { col: col, linhas: v.slice(h + 1) };
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
  var m = String(v || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
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
  Logger.log('produtos %s · vendas %s · compras %s · metas %s · brindes excluídos %s',
    d.produtos.length, d.vendas.length, d.compras.length, d.metas.length, d.brindes_excluidos);
}
