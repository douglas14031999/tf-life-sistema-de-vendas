/**
 * TF Life – Backend (Google Apps Script) – versão de alta performance
 * Cole em Extensões > Apps Script e publique em Implantar > Gerenciar implantações > Editar > Nova versão.
 * (Mantenha "Executar como: Eu" e "Quem tem acesso: Qualquer pessoa" para a URL /exec permanecer a mesma.)
 *
 * DICA: execute UMA VEZ a função instalarGatilhos() no editor do Apps Script e autorize.
 * Ela cria um gatilho de 5 em 5 min que apaga fotos do Drive em segundo plano e mantém o cache da loja aquecido.
 */
var SS = SpreadsheetApp.getActive();
var PROD = ['id','nome','categoria','descricao','descricao_completa','preco','estoque','etiqueta','ativo','imagem','cores','tamanhos','variacoes'];
var MAX_TENT = 5, BLOQUEIO_S = 900, SESSAO_S = 21600;
var MAX_PED = 200;                 // pedidos recentes enviados ao painel (resumoPed calcula os totais de todos)
var SHARE_POR_ARQUIVO = true;      // compartilhamento das imagens do Drive
var CLIENTE_TEL_COMPLETO = true;   // true = busca por telefone devolve nome e endereço completo para preenchimento
var LIMITE_BUSCA_TEL_MIN = 150;    // limite global de segurança por minuto
var _memo = {};                    // reaproveitamento de leitura dentro da mesma requisição
var _SHEETS_MAP = null;

function getSheet(nome) {
  if (!_SHEETS_MAP) {
    _SHEETS_MAP = {};
    var sheets = SS.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      _SHEETS_MAP[sheets[i].getName()] = sheets[i];
    }
  }
  return _SHEETS_MAP[nome] || null;
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function bool(v) { return v === true || /^(true|sim|1)$/i.test(String(v)); }
function safe(v) { return typeof v === 'string' && /^[=+\-@]/.test(v) ? ' ' + v : v; } // evita fórmulas injetadas
function onlyDigits(v) { return String(v || '').replace(/\D/g, ''); }

/* ---------- Leitura de abas (com memória e cache) ---------- */
function rd(nome) {
  if (_memo[nome]) return _memo[nome];
  var s = getSheet(nome), res = [];
  if (s) {
    var v = s.getDataRange().getValues();
    if (v && v.length >= 2) {
      var h = v.shift();
      res = v.map(function (r, i) { var o = { _row: i + 2 }; h.forEach(function (k, j) { o[k] = r[j]; }); return o; });
    }
  }
  _memo[nome] = res;
  return res;
}
function inv(nome) { delete _memo[nome]; }

// Abas pequenas e estáveis com cache de 5 min no CacheService
function rdC(nome) {
  var c = CacheService.getScriptCache(), j = c.get('rd_' + nome);
  if (j) { try { return JSON.parse(j); } catch (x) {} }
  var r = rd(nome);
  try { c.put('rd_' + nome, JSON.stringify(r), 300); } catch (x) {}
  return r;
}
function invC(nome) { inv(nome); try { CacheService.getScriptCache().remove('rd_' + nome); } catch (x) {} }

function cfg() {
  var c = {}; rdC('Configuracoes').forEach(function (r) { if (r.chave) c[r.chave] = r.valor; });
  c.whatsapp = onlyDigits(c.whatsapp);
  return c;
}
function parseVars(v) {
  if (!v) return [];
  try { var a = typeof v === 'string' ? JSON.parse(v) : v; return Array.isArray(a) ? a : []; } catch (e) { return []; }
}
function prod(p) {
  return { id: String(p.id), nome: p.nome, categoria: p.categoria, descricao: p.descricao, descricao_completa: p.descricao_completa,
    preco: Number(p.preco) || 0, estoque: parseInt(p.estoque, 10) || 0, etiqueta: p.etiqueta, ativo: bool(p.ativo),
    imagem: p.imagem, cores: p.cores, tamanhos: p.tamanhos, variacoes: parseVars(p.variacoes) };
}
function produtos() { return rd('Produtos').filter(function (p) { return p.id && p.nome; }).map(prod); }
function categorias() {
  return rdC('Categorias').filter(function (c) { return c.nome; })
    .map(function (c) { return { nome: c.nome, ordem: Number(c.ordem) || 0, ativa: bool(c.ativa) }; })
    .sort(function (a, b) { return a.ordem - b.ordem; });
}

/* ---------- Cache público da loja (comprimido em blocos < 100 KB) ---------- */
function cachePutBig(key, txt, ttl) {
  try {
    var b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(txt, 'application/json', 'd.json')).getBytes());
    var CH = 90000, n = Math.ceil(b64.length / CH), o = {};
    for (var i = 0; i < n; i++) o[key + '_' + i] = b64.substr(i * CH, CH);
    o[key + '_n'] = String(n);
    CacheService.getScriptCache().putAll(o, ttl);
  } catch (x) { Logger.log('cachePutBig: ' + x); }
}
function cacheGetBig(key) {
  try {
    var c = CacheService.getScriptCache(), n = parseInt(c.get(key + '_n'), 10);
    if (!n) return null;
    var ks = []; for (var i = 0; i < n; i++) ks.push(key + '_' + i);
    var got = c.getAll(ks), b64 = '';
    for (var j = 0; j < n; j++) { if (!got[ks[j]]) return null; b64 += got[ks[j]]; }
    return Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip')).getDataAsString();
  } catch (x) { return null; }
}
function limparCache() {
  try {
    var c = CacheService.getScriptCache(), n = parseInt(c.get('loja_n'), 10) || 0, ks = ['loja_n', 'loja_publica', 'admin_tudo'];
    for (var i = 0; i < n; i++) ks.push('loja_' + i);
    c.removeAll(ks);
  } catch (x) {}
}
function lojaTexto() {
  var t = cacheGetBig('loja');
  if (t) return t;
  var c = cfg();
  t = JSON.stringify({
    ok: 1,
    config: { whatsapp: c.whatsapp, faixa: c.faixa || '' },
    categorias: categorias().filter(function (x) { return x.ativa; }).map(function (x) { return x.nome; }),
    produtos: produtos().filter(function (p) { return p.ativo; })
  });
  cachePutBig('loja', t, 900);
  return t;
}
function bumpPV() { try { CacheService.getScriptCache().put('pv', String(Date.now()), 21600); } catch (x) {} }
function versao() { return { ok: 1, pv: String(CacheService.getScriptCache().get('pv') || '0') }; }

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.a === 'cliente' || p.a === 'buscarCliente') {
    return out(buscarCliente({ telefone: onlyDigits(p.tel || p.telefone), device_id: String(p.dev || p.device_id || p.dispositivo || '').trim() }));
  }
  if (p.a === 'cupom' || p.a === 'validarCupom') {
    return out(validarCupom({ codigo: String(p.codigo || p.cod || '') }));
  }
  return ContentService.createTextOutput(lojaTexto()).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Ações ---------- */
var PUBLICAS = { login: 1, pedido: 1, buscarCliente: 1, validarCupom: 1 };
var SEM_LOCK = { login: 1, tudo: 1, pedidos: 1, versao: 1, buscarCliente: 1, validarCupom: 1, imagem: 1 };

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents), f = ACOES[d.a];
    if (!f) return out({ ok: 0, erro: 'Ação inválida.' });
    if (!PUBLICAS[d.a]) {
      var u = sessao(d.t);
      if (!u) return out({ ok: 0, erro: 'Sessão expirada. Entre de novo.', sessao: 1 });
      if (bool(u.trocar_senha) && d.a !== 'senha') return out({ ok: 0, erro: 'Defina uma nova senha primeiro.', trocar: 1 });
      d.u = u;
    }
    if (SEM_LOCK[d.a]) return out(f(d));
    var lock = LockService.getScriptLock(); lock.waitLock(15000);
    try { return out(f(d)); } finally { lock.releaseLock(); }
  } catch (x) { return out({ ok: 0, erro: String(x.message || x) }); }
}

/* ---------- Login e sessão ---------- */
function hash(u, s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'tflife|' + u + '|' + s)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function sessao(t) {
  if (!t) return null;
  var c = CacheService.getScriptCache(), g = c.getAll(['uv', 's_' + t]), uv = g.uv || '0';
  if (g['s_' + t]) { try { var o = JSON.parse(g['s_' + t]); if (o.v === uv) return o.u; } catch (x) {} }
  var user = c.get('t_' + t); if (!user) return null;
  var u = rd('Usuarios').filter(function (r) { return r.usuario === user && bool(r.ativo); })[0];
  if (u) { try { c.put('s_' + t, JSON.stringify({ v: uv, u: u }), 300); } catch (x) {} }
  return u || null;
}
function bumpUV() { try { CacheService.getScriptCache().put('uv', String(Date.now()), 21600); } catch (x) {} }

function login(d) {
  var cache = CacheService.getScriptCache(), usuario = String(d.usuario || '').trim().toLowerCase(), chave = 'f_' + usuario;
  if (Number(cache.get(chave) || 0) >= MAX_TENT) return { ok: 0, erro: 'Muitas tentativas. Aguarde 15 minutos e tente de novo.' };
  var u = rd('Usuarios').filter(function (r) { return String(r.usuario).toLowerCase() === usuario && bool(r.ativo); })[0];
  if (!u || u.senha_hash !== hash(usuario, String(d.senha || ''))) {
    cache.put(chave, String(Number(cache.get(chave) || 0) + 1), BLOQUEIO_S);
    return { ok: 0, erro: 'Usuário ou senha incorretos.' };
  }
  cache.remove(chave);
  var t = Utilities.getUuid() + Utilities.getUuid(), trocar = bool(u.trocar_senha);
  cache.put('t_' + t, u.usuario, SESSAO_S);
  try { cache.put('s_' + t, JSON.stringify({ v: cache.get('uv') || '0', u: u }), 300); } catch (x) {}
  var r = { ok: 1, token: t, nome: u.nome, trocar: trocar };
  // OTIMIZAÇÃO CRÍTICA: Os dados completos do painel já vão no retorno do login!
  if (!trocar) {
    try { r.dados = tudo(); } catch (xTudo) { Logger.log('Erro ao carregar tudo no login: ' + xTudo); }
  }
  return r;
}
function senha(d) {
  var u = d.u, nova = String(d.nova || '');
  if (u.senha_hash !== hash(u.usuario, String(d.atual || ''))) return { ok: 0, erro: 'Senha atual incorreta.' };
  if (nova.length < 8) return { ok: 0, erro: 'A nova senha precisa ter pelo menos 8 caracteres.' };
  if (nova === d.atual) return { ok: 0, erro: 'Escolha uma senha diferente da atual.' };
  var s = getSheet('Usuarios');
  if (s) {
    s.getRange(u._row, 3).setValue(hash(u.usuario, nova));
    s.getRange(u._row, 5).setValue(false);
  }
  try { CacheService.getScriptCache().remove('s_' + d.t); } catch (x) {}
  inv('Usuarios'); bumpUV();
  return { ok: 1 };
}

/* ---------- Painel ---------- */
function tudo() {
  return {
    ok: 1,
    produtos: produtos(),
    categorias: categorias(),
    config: cfg(),
    usuarios: rd('Usuarios').filter(function (r) { return r.usuario; })
      .map(function (r) { return { usuario: r.usuario, nome: r.nome, ativo: bool(r.ativo) }; }),
    cupons: cupons(),
    pedidos: pedidosList(MAX_PED),
    resumoPed: resumoPedidos()
  };
}

function linhaLivre(s) {
  var last = s.getLastRow();
  if (last >= 2) {
    var v = s.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < v.length; i++) if (v[i][0] === '') return i + 2;
  }
  var nxt = Math.max(last + 1, 2);
  if (nxt > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), 1);
  return nxt;
}
function acharLinha(s, valor) {
  var last = s.getLastRow(); if (last < 2) return 0;
  var r = s.getRange(2, 1, last - 1, 1).createTextFinder(String(valor)).matchEntireCell(true).matchCase(true).findNext();
  return r ? r.getRow() : 0;
}

/* ---------- Fotos no Drive ---------- */
function idDrive(url) {
  if (!url || typeof url !== 'string') return '';
  var m = /[?&]id=([a-zA-Z0-9_-]+)/.exec(url) || /\/d\/([a-zA-Z0-9_-]+)/.exec(url);
  return m ? m[1] : '';
}
function excluirArquivoDrive(url) {
  var id = idDrive(url); if (!id) return;
  try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { Logger.log('Drive: ' + e); }
}
function enfileirarExclusao(url) {
  var id = idDrive(url); if (!id) return;
  var P = PropertiesService.getScriptProperties(), fila = [];
  try { fila = JSON.parse(P.getProperty('fila_drive') || '[]'); } catch (x) {}
  if (fila.length >= 200) { excluirArquivoDrive(url); return; }
  if (fila.indexOf(id) === -1) fila.push(id);
  P.setProperty('fila_drive', JSON.stringify(fila));
}
function processarFilaDrive() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    var P = PropertiesService.getScriptProperties(), fila = [];
    try { fila = JSON.parse(P.getProperty('fila_drive') || '[]'); } catch (x) {}
    if (!fila.length) return;
    fila.forEach(function (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { Logger.log('Drive: ' + e); } });
    P.deleteProperty('fila_drive');
  } finally { lock.releaseLock(); }
}
function pastaImagens() {
  var P = PropertiesService.getScriptProperties(), id = P.getProperty('pasta_id');
  if (id) { try { return DriveApp.getFolderById(id); } catch (x) {} }
  var it = DriveApp.getFoldersByName('TF Life - Imagens'), f = it.hasNext() ? it.next() : DriveApp.createFolder('TF Life - Imagens');
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (x) {}
  P.setProperty('pasta_id', f.getId());
  return f;
}
function imagem(d) {
  var m = /^data:(image\/(jpeg|png|webp));base64,(.+)$/.exec(String(d.dados || ''));
  if (!m) return { ok: 0, erro: 'Formato de imagem não aceito (use JPG, PNG ou WebP).' };
  var bytes = Utilities.base64Decode(m[3]);
  if (bytes.length > 3 * 1024 * 1024) return { ok: 0, erro: 'Imagem acima de 3 MB.' };
  var f = pastaImagens().createFile(Utilities.newBlob(bytes, m[1], 'produto-' + Date.now() + '.' + m[2].split('/')[1]));
  if (SHARE_POR_ARQUIVO) f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: 1, url: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w900' };
}

/* ---------- Gatilho de manutenção (rodar instalarGatilhos() uma vez no editor) ---------- */
function manutencao() {
  try { processarFilaDrive(); } catch (x) { Logger.log(x); }
  try { lojaTexto(); } catch (x) { Logger.log(x); }
}
function instalarGatilhos() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'manutencao') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('manutencao').timeBased().everyMinutes(5).create();
}

/* ---------- Produtos, categorias, usuários, config ---------- */
function salvarProduto(d) {
  var p = d.p || {}, s = getSheet('Produtos');
  var nome = String(p.nome || '').trim(), preco = Number(p.preco), est = parseInt(p.estoque, 10);
  var categoria = String(p.categoria || '').trim();
  if (!nome) return { ok: 0, erro: 'Informe o nome do produto.' };
  if (!categoria) return { ok: 0, erro: 'Selecione a categoria do produto.' };
  var catValida = categorias().filter(function (c) { return String(c.nome).trim().toLowerCase() === categoria.toLowerCase(); })[0];
  if (!catValida) return { ok: 0, erro: 'Escolha uma categoria válida.' };
  categoria = catValida.nome;
  if (!(preco >= 0)) return { ok: 0, erro: 'Preço inválido.' };
  if (!(est >= 0)) return { ok: 0, erro: 'Estoque inválido.' };
  var img = String(p.imagem || ''); if (img && !/^https:\/\//.test(img)) return { ok: 0, erro: 'Imagem inválida.' };
  var id = p.id || 'p' + Date.now().toString(36);
  var ex = rd('Produtos').filter(function (r) { return String(r.id) === id; })[0];

  var varsJson = '', novasImagens = [];
  if (img) novasImagens.push(img);
  var listaV = parseVars(p.variacoes);
  if (listaV.length) {
    varsJson = JSON.stringify(listaV);
    listaV.forEach(function (v) { if (v && v.imagem) novasImagens.push(v.imagem); });
    if (!img && listaV[0].imagem) img = listaV[0].imagem;
  }

  if (ex) {
    var antigas = [];
    if (ex.imagem) antigas.push(ex.imagem);
    parseVars(ex.variacoes).forEach(function (v) { if (v && v.imagem) antigas.push(v.imagem); });
    antigas.forEach(function (f) { if (novasImagens.indexOf(f) === -1) enfileirarExclusao(f); });
  }

  var h = s.getRange(1, 1, 1, Math.max(s.getLastColumn(), PROD.length)).getValues()[0];
  for (var ci = 0; ci < PROD.length; ci++) {
    if (h[ci] !== PROD[ci]) { s.getRange(1, 1, 1, PROD.length).setValues([PROD]); break; }
  }

  var linha = ex ? ex._row : linhaLivre(s);
  var vals = [id, nome, categoria, p.descricao, p.descricao_completa, preco, est, p.etiqueta, !!p.ativo, img, p.cores, p.tamanhos, varsJson]
    .map(function (v) { return safe(typeof v === 'string' ? v.trim() : v); });
  s.getRange(linha, 1, 1, PROD.length).setValues([vals]);
  inv('Produtos'); limparCache();
  return { ok: 1, id: id };
}
function excluirProduto(d) {
  var r = rd('Produtos').filter(function (x) { return String(x.id) === d.id; })[0];
  if (!r) return { ok: 0, erro: 'Produto não encontrado.' };
  if (r.imagem) enfileirarExclusao(r.imagem);
  parseVars(r.variacoes).forEach(function (v) { if (v && v.imagem) enfileirarExclusao(v.imagem); });
  var s = getSheet('Produtos');
  if (s) s.getRange(r._row, 1, 1, PROD.length).clearContent();
  inv('Produtos'); limparCache();
  return { ok: 1 };
}
function salvarCategoria(d) {
  var nome = String(d.nome || '').trim(), antigo = d.antigo, s = getSheet('Categorias');
  if (!nome) return { ok: 0, erro: 'Informe o nome da categoria.' };
  var todas = rd('Categorias');
  if (todas.some(function (c) { return c.nome === nome && c.nome !== antigo; })) return { ok: 0, erro: 'Já existe uma categoria com esse nome.' };
  var ex = todas.filter(function (c) { return c.nome === antigo; })[0];
  var linha = ex ? ex._row : linhaLivre(s);
  s.getRange(linha, 1, 1, 3).setValues([[safe(nome), Number(d.ordem) || 0, !!d.ativa]]);
  if (ex && antigo !== nome) {
    var ps = getSheet('Produtos');
    if (ps) rd('Produtos').forEach(function (p) { if (p.categoria === antigo) ps.getRange(p._row, 3).setValue(nome); });
    inv('Produtos');
  }
  invC('Categorias'); limparCache();
  return { ok: 1 };
}
function excluirCategoria(d) {
  if (produtos().some(function (p) { return p.categoria === d.nome; })) return { ok: 0, erro: 'Há produtos nessa categoria. Mude os produtos de categoria ou desative-a.' };
  var c = rd('Categorias').filter(function (x) { return x.nome === d.nome; })[0];
  if (!c) return { ok: 0, erro: 'Categoria não encontrada.' };
  var s = getSheet('Categorias');
  if (s) s.getRange(c._row, 1, 1, 3).clearContent();
  invC('Categorias'); limparCache();
  return { ok: 1 };
}
function salvarUsuario(d) {
  var s = getSheet('Usuarios'), usuario = String(d.usuario || '').trim().toLowerCase(), nome = String(d.nome || '').trim();
  if (!/^[a-z0-9._-]{3,30}$/.test(usuario)) return { ok: 0, erro: 'Usuário: 3 a 30 letras minúsculas, números, ponto, hífen ou _.' };
  if (!nome) return { ok: 0, erro: 'Informe o nome.' };
  var ex = rd('Usuarios').filter(function (r) { return String(r.usuario).toLowerCase() === usuario; })[0];
  var nova = String(d.senha || '');
  if (!ex && nova.length < 8) return { ok: 0, erro: 'A senha temporária precisa ter pelo menos 8 caracteres.' };
  if (nova && nova.length < 8) return { ok: 0, erro: 'A senha precisa ter pelo menos 8 caracteres.' };
  if (ex && usuario === d.u.usuario && !d.ativo) return { ok: 0, erro: 'Você não pode desativar o próprio usuário.' };
  var linha = ex ? ex._row : linhaLivre(s);
  s.getRange(linha, 1, 1, 2).setValues([[usuario, safe(nome)]]);
  s.getRange(linha, 4).setValue(!!d.ativo);
  if (nova) { s.getRange(linha, 3).setValue(hash(usuario, nova)); s.getRange(linha, 5).setValue(true); }
  inv('Usuarios'); bumpUV();
  return { ok: 1 };
}
function salvarConfig(d) {
  var s = getSheet('Configuracoes'), zap = onlyDigits(d.whatsapp);
  if (zap.length < 12) return { ok: 0, erro: 'WhatsApp: use 55 + DDD + número, só dígitos.' };
  rd('Configuracoes').forEach(function (r) {
    if (r.chave === 'whatsapp') { var c = s.getRange(r._row, 2); c.setNumberFormat('@'); c.setValue(zap); }
    if (r.chave === 'faixa') s.getRange(r._row, 2).setValue(safe(String(d.faixa || '').trim()));
  });
  invC('Configuracoes'); limparCache();
  return { ok: 1 };
}

/* ---------- Pedidos e clientes ---------- */
function getOuCriarAba(nome, cabecalhos) {
  var s = getSheet(nome);
  if (!s) {
    s = SS.insertSheet(nome);
    s.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]).setFontWeight('bold');
    try { s.setFrozenRows(1); } catch (x) {}
    _SHEETS_MAP = null;
  }
  return s;
}
function getAbaClientes() { return getSheet('Clientes') || getSheet('Usuarios_Clientes') || getSheet('Clientes_Loja'); }
function garantirAbaClientes() {
  var s = getAbaClientes();
  if (!s) {
    s = getOuCriarAba('Clientes', ['telefone', 'nome', 'endereco', 'data_cadastro', 'ultimo_pedido', 'total_pedidos', 'dispositivos']);
    try { s.getRange('A:A').setNumberFormat('@'); s.getRange('G:G').setNumberFormat('@'); } catch (x) {}
  }
  return s;
}
function garantirAbaPedidos() {
  var s = getSheet('Pedidos');
  if (!s) {
    s = getOuCriarAba('Pedidos', ['id_pedido', 'data_hora', 'telefone', 'nome', 'endereco', 'forma_pagamento', 'itens', 'total', 'status']);
    try { s.getRange('C:C').setNumberFormat('@'); } catch (x) {}
  }
  return s;
}
function formatarDataHoraSheet(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var tz = 'America/Sao_Paulo';
    try { tz = Session.getScriptTimeZone() || tz; } catch (x) {}
    return Utilities.formatDate(val, tz, 'dd/MM/yyyy HH:mm:ss');
  }
  return String(val);
}
function mapPedido(p) {
  return { id: String(p.id_pedido), data_hora: formatarDataHoraSheet(p.data_hora), telefone: String(p.telefone || ''), nome: String(p.nome || ''),
    endereco: String(p.endereco || ''), forma_pagamento: String(p.forma_pagamento || ''), itens: String(p.itens || ''),
    total: Number(p.total) || 0, status: String(p.status || 'Recebido') };
}

// Lê apenas os últimos N pedidos para máxima velocidade
function pedidosList(max) {
  var s = getSheet('Pedidos'); if (!s) return [];
  var last = s.getLastRow(); if (last < 2) return [];
  var n = Math.min(max || MAX_PED, last - 1);
  var h = s.getRange(1, 1, 1, 9).getValues()[0];
  return s.getRange(last - n + 1, 1, n, 9).getValues().map(function (r) { var o = {}; h.forEach(function (k, j) { o[k] = r[j]; }); return o; })
    .filter(function (p) { return p.id_pedido; }).map(mapPedido).reverse();
}

// Totais de TODOS os pedidos lendo apenas colunas total e status
function resumoPedidos() {
  var s = getSheet('Pedidos'), r = { total: 0, faturamento: 0, pendentes: 0 };
  if (!s || s.getLastRow() < 2) return r;
  s.getRange(2, 8, s.getLastRow() - 1, 2).getValues().forEach(function (x) {
    if (x[0] === '' && !x[1]) return;
    var st = x[1] || 'Recebido'; r.total++;
    if (st !== 'Cancelado') r.faturamento += Number(x[0]) || 0;
    if (st === 'Recebido' || st === 'Em preparo' || st === 'Pronto / Enviado') r.pendentes++;
  });
  r.faturamento = Number(r.faturamento.toFixed(2));
  return r;
}
function pedidosListAction() { return { ok: 1, pedidos: pedidosList(MAX_PED), resumoPed: resumoPedidos() }; }

function atualizarStatusPedido(d) {
  var id = String(d.id || '').trim(), st = String(d.status || '').trim();
  if (!id || !st) return { ok: 0, erro: 'Informe o ID e o novo status.' };
  var s = getSheet('Pedidos'); if (!s) return { ok: 0, erro: 'Aba Pedidos não encontrada.' };
  var lin = acharLinha(s, id); if (!lin) return { ok: 0, erro: 'Pedido não encontrado.' };
  s.getRange(lin, 9).setValue(safe(st)); bumpPV();
  return { ok: 1 };
}
function excluirPedido(d) {
  var id = String(d.id || '').trim(); if (!id) return { ok: 0, erro: 'ID do pedido não informado.' };
  var s = getSheet('Pedidos'); if (!s) return { ok: 0, erro: 'Aba Pedidos não encontrada.' };
  var lin = acharLinha(s, id); if (!lin) return { ok: 0, erro: 'Pedido não encontrado.' };
  s.deleteRow(lin); bumpPV();
  return { ok: 1 };
}

function mesmoTel(a, b) {
  if (!a || !b) return false;
  return a === b || (a.length >= 10 && b.length >= 10 && (a.slice(-10) === b.slice(-10) || a.slice(-11) === b.slice(-11)));
}
function limiteBuscaOk() {
  var c = CacheService.getScriptCache(), k = 'bt_' + Math.floor(Date.now() / 60000), n = Number(c.get(k) || 0);
  if (n >= LIMITE_BUSCA_TEL_MIN) return false;
  c.put(k, String(n + 1), 120); return true;
}
function buscarCliente(d) {
  var tel = onlyDigits(d.telefone || d.tel), dev = String(d.device_id || d.dispositivo || d.dev || '').trim();
  if (tel.length < 8 && !dev) return { ok: 1, encontrado: false };
  var aba = getAbaClientes(); if (!aba) return { ok: 1, encontrado: false };
  var lista = rd(aba.getName()), achou = null, porDisp = false;
  if (dev) {
    achou = lista.filter(function (r) {
      return String(r.dispositivos || '').split(/[,;\s]+/).map(function (s) { return s.trim(); }).filter(Boolean).indexOf(dev) !== -1;
    })[0];
    porDisp = !!achou;
  }
  if (!achou && tel.length >= 8) {
    if (!limiteBuscaOk()) return { ok: 1, encontrado: false };
    achou = lista.filter(function (r) { return mesmoTel(onlyDigits(r.telefone), tel); })[0];
  }
  if (achou && (achou.nome || achou.telefone)) {
    var endRaw = String(achou.endereco || '').trim();
    var arrEnd = endRaw ? endRaw.split(/\s*\|\s*|\s*;\s*|\n+/).map(function (s) { return s.trim(); }).filter(Boolean) : [];
    var endPrincipal = arrEnd[0] || endRaw;

    if (porDisp || CLIENTE_TEL_COMPLETO) {
      return {
        ok: 1,
        encontrado: true,
        telefone: String(achou.telefone || ''),
        nome: String(achou.nome || ''),
        endereco: endPrincipal,
        enderecos: arrEnd
      };
    }
    return { ok: 1, encontrado: true, nome: String(achou.nome || '').trim().split(/\s+/)[0] };
  }
  return { ok: 1, encontrado: false };
}
function upsertCliente(telefone, nome, endereco, agora, dispositivo) {
  if (!telefone) return;
  var s = garantirAbaClientes(), telLimpo = onlyDigits(telefone);
  var headers = s.getRange(1, 1, 1, Math.max(s.getLastColumn(), 1)).getValues()[0].map(function (x) { return String(x || '').trim().toLowerCase(); });
  var colTel = headers.indexOf('telefone') + 1 || 1, colNome = headers.indexOf('nome') + 1 || 2, colEnd = headers.indexOf('endereco') + 1 || 3;
  var colCad = headers.indexOf('data_cadastro') + 1 || 4, colUlt = headers.indexOf('ultimo_pedido') + 1 || 5, colTot = headers.indexOf('total_pedidos') + 1 || 6;
  var colDisp = headers.indexOf('dispositivos') + 1;
  if (!colDisp) {
    colDisp = headers.length + 1;
    s.getRange(1, colDisp).setValue('dispositivos').setFontWeight('bold');
    try { s.getRange(1, colDisp, s.getMaxRows(), 1).setNumberFormat('@'); } catch (x) {}
  }
  var ncol = Math.max(headers.length, colDisp), dev = String(dispositivo || '').trim();
  var ex = rd(s.getName()).filter(function (r) { return mesmoTel(onlyDigits(r.telefone), telLimpo); })[0];
  var novoEnd = String(endereco || '').trim();

  if (ex) {
    var row = s.getRange(ex._row, 1, 1, ncol).getValues()[0];
    if (nome) row[colNome - 1] = safe(nome);
    if (novoEnd) {
      var endAtual = String(row[colEnd - 1] || '').trim();
      var arrEnd = endAtual ? endAtual.split(/\s*\|\s*|\s*;\s*|\n+/).map(function (s) { return s.trim(); }).filter(Boolean) : [];
      arrEnd = arrEnd.filter(function (e) { return e.toLowerCase() !== novoEnd.toLowerCase(); });
      arrEnd.unshift(novoEnd);
      if (arrEnd.length > 5) arrEnd = arrEnd.slice(0, 5);
      row[colEnd - 1] = safe(arrEnd.join(' | '));
    }
    row[colUlt - 1] = agora;
    row[colTot - 1] = (parseInt(ex.total_pedidos, 10) || 0) + 1;
    if (dev) {
      var arr = String(row[colDisp - 1] || '').split(/[,;\s]+/).map(function (x) { return x.trim(); }).filter(Boolean);
      if (arr.indexOf(dev) === -1) { arr.push(dev); row[colDisp - 1] = arr.join(', '); }
    }
    s.getRange(ex._row, 1, 1, ncol).setValues([row]);
  } else {
    var lin = linhaLivre(s), nv = []; for (var i = 0; i < ncol; i++) nv.push('');
    nv[colTel - 1] = telLimpo; nv[colNome - 1] = safe(nome || ''); nv[colEnd - 1] = safe(novoEnd);
    nv[colCad - 1] = agora; nv[colUlt - 1] = agora; nv[colTot - 1] = 1;
    if (dev) nv[colDisp - 1] = dev;
    try { s.getRange(lin, colTel).setNumberFormat('@'); } catch (x) {}
    s.getRange(lin, 1, 1, ncol).setValues([nv]);
  }
}

/* ---------- Cupons ---------- */
function garantirAbaCupons() {
  var s = getSheet('Cupons');
  if (!s) {
    s = getOuCriarAba('Cupons', ['codigo', 'tipo', 'valor', 'limite_usos', 'usos_atuais', 'validade', 'ativo']);
    try { s.getRange('A:A').setNumberFormat('@'); s.getRange('F:F').setNumberFormat('@'); } catch (x) {}
  }
  return s;
}
function formatarDataCupom(val) {
  if (!val) return '';
  if (typeof val === 'object' && val.getTime) return Utilities.formatDate(val, 'America/Sao_Paulo', 'yyyy-MM-dd');
  var s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var p = s.split(/[-/]/);
  if (p.length === 3) {
    if (p[0].length === 4) return p[0] + '-' + ('0' + p[1]).slice(-2) + '-' + ('0' + p[2]).slice(-2);
    return p[2] + '-' + ('0' + p[1]).slice(-2) + '-' + ('0' + p[0]).slice(-2);
  }
  return s;
}
function couponExpirou(v) {
  if (!v) return false;
  var dt = null;
  if (typeof v === 'object' && v.getTime) dt = new Date(v.getFullYear(), v.getMonth(), v.getDate(), 23, 59, 59);
  else if (typeof v === 'string') {
    var p = v.trim().split(/[-/]/);
    if (p.length === 3) dt = p[0].length === 4
      ? new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10), 23, 59, 59)
      : new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10), 23, 59, 59);
  }
  if (!dt || isNaN(dt.getTime())) return false;
  return Date.now() > dt.getTime();
}
function cupons() {
  return rd('Cupons').filter(function (c) { return c.codigo; }).map(function (c) {
    return { codigo: String(c.codigo).toUpperCase(), tipo: c.tipo === 'porcentagem' ? 'porcentagem' : 'fixo', valor: Number(c.valor) || 0,
      limite_usos: parseInt(c.limite_usos, 10) || 0, usos_atuais: parseInt(c.usos_atuais, 10) || 0,
      validade: formatarDataCupom(c.validade), ativo: bool(c.ativo) };
  });
}
function validarCupom(d) {
  var cod = String(d.codigo || d.cupom || '').trim().toUpperCase();
  if (!cod) return { ok: 1, valido: false, erro: 'Informe o código do cupom.' };
  var c = rd('Cupons').filter(function (r) { return String(r.codigo || '').trim().toUpperCase() === cod; })[0];
  if (!c || !bool(c.ativo)) return { ok: 1, valido: false, erro: 'Cupom desativado ou não existe.' };
  var lim = parseInt(c.limite_usos, 10) || 0, usos = parseInt(c.usos_atuais, 10) || 0;
  if (lim > 0 && usos >= lim) return { ok: 1, valido: false, erro: 'Este cupom já atingiu o limite de usos e está esgotado.' };
  if (c.validade && couponExpirou(c.validade)) return { ok: 1, valido: false, erro: 'Este cupom expirou e está indisponível.' };
  return { ok: 1, valido: true, codigo: String(c.codigo).toUpperCase(), tipo: c.tipo === 'porcentagem' ? 'porcentagem' : 'fixo', valor: Number(c.valor) || 0 };
}
function salvarCupom(d) {
  var c = d.c || d, s = garantirAbaCupons();
  var codigo = String(c.codigo || '').trim().toUpperCase().replace(/[^A-Z0-9_-]/g, '');
  if (codigo.length < 2) return { ok: 0, erro: 'Código do cupom inválido. Use ao menos 2 letras ou números.' };
  var tipo = c.tipo === 'porcentagem' ? 'porcentagem' : 'fixo', valor = Number(c.valor);
  if (!(valor > 0)) return { ok: 0, erro: 'Informe um valor de desconto válido maior que zero.' };
  if (tipo === 'porcentagem' && valor > 100) return { ok: 0, erro: 'Desconto em porcentagem não pode ultrapassar 100%.' };
  var ex = rd('Cupons').filter(function (r) { return String(r.codigo || '').trim().toUpperCase() === codigo; })[0];
  var linha = ex ? ex._row : linhaLivre(s), usos = ex ? (parseInt(ex.usos_atuais, 10) || 0) : 0;
  try { s.getRange(linha, 1).setNumberFormat('@'); s.getRange(linha, 6).setNumberFormat('@'); } catch (x) {}
  s.getRange(linha, 1, 1, 7).setValues([[codigo, tipo, valor, parseInt(c.limite_usos, 10) || 0, usos, c.validade ? formatarDataCupom(c.validade) : '', bool(c.ativo)]]);
  inv('Cupons');
  return { ok: 1, codigo: codigo };
}
function excluirCupom(d) {
  var codigo = String(d.codigo || '').trim().toUpperCase(), s = getSheet('Cupons');
  if (!s) return { ok: 0, erro: 'Aba de cupons não encontrada.' };
  var ex = rd('Cupons').filter(function (r) { return String(r.codigo || '').trim().toUpperCase() === codigo; })[0];
  if (!ex) return { ok: 0, erro: 'Cupom não encontrado.' };
  s.getRange(ex._row, 1, 1, 7).clearContent(); inv('Cupons');
  return { ok: 1 };
}

/* ---------- Pedido da loja (estoque, cupom, registro) ---------- */
function precoItem(p, it) {
  var pr = Number(p.preco) || 0, vars = parseVars(p.variacoes);
  for (var i = 0; i < vars.length; i++) {
    var v = vars[i];
    if ((it.varId && v.id === it.varId) || (it.varNome && v.nome === it.varNome)) {
      var vp = Number(v.preco);
      if (v.preco !== '' && v.preco != null && !isNaN(vp) && vp >= 0) pr = vp;
      break;
    }
  }
  return pr;
}
function pedido(d) {
  var sProd = getSheet('Produtos'), pm = {};
  rd('Produtos').forEach(function (r) { if (r.id !== '') pm[String(r.id)] = r; });

  var itens = [], mapa = {}, falta = [];
  (d.itens || []).slice(0, 40).forEach(function (i) {
    var q = parseInt(i.q, 10), id = String(i.id || '');
    if (!(q > 0 && q <= 50)) return;
    var p = pm[id];
    if (!p || !bool(p.ativo)) { if (falta.indexOf('produto indisponível') < 0) falta.push('produto indisponível'); return; }
    mapa[id] = (mapa[id] || 0) + q;
    itens.push({ id: id, q: q, varId: i.varId, varNome: i.varNome, cor: i.cor, tam: i.tam });
  });
  var ids = Object.keys(mapa);
  if (!ids.length && !falta.length) return { ok: 0, erro: 'Carrinho vazio.' };
  ids.forEach(function (id) {
    var est = parseInt(pm[id].estoque, 10) || 0;
    if (est < mapa[id]) falta.push(pm[id].nome + ' (restam ' + est + ')');
  });
  if (falta.length) return { ok: 0, erro: 'Sem estoque suficiente: ' + falta.join(', ') + '.', atualizar: 1 };

  // Baixa de estoque
  ids.forEach(function (id) {
    var p = pm[id];
    sProd.getRange(p._row, 7).setValue(Math.max(0, (parseInt(p.estoque, 10) || 0) - mapa[id]));
    var vars = parseVars(p.variacoes), mudou = false;
    if (vars.length) {
      itens.filter(function (it) { return it.id === id && (it.varId || it.varNome); }).forEach(function (it) {
        vars.forEach(function (v) {
          if ((it.varId && v.id === it.varId) || (it.varNome && v.nome === it.varNome)) {
            v.estoque = Math.max(0, (parseInt(v.estoque, 10) || 0) - it.q);
            if (v.tamanhos && it.tam && v.tamanhos[it.tam] !== undefined) v.tamanhos[it.tam] = Math.max(0, (parseInt(v.tamanhos[it.tam], 10) || 0) - it.q);
            mudou = true;
          }
        });
      });
      if (mudou) sProd.getRange(p._row, 13).setValue(JSON.stringify(vars));
    }
  });

  var resumo = [], total = 0;
  itens.forEach(function (it) {
    var p = pm[it.id], sub = it.q * precoItem(p, it);
    total += sub;
    var det = [it.varNome ? 'Opção: ' + it.varNome : '', it.cor, it.tam].filter(Boolean).join('/');
    resumo.push(it.q + 'x ' + p.nome + (det ? ' (' + det + ')' : '') + ' [R$ ' + sub.toFixed(2) + ']');
  });

  var cupomCod = String(d.cupom || '').trim().toUpperCase(), desconto = 0;
  if (cupomCod) {
    var rv = validarCupom({ codigo: cupomCod });
    if (rv.ok && rv.valido) {
      desconto = rv.tipo === 'porcentagem' ? Number((total * (rv.valor / 100)).toFixed(2)) : Math.min(total, Number(rv.valor));
      try {
        var sCup = getSheet('Cupons');
        var exCup = rd('Cupons').filter(function (r) { return String(r.codigo || '').trim().toUpperCase() === cupomCod; })[0];
        if (sCup && exCup) { sCup.getRange(exCup._row, 5).setValue((parseInt(exCup.usos_atuais, 10) || 0) + 1); inv('Cupons'); }
      } catch (xc) { Logger.log('Cupom: ' + xc); }
    }
  }
  var totalFinal = Math.max(0, Number((total - desconto).toFixed(2)));

  var tz = 'America/Sao_Paulo';
  try { tz = Session.getScriptTimeZone() || tz; } catch (x) {}
  var agora = new Date(), agoraStr = Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm:ss');
  var tel = onlyDigits(d.telefone), nomeC = String(d.nome || '').trim(), endC = String(d.endereco || '').trim();
  var pag = String(d.pagamento || d.forma_pagamento || '').trim(), disp = String(d.dispositivo || d.device_id || '').trim();

  var sPed = garantirAbaPedidos(), lin = linhaLivre(sPed);
  var numPedido = Math.max(1, Math.max(sPed.getLastRow(), lin - 1));
  var idPedido = 'PED-' + ('0000' + numPedido).slice(-4);
  var txtItens = resumo.join('; ');
  if (cupomCod && desconto > 0) txtItens += ' [CUPOM: ' + cupomCod + ' (-R$ ' + desconto.toFixed(2) + ')]';

  try {
    try { sPed.getRange(lin, 3).setNumberFormat('@'); } catch (x) {}
    sPed.getRange(lin, 1, 1, 9).setValues([[idPedido, agoraStr, tel, safe(nomeC), safe(endC), safe(pag), safe(txtItens), totalFinal, 'Recebido']]);
  } catch (xp) { Logger.log('Pedido: ' + xp); }
  if (tel) { try { upsertCliente(tel, nomeC, endC, agoraStr, disp); } catch (xc2) { Logger.log('Cliente: ' + xc2); } }

  inv('Produtos'); limparCache(); bumpPV();
  return { ok: 1, id_pedido: idPedido, num_pedido: numPedido, desconto: desconto, total: totalFinal };
}

var ACOES = {
  login: login, senha: senha, tudo: tudo, versao: versao,
  salvarProduto: salvarProduto, excluirProduto: excluirProduto,
  salvarCategoria: salvarCategoria, excluirCategoria: excluirCategoria,
  salvarUsuario: salvarUsuario, salvarConfig: salvarConfig, imagem: imagem,
  pedido: pedido, pedidos: pedidosListAction,
  atualizarStatusPedido: atualizarStatusPedido, excluirPedido: excluirPedido,
  buscarCliente: buscarCliente, validarCupom: validarCupom,
  salvarCupom: salvarCupom, excluirCupom: excluirCupom
};
