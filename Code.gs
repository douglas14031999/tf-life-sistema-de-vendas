/**
 * TF Life – Backend (Google Apps Script)
 * Cole este arquivo em Extensões > Apps Script da planilha "TF Life – Loja (dados)"
 * e publique em Implantar > Nova implantação > App da Web
 *   Executar como: Eu | Quem tem acesso: Qualquer pessoa
 * Copie a URL /exec gerada e cole na variável API das páginas HTML.
 */
var SS = SpreadsheetApp.getActive();
var PROD = ['id','nome','categoria','descricao','descricao_completa','preco','estoque','etiqueta','ativo','imagem','cores','tamanhos','variacoes'];
var MAX_TENT = 5, BLOQUEIO_S = 900, SESSAO_S = 21600;

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function bool(v) { return v === true || /^(true|sim|1)$/i.test(String(v)); }
function safe(v) { return typeof v === 'string' && /^[=+\-@]/.test(v) ? ' ' + v : v; } // evita fórmulas injetadas
function rd(nome) {
  var s = SS.getSheetByName(nome);
  if (!s) return [];
  var v = s.getDataRange().getValues();
  if (!v || v.length < 2) return [];
  var h = v.shift();
  return v.map(function (r, i) { var o = { _row: i + 2 }; h.forEach(function (k, j) { o[k] = r[j]; }); return o; });
}
function cfg() {
  var c = {}; rd('Configuracoes').forEach(function (r) { if (r.chave) c[r.chave] = r.valor; });
  c.whatsapp = String(c.whatsapp || '').replace(/\D/g, '');
  return c;
}
function prod(p) {
  var vars = [];
  if (p.variacoes) {
    try {
      vars = typeof p.variacoes === 'string' ? JSON.parse(p.variacoes) : p.variacoes;
      if (!Array.isArray(vars)) vars = [];
    } catch (e) { vars = []; }
  }
  return { id: String(p.id), nome: p.nome, categoria: p.categoria, descricao: p.descricao, descricao_completa: p.descricao_completa,
    preco: Number(p.preco) || 0, estoque: parseInt(p.estoque, 10) || 0, etiqueta: p.etiqueta, ativo: bool(p.ativo),
    imagem: p.imagem, cores: p.cores, tamanhos: p.tamanhos, variacoes: vars };
}
function produtos() { return rd('Produtos').filter(function (p) { return p.id && p.nome; }).map(prod); }
function categorias() {
  return rd('Categorias').filter(function (c) { return c.nome; })
    .map(function (c) { return { nome: c.nome, ordem: Number(c.ordem) || 0, ativa: bool(c.ativa) }; })
    .sort(function (a, b) { return a.ordem - b.ordem; });
}

/* ---------- Loja (leitura pública com Cache em Memória) ---------- */
function limparCache() {
  try { CacheService.getScriptCache().remove('loja_publica'); } catch (x) {}
}

function doGet(e) {
  // Busca rápida de cliente por telefone ou por dispositivo via GET (?a=cliente&tel=... ou ?a=cliente&dev=...)
  if (e && e.parameter && (e.parameter.a === 'cliente' || e.parameter.a === 'buscarCliente')) {
    var tel = String(e.parameter.tel || e.parameter.telefone || '').replace(/\D/g, '');
    var dev = String(e.parameter.dev || e.parameter.device_id || e.parameter.dispositivo || '').trim();
    return out(buscarCliente({ telefone: tel, device_id: dev }));
  }

  // Validação rápida de cupom via GET (?a=cupom&codigo=DE20)
  if (e && e.parameter && (e.parameter.a === 'cupom' || e.parameter.a === 'validarCupom')) {
    var cod = String(e.parameter.codigo || e.parameter.cod || '');
    return out(validarCupom({ codigo: cod }));
  }

  var cache = CacheService.getScriptCache();
  var cJson = cache.get('loja_publica');
  if (cJson) {
    return ContentService.createTextOutput(cJson).setMimeType(ContentService.MimeType.JSON);
  }
  var c = cfg();
  var d = {
    ok: 1,
    config: { whatsapp: c.whatsapp, faixa: c.faixa || '' },
    categorias: categorias().filter(function (x) { return x.ativa; }).map(function (x) { return x.nome; }),
    produtos: produtos().filter(function (p) { return p.ativo; })
  };
  var txt = JSON.stringify(d);
  try { cache.put('loja_publica', txt, 600); } catch (x) {} // Cache de 10 min, limpo automaticamente ao salvar
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Ações ---------- */
var PUBLICAS = { login: 1, pedido: 1, buscarCliente: 1, validarCupom: 1 };
var READ_ONLY = { login: 1, tudo: 1, pedidos: 1, buscarCliente: 1, validarCupom: 1 };

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
    if (READ_ONLY[d.a]) {
      return out(f(d));
    }
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
  var user = CacheService.getScriptCache().get('t_' + t); if (!user) return null;
  var u = rd('Usuarios').filter(function (r) { return r.usuario === user && bool(r.ativo); })[0];
  return u || null;
}
function login(d) {
  var cache = CacheService.getScriptCache(), usuario = String(d.usuario || '').trim().toLowerCase(), chave = 'f_' + usuario;
  if (Number(cache.get(chave) || 0) >= MAX_TENT) return { ok: 0, erro: 'Muitas tentativas. Aguarde 15 minutos e tente de novo.' };
  var u = rd('Usuarios').filter(function (r) { return String(r.usuario).toLowerCase() === usuario && bool(r.ativo); })[0];
  if (!u || u.senha_hash !== hash(usuario, String(d.senha || ''))) {
    cache.put(chave, String(Number(cache.get(chave) || 0) + 1), BLOQUEIO_S);
    return { ok: 0, erro: 'Usuário ou senha incorretos.' };
  }
  cache.remove(chave);
  var t = Utilities.getUuid() + Utilities.getUuid();
  cache.put('t_' + t, u.usuario, SESSAO_S);
  return { ok: 1, token: t, nome: u.nome, trocar: bool(u.trocar_senha) };
}
function senha(d) {
  var u = d.u, nova = String(d.nova || '');
  if (u.senha_hash !== hash(u.usuario, String(d.atual || ''))) return { ok: 0, erro: 'Senha atual incorreta.' };
  if (nova.length < 8) return { ok: 0, erro: 'A nova senha precisa ter pelo menos 8 caracteres.' };
  if (nova === d.atual) return { ok: 0, erro: 'Escolha uma senha diferente da atual.' };
  var s = SS.getSheetByName('Usuarios');
  s.getRange(u._row, 3).setValue(hash(u.usuario, nova));
  s.getRange(u._row, 5).setValue(false);
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
    pedidos: pedidosList()
  };
}
function linhaLivre(s) {
  var v = s.getRange(1, 1, s.getMaxRows(), 1).getValues();
  for (var i = 1; i < v.length; i++) if (!v[i][0]) return i + 1;
  s.insertRowAfter(s.getMaxRows()); return s.getMaxRows();
}

/**
 * Move arquivo de imagem do Google Drive para a lixeira para liberar espaço.
 */
function excluirArquivoDrive(url) {
  if (!url || typeof url !== 'string') return;
  try {
    var m = /[?&]id=([a-zA-Z0-9_-]+)/.exec(url) || /\/d\/([a-zA-Z0-9_-]+)/.exec(url);
    if (m && m[1]) {
      var file = DriveApp.getFileById(m[1]);
      if (file) {
        file.setTrashed(true);
      }
    }
  } catch (e) {
    Logger.log('Erro ao mover arquivo do Drive para a lixeira: ' + e);
  }
}

function salvarProduto(d) {
  var p = d.p || {}, s = SS.getSheetByName('Produtos');
  var nome = String(p.nome || '').trim(), preco = Number(p.preco), est = parseInt(p.estoque, 10);
  var categoria = String(p.categoria || '').trim();
  if (!nome) return { ok: 0, erro: 'Informe o nome do produto.' };
  if (!categoria) return { ok: 0, erro: 'Selecione a categoria do produto.' };
  var cats = categorias();
  var catValida = cats.filter(function (c) { return String(c.nome).trim().toLowerCase() === categoria.toLowerCase(); })[0];
  if (!catValida) return { ok: 0, erro: 'Escolha uma categoria válida.' };
  categoria = catValida.nome;
  if (!(preco >= 0)) return { ok: 0, erro: 'Preço inválido.' };
  if (!(est >= 0)) return { ok: 0, erro: 'Estoque inválido.' };
  var img = String(p.imagem || ''); if (img && !/^https:\/\//.test(img)) return { ok: 0, erro: 'Imagem inválida.' };
  var id = p.id || 'p' + Date.now().toString(36);
  var ex = rd('Produtos').filter(function (r) { return String(r.id) === id; })[0];

  // Trata variações (múltiplas imagens com preço e estoque)
  var varsJson = '';
  var novasImagens = [];
  if (img) novasImagens.push(img);
  if (p.variacoes) {
    var listaV = typeof p.variacoes === 'string' ? JSON.parse(p.variacoes) : p.variacoes;
    if (Array.isArray(listaV) && listaV.length > 0) {
      varsJson = JSON.stringify(listaV);
      listaV.forEach(function (v) {
        if (v && v.imagem) novasImagens.push(v.imagem);
      });
      if (!img && listaV[0].imagem) img = listaV[0].imagem;
    }
  }

  // Se fotos anteriores foram removidas, apaga do Google Drive para economizar espaço
  if (ex) {
    var fotosAntigas = [];
    if (ex.imagem) fotosAntigas.push(ex.imagem);
    if (ex.variacoes) {
      try {
        var vAnt = typeof ex.variacoes === 'string' ? JSON.parse(ex.variacoes) : ex.variacoes;
        if (Array.isArray(vAnt)) {
          vAnt.forEach(function (v) { if (v && v.imagem) fotosAntigas.push(v.imagem); });
        }
      } catch (e) {}
    }
    fotosAntigas.forEach(function (f) {
      if (novasImagens.indexOf(f) === -1) {
        excluirArquivoDrive(f);
      }
    });
  }

  // Garante cabeçalho com a coluna 'variacoes'
  var h = s.getRange(1, 1, 1, Math.max(s.getLastColumn(), PROD.length)).getValues()[0];
  var alterouCabecalho = false;
  for (var ci = 0; ci < PROD.length; ci++) {
    if (h[ci] !== PROD[ci]) { alterouCabecalho = true; break; }
  }
  if (alterouCabecalho) {
    s.getRange(1, 1, 1, PROD.length).setValues([PROD]);
  }

  var linha = ex ? ex._row : linhaLivre(s);
  var vals = [id, nome, categoria, p.descricao, p.descricao_completa, preco, est, p.etiqueta, !!p.ativo, img, p.cores, p.tamanhos, varsJson]
    .map(function (v) { return safe(typeof v === 'string' ? v.trim() : v); });
  s.getRange(linha, 1, 1, PROD.length).setValues([vals]);
  limparCache();
  return { ok: 1, id: id };
}
function excluirProduto(d) {
  var r = rd('Produtos').filter(function (x) { return String(x.id) === d.id; })[0];
  if (!r) return { ok: 0, erro: 'Produto não encontrado.' };
  // Apaga todas as fotos do produto no Google Drive para economizar espaço
  if (r.imagem) {
    excluirArquivoDrive(r.imagem);
  }
  if (r.variacoes) {
    try {
      var vars = typeof r.variacoes === 'string' ? JSON.parse(r.variacoes) : r.variacoes;
      if (Array.isArray(vars)) {
        vars.forEach(function (v) { if (v && v.imagem) excluirArquivoDrive(v.imagem); });
      }
    } catch (e) {}
  }
  SS.getSheetByName('Produtos').getRange(r._row, 1, 1, PROD.length).clearContent();
  limparCache();
  return { ok: 1 };
}
function salvarCategoria(d) {
  var nome = String(d.nome || '').trim(), antigo = d.antigo, s = SS.getSheetByName('Categorias');
  if (!nome) return { ok: 0, erro: 'Informe o nome da categoria.' };
  var todas = rd('Categorias');
  if (todas.some(function (c) { return c.nome === nome && c.nome !== antigo; })) return { ok: 0, erro: 'Já existe uma categoria com esse nome.' };
  var ex = todas.filter(function (c) { return c.nome === antigo; })[0];
  var linha = ex ? ex._row : linhaLivre(s);
  s.getRange(linha, 1, 1, 3).setValues([[safe(nome), Number(d.ordem) || 0, !!d.ativa]]);
  if (ex && antigo !== nome) {
    var ps = SS.getSheetByName('Produtos');
    rd('Produtos').forEach(function (p) { if (p.categoria === antigo) ps.getRange(p._row, 3).setValue(nome); });
  }
  return { ok: 1 };
}
function excluirCategoria(d) {
  if (produtos().some(function (p) { return p.categoria === d.nome; })) return { ok: 0, erro: 'Há produtos nessa categoria. Mude os produtos de categoria ou desative-a.' };
  var c = rd('Categorias').filter(function (x) { return x.nome === d.nome; })[0];
  if (!c) return { ok: 0, erro: 'Categoria não encontrada.' };
  SS.getSheetByName('Categorias').getRange(c._row, 1, 1, 3).clearContent();
  return { ok: 1 };
}
function salvarUsuario(d) {
  var s = SS.getSheetByName('Usuarios'), usuario = String(d.usuario || '').trim().toLowerCase(), nome = String(d.nome || '').trim();
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
  if (nova) { s.getRange(linha, 3).setValue(hash(usuario, nova)); s.getRange(linha, 5).setValue(true); } // senha temporária: troca no próximo acesso
  return { ok: 1 };
}
function salvarConfig(d) {
  var s = SS.getSheetByName('Configuracoes'), zap = String(d.whatsapp || '').replace(/\D/g, '');
  if (zap.length < 12) return { ok: 0, erro: 'WhatsApp: use 55 + DDD + número, só dígitos.' };
  rd('Configuracoes').forEach(function (r) {
    if (r.chave === 'whatsapp') { var c = s.getRange(r._row, 2); c.setNumberFormat('@'); c.setValue(zap); }
    if (r.chave === 'faixa') s.getRange(r._row, 2).setValue(safe(String(d.faixa || '').trim()));
  });
  limparCache();
  return { ok: 1 };
}
function imagem(d) {
  var m = /^data:(image\/(jpeg|png|webp));base64,(.+)$/.exec(String(d.dados || ''));
  if (!m) return { ok: 0, erro: 'Formato de imagem não aceito (use JPG, PNG ou WebP).' };
  var bytes = Utilities.base64Decode(m[3]);
  if (bytes.length > 3 * 1024 * 1024) return { ok: 0, erro: 'Imagem acima de 3 MB.' };
  var it = DriveApp.getFoldersByName('TF Life - Imagens'), pasta = it.hasNext() ? it.next() : DriveApp.createFolder('TF Life - Imagens');
  var f = pasta.createFile(Utilities.newBlob(bytes, m[1], 'produto-' + Date.now() + '.' + m[2].split('/')[1]));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: 1, url: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w900' };
}

/* ---------- Gestão de Clientes e Pedidos ---------- */
function getOuCriarAba(nome, cabecalhos) {
  var s = SS.getSheetByName(nome);
  if (!s) {
    s = SS.insertSheet(nome);
    s.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    s.getRange(1, 1, 1, cabecalhos.length).setFontWeight('bold');
    try { s.setFrozenRows(1); } catch (x) {}
  }
  return s;
}

function getAbaClientes() {
  return SS.getSheetByName('Clientes') || SS.getSheetByName('Usuarios_Clientes') || SS.getSheetByName('Clientes_Loja');
}

function garantirAbaClientes() {
  var s = getAbaClientes();
  if (!s) {
    s = getOuCriarAba('Clientes', ['telefone', 'nome', 'endereco', 'data_cadastro', 'ultimo_pedido', 'total_pedidos', 'dispositivos']);
    try { s.getRange('A:A').setNumberFormat('@'); } catch (x) {}
    try { s.getRange('G:G').setNumberFormat('@'); } catch (x) {}
  } else {
    try {
      var lastCol = Math.max(s.getLastColumn(), 1);
      var headers = s.getRange(1, 1, 1, lastCol).getValues()[0].map(function (x) { return String(x || '').trim().toLowerCase(); });
      if (headers.indexOf('dispositivos') === -1) {
        var colDisp = headers.length + 1;
        s.getRange(1, colDisp).setValue('dispositivos').setFontWeight('bold');
        try { s.getRange(1, colDisp, s.getMaxRows(), 1).setNumberFormat('@'); } catch (x) {}
      }
    } catch (eH) {
      Logger.log('Erro ao garantir coluna dispositivos: ' + eH);
    }
  }
  return s;
}

function garantirAbaPedidos() {
  var s = SS.getSheetByName('Pedidos');
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

function pedidosList() {
  var s = SS.getSheetByName('Pedidos');
  if (!s) return [];
  return rd('Pedidos').filter(function (r) { return r.id_pedido; }).map(function (p) {
    return {
      id: String(p.id_pedido),
      data_hora: formatarDataHoraSheet(p.data_hora),
      telefone: String(p.telefone || ''),
      nome: String(p.nome || ''),
      endereco: String(p.endereco || ''),
      forma_pagamento: String(p.forma_pagamento || ''),
      itens: String(p.itens || ''),
      total: Number(p.total) || 0,
      status: String(p.status || 'Recebido')
    };
  }).reverse();
}

function pedidosListAction() {
  return { ok: 1, pedidos: pedidosList() };
}

function atualizarStatusPedido(d) {
  var id = String(d.id || '').trim();
  var novoStatus = String(d.status || '').trim();
  if (!id || !novoStatus) return { ok: 0, erro: 'Informe o ID e o novo status.' };
  var s = SS.getSheetByName('Pedidos');
  if (!s) return { ok: 0, erro: 'Aba Pedidos não encontrada.' };
  var ex = rd('Pedidos').filter(function (r) { return String(r.id_pedido) === id; })[0];
  if (!ex) return { ok: 0, erro: 'Pedido não encontrado.' };
  s.getRange(ex._row, 9).setValue(safe(novoStatus));
  return { ok: 1 };
}

function excluirPedido(d) {
  var id = String(d.id || '').trim();
  if (!id) return { ok: 0, erro: 'ID do pedido não informado.' };
  var s = SS.getSheetByName('Pedidos');
  if (!s) return { ok: 0, erro: 'Aba Pedidos não encontrada.' };
  var ex = rd('Pedidos').filter(function (r) { return String(r.id_pedido) === id; })[0];
  if (!ex) return { ok: 0, erro: 'Pedido não encontrado.' };
  s.deleteRow(ex._row);
  return { ok: 1 };
}

function buscarCliente(d) {
  var tel = String(d.telefone || d.tel || '').replace(/\D/g, '');
  var dev = String(d.device_id || d.dispositivo || d.dev || '').trim();
  if (tel.length < 8 && !dev) return { ok: 1, encontrado: false };
  var aba = getAbaClientes();
  if (!aba) return { ok: 1, encontrado: false };

  var lista = rd(aba.getName());
  var achou = null;

  // 1. Prioridade: busca por dispositivo cadastrado (se informado)
  if (dev) {
    achou = lista.filter(function (r) {
      var disps = String(r.dispositivos || '').trim();
      if (!disps) return false;
      var arr = disps.split(/[,;\s]+/).map(function (s) { return s.trim(); }).filter(Boolean);
      return arr.indexOf(dev) !== -1;
    })[0];
  }

  // 2. Se não encontrou por dispositivo (ou se veio telefone direto), busca por telefone
  if (!achou && tel.length >= 8) {
    achou = lista.filter(function (r) {
      var rTel = String(r.telefone || '').replace(/\D/g, '');
      if (!rTel) return false;
      return rTel === tel ||
        (rTel.length >= 10 && tel.length >= 10 && (
          rTel.slice(-10) === tel.slice(-10) ||
          rTel.slice(-11) === tel.slice(-11)
        ));
    })[0];
  }

  if (achou && (achou.nome || achou.telefone)) {
    return {
      ok: 1,
      encontrado: true,
      telefone: String(achou.telefone || ''),
      nome: String(achou.nome || ''),
      endereco: String(achou.endereco || '')
    };
  }
  return { ok: 1, encontrado: false };
}

function upsertCliente(telefone, nome, endereco, agora, dispositivo) {
  if (!telefone) return;
  var s = garantirAbaClientes();
  var telLimpo = String(telefone).replace(/\D/g, '');
  var hRange = s.getRange(1, 1, 1, Math.max(s.getLastColumn(), 1));
  var headers = hRange.getValues()[0].map(function (x) { return String(x || '').trim().toLowerCase(); });

  var colTel = headers.indexOf('telefone') + 1 || 1;
  var colNome = headers.indexOf('nome') + 1 || 2;
  var colEnd = headers.indexOf('endereco') + 1 || 3;
  var colDataCad = headers.indexOf('data_cadastro') + 1 || 4;
  var colUltPed = headers.indexOf('ultimo_pedido') + 1 || 5;
  var colTotPed = headers.indexOf('total_pedidos') + 1 || 6;
  var colDisp = headers.indexOf('dispositivos') + 1;
  if (!colDisp) {
    colDisp = headers.length + 1;
    s.getRange(1, colDisp).setValue('dispositivos').setFontWeight('bold');
    try { s.getRange(1, colDisp, s.getMaxRows(), 1).setNumberFormat('@'); } catch (x) {}
  }

  var lista = rd(s.getName());
  var ex = lista.filter(function (r) {
    var rTel = String(r.telefone || '').replace(/\D/g, '');
    return rTel === telLimpo ||
      (rTel.length >= 10 && telLimpo.length >= 10 && (
        rTel.slice(-10) === telLimpo.slice(-10) ||
        rTel.slice(-11) === telLimpo.slice(-11)
      ));
  })[0];

  var devLimpo = String(dispositivo || '').trim();

  if (ex) {
    var nLinha = ex._row;
    if (nome) s.getRange(nLinha, colNome).setValue(safe(nome));
    if (endereco) s.getRange(nLinha, colEnd).setValue(safe(endereco));
    s.getRange(nLinha, colUltPed).setValue(agora);
    var qtd = (parseInt(ex.total_pedidos, 10) || 0) + 1;
    s.getRange(nLinha, colTotPed).setValue(qtd);

    if (devLimpo) {
      var dispsAtuais = String(ex.dispositivos || '').trim();
      var arr = dispsAtuais ? dispsAtuais.split(/[,;\s]+/).map(function (s) { return s.trim(); }).filter(Boolean) : [];
      if (arr.indexOf(devLimpo) === -1) {
        arr.push(devLimpo);
        s.getRange(nLinha, colDisp).setValue(arr.join(', '));
      }
    }
  } else {
    var novaLinha = linhaLivre(s);
    var cTel = s.getRange(novaLinha, colTel);
    try { cTel.setNumberFormat('@'); } catch (x) {}
    cTel.setValue(telLimpo);
    s.getRange(novaLinha, colNome).setValue(safe(nome || ''));
    s.getRange(novaLinha, colEnd).setValue(safe(endereco || ''));
    s.getRange(novaLinha, colDataCad).setValue(agora);
    s.getRange(novaLinha, colUltPed).setValue(agora);
    s.getRange(novaLinha, colTotPed).setValue(1);
    if (devLimpo) {
      s.getRange(novaLinha, colDisp).setValue(devLimpo);
    }
  }
}

/* ---------- Gestão de Cupons de Desconto ---------- */
function garantirAbaCupons() {
  var s = SS.getSheetByName('Cupons');
  if (!s) {
    s = getOuCriarAba('Cupons', ['codigo', 'tipo', 'valor', 'limite_usos', 'usos_atuais', 'validade', 'ativo']);
    try { s.getRange('A:A').setNumberFormat('@'); } catch (x) {}
    try { s.getRange('F:F').setNumberFormat('@'); } catch (x) {}
  }
  return s;
}

function formatarDataCupom(val) {
  if (!val) return '';
  if (typeof val === 'object' && val.getTime) {
    return Utilities.formatDate(val, 'America/Sao_Paulo', 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var partes = s.split(/[-/]/);
  if (partes.length === 3) {
    if (partes[0].length === 4) return partes[0] + '-' + ('0' + partes[1]).slice(-2) + '-' + ('0' + partes[2]).slice(-2);
    return partes[2] + '-' + ('0' + partes[1]).slice(-2) + '-' + ('0' + partes[0]).slice(-2);
  }
  return s;
}

function couponExpirou(valStr) {
  if (!valStr) return false;
  var dVal = null;
  if (typeof valStr === 'object' && valStr.getTime) {
    dVal = new Date(valStr.getFullYear(), valStr.getMonth(), valStr.getDate(), 23, 59, 59);
  } else if (typeof valStr === 'string') {
    var partes = valStr.trim().split(/[-/]/);
    if (partes.length === 3) {
      if (partes[0].length === 4) {
        dVal = new Date(parseInt(partes[0], 10), parseInt(partes[1], 10) - 1, parseInt(partes[2], 10), 23, 59, 59);
      } else {
        dVal = new Date(parseInt(partes[2], 10), parseInt(partes[1], 10) - 1, parseInt(partes[0], 10), 23, 59, 59);
      }
    }
  }
  if (!dVal || isNaN(dVal.getTime())) return false;
  return (new Date()).getTime() > dVal.getTime();
}

function cupons() {
  var s = SS.getSheetByName('Cupons');
  if (!s) return [];
  return rd('Cupons').filter(function (c) { return c.codigo; })
    .map(function (c) {
      return {
        codigo: String(c.codigo).toUpperCase(),
        tipo: c.tipo === 'porcentagem' ? 'porcentagem' : 'fixo',
        valor: Number(c.valor) || 0,
        limite_usos: parseInt(c.limite_usos, 10) || 0,
        usos_atuais: parseInt(c.usos_atuais, 10) || 0,
        validade: formatarDataCupom(c.validade),
        ativo: bool(c.ativo)
      };
    });
}

function validarCupom(d) {
  var cod = String(d.codigo || d.cupom || '').trim().toUpperCase();
  if (!cod) return { ok: 1, valido: false, erro: 'Informe o código do cupom.' };
  var aba = SS.getSheetByName('Cupons');
  if (!aba) return { ok: 1, valido: false, erro: 'Cupom desativado ou não existe.' };
  var lista = rd('Cupons');
  var c = lista.filter(function (r) { return String(r.codigo || '').trim().toUpperCase() === cod; })[0];
  if (!c) return { ok: 1, valido: false, erro: 'Cupom desativado ou não existe.' };
  if (!bool(c.ativo)) return { ok: 1, valido: false, erro: 'Cupom desativado ou não existe.' };

  var lim = parseInt(c.limite_usos, 10) || 0;
  var usos = parseInt(c.usos_atuais, 10) || 0;
  if (lim > 0 && usos >= lim) {
    return { ok: 1, valido: false, erro: 'Este cupom já atingiu o limite de usos e está esgotado.' };
  }

  if (c.validade && couponExpirou(c.validade)) {
    return { ok: 1, valido: false, erro: 'Este cupom expirou e está indisponível.' };
  }

  return {
    ok: 1,
    valido: true,
    codigo: String(c.codigo).toUpperCase(),
    tipo: c.tipo === 'porcentagem' ? 'porcentagem' : 'fixo',
    valor: Number(c.valor) || 0
  };
}

function salvarCupom(d) {
  var c = d.c || d;
  var s = garantirAbaCupons();
  var codigo = String(c.codigo || '').trim().toUpperCase().replace(/[^A-Z0-9_-]/g, '');
  if (codigo.length < 2) return { ok: 0, erro: 'Código do cupom inválido. Use ao menos 2 letras ou números.' };
  var tipo = c.tipo === 'porcentagem' ? 'porcentagem' : 'fixo';
  var valor = Number(c.valor);
  if (!(valor > 0)) return { ok: 0, erro: 'Informe um valor de desconto válido maior que zero.' };
  if (tipo === 'porcentagem' && valor > 100) return { ok: 0, erro: 'Desconto em porcentagem não pode ultrapassar 100%.' };

  var limite = parseInt(c.limite_usos, 10) || 0;
  var validade = c.validade ? formatarDataCupom(c.validade) : '';
  var ativo = bool(c.ativo);

  var lista = rd('Cupons');
  var ex = lista.filter(function (r) { return String(r.codigo || '').trim().toUpperCase() === codigo; })[0];
  var linha = ex ? ex._row : linhaLivre(s);
  var usosAtuais = ex ? (parseInt(ex.usos_atuais, 10) || 0) : 0;

  var celCod = s.getRange(linha, 1);
  try { celCod.setNumberFormat('@'); } catch (x) {}
  celCod.setValue(codigo);

  var celVal = s.getRange(linha, 6);
  try { celVal.setNumberFormat('@'); } catch (x) {}

  s.getRange(linha, 2, 1, 6).setValues([[
    tipo,
    valor,
    limite,
    usosAtuais,
    validade,
    ativo
  ]]);

  return { ok: 1, codigo: codigo };
}

function excluirCupom(d) {
  var codigo = String(d.codigo || '').trim().toUpperCase();
  var s = SS.getSheetByName('Cupons');
  if (!s) return { ok: 0, erro: 'Aba de cupons não encontrada.' };
  var lista = rd('Cupons');
  var ex = lista.filter(function (r) { return String(r.codigo || '').trim().toUpperCase() === codigo; })[0];
  if (!ex) return { ok: 0, erro: 'Cupom não encontrado.' };
  s.getRange(ex._row, 1, 1, 7).clearContent();
  return { ok: 1 };
}

/* ---------- Pedido (baixa de estoque, cupom, registro em Pedidos e Clientes) ---------- */
function pedido(d) {
  var itens = (d.itens || []).slice(0, 40);
  var sProd = SS.getSheetByName('Produtos');
  var ps = rd('Produtos');
  var mapa = {};

  itens.forEach(function (i) {
    var q = parseInt(i.q, 10);
    if (q > 0 && q <= 50) mapa[i.id] = (mapa[i.id] || 0) + q;
  });

  var ids = Object.keys(mapa), falta = [];
  if (!ids.length) return { ok: 0, erro: 'Carrinho vazio.' };

  ids.forEach(function (id) {
    var p = ps.filter(function (r) { return String(r.id) === id; })[0];
    if (!p || !bool(p.ativo)) falta.push('produto indisponível');
    else if ((parseInt(p.estoque, 10) || 0) < mapa[id]) falta.push(p.nome + ' (restam ' + (parseInt(p.estoque, 10) || 0) + ')');
  });

  if (falta.length) return { ok: 0, erro: 'Sem estoque suficiente: ' + falta.join(', ') + '.', atualizar: 1 };

  // Dá baixa no estoque
  ids.forEach(function (id) {
    var p = ps.filter(function (r) { return String(r.id) === id; })[0];
    var novoEst = Math.max(0, (parseInt(p.estoque, 10) || 0) - mapa[id]);
    sProd.getRange(p._row, 7).setValue(novoEst);
    if (p.variacoes) {
      try {
        var vars = typeof p.variacoes === 'string' ? JSON.parse(p.variacoes) : p.variacoes;
        if (Array.isArray(vars)) {
          var mudouVars = false;
          itens.filter(function (it) { return it.id === id && (it.varId || it.varNome); }).forEach(function (it) {
            vars.forEach(function (v) {
              if ((it.varId && v.id === it.varId) || (it.varNome && v.nome === it.varNome)) {
                var qtdBaixa = parseInt(it.q, 10) || 0;
                v.estoque = Math.max(0, (parseInt(v.estoque, 10) || 0) - qtdBaixa);
                if (v.tamanhos && it.tam && v.tamanhos[it.tam] !== undefined) {
                  v.tamanhos[it.tam] = Math.max(0, (parseInt(v.tamanhos[it.tam], 10) || 0) - qtdBaixa);
                }
                mudouVars = true;
              }
            });
          });
          if (mudouVars) {
            sProd.getRange(p._row, 13).setValue(JSON.stringify(vars));
          }
        }
      } catch (x) {}
    }
  });

  // Monta resumo dos itens e valor total
  var resumoItens = [];
  var totalCalculado = 0;
  itens.forEach(function (it) {
    var p = ps.filter(function (r) { return String(r.id) === it.id; })[0];
    var precoUnit = Number(it.preco);
    if (!precoUnit || isNaN(precoUnit)) precoUnit = p ? (Number(p.preco) || 0) : 0;
    var qtd = parseInt(it.q, 10) || 1;
    var sub = qtd * precoUnit;
    totalCalculado += sub;
    var det = [it.varNome ? 'Opção: ' + it.varNome : '', it.cor, it.tam].filter(Boolean).join('/');
    var nomeProd = p ? p.nome : (it.nome || it.id);
    resumoItens.push(qtd + 'x ' + nomeProd + (det ? ' (' + det + ')' : '') + ' [R$ ' + sub.toFixed(2) + ']');
  });

  // Validação e aplicação do cupom
  var cupomCod = String(d.cupom || '').trim().toUpperCase();
  var descontoCupom = 0;
  if (cupomCod) {
    var resV = validarCupom({ codigo: cupomCod });
    if (resV.ok && resV.valido) {
      if (resV.tipo === 'porcentagem') {
        descontoCupom = Number((totalCalculado * (resV.valor / 100)).toFixed(2));
      } else {
        descontoCupom = Math.min(totalCalculado, Number(resV.valor));
      }
      try {
        var sCup = SS.getSheetByName('Cupons');
        var exCup = rd('Cupons').filter(function (r) { return String(r.codigo || '').trim().toUpperCase() === cupomCod; })[0];
        if (sCup && exCup) {
          sCup.getRange(exCup._row, 5).setValue((parseInt(exCup.usos_atuais, 10) || 0) + 1);
        }
      } catch (xCup) {
        Logger.log('Erro ao incrementar cupom: ' + xCup);
      }
    }
  }

  var totalFinal = Math.max(0, Number((totalCalculado - descontoCupom).toFixed(2)));

  var tz = 'America/Sao_Paulo';
  try { tz = Session.getScriptTimeZone() || tz; } catch (x) {}
  var agoraStr = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss');
  var telLimpo = String(d.telefone || '').replace(/\D/g, '');
  var nomeCliente = String(d.nome || '').trim();
  var endCliente = String(d.endereco || '').trim();
  var formaPag = String(d.pagamento || d.forma_pagamento || '').trim();
  var dispCliente = String(d.dispositivo || d.device_id || '').trim();
  var idPedido = 'PED-' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd-HHmmss') + '-' + Math.floor(100 + Math.random() * 900);

  var resumoGravado = resumoItens.join('; ');
  if (cupomCod && descontoCupom > 0) {
    resumoGravado += ' [CUPOM: ' + cupomCod + ' (-R$ ' + descontoCupom.toFixed(2) + ')]';
  }

  // 1. Grava o pedido na aba Pedidos
  try {
    var sPed = garantirAbaPedidos();
    var linPed = linhaLivre(sPed);
    var celTel = sPed.getRange(linPed, 3);
    try { celTel.setNumberFormat('@'); } catch (x) {}
    celTel.setValue(telLimpo);
    sPed.getRange(linPed, 1, 1, 2).setValues([[idPedido, agoraStr]]);
    sPed.getRange(linPed, 4, 1, 6).setValues([[
      safe(nomeCliente),
      safe(endCliente),
      safe(formaPag),
      safe(resumoGravado),
      Number(totalFinal.toFixed(2)),
      'Recebido'
    ]]);
  } catch (xPed) {
    Logger.log('Erro ao salvar pedido: ' + xPed);
  }

  // 2. Grava ou atualiza o cliente na aba Clientes (usuários da loja)
  if (telLimpo) {
    try {
      upsertCliente(telLimpo, nomeCliente, endCliente, agoraStr, dispCliente);
    } catch (xCli) {
      Logger.log('Erro ao salvar cliente: ' + xCli);
    }
  }

  limparCache();
  return { ok: 1, id_pedido: idPedido, desconto: descontoCupom, total: totalFinal };
}

var ACOES = {
  login: login,
  senha: senha,
  tudo: tudo,
  salvarProduto: salvarProduto,
  excluirProduto: excluirProduto,
  salvarCategoria: salvarCategoria,
  excluirCategoria: excluirCategoria,
  salvarUsuario: salvarUsuario,
  salvarConfig: salvarConfig,
  imagem: imagem,
  pedido: pedido,
  pedidos: pedidosListAction,
  atualizarStatusPedido: atualizarStatusPedido,
  excluirPedido: excluirPedido,
  buscarCliente: buscarCliente,
  validarCupom: validarCupom,
  salvarCupom: salvarCupom,
  excluirCupom: excluirCupom
};
