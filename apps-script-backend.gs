// ============================================================
// MORFE ERP - Google Apps Script (Backend API)
// Cole este codigo no Google Apps Script e faca deploy
// ============================================================

var SHEET_NAME = 'Morfe ERP';

// Tabelas do sistema
var TABELAS = [
  'config', 'produtos', 'clientes', 'fornecedores', 'materias',
  'vendas', 'compras', 'producao', 'pedidos', 'kits', 'fichas',
  'custos_fixos', 'areceber', 'apagar', 'consignados'
];

// ============================================================
// AUTENTICACAO - so login com Google. A senha fixa foi removida (era
// exposta em texto puro no index.html, que e publico) depois de
// confirmar que o login com Google funciona de verdade.
// ============================================================
function validarAcesso(params) {
  if (params.token) {
    var resultado = validarTokenGoogle(params.token);
    if (resultado.ok) return { ok: true, metodo: 'google', email: resultado.email };
    return { ok: false, erro: resultado.erro };
  }
  return { ok: false, erro: 'Nao autenticado' };
}

// Valida um token de login do Google (ID token) chamando o endpoint oficial
// do Google, que confirma a assinatura do token e devolve os dados nele.
// So aceita se: o token for valido e nao estiver vencido, o "audience"
// (client ID) bater com o configurado, e o email bater com o autorizado.
//
// Para funcionar, configure em Extensoes > Apps Script > Configuracoes do
// projeto > Propriedades do script (Script Properties), duas chaves:
//   GOOGLE_CLIENT_ID   = o Client ID criado no Google Cloud Console
//   AUTHORIZED_EMAIL   = o unico email que pode acessar o sistema
function validarTokenGoogle(idToken) {
  var props = PropertiesService.getScriptProperties();
  var clientIdEsperado = props.getProperty('GOOGLE_CLIENT_ID');
  var emailAutorizado = props.getProperty('AUTHORIZED_EMAIL');

  if (!clientIdEsperado || !emailAutorizado) {
    return { ok: false, erro: 'Login com Google ainda nao foi configurado no backend (faltam as Script Properties GOOGLE_CLIENT_ID e AUTHORIZED_EMAIL)' };
  }

  // Login ja conferido ha pouco: responde na hora, sem perguntar de novo ao Google.
  // Guarda so um resumo (hash) do token, nunca o token em si.
  var cache = CacheService.getScriptCache();
  var chaveCache = 'tok_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(idToken), Utilities.Charset.UTF_8));
  var emailGuardado = cache.get(chaveCache);
  if (emailGuardado && emailGuardado.toLowerCase() === String(emailAutorizado).toLowerCase()) {
    return { ok: true, email: emailGuardado };
  }

  try {
    var resposta = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true }
    );
    var dados = JSON.parse(resposta.getContentText());

    if (dados.error) return { ok: false, erro: 'Token invalido ou expirado, faca login novamente' };
    if (dados.aud !== clientIdEsperado) return { ok: false, erro: 'Token nao pertence a este aplicativo' };
    if (dados.email_verified !== 'true' && dados.email_verified !== true) return { ok: false, erro: 'Email do Google nao verificado' };
    if (String(dados.email).toLowerCase() !== String(emailAutorizado).toLowerCase()) return { ok: false, erro: 'Este email do Google nao tem acesso a este sistema' };

    // Lembra ate 1 minuto antes do token vencer (no maximo 50 min).
    var segundosRestantes = Math.floor(Number(dados.exp) - Date.now() / 1000) - 60;
    if (segundosRestantes > 30) cache.put(chaveCache, String(dados.email), Math.min(segundosRestantes, 3000));

    return { ok: true, email: dados.email };
  } catch (err) {
    return { ok: false, erro: 'Erro ao validar login do Google: ' + err.toString() };
  }
}

// GET: Ler dados
//
// Usa o mesmo LockService do doPost. Sem essa trava, uma leitura podia
// cair bem no meio de uma gravacao (salvarTabela faz clearContents() e
// SO DEPOIS setValues() — nao e uma operacao atomica), e ver a aba
// momentaneamente VAZIA no meio do caminho. Foi exatamente esse sintoma
// que fez o app ganhar aquela protecao de "tabela veio vazia, mantendo
// dados locais" no client — essa trava ataca a causa, nao so o sintoma.
function doGet(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return jsonResponse({ ok: false, erro: 'Sistema ocupado gravando outra coisa, tente de novo em alguns segundos' });
  }
  try {
    var params = e.parameter;

    var acesso = validarAcesso(params);
    if (!acesso.ok) {
      return jsonResponse({ ok: false, erro: acesso.erro || 'Nao autorizado' });
    }

    var acao = params.acao || 'ler_tudo';
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (acao === 'ler_tudo') {
      var dados = {};
      for (var i = 0; i < TABELAS.length; i++) {
        dados[TABELAS[i]] = lerTabela(ss, TABELAS[i]);
      }
      // Fica fora de TABELAS: so o Apps Script grava nela, o 'salvar_tudo' do app nunca a sobrescreve.
      dados.mp_pagamentos = lerTabela(ss, MP_ABA);
      return jsonResponse({ ok: true, dados: dados });
    }

    if (acao === 'ler') {
      var tabela = params.tabela;
      if (!tabela) return jsonResponse({ ok: false, erro: 'Tabela nao informada' });
      return jsonResponse({ ok: true, dados: lerTabela(ss, tabela) });
    }

    return jsonResponse({ ok: false, erro: 'Acao nao reconhecida' });

  } catch (err) {
    return jsonResponse({ ok: false, erro: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// POST: Salvar dados
//
// A trava (LockService) e a correcao do bug de estoque "somando em vez de
// subtrair": ajustarEstoqueLote (e salvarTabela em geral) fazem ler a aba
// inteira -> mexer em memoria -> regravar a aba inteira. Sem lock, dois
// doPost concorrentes (ex: dois pedidos salvos quase ao mesmo tempo, de
// dois aparelhos/abas diferentes, ou um Pedido e uma Compra) podiam ler o
// estoque ao mesmo tempo, cada um calcular seu resultado em cima do MESMO
// valor antigo, e o que gravasse por ultimo sobrescrevia o outro por
// inteiro — apagando uma baixa de estoque que tinha acabado de acontecer.
// Com a trava, cada doPost roda sozinho do inicio ao fim antes do proximo
// comecar, entao sempre le o valor mais atual de verdade.
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // espera ate 30s por outra gravacao em andamento
  } catch (err) {
    return jsonResponse({ ok: false, erro: 'Sistema ocupado gravando outra coisa, tente de novo em alguns segundos' });
  }

  try {
    var body = JSON.parse(e.postData.contents);

    var acesso = validarAcesso(body);
    if (!acesso.ok) {
      return jsonResponse({ ok: false, erro: acesso.erro || 'Nao autorizado' });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var acao = body.acao;
    var tabela = body.tabela;
    var dados = body.dados;
    var id = body.id;


    // Leitura de tudo por POST: o login vai no corpo da mensagem, nao no endereco.
    if (acao === 'ler_tudo') {
      var tudo = {};
      for (var j = 0; j < TABELAS.length; j++) tudo[TABELAS[j]] = lerTabela(ss, TABELAS[j]);
      tudo.mp_pagamentos = lerTabela(ss, MP_ABA);
      return jsonResponse({ ok: true, dados: tudo });
    }

    if (acao === 'salvar_tudo') {
      for (var i = 0; i < TABELAS.length; i++) {
        var t = TABELAS[i];
        if (dados[t] !== undefined) {
          salvarTabela(ss, t, dados[t]);
        }
      }
      return jsonResponse({ ok: true, msg: 'Sync completo realizado' });
    }

    if (acao === 'upsert') {
      if (!tabela || !dados) return jsonResponse({ ok: false, erro: 'Dados incompletos' });
      upsertRegistro(ss, tabela, dados);
      return jsonResponse({ ok: true, msg: 'Salvo com sucesso' });
    }

    if (acao === 'excluir') {
      if (!tabela || !id) return jsonResponse({ ok: false, erro: 'Dados incompletos' });
      excluirRegistro(ss, tabela, id);
      return jsonResponse({ ok: true, msg: 'Excluido com sucesso' });
    }

    if (acao === 'salvar_tabela') {
      if (!tabela || !dados) return jsonResponse({ ok: false, erro: 'Dados incompletos' });
      salvarTabela(ss, tabela, dados);
      return jsonResponse({ ok: true, msg: 'Tabela salva' });
    }

    // Ajuste ATOMICO de estoque: em vez do aparelho mandar o valor final pronto,
    // manda so a MUDANCA (delta) e o proprio Apps Script soma/subtrai em cima do
    // valor mais atual que ele tem na planilha. Isso evita que dois aparelhos
    // calculando ao mesmo tempo (um com dado desatualizado) apaguem a baixa um do outro.
    // (E agora, com o LockService acima, essa garantia e real de verdade — antes
    // o "mais atual" podia nao ser tao atual assim, por causa da concorrencia.)
    if (acao === 'ajustar_estoque') {
      if (!dados || !dados.itens) return jsonResponse({ ok: false, erro: 'Dados incompletos' });
      var resultado = ajustarEstoqueLote(ss, dados.itens);
      return jsonResponse({ ok: true, msg: 'Estoque ajustado', resultado: resultado });
    }

    if (acao === 'mp_gerar_pix') return jsonResponse({ ok: true, dados: mpGerarPix(dados || {}) });
    if (acao === 'mp_gerar_link') return jsonResponse({ ok: true, dados: mpGerarLink(dados || {}) });
    if (acao === 'mp_sincronizar') {
      mpBuscarPagamentos(ss, 3);
      return jsonResponse({ ok: true, dados: lerTabela(ss, MP_ABA) });
    }

    return jsonResponse({ ok: false, erro: 'Acao nao reconhecida' });

  } catch (err) {
    return jsonResponse({ ok: false, erro: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// Ajusta o estoque de varios insumos de uma vez, somando/subtraindo (delta) em
// cima do valor ATUAL da planilha (le a planilha uma vez, aplica todos os deltas,
// salva uma vez). itens = [{id: 'MP-001', delta: -5}, {id: 'MP-002', delta: 3}, ...]
function ajustarEstoqueLote(ss, itens) {
  var registros = lerTabela(ss, 'materias');
  var porId = {};
  for (var i = 0; i < registros.length; i++) { porId[registros[i].id] = registros[i]; }

  var resultado = [];
  for (var j = 0; j < itens.length; j++) {
    var it = itens[j];
    var reg = porId[it.id];
    if (!reg) { resultado.push({ id: it.id, ok: false }); continue; }
    reg.estoque = (Number(reg.estoque) || 0) + Number(it.delta || 0);
    resultado.push({ id: it.id, ok: true, novoEstoque: reg.estoque });
  }

  salvarTabela(ss, 'materias', registros);
  return resultado;
}

// Helpers
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, nome) {
  var sheet = ss.getSheetByName(nome);
  if (!sheet) {
    sheet = ss.insertSheet(nome);
  }
  return sheet;
}

function lerTabela(ss, tabela) {
  var sheet = ss.getSheetByName(tabela);
  if (!sheet) return [];

  var dados = sheet.getDataRange().getValues();
  if (dados.length < 2) return [];

  var headers = dados[0];
  var resultado = [];
  var fuso = ss.getSpreadsheetTimeZone();

  for (var r = 1; r < dados.length; r++) {
    var row = dados[r];
    if (row[0] === '') continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      var val = row[c];
      if (Object.prototype.toString.call(val) === '[object Date]') {
        // O Google Sheets as vezes converte texto tipo "2026-07-16" pra
        // data de verdade sozinho. Aqui a gente devolve sempre como texto
        // YYYY-MM-DD, pra bater com o formato que o app espera.
        val = Utilities.formatDate(val, fuso, 'yyyy-MM-dd');
      } else if (typeof val === 'string' && (val.charAt(0) === '[' || val.charAt(0) === '{')) {
        try { val = JSON.parse(val); } catch(e) {}
      }
      obj[headers[c]] = val;
    }
    resultado.push(obj);
  }
  return resultado;
}

function salvarTabela(ss, tabela, registros) {
  var sheet = getOrCreateSheet(ss, tabela);
  sheet.clearContents();

  if (!registros || registros.length === 0) return;

  var camposSet = {};
  var headers = [];
  for (var i = 0; i < registros.length; i++) {
    var keys = Object.keys(registros[i]);
    for (var k = 0; k < keys.length; k++) {
      if (!camposSet[keys[k]]) {
        camposSet[keys[k]] = true;
        headers.push(keys[k]);
      }
    }
  }

  var rows = [headers];

  for (var i = 0; i < registros.length; i++) {
    var row = [];
    for (var h = 0; h < headers.length; h++) {
      var val = registros[i][headers[h]];
      if (Array.isArray(val) || (typeof val === 'object' && val !== null)) {
        row.push(JSON.stringify(val));
      } else {
        row.push(val !== undefined ? val : '');
      }
    }
    rows.push(row);
  }

  var range = sheet.getRange(1, 1, rows.length, headers.length);
  // Forca formato de TEXTO em todas as celulas ANTES de escrever, pra
  // impedir o Google Sheets de "detectar" datas sozinho e converter
  // campos tipo "2026-07-16" pra um valor de data de verdade.
  range.setNumberFormat('@');
  range.setValues(rows);
}

function upsertRegistro(ss, tabela, registro) {
  var registros = lerTabela(ss, tabela);
  var idx = -1;
  for (var i = 0; i < registros.length; i++) {
    if (registros[i].id === registro.id) { idx = i; break; }
  }
  if (idx >= 0) {
    registros[idx] = registro;
  } else {
    registros.push(registro);
  }
  salvarTabela(ss, tabela, registros);
}

function excluirRegistro(ss, tabela, id) {
  var registros = lerTabela(ss, tabela);
  var novos = [];
  for (var i = 0; i < registros.length; i++) {
    if (registros[i].id !== id) novos.push(registros[i]);
  }
  salvarTabela(ss, tabela, novos);
}

// Config inicial - criar estrutura se nao existir
// RODE ESTA FUNCAO PRIMEIRO (uma vez so)
function criarEstrutura() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.rename(SHEET_NAME);

  for (var i = 0; i < TABELAS.length; i++) {
    var t = TABELAS[i];
    if (!ss.getSheetByName(t)) {
      ss.insertSheet(t);
      Logger.log('Aba criada: ' + t);
    }
  }

  var configExistente = lerTabela(ss, 'config');
  if (configExistente.length === 0) {
    var categoriasArr = ['Caneca','Azulejo','Camiseta','Almofada','Quadro','Chaveiro','Mouse Pad','Squeeze','Bolsa','Porta Retrato','Outros'];
    var configPadrao = [{
      id: 1,
      empresa: 'Morfe Personalizados',
      whatsapp: '5592994875346',
      instagram: '@morfepersonalizados',
      email: 'morfepersonalizados@gmail.com',
      cidade: 'Manaus - AM',
      meta_mensal: 5000,
      meta_anual: 60000,
      taxa_pix: 0,
      taxa_debito: 1.5,
      taxa_credito1: 2.5,
      margem_padrao: 40,
      reserva: 10,
      categorias: JSON.stringify(categoriasArr)
    }];
    salvarTabela(ss, 'config', configPadrao);
    Logger.log('Config padrao criada');
  }

  Logger.log('Estrutura criada com sucesso!');
}

// ============================================================
// MERCADO PAGO
// A chave fica SO aqui, nas Propriedades do script:
//   MP_ACCESS_TOKEN    = Access Token de producao (APP_USR-...)
//   MP_EMAIL_PAGADOR   = um email seu, DIFERENTE do email da conta
//                        Mercado Pago (o PIX exige um email de pagador;
//                        usado so quando o cliente nao tem email)
// Este codigo so LE pagamentos e CRIA cobrancas. Nao devolve dinheiro,
// nao transfere e nao saca.
// ============================================================
var MP_API = 'https://api.mercadopago.com';
var MP_ABA = 'mp_pagamentos';
var MP_MAX_LINHAS = 600;
var MP_ABA_COBRANCAS = 'mp_cobrancas';

function mpToken() {
  var t = PropertiesService.getScriptProperties().getProperty('MP_ACCESS_TOKEN');
  if (!t) throw new Error('Mercado Pago ainda nao configurado (falta MP_ACCESS_TOKEN nas Propriedades do script)');
  return t;
}

function mpChamar(metodo, caminho, corpo, chaveIdem) {
  var opcoes = {
    method: metodo,
    muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + mpToken() }
  };
  if (corpo) { opcoes.contentType = 'application/json'; opcoes.payload = JSON.stringify(corpo); }
  if (chaveIdem) opcoes.headers['X-Idempotency-Key'] = chaveIdem;
  var r = UrlFetchApp.fetch(MP_API + caminho, opcoes);
  var codigo = r.getResponseCode();
  var dados = {};
  try { dados = JSON.parse(r.getContentText()); } catch (e) {}
  if (codigo < 200 || codigo >= 300) {
    throw new Error('Mercado Pago respondeu ' + codigo + ': ' + (dados.message || dados.error || r.getContentText().slice(0, 200)));
  }
  return dados;
}

function mpDataIso(d) {
  return Utilities.formatDate(d, 'America/Manaus', "yyyy-MM-dd'T'HH:mm:ss.SSSXXX");
}

// Cobranca PIX com o numero do pedido (API de Orders): quando o cliente paga, o ERP sabe de qual pedido e.
// A cobranca fica anotada na aba mp_cobrancas para a busca conferir se foi paga.
function mpGerarPix(d) {
  var valor = Math.round(Number(d.valor) * 100) / 100;
  if (!(valor > 0)) throw new Error('Valor invalido');
  if (!d.pedido_id) throw new Error('Pedido nao informado');
  var email = d.email || PropertiesService.getScriptProperties().getProperty('MP_EMAIL_PAGADOR');
  if (!email) throw new Error('Falta MP_EMAIL_PAGADOR nas Propriedades do script');
  var vence = new Date(Date.now() + 3 * 24 * 3600 * 1000);
  var o = mpChamar('post', '/v1/orders', {
    type: 'online',
    processing_mode: 'automatic',
    total_amount: valor.toFixed(2),
    external_reference: String(d.pedido_id).slice(0, 64),
    config: { location: { state_id: 'BR-AM' } },
    transactions: { payments: [{ amount: valor.toFixed(2), payment_method: { id: 'pix', type: 'bank_transfer' }, expiration_time: 'P3D' }] },
    payer: { email: email }
  }, Utilities.getUuid());
  var pg = (o.transactions && o.transactions.payments && o.transactions.payments[0]) || {};
  var pm = pg.payment_method || {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cobs = lerTabela(ss, MP_ABA_COBRANCAS);
  cobs.unshift({ id: String(o.id), pedido: String(d.pedido_id), valor: valor, criado: mpDataIso(new Date()), status: o.status || '' });
  salvarTabela(ss, MP_ABA_COBRANCAS, cobs.slice(0, 300));
  return { id: String(o.id), valor: valor, qr_code: pm.qr_code, qr_code_base64: pm.qr_code_base64, ticket_url: pm.ticket_url, vence: mpDataIso(vence) };
}

var MP_STATUS_ORDER = { processed: 'approved', refunded: 'refunded', charged_back: 'charged_back', canceled: 'cancelled', failed: 'cancelled' };

// Confere as cobrancas PIX criadas nos ultimos 31 dias que ainda nao terminaram.
function mpConferirCobrancas(ss) {
  var cobs = lerTabela(ss, MP_ABA_COBRANCAS);
  var linhas = [], limite = Date.now() - 31 * 24 * 3600 * 1000, mudou = false;
  for (var i = 0; i < cobs.length; i++) {
    var c = cobs[i];
    if (['processed', 'canceled', 'failed', 'refunded'].indexOf(c.status) >= 0) continue;
    if (new Date(c.criado).getTime() < limite) continue;
    var o;
    try { o = mpChamar('get', '/v1/orders/' + encodeURIComponent(c.id)); } catch (e) { continue; }
    if (o.status !== c.status) { c.status = o.status; mudou = true; }
    var st = MP_STATUS_ORDER[o.status];
    if (!st) continue;
    var pg = (o.transactions && o.transactions.payments && o.transactions.payments[0]) || {};
    var bruto = o.last_updated_date || o.created_date, dt = new Date(bruto);
    // A order pode vir em UTC; a busca de pagamentos vem no horario de Manaus. Usa sempre Manaus.
    var quando = mpDataIso(bruto && !isNaN(dt.getTime()) ? dt : new Date());
    linhas.push({
      id: String(pg.reference_id || pg.id || o.id),
      data: String(quando).slice(0, 10),
      hora: String(quando).slice(11, 16),
      status: st,
      valor: Number(pg.amount || o.total_amount) || 0,
      liquido: 0,
      taxa: 0,
      metodo: 'pix',
      tipo: 'bank_transfer',
      operacao: 'order',
      referencia: o.external_reference || c.pedido,
      descricao: 'PIX ' + c.pedido,
      pagador: '',
      atualizado: String(quando),
      fonte: 'order'
    });
  }
  if (mudou) salvarTabela(ss, MP_ABA_COBRANCAS, cobs);
  return linhas;
}

// Link de pagamento (Checkout Pro). Com so_cartao, aceita so cartao (o PIX gratis e pela chave).
// referencia pode vir como "PEDIDO|valor do pedido" quando o cliente paga a taxa do cartao.
function mpGerarLink(d) {
  var valor = Math.round(Number(d.valor) * 100) / 100;
  if (!(valor > 0)) throw new Error('Valor invalido');
  if (!d.pedido_id) throw new Error('Pedido nao informado');
  var vence = new Date(Date.now() + 7 * 24 * 3600 * 1000);
  var pref = mpChamar('post', '/checkout/preferences', {
    items: [{ title: String(d.descricao || ('Pedido ' + d.pedido_id)).slice(0, 200), quantity: 1, unit_price: valor, currency_id: 'BRL' }],
    external_reference: String(d.referencia || d.pedido_id),
    payment_methods: d.so_cartao ? { excluded_payment_types: [{ id: 'ticket' }, { id: 'bank_transfer' }, { id: 'atm' }] } : undefined,
    expires: true,
    expiration_date_to: mpDataIso(vence)
  }, Utilities.getUuid());
  return { id: String(pref.id), valor: valor, link: pref.init_point, vence: mpDataIso(vence) };
}

// Numero da sua conta no Mercado Pago (guardado na primeira vez) para saber se o dinheiro entrou ou saiu.
function mpMeuId() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('MP_USER_ID');
  if (!id) {
    try { id = String(mpChamar('get', '/users/me').id || ''); } catch (e) { id = ''; }
    if (id) props.setProperty('MP_USER_ID', id);
  }
  return id;
}

function mpLinhaPagamento(p, meuId) {
  var taxa = 0;
  // So as taxas pagas por voce; juros de parcelamento pagos pelo cliente nao entram.
  (p.fee_details || []).forEach(function (f) { if (f.fee_payer !== 'payer') taxa += Number(f.amount) || 0; });
  var quando = p.date_approved || p.date_created || '';
  var pagador = p.payer || {};
  var nome = [pagador.first_name, pagador.last_name].filter(function (x) { return x; }).join(' ');
  return {
    id: String(p.id),
    data: String(quando).slice(0, 10),
    hora: String(quando).slice(11, 16),
    status: p.status || '',
    valor: Number(p.transaction_amount) || 0,
    liquido: Number(p.transaction_details && p.transaction_details.net_received_amount) || 0,
    taxa: Math.round(taxa * 100) / 100,
    metodo: p.payment_method_id || '',
    tipo: p.payment_type_id || '',
    operacao: p.operation_type || '',
    referencia: p.external_reference || '',
    descricao: p.description || '',
    pagador: nome || pagador.email || '',
    atualizado: p.date_last_updated || '',
    // 'saida' = voce pagou alguem (ex.: Uber, compra); 'entrada' = alguem te pagou
    sentido: meuId && p.collector_id != null ? (String(p.collector_id) === meuId ? 'entrada' : 'saida') : ''
  };
}

// Busca no Mercado Pago o que mudou nos ultimos dias e guarda na aba mp_pagamentos.
function mpBuscarPagamentos(ss, dias) {
  var novos = [], meuId = mpMeuId();
  for (var offset = 0; offset < 500; offset += 50) {
    var r = mpChamar('get', '/v1/payments/search?sort=date_last_updated&criteria=desc&range=date_last_updated' +
      '&begin_date=NOW-' + (dias || 3) + 'DAYS&end_date=NOW&limit=50&offset=' + offset);
    var lista = r.results || [];
    for (var i = 0; i < lista.length; i++) novos.push(mpLinhaPagamento(lista[i], meuId));
    if (lista.length < 50) break;
  }
  var cobrancas = mpConferirCobrancas(ss);
  var atuais = lerTabela(ss, MP_ABA);
  var porId = {};
  atuais.forEach(function (x) { porId[String(x.id)] = x; });
  // A busca de pagamentos tem as taxas; se ela trouxer o mesmo pagamento com o mesmo numero, vale a dela.
  // Se vier com numero diferente, o ERP percebe que e o mesmo PIX e lanca uma vez so.
  cobrancas.forEach(function (x) { if (!porId[x.id] || porId[x.id].fonte === 'order') porId[x.id] = x; });
  novos.forEach(function (x) { porId[x.id] = x; });
  var todos = Object.keys(porId).map(function (k) { return porId[k]; });
  todos.sort(function (a, b) { return String(b.atualizado).localeCompare(String(a.atualizado)); });
  salvarTabela(ss, MP_ABA, todos.slice(0, MP_MAX_LINHAS));
  return novos.length;
}

// Rodada pelo gatilho de tempo (a cada 5 minutos).
function mpSincronizar() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    mpBuscarPagamentos(SpreadsheetApp.getActiveSpreadsheet(), 3);
  } finally {
    lock.releaseLock();
  }
}

// RODE ESTA FUNCAO UMA VEZ depois de colar a chave nas Propriedades do script.
// Ela confere a chave, busca os ultimos 30 dias e liga a busca a cada 5 minutos.
function instalarMercadoPago() {
  var eu = mpChamar('get', '/users/me');
  if (eu.id) PropertiesService.getScriptProperties().setProperty('MP_USER_ID', String(eu.id));
  Logger.log('Conta Mercado Pago conectada: ' + (eu.nickname || eu.email || eu.id));
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'mpSincronizar') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('mpSincronizar').timeBased().everyMinutes(5).create();
  var n = mpBuscarPagamentos(SpreadsheetApp.getActiveSpreadsheet(), 30);
  Logger.log('Pagamentos encontrados nos ultimos 30 dias: ' + n);
  Logger.log('Pronto! A busca automatica roda a cada 5 minutos.');
}
