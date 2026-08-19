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
