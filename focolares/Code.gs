// Google Apps Script — Festa da Gastronomia Italiana 2026
// Deploy: Extensões > Apps Script > Implantar > Novo > Web App
//   Executar como: Eu mesmo
//   Quem tem acesso: Qualquer pessoa
// Copie a URL gerada e cole em SCRIPT_URL no HTML.
// IMPORTANTE: após alterar este arquivo, reimplante como NOVA VERSÃO.
//
// Propriedades necessárias (Configurações do projeto > Propriedades do script):
//   ASAAS_API_KEY  =  
//   ASAAS_BASE_URL = https://sandbox.asaas.com/api/v3  (teste)
//                    ou https://api.asaas.com/v3       (produção)

// Manter em sincronia com config.js
var PRECO = { adulto: 175, criancaPagante: 90 };
var SHEET_NAME = 'Inscrições';
var SHEET_ASAAS = 'Cobranças Asaas';
var MAX_PARCELAS = 3;
var FORMAS_PAGAMENTO = {
  pix:    'PIX',
  cartao: 'Cartão de crédito',
  balcao: 'Balcão — Padaria Espiga Dourada',
  isento: 'Isento (sem valor a pagar)'
};

var DIAS_LABEL    = { sabado: 'Sábado',  domingo: 'Domingo' };
var PREFIXOS_DIA  = { sabado: 'FGI_S_',  domingo: 'FGI_D_' };
var COL_DIA       = 11;
var COL_CODIGO    = 12;

function doPost(e) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return jsonResponse({ status: 'error', message: 'Servidor ocupado. Tente novamente em instantes.' });
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp',
        'Nome',
        'Telefone',
        'Adultos',
        'Crianças Pagantes',
        'Crianças Não Pagantes',
        'Total Convites',
        'Total Valor (R$)',
        'Forma de Pagamento',
        'CPF',
        'Dia',
        'Código'
      ]);
    } else {
      // Planilhas criadas antes das colunas Forma de Pagamento / CPF / Dia / Código
      if (sheet.getRange(1, 9).getValue()          === '') sheet.getRange(1, 9).setValue('Forma de Pagamento');
      if (sheet.getRange(1, 10).getValue()         === '') sheet.getRange(1, 10).setValue('CPF');
      if (sheet.getRange(1, COL_DIA).getValue()    === '') sheet.getRange(1, COL_DIA).setValue('Dia');
      if (sheet.getRange(1, COL_CODIGO).getValue() === '') sheet.getRange(1, COL_CODIGO).setValue('Código');
    }

    var data = JSON.parse(e.postData.contents);

    var diaKey = classificarDia_(data.dia);
    if (!diaKey) {
      return jsonResponse({ status: 'error', message: 'Escolha em qual dia da festa você vai (Sábado ou Domingo).' });
    }

    var adultos   = toQty(data.adultos);
    var pagantes  = toQty(data.criancas_pagantes);
    var gratuitas = toQty(data.criancas_nao_pagantes);

    if (adultos + pagantes + gratuitas === 0) {
      return jsonResponse({ status: 'error', message: 'Nenhum convite selecionado.' });
    }

    var totalValor = adultos * PRECO.adulto + pagantes * PRECO.criancaPagante;
    var pagamento = totalValor === 0 ? 'isento' : String(data.forma_pagamento || '');
    if (!FORMAS_PAGAMENTO[pagamento]) {
      return jsonResponse({ status: 'error', message: 'Forma de pagamento inválida.' });
    }

    var nome     = String(data.nome || '');
    var telefone = String(data.telefone || '');
    var cpf      = digits_(data.cpf);

    // CPF obrigatório apenas para cartão (Asaas exige cpfCnpj no customer)
    if (pagamento === 'cartao' && !isCpfValido_(cpf)) {
      return jsonResponse({ status: 'error', message: 'CPF inválido para pagamento com cartão.' });
    }

    var codigo = proximoCodigoDia_(sheet, diaKey);

    sheet.appendRow([
      new Date(),
      nome,
      telefone,
      adultos,
      pagantes,
      gratuitas,
      adultos + pagantes + gratuitas,
      totalValor,
      FORMAS_PAGAMENTO[pagamento],
      cpf || '',
      DIAS_LABEL[diaKey],
      codigo
    ]);

    // Código é interno (uso administrativo na planilha) — não vai na resposta.
    var resposta = { status: 'ok' };

    // Cartão: cria a cobrança no Asaas e devolve a URL de pagamento.
    // Se algo falhar (API fora, credenciais erradas), a inscrição já está
    // gravada — devolvemos aviso pro cliente pedir o link pelo WhatsApp.
    if (pagamento === 'cartao') {
      var parcelas = Math.min(MAX_PARCELAS, Math.max(1, parseInt(data.parcelas, 10) || 1));
      try {
        var invoiceUrl = criarCobrancaAsaas_({
          nome:      nome,
          telefone:  telefone,
          cpf:       cpf,
          adultos:   adultos,
          pagantes:  pagantes,
          gratuitas: gratuitas,
          totalValor: totalValor,
          parcelas:  parcelas
        });
        resposta.paymentUrl = invoiceUrl;
        resposta.parcelas   = parcelas;
      } catch (asaasErr) {
        // Não expor detalhes do erro ao cliente; ficam nos logs pra debug.
        Logger.log('Asaas erro: ' + asaasErr + (asaasErr && asaasErr.stack ? '\n' + asaasErr.stack : ''));
        resposta.paymentError = true;
      }
    }

    return jsonResponse(resposta);

  } catch (err) {
    return jsonResponse({ status: 'error', message: String(err) });

  } finally {
    lock.releaseLock();
  }
}

function toQty(v) {
  var n = parseInt(v, 10);
  if (isNaN(n) || n < 0) return 0;
  return Math.min(n, 99);
}

function digits_(v) {
  return String(v == null ? '' : v).replace(/\D/g, '');
}

// Validação padrão de CPF (11 dígitos + dígitos verificadores).
function isCpfValido_(cpf) {
  var d = digits_(cpf);
  if (d.length !== 11) return false;
  if (/^(\d)\1{10}$/.test(d)) return false;
  for (var t = 9; t < 11; t++) {
    var soma = 0;
    for (var i = 0; i < t; i++) soma += parseInt(d.charAt(i), 10) * ((t + 1) - i);
    var dv = (soma * 10) % 11;
    if (dv === 10) dv = 0;
    if (dv !== parseInt(d.charAt(t), 10)) return false;
  }
  return true;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Código identificador (para alocação de mesas, controle de entrada, etc.) ──
// Formato: FGI_S_001 (sábado) / FGI_D_001 (domingo). Numeração sequencial por
// dia, 3 dígitos. Preenchido no momento da inscrição, sob o mesmo lock do doPost.

function classificarDia_(v) {
  var s = String(v == null ? '' : v).trim().toLowerCase();
  if (s === 'sabado' || s === 'sábado' || s === 's') return 'sabado';
  if (s === 'domingo' || s === 'd') return 'domingo';
  return null;
}

function proximoCodigoDia_(sheet, diaKey) {
  var prefixo = PREFIXOS_DIA[diaKey];
  var lastRow = sheet.getLastRow();
  var max = 0;
  if (lastRow >= 2) {
    var codigos = sheet.getRange(2, COL_CODIGO, lastRow - 1, 1).getValues();
    for (var i = 0; i < codigos.length; i++) {
      var c = String(codigos[i][0] || '');
      if (c.indexOf(prefixo) === 0) {
        var n = parseInt(c.slice(prefixo.length), 10);
        if (!isNaN(n) && n > max) max = n;
      }
    }
  }
  return prefixo + ('000' + (max + 1)).slice(-3);
}

// Rode UMA VEZ pelo editor do Apps Script para preencher a coluna "Código" das
// linhas antigas. Antes, preencha manualmente a coluna "Dia" (col K) com
// "Sábado" ou "Domingo" nas linhas existentes. Linhas sem Dia ficam sem código
// e o total pulado é reportado no Logger.
// A numeração respeita códigos já existentes (não sobrescreve) e usa a ordem
// do Timestamp para atribuir os novos.
function preencherCodigosExistentes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Aba "' + SHEET_NAME + '" não encontrada.');

  if (sheet.getRange(1, COL_DIA).getValue()    === '') sheet.getRange(1, COL_DIA).setValue('Dia');
  if (sheet.getRange(1, COL_CODIGO).getValue() === '') sheet.getRange(1, COL_CODIGO).setValue('Código');

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Sem linhas para preencher.'); return; }

  var qtd        = lastRow - 1;
  var timestamps = sheet.getRange(2, 1,          qtd, 1).getValues();
  var dias       = sheet.getRange(2, COL_DIA,    qtd, 1).getValues();
  var codigos    = sheet.getRange(2, COL_CODIGO, qtd, 1).getValues();

  var maxPorDia = { sabado: 0, domingo: 0 };
  for (var i = 0; i < qtd; i++) {
    var c = String(codigos[i][0] || '');
    ['sabado', 'domingo'].forEach(function(k) {
      var p = PREFIXOS_DIA[k];
      if (c.indexOf(p) === 0) {
        var n = parseInt(c.slice(p.length), 10);
        if (!isNaN(n) && n > maxPorDia[k]) maxPorDia[k] = n;
      }
    });
  }

  var ordem = [];
  for (var i = 0; i < qtd; i++) ordem.push(i);
  ordem.sort(function(a, b) {
    var ta = new Date(timestamps[a][0]).getTime();
    var tb = new Date(timestamps[b][0]).getTime();
    return ta - tb;
  });

  var preenchidas = 0, puladas = 0;
  ordem.forEach(function(rowIdx) {
    if (codigos[rowIdx][0]) return;
    var key = classificarDia_(dias[rowIdx][0]);
    if (!key) { puladas++; return; }
    maxPorDia[key]++;
    codigos[rowIdx][0] = PREFIXOS_DIA[key] + ('000' + maxPorDia[key]).slice(-3);
    preenchidas++;
  });

  sheet.getRange(2, COL_CODIGO, qtd, 1).setValues(codigos);
  Logger.log('Códigos preenchidos: ' + preenchidas + '. Linhas puladas por falta de Dia: ' + puladas + '.');
}

// Rode ESTA função UMA VEZ pelo editor do Apps Script (botão Executar / Run).
// Ela dispara o pop-up de autorização pedindo o escopo external_request,
// necessário para o UrlFetchApp.fetch conseguir chamar a API do Asaas.
// Depois de autorizar, o web app funciona normalmente. Pode remover depois.
function autorizarPermissoesExternas() {
  var res = UrlFetchApp.fetch('https://sandbox.asaas.com/api/v3', { muteHttpExceptions: true });
  Logger.log('Autorização OK. Status: ' + res.getResponseCode());
}

// ── Asaas ──────────────────────────────────────────────────────────────────

function criarCobrancaAsaas_(p) {
  var props   = PropertiesService.getScriptProperties();
  var apiKey  = props.getProperty('ASAAS_API_KEY');
  var baseUrl = props.getProperty('ASAAS_BASE_URL');
  if (!apiKey || !baseUrl) {
    throw new Error('ASAAS_API_KEY/ASAAS_BASE_URL não configurados no Apps Script.');
  }
  baseUrl = baseUrl.replace(/\/+$/, '');

  var customerId = asaasUpsertCliente_(baseUrl, apiKey, p.nome, p.telefone, p.cpf);

  var descPartes = [];
  if (p.adultos)   descPartes.push(p.adultos   + ' adulto'    + (p.adultos   > 1 ? 's' : ''));
  if (p.pagantes)  descPartes.push(p.pagantes  + ' criança'   + (p.pagantes  > 1 ? 's' : '') + ' pagante' + (p.pagantes > 1 ? 's' : ''));
  if (p.gratuitas) descPartes.push(p.gratuitas + ' criança'   + (p.gratuitas > 1 ? 's' : '') + ' gratuita' + (p.gratuitas > 1 ? 's' : ''));
  var descricao = 'Festa Gastronomia Italiana 2026 — ' + descPartes.join(', ');

  var payload = {
    customer:          customerId,
    billingType:       'CREDIT_CARD',
    dueDate:           dataVencimento_(3),
    description:       descricao,
    externalReference: (p.telefone || '') + '|' + Date.now()
  };

  if (p.parcelas > 1) {
    // totalValue faz o Asaas distribuir os centavos automaticamente entre as parcelas
    payload.installmentCount = p.parcelas;
    payload.totalValue       = p.totalValor;
  } else {
    payload.value = p.totalValor;
  }

  var resp = asaasFetch_(baseUrl + '/payments', apiKey, 'post', payload);
  var invoiceUrl = resp.invoiceUrl;
  if (!invoiceUrl) {
    throw new Error('Asaas não retornou invoiceUrl. Resposta: ' + JSON.stringify(resp));
  }

  registrarCobrancaAsaas_({
    nome:       p.nome,
    telefone:   p.telefone,
    cpf:        p.cpf,
    totalValor: p.totalValor,
    parcelas:   p.parcelas,
    customerId: customerId,
    paymentId:  resp.id || resp.installment || '',
    invoiceUrl: invoiceUrl
  });

  return invoiceUrl;
}

function asaasUpsertCliente_(baseUrl, apiKey, nome, telefone, cpf) {
  var telDigitos = digits_(telefone);
  var cpfDigitos = digits_(cpf);

  // Se já existe cliente com esse CPF, reusa (Asaas indexa cpfCnpj).
  if (cpfDigitos) {
    var busca = asaasFetch_(
      baseUrl + '/customers?cpfCnpj=' + encodeURIComponent(cpfDigitos),
      apiKey, 'get', null
    );
    if (busca && busca.data && busca.data.length && busca.data[0].id) {
      return busca.data[0].id;
    }
  }

  var body = { name: nome || 'Sem nome' };
  if (telDigitos) body.mobilePhone = telDigitos;
  if (cpfDigitos) body.cpfCnpj = cpfDigitos;

  var resp;
  try {
    resp = asaasFetch_(baseUrl + '/customers', apiKey, 'post', body);
  } catch (err) {
    // Se o Asaas rejeitar por causa do telefone, tenta de novo sem ele —
    // pagamento não depende de celular, só afeta notificações.
    if (body.mobilePhone && /celular|telefone|mobile|phone/i.test(String(err))) {
      Logger.log('Asaas rejeitou mobilePhone (' + telDigitos + '), tentando sem: ' + err);
      delete body.mobilePhone;
      resp = asaasFetch_(baseUrl + '/customers', apiKey, 'post', body);
    } else {
      throw err;
    }
  }

  if (!resp.id) {
    throw new Error('Asaas não retornou id do cliente. Resposta: ' + JSON.stringify(resp));
  }
  return resp.id;
}

function asaasFetch_(url, apiKey, method, body) {
  var options = {
    method:             method,
    contentType:        'application/json',
    muteHttpExceptions: true,
    headers: {
      'access_token': apiKey,
      'User-Agent':   'festa-italiana-mariapolisginetta'
    }
  };
  if (body) options.payload = JSON.stringify(body);

  var res  = UrlFetchApp.fetch(url, options);
  var code = res.getResponseCode();
  var text = res.getContentText();
  var json;
  try { json = JSON.parse(text); } catch (_) { json = null; }

  if (code < 200 || code >= 300) {
    var msg = json && json.errors && json.errors[0] && json.errors[0].description
      ? json.errors[0].description
      : ('HTTP ' + code + ' — ' + text.slice(0, 200));
    throw new Error('Asaas: ' + msg);
  }
  return json || {};
}

function dataVencimento_(diasAFrente) {
  var d = new Date();
  d.setDate(d.getDate() + diasAFrente);
  var mm = ('0' + (d.getMonth() + 1)).slice(-2);
  var dd = ('0' + d.getDate()).slice(-2);
  return d.getFullYear() + '-' + mm + '-' + dd;
}

function registrarCobrancaAsaas_(r) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_ASAAS) || ss.insertSheet(SHEET_ASAAS);
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'Timestamp', 'Nome', 'Telefone', 'CPF', 'Total (R$)', 'Parcelas',
      'Customer ID', 'Payment ID', 'Invoice URL'
    ]);
  }
  sh.appendRow([
    new Date(), r.nome, r.telefone, r.cpf || '', r.totalValor, r.parcelas,
    r.customerId, r.paymentId, r.invoiceUrl
  ]);
}
