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
        'CPF'
      ]);
    } else {
      // Planilhas criadas antes das colunas Forma de Pagamento / CPF
      if (sheet.getRange(1, 9).getValue()  === '') sheet.getRange(1, 9).setValue('Forma de Pagamento');
      if (sheet.getRange(1, 10).getValue() === '') sheet.getRange(1, 10).setValue('CPF');
    }

    var data = JSON.parse(e.postData.contents);

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
      cpf || ''
    ]);

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
