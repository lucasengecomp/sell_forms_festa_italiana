// Google Apps Script — Festa da Gastronomia Italiana 2026
// Deploy: Extensões > Apps Script > Implantar > Novo > Web App
//   Executar como: Eu mesmo
//   Quem tem acesso: Qualquer pessoa
// Copie a URL gerada e cole em SCRIPT_URL no HTML.
// IMPORTANTE: após alterar este arquivo, reimplante como NOVA VERSÃO.

// Manter em sincronia com config.js
var PRECO = { adulto: 165, criancaPagante: 85 };
var SHEET_NAME = 'Inscrições';
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
        'Forma de Pagamento'
      ]);
    } else if (sheet.getRange(1, 9).getValue() === '') {
      // Planilhas criadas antes da coluna de pagamento
      sheet.getRange(1, 9).setValue('Forma de Pagamento');
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

    sheet.appendRow([
      new Date(),
      String(data.nome || ''),
      String(data.telefone || ''),
      adultos,
      pagantes,
      gratuitas,
      adultos + pagantes + gratuitas,
      totalValor,
      FORMAS_PAGAMENTO[pagamento]
    ]);

    return jsonResponse({ status: 'ok' });

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

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
