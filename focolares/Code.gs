// Google Apps Script — Festa da Gastronomia Italiana 2026
// Deploy: Extensões > Apps Script > Implantar > Novo > Web App
//   Executar como: Eu mesmo
//   Quem tem acesso: Qualquer pessoa
// Copie a URL gerada e cole em SCRIPT_URL no HTML.

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(5000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp',
        'Nome',
        'Telefone',
        'Adultos',
        'Crianças Pagantes',
        'Crianças Não Pagantes',
        'Total Convites',
        'Total Valor (R$)'
      ]);
    }

    var data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date(data.timestamp || Date.now()),
      data.nome              || '',
      data.telefone          || '',
      Number(data.adultos)               || 0,
      Number(data.criancas_pagantes)     || 0,
      Number(data.criancas_nao_pagantes) || 0,
      Number(data.total_convites)        || 0,
      Number(data.total_valor)           || 0
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}
