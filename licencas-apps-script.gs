/**********************************************************
 * ClimaPro — Servidor de licenças (Google Apps Script)
 * Cole este código inteiro no editor de script da planilha.
 * A planilha precisa ter uma aba chamada "licencas" com os
 * cabeçalhos na linha 1:
 * chave | email | nome | status | ultimo_acesso | acessos
 **********************************************************/

var ABA = 'licencas';

/* ---------- API consultada pelo app ---------- */
function doGet(e) {
  var chave = String((e.parameter && e.parameter.chave) || '').trim().toUpperCase();
  var resp = { ok: false, motivo: 'invalida' };

  if (chave) {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ABA);
    var dados = sh.getDataRange().getValues(); // linha 0 = cabeçalho
    for (var i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim().toUpperCase() === chave) {
        var status = String(dados[i][3]).trim().toUpperCase();
        if (status === 'ATIVO') {
          resp = { ok: true };
          // registra o uso: ultimo_acesso (col E) e contador (col F)
          sh.getRange(i + 1, 5).setValue(new Date());
          sh.getRange(i + 1, 6).setValue((Number(dados[i][5]) || 0) + 1);
        } else {
          resp = { ok: false, motivo: 'inativa' };
        }
        break;
      }
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify(resp))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Menu dentro da planilha ---------- */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ClimaPro')
    .addItem('Gerar chave na linha selecionada', 'gerarChave')
    .addToUi();
}

/* Gera uma chave CLIMA-XXXX-XXXX na coluna A da linha selecionada
   e marca o status como ATIVO. Preencha email/nome à vontade. */
function gerarChave() {
  var sh = SpreadsheetApp.getActiveSheet();
  var linha = sh.getActiveRange().getRow();
  if (linha < 2) { SpreadsheetApp.getUi().alert('Selecione uma linha abaixo do cabeçalho.'); return; }
  var alfa = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // sem 0/O/1/I para evitar confusão
  function bloco() {
    var s = '';
    for (var i = 0; i < 4; i++) s += alfa.charAt(Math.floor(Math.random() * alfa.length));
    return s;
  }
  sh.getRange(linha, 1).setValue('CLIMA-' + bloco() + '-' + bloco());
  sh.getRange(linha, 4).setValue('ATIVO');
}
