// ═══════════════════════════════════════════════════════════════
// INSTRUCOES DE CONFIGURACAO
// ═══════════════════════════════════════════════════════════════
//
// 1. Abra Google Sheets (sheets.google.com) e crie uma planilha nova
// 2. Va em Extensoes > Apps Script
// 3. Apague o conteudo e cole TODO este codigo
// 4. Clique em "Implantar" > "Nova implantacao"
// 5. Tipo: "App da Web"
// 6. Executar como: "Eu" (sua conta)
// 7. Quem tem acesso: "Qualquer pessoa"
// 8. Clique em "Implantar" e autorize
// 9. Copie a URL gerada e cole na aba Config do app de treino
//
// PRONTO! Seus treinos serao salvos automaticamente na planilha.
// ═══════════════════════════════════════════════════════════════

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Treinos');
  if (!sheet) {
    sheet = ss.insertSheet('Treinos');
    sheet.appendRow(['Data', 'Treino', 'Titulo', 'Exercicio', 'Concluido', 'Carga (kg)', 'Repeticoes']);
    sheet.getRange('1:1').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getOrCreateSheet();

    if (data.action === 'save') {
      var entry = data.entry;
      for (var i = 0; i < entry.items.length; i++) {
        var item = entry.items[i];
        sheet.appendRow([
          entry.date,
          entry.training,
          entry.title,
          item.name,
          item.done ? 'Sim' : 'Nao',
          item.weight || '',
          item.reps || ''
        ]);
      }
      return ContentService.createTextOutput(
        JSON.stringify({ status: 'ok', action: 'save' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === 'sync') {
      sheet.clear();
      sheet.appendRow(['Data', 'Treino', 'Titulo', 'Exercicio', 'Concluido', 'Carga (kg)', 'Repeticoes']);
      sheet.getRange('1:1').setFontWeight('bold');
      var history = data.history;
      for (var h = 0; h < history.length; h++) {
        var entry = history[h];
        for (var i = 0; i < entry.items.length; i++) {
          var item = entry.items[i];
          sheet.appendRow([
            entry.date,
            entry.training,
            entry.title,
            item.name,
            item.done ? 'Sim' : 'Nao',
            item.weight || '',
            item.reps || ''
          ]);
        }
      }
      return ContentService.createTextOutput(
        JSON.stringify({ status: 'ok', action: 'sync', count: history.length })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: 'Acao desconhecida' })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var sheet = getOrCreateSheet();
    var lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return ContentService.createTextOutput(
        JSON.stringify({ status: 'ok', history: [] })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    var groups = {};
    var order = [];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var key = String(row[0]) + '|' + String(row[1]);
      if (!groups[key]) {
        groups[key] = {
          date: String(row[0]),
          training: String(row[1]),
          title: String(row[2]),
          items: []
        };
        order.push(key);
      }
      groups[key].items.push({
        name: String(row[3]),
        done: row[4] === 'Sim',
        weight: String(row[5] || ''),
        reps: String(row[6] || '')
      });
    }

    var history = [];
    for (var j = 0; j < order.length; j++) {
      history.push(groups[order[j]]);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ status: 'ok', history: history })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: err.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
