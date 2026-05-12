function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    // Ajouter l'en-tête si la feuille est vide
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Date', 'Prénom', 'Nom', 'Téléphone', 'Enfant(s)',
        'Type', 'Stand', 'Créneau', 'Gâteau'
      ]);
      sheet.getRange(1, 1, 1, 9).setFontWeight('bold');
    }

    var now = new Date();
    var dateStr = Utilities.formatDate(now, 'Europe/Paris', 'dd/MM/yyyy HH:mm');

    sheet.appendRow([
      dateStr,
      data.prenom || '',
      data.nom    || '',
      data.tel    || '',
      data.enfant || '',
      data.type   || '',
      data.stand  || '',
      data.creneau|| '',
      data.gateau || ''
    ]);

    // Renvoyer les inscriptions à jour après l'ajout
    var result = getInscriptions_();
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', inscriptions: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var result = getInscriptions_();
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', inscriptions: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getInscriptions_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var result = [];

  data.forEach(function(row) {
    if (!row[1] && !row[2]) return; // ligne vide
    result.push({
      date:    row[0] ? Utilities.formatDate(new Date(row[0]), 'Europe/Paris', 'dd/MM HH:mm') : '',
      prenom:  row[1] || '',
      nom:     row[2] || '',
      type:    row[5] || '',
      stand:   row[6] || '',
      creneau: row[7] || '',
      gateau:  row[8] || ''
    });
  });

  return result;
}
