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
      data.nom || '',
      data.tel || '',
      data.enfant || '',
      data.type || '',
      data.stand || '',
      data.creneau || '',
      data.gateau || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'Script actif' }))
    .setMimeType(ContentService.MimeType.JSON);
}
