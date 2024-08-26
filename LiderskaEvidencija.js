function LiderskaEvidencija() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultSheet = ss.getSheetByName('Лид. Евид.');

  if (!resultSheet) {
    SpreadsheetApp.getUi().alert('List "Лид. Евид." nije pronađen.');
    return;
  }

  const sheetMappings = {
    'Јануар': 'I Јан', 'Фебруар': 'I Феб', 'Март': 'I Мар', 'Април': 'II Апр', 'Мај': 'II Мај', 'Јун': 'II Јун',
    'Јул': 'III Јул', 'Август': 'III Авг', 'Септембар': 'III Сеп', 'Октобар': 'IV Окт', 'Новембар': 'IV Нов', 'Децембар': 'IV Дец',
    'Januar': 'I Јан', 'Februar': 'I Феб', 'Mart': 'I Мар', 'April': 'II Апр', 'Maj': 'II Мај', 'Jun': 'II Јун',
    'Jul': 'III Јул', 'Avgust': 'III Авг', 'Septembar': 'III Сеп', 'Oktobar': 'IV Окт', 'Novembar': 'IV Нов', 'Decembar': 'IV Дец',
    'Јан': 'I Јан', 'Феб': 'I Феб', 'Мар': 'I Мар', 'Апр': 'II Апр', 'Мај': 'II Мај', 'Јун': 'II Јун',
    'Јул': 'III Јул', 'Авг': 'III Авг', 'Сеп': 'III Сеп', 'Окт': 'IV Окт', 'Нов': 'IV Нов', 'Дец': 'IV Дец',
    'Jan': 'I Јан', 'Feb': 'I Феб', 'Mar': 'I Мар', 'Apr': 'II Апр', 'Maj': 'II Мај', 'Jun': 'II Јун',
    'Jul': 'III Јул', 'Avg': 'III Авг', 'Sep': 'III Сеп', 'Okt': 'IV Окт', 'Nov': 'IV Нов', 'Dec': 'IV Дец'
  };

  const ui = SpreadsheetApp.getUi();
  const monthName = ui.prompt('Unesi naziv meseca:').getResponseText().trim();
  if (!monthName) {
    ui.alert('Naziv meseca nije unet. Skripta se ne može nastaviti.');
    return;
  }

  const sheetName = sheetMappings[monthName];
  if (!sheetName) {
    ui.alert('Nepoznat naziv meseca. Unesi validan naziv meseca.');
    return;
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    ui.alert('List za zadati mesec nije pronađen.');
    return;
  }

  resultSheet.getRange('AK3:AM').clearContent();

  const users = resultSheet.getRange('C3:C').getValues().flat().filter(name => name);

  // Kreiraj mapu za rezultate
  const resultsMap = new Map();
  users.forEach(user => {
    resultsMap.set(user, { 'ВБГ': 0, 'КО': 0, 'КВ': 0 });
  });

  const data = sheet.getDataRange().getValues();
  data.forEach(row => {
    const user = users.find(user => row.includes(user));
    if (user) {
      row.forEach(cell => {
        if (typeof cell === 'string') {
          const counts = resultsMap.get(user);
          if (cell.includes('КВ') || cell.includes('КЧ') || cell.includes('ЗКВ') || cell.includes('ЗКЧ')) counts['КВ']++;
          if (cell.includes('ВБГ') || cell.includes('ДОК')) counts['ВБГ']++;
          if (cell.includes('КО')) counts['КО']++;
        }
      });
    }
  });

  const resultValues = [];
  users.forEach(user => {
    const userRow = resultSheet.createTextFinder(user).findNext();
    if (userRow) {
      const row = userRow.getRow();
      const counts = resultsMap.get(user);
      resultValues.push({
        row: row,
        values: [
          counts['ВБГ'] > 0 ? counts['ВБГ'] : '',
          counts['КО'] > 0 ? counts['КО'] : '',
          counts['КВ'] > 0 ? counts['КВ'] : ''
        ]
      });
    }
  });

  // Upisi sve rezultate u jednom potezu
  resultValues.forEach(({ row, values }) => {
    resultSheet.getRange(`AK${row}:AM${row}`).setValues([values]);
  });

  ui.alert('Obrada završena', 'Podaci su uspešno ažurirani.', ui.ButtonSet.OK);
}
