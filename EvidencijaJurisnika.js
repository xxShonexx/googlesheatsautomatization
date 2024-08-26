function EvidencijaJurisnici() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ['I Јан', 'I Феб', 'I Мар', 'II Апр', 'II Мај', 'II Јун', 'III Јул', 'III Авг', 'III Сеп', 'IV Окт', 'IV Нов', 'IV Дец'];
  const rezultatSheet = ss.getSheetByName('Јуришници'); // Мењање имена листа у који уписује
  
  if (!rezultatSheet) {
    Logger.log('Лист Јуришници не постоји.');
    return;
  }

  rezultatSheet.getRange('B2:M').clearContent();  // Брисанје претходних садржаја у опсегу

  const naslov = ['ЈУР', 'КО', 'ВБГ', 'БОЛ', 'ИНС'];
  const targetColor = '#212121';
  
  const namesSheet = ss.getSheetByName('ВЕС'); // Замена 'ВЕС' са стварним именом листа
  if (!namesSheet) {
    Logger.log('Лист ВЕС не постоји.');
    return;
  }
  
  const nameRange = namesSheet.getRange('I20:I37');
  const users = nameRange.getValues().flat();  // Флаттенује у један низ

  if (users.length === 0) {
    Logger.log('Нема имена за обраду.');
    return;
  }
  
  let currentRow = 1;  // Почетни ред за уписивање резултата

  users.forEach(user => {
    if (user) {  // Проверити да ли је име присутно
      rezultatSheet.getRange(currentRow, 1).setValue(user);
      upisiNaslovJurisnici(rezultatSheet, naslov, currentRow + 1);
      upisiImenaSheetovaJurisnici(rezultatSheet, sheetNames, currentRow);

      sheetNames.forEach((sheetName, colIndex) => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          Logger.log('Nedostaje ' + sheetName);
          return;
        }

        const { counts, insCount } = izracunajRezultateJurisnici(sheet, user, targetColor);
        upisiRezultateJurisnici(rezultatSheet, counts, insCount, currentRow + 1, colIndex + 2);
      });

      currentRow += naslov.length + 2;  // Размак између редова
    }
  });
}

// Функција за уписивање наслова (специфична за Јуришници)
function upisiNaslovJurisnici(sheet, naslov, startRow) {
  naslov.forEach((name, index) => {
    sheet.getRange(startRow + index, 1).setValue(name);
  });
}

// Функција за уписивање имена листова (специфична за Јуришници)
function upisiImenaSheetovaJurisnici(sheet, sheetNames, row) {
  sheetNames.forEach((name, index) => {
    sheet.getRange(row, 2 + index).setValue(name);
  });
}

// Функција за рачунање резултата (специфична за Јуришници)
function izracunajRezultateJurisnici(sheet, user, targetColor) {
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const backgrounds = dataRange.getBackgrounds();
  const counts = { 'ЈУР': 0, 'КО': 0, 'ВБГ': 0, 'БОЛ': 0 };
  let insCount = 0;

  data.forEach((row, rowIndex) => {
    if (row.some(cell => typeof cell === 'string' && cell.includes(user))) {
      row.forEach((cell, colIndex) => {
        if (backgrounds[rowIndex][colIndex] === targetColor && typeof cell === 'string') {
          Object.keys(counts).forEach(key => {
            if (cell.includes(key)) {
              counts[key]++;
            }
          });
          if (cell.includes('ИНС')) insCount++;
        }
      });
    }
  });

  return { counts, insCount };
}

// Функција за уписивање резултата (специфична за Јуришници)
function upisiRezultateJurisnici(sheet, counts, insCount, startRow, col) {
  if (startRow && col) { 
    let resultRow = startRow;
    Object.values(counts).forEach(count => {
      sheet.getRange(resultRow++, col).setValue(count);
    });
    sheet.getRange(resultRow, col).setValue(insCount);
  } 
}
