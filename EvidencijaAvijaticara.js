function EvidencijaAvijaticara() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ['I Јан', 'I Феб', 'I Мар', 'II Апр', 'II Мај', 'II Јун', 'III Јул', 'III Авг', 'III Сеп', 'IV Окт', 'IV Нов', 'IV Дец'];
  const rezultatSheet = ss.getSheetByName('Авијатичари'); // Мењање имена листа у који уписује
  
  if (!rezultatSheet) {
    Logger.log('Лист Авијатичари не постоји.');
    return;
  }

  rezultatSheet.getRange('B2:M').clearContent();  // Брисанје претходних садржаја у опсегу

  const naslov = ['АВИ', 'ВБГ', 'ИНС'];
  const targetColor = '#073763';
  
  const namesSheet = ss.getSheetByName('ВЕС'); // Замена 'ВЕС' са стварним именом листа
  if (!namesSheet) {
    Logger.log('Лист ВЕС не постоји.');
    return;
  }
  
  const nameRange = namesSheet.getRange('D4:D10');
  const users = nameRange.getValues().flat();  // Флаттенује у један низ

  if (users.length === 0) {
    Logger.log('Нема имена за обраду.');
    return;
  }
  
  let currentRow = 1;  // Почетни ред за уписивање резултата

  users.forEach(user => {
    if (user) {  // Проверити да ли је име присутно
      rezultatSheet.getRange(currentRow, 1).setValue(user);
      upisiNaslovAvijaticari(rezultatSheet, naslov, currentRow + 1);
      upisiImenaSheetovaAvijaticari(rezultatSheet, sheetNames, currentRow);

      sheetNames.forEach((sheetName, colIndex) => {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          Logger.log('Nedostaje ' + sheetName);
          return;
        }

        const { counts, insCount } = izracunajRezultateAvijaticari(sheet, user, targetColor);
        upisiRezultateAvijaticari(rezultatSheet, counts, insCount, currentRow + 1, colIndex + 2);
      });

      currentRow += naslov.length + 2;  // Размак између редова
    }
  });
}

// Функција за уписивање наслова (специфична за Авијатичари)
function upisiNaslovAvijaticari(sheet, naslov, startRow) {
  naslov.forEach((name, index) => {
    sheet.getRange(startRow + index, 1).setValue(name);
  });
}

// Функција за уписивање имена листова (специфична за Авијатичари)
function upisiImenaSheetovaAvijaticari(sheet, sheetNames, row) {
  sheetNames.forEach((name, index) => {
    sheet.getRange(row, 2 + index).setValue(name);
  });
}

// Функција за рачунање резултата (специфична за Авијатичари)
function izracunajRezultateAvijaticari(sheet, user, targetColor) {
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const backgrounds = dataRange.getBackgrounds();
  const counts = { 'АВИ': 0, 'ВБГ': 0, };
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

// Функција за уписивање резултата (специфична за Авијатичари)
function upisiRezultateAvijaticari(sheet, counts, insCount, startRow, col) {
  if (startRow && col) { 
    let resultRow = startRow;
    Object.values(counts).forEach(count => {
      sheet.getRange(resultRow++, col).setValue(count);
    });
    sheet.getRange(resultRow, col).setValue(insCount);
  } 
}
