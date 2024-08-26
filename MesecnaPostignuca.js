function MesecnaPostignuca() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resultSheet = ss.getSheetByName('Месечна постигнућа');

  if (!resultSheet) {
    SpreadsheetApp.getUi().alert('List "Месечна Постигнућа" није пронађен.');
    return;
  }

  const sheetMappings = {
    'Јануар': 'I Јан', 'Фебруар': 'I Феб', 'Март': 'I Мар', 'Април': 'II Апр', 'Мај': 'II Мај', 'Јун': 'II Јун',
    'Јул': 'III Јул', 'Август': 'III Авг', 'Септембар': 'III Сеп', 'Октобар': 'IV Окт', 'Новембар': 'IV Нов', 'Децембар': 'IV Дец'
  };

  const ui = SpreadsheetApp.getUi();
  const monthName = ui.prompt('Унеси назив месеца:').getResponseText().trim();
  if (!monthName) {
    ui.alert('Назив месеца није унет. Скрипта се не може наставити.');
    return;
  }

  const sheetName = sheetMappings[monthName];
  if (!sheetName) {
    ui.alert('Непознат назив месеца. Унеси валидан назив месеца.');
    return;
  }

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    ui.alert('Лист за задати месец није пронађен.');
    return;
  }

  resultSheet.getRange('B2').clearContent();

  const usersRange = resultSheet.getRange('A2:A');
  const users = usersRange.getValues().flat().filter(name => name);

  if (users.length === 0) {
    ui.alert('Нема података у колони А.');
    return;
  }

  const resultsMap = new Map();
  users.forEach(user => {
    resultsMap.set(user, { 'ЗЕВ': 0, 'КВ': 0, 'КО': 0, 'ВБГ': 0, 'АВИ': 0, 'АРТ': 0, 'ТЕН': 0, 'БОЛ': 0, 'ЈУР': 0, 'ЗЛА': 0, 'ПЕШ': 0 });
  });

  const data = sheet.getDataRange().getValues();
  const backgrounds = sheet.getDataRange().getBackgrounds(); // Додајемо позадине
  const borderColors = sheet.getDataRange().getBorders(); // Додајемо боје ивица

  data.forEach((row, rowIndex) => {
    const user = users.find(user => row.includes(user));
    if (user) {
      row.forEach((cell, colIndex) => {
        const bgColor = backgrounds[rowIndex][colIndex]; // Узимамо боју позадине
        const borderColor = borderColors[rowIndex][colIndex]; // Узимамо боју ивица
        const counts = resultsMap.get(user);

        // Проверите активности
        if (typeof cell === 'string') {
          if (cell.includes('ЗЕВ')) counts['ЗЕВ']++;
          if (cell.includes('КВ')) counts['КВ']++;
          if (cell.includes('КО')) counts['КО']++;
          if (cell.includes('ВБГ')) counts['ВБГ']++;
          if (cell.includes('АВИ')) counts['АВИ']++;
          if (cell.includes('АРТ')) counts['АРТ']++;
          if (cell.includes('БОЛ')) counts['БОЛ']++;
          if (cell.includes('ЗЛА')) counts['ЗЛА']++;
          if (cell.includes('ПЕШ')) counts['ПЕШ']++;
        }

        // Ако је боја позадине #212121, повећавамо број за „Јуришнике“
        if (bgColor === '#212121') {
          counts['ЈУР']++;
        }

        // Ако је боја позадине #434343 и садржи „ТЕН“ или „ВБГ“, повећавамо број за „ТЕН“
        if (bgColor === '#434343' && typeof cell === 'string' && (cell.includes('ТЕН') || cell.includes('ВБГ'))) {
          counts['ТЕН']++;
        }
        
        // Ако је боја позадине #073763 и садржи „АВИ“ или „ВБГ“, повећавамо број за „АВИ“
          if (bgColor === '#073763' && typeof cell === 'string' && (cell.includes('АВИ') || cell.includes('ВБГ'))) {
            counts['АВИ']++;
          }

        // Ако је боја ивица #ffff00, повећавамо број за „ЗЛА“
        if (borderColor === '#ffff00') {
          counts['ЗЛА']++;
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
          counts['ЗЕВ'] > 0 ? counts['ЗЕВ'] : '',
          counts['КО'] > 0 ? counts['КО'] : '',
          counts['КВ'] > 0 ? counts['КВ'] : '',
          counts['ВБГ'] > 0 ? counts['ВБГ'] : '',
          counts['АВИ'] > 0 ? counts['АВИ'] : '',
          counts['АРТ'] > 0 ? counts['АРТ'] : '',
          counts['ТЕН'] > 0 ? counts['ТЕН'] : '',
          counts['БОЛ'] > 0 ? counts['БОЛ'] : '',
          counts['ЈУР'] > 0 ? counts['ЈУР'] : '',
          counts['ЗЛА'] > 0 ? counts['ЗЛА'] : '',
          counts['ПЕШ'] > 0 ? counts['ПЕШ'] : ''
        ]
      });
    } else {
      Logger.log(`Корисник ${user} није пронађен у резултатима.`);
    }
  });

  resultValues.forEach(({ row, values }) => {
    resultSheet.getRange(`B${row}:L${row}`).setValues([values]);
  });

  ui.alert('Обрада завршена', 'Подаци су успешно ажурирани.', ui.ButtonSet.OK);
}
