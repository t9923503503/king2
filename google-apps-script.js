/**
 * Google Apps Script — King of the Court
 * Экспорт итоговой таблицы турнира в Google Sheets
 *
 * КАК РАЗВЕРНУТЬ:
 * 1. Откройте Google Sheets → Extensions → Apps Script
 * 2. Вставьте этот код (замените содержимое Code.gs)
 * 3. Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Нажмите Deploy, скопируйте URL
 * 5. Вставьте URL в поле приложения (Статистика → Экспорт в Google Sheets)
 */

// ID вашей таблицы (из URL: .../spreadsheets/d/ВАШ_ID/...)
// Оставьте пустым — скрипт запишет в ту же таблицу, к которой прикреплён
const SPREADSHEET_ID = '';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    writeResults(data);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function writeResults(data) {
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  // Название листа = название турнира + дата
  const sheetName = `${data.tournament} ${data.date}`.slice(0, 100);

  // Удалить старый лист с тем же именем, если есть
  const existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);

  const sheet = ss.insertSheet(sheetName);

  // ── Заголовок ──────────────────────────────────────────────
  sheet.getRange('A1').setValue(data.tournament);
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold');
  sheet.getRange('B1').setValue(data.date);
  sheet.getRange('C1').setValue('Экспорт: ' + new Date(data.exportedAt).toLocaleString('ru'));
  sheet.getRange('A1:C1').setBackground('#1a3a1a').setFontColor('#6ABF69');

  // ── Шапка таблицы ──────────────────────────────────────────
  const headers = ['#', '±', 'Имя', 'Пол', 'Корт', 'Очки', 'Раундов', 'Среднее', 'Лучший раунд'];
  const headerRow = sheet.getRange(3, 1, 1, headers.length);
  headerRow.setValues([headers]);
  headerRow.setFontWeight('bold').setBackground('#0d2a1a').setFontColor('#6ABF69');

  // ── Данные ─────────────────────────────────────────────────
  if (data.rows && data.rows.length > 0) {
    const values = data.rows.map(r => [
      r.place,
      r.tied,
      r.name,
      r.gender,
      r.court,
      r.totalPts,
      r.rPlayed,
      r.avg,
      r.bestRound,
    ]);
    sheet.getRange(4, 1, values.length, headers.length).setValues(values);

    // Цвет строк (чередование)
    for (let i = 0; i < values.length; i++) {
      const bg = i % 2 === 0 ? '#111122' : '#0d0d1a';
      sheet.getRange(4 + i, 1, 1, headers.length).setBackground(bg).setFontColor('#e8e8f0');
    }

    // Топ-3 золото/серебро/бронза
    const medals = ['#8B6914', '#6b6b6b', '#5a3800'];
    for (let i = 0; i < Math.min(3, values.length); i++) {
      sheet.getRange(4 + i, 1, 1, headers.length)
        .setBackground(medals[i] || '#111122')
        .setFontWeight(i === 0 ? 'bold' : 'normal');
    }
  } else {
    sheet.getRange(4, 1).setValue('Нет данных');
  }

  // ── Ширина столбцов ────────────────────────────────────────
  sheet.setColumnWidth(1, 40);   // #
  sheet.setColumnWidth(2, 30);   // ±
  sheet.setColumnWidth(3, 160);  // Имя
  sheet.setColumnWidth(4, 45);   // Пол
  sheet.setColumnWidth(5, 110);  // Корт
  sheet.setColumnWidth(6, 60);   // Очки
  sheet.setColumnWidth(7, 70);   // Раундов
  sheet.setColumnWidth(8, 70);   // Среднее
  sheet.setColumnWidth(9, 90);   // Лучший раунд

  // Закрепить шапку
  sheet.setFrozenRows(3);

  // Активировать новый лист
  ss.setActiveSheet(sheet);
}
