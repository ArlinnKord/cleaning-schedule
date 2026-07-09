const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// ==============================
// НАСТРОЙКА ГОРНИЧНЫХ
// ==============================
// Впишите сюда номера для каждой горничной
const HOUSEKEEPERS = [
  { name: 'Горничная 1', color: 'FF00B050', rooms: [101, 102, 103, 201, 202, 203, 205, 207, 210, 213, 215, 219, 225] },
  { name: 'Горничная 2', color: 'FF0070C0', rooms: [104, 105, 106, 107, 108, 109, 110, 112, 113, 114, 115, 116] },
  { name: 'Горничная 3', color: 'FFFF0000', rooms: [204, 206, 208, 209, 211, 212, 214, 216, 217, 218, 220, 221, 222, 223, 224] },
];

// ==============================
// АВТОПОИСК ВХОДНОГО ФАЙЛА
// ==============================
const IN_DIR = path.join(__dirname, 'готовые_отчёты_уборка');
const OUT_DIR = path.join(__dirname, 'отчёты_горничным');
if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

function findInputFile() {
  const files = fs.readdirSync(IN_DIR)
    .filter(f => f.startsWith('виды-уборок-') && f.endsWith('.xlsx') && !f.includes('~$'))
    .sort()
    .reverse();
  if (files.length === 0) {
    console.error('❌ Не найден файл виды-уборок-ДД.ММ.ГГ.xlsx в папке готовые_отчёты_уборка');
    process.exit(1);
  }
  return path.join(IN_DIR, files[0]);
}

const INPUT_FILE = findInputFile();

const dateMatch = path.basename(INPUT_FILE).match(/виды-уборок-(\d{2}\.\d{2}\.\d{2})/);
const FILE_DATE = dateMatch ? dateMatch[1] : '';

const OUTPUT_FILE = path.join(OUT_DIR, `отчёт-горничным-${FILE_DATE}.xlsx`);

console.log(`📥 ${path.basename(INPUT_FILE)}`);
console.log(`📤 ${path.basename(OUTPUT_FILE)}`);

// ==============================
// ЧТЕНИЕ ДАННЫХ
// ==============================
const srcWb = XLSX.readFile(INPUT_FILE);
const srcWs = srcWb.Sheets['Уборки на сегодня'];
const rows = XLSX.utils.sheet_to_json(srcWs, { defval: '', header: 1 });

// Первая строка — заголовок
const headerRow = rows[0];
const dataRows = rows.slice(1).filter(r => r[0]);

// Цвета для колонок типов уборок (как в исходнике)
const TYPE_COLORS = {
  1: { bg: 'FFFCE4D6', font: 'FFC65911' }, // 10 мин — оранжевый
  2: { bg: 'FFD9EAD3', font: 'FF38761D' }, // выезд/заезд — зелёный
  3: { bg: 'FFF4CCCC', font: 'FFCC0000' }, // выезд — красный
  4: { bg: 'FFD9D9D9', font: 'FF595959' }, // 20 мин — серый
};

// ==============================
// ФОРМИРОВАНИЕ ОТЧЁТА
// ==============================
async function buildReport() {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Отчёт по уборкам';

  // Стиль шапки
  const headerStyle = {
    font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 },
    fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2F5496' } },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: {
      top: { style: 'thin' }, bottom: { style: 'thin' },
      left: { style: 'thin' }, right: { style: 'thin' },
    },
  };

  // Базовый стиль ячейки
  function baseStyle() {
    return {
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: {
        top: { style: 'thin' }, bottom: { style: 'thin' },
        left: { style: 'thin' }, right: { style: 'thin' },
      },
    };
  }

  // Стиль для комментария (с выравниванием влево)
  function commentStyle() {
    const s = baseStyle();
    s.alignment.horizontal = 'left';
    return s;
  }

  // Заголовки как в исходнике
  const HEADERS = ['Номер', '10 минут', 'Выезд/Заезд', 'Выезд', '20 минут', 'Комментарий', 'Кол-во чел'];

  for (const hk of HOUSEKEEPERS) {
    const sheet = workbook.addWorksheet(hk.name, {
      views: [{ state: 'frozen', ySplit: 1 }],
    });

    // Шапка
    const hRow = sheet.addRow(HEADERS);
    hRow.height = 28;
    HEADERS.forEach((_, i) => {
      const cell = hRow.getCell(i + 1);
      Object.assign(cell, headerStyle);
    });

    // Ширина колонок
    sheet.getColumn(1).width = 8;   // Номер
    sheet.getColumn(2).width = 10;  // 10 минут
    sheet.getColumn(3).width = 12;  // Выезд/Заезд
    sheet.getColumn(4).width = 10;  // Выезд
    sheet.getColumn(5).width = 10;  // 20 минут
    sheet.getColumn(6).width = 45;  // Комментарий
    sheet.getColumn(7).width = 12;  // Кол-во чел

    const roomSet = new Set(hk.rooms);
    const highlightColor = hk.color;

    for (const row of dataRows) {
      const room = Number(row[0]);
      const isMyRoom = roomSet.has(room);

      // Добавляем строку с теми же данными
      const excelRow = sheet.addRow(row.slice(0, 7));
      excelRow.height = 22;

      // Если номер belongs to this housekeeper — подсветить всю строку
      if (isMyRoom) {
        const lightColor = hk.color.slice(0, 4) + '1A' + hk.color.slice(6);

        for (let c = 1; c <= 7; c++) {
          const cell = excelRow.getCell(c);
          Object.assign(cell, c === 6 ? commentStyle() : baseStyle());
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: lightColor },
          };
          cell.font = { bold: true };
        }
      } else {
        // Обычное форматирование
        for (let c = 1; c <= 7; c++) {
          const cell = excelRow.getCell(c);
          Object.assign(cell, c === 6 ? commentStyle() : baseStyle());

          // row — 0-индексированный массив: row[0]=Номер, row[1]=10мин, row[2]=Выезд/Заезд, row[3]=Выезд, row[4]=20мин
          const dataIdx = c - 1;
          if (dataIdx >= 1 && dataIdx <= 4 && row[dataIdx]) {
            const tc = TYPE_COLORS[dataIdx];
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: tc.bg } };
            cell.font = { color: { argb: tc.font }, bold: true };
          }
        }
      }
    }

    // Итоговая строка
    const totalRow = sheet.addRow([]);
    totalRow.height = 24;

    const myTasks = dataRows.filter(r => roomSet.has(Number(r[0])));
    let totalMin = 0;
    for (const t of myTasks) {
      if (t[1]) totalMin += 10;
      else if (t[2] || t[3]) totalMin += 40;
      else if (t[4]) totalMin += 20;
    }

    const totalLabel = `Итого: ${myTasks.length} задач / ${totalMin} мин`;
    const totalCell = totalRow.getCell(1);
    totalCell.value = totalLabel;
    totalCell.font = { bold: true, size: 11 };
    totalCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hk.color } };
    totalCell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    Object.assign(totalCell, baseStyle());
    totalCell.alignment = { horizontal: 'left', vertical: 'center' };
    sheet.mergeCells(`A${totalRow.number}:G${totalRow.number}`);
  }

  await workbook.xlsx.writeFile(OUTPUT_FILE);
  console.log(`✅ Готово: ${path.basename(OUTPUT_FILE)}`);

  // Сводка
  console.log('\n=== СВОДКА ===');
  for (const hk of HOUSEKEEPERS) {
    const roomSet = new Set(hk.rooms);
    const tasks = dataRows.filter(r => roomSet.has(Number(r[0])));
    let total = 0;
    const taskTypes = {};
    for (const t of tasks) {
      let mins = 0;
      let type = '';
      if (t[1]) { mins = 10; type = '10'; }
      else if (t[2]) { mins = 40; type = 'выезд/заезд'; }
      else if (t[3]) { mins = 40; type = 'выезд'; }
      else if (t[4]) { mins = 20; type = '20'; }
      total += mins;
      taskTypes[type] = (taskTypes[type] || 0) + 1;
    }
    const typeStr = Object.entries(taskTypes).map(([k, v]) => `${k}: ${v}`).join(', ');
    console.log(`  ${hk.name}: ${tasks.length} задач, ${total} мин`);
    console.log(`    ${typeStr}`);
  }
}

buildReport().catch(err => {
  console.error('Ошибка:', err);
  process.exit(1);
});
