const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const readline = require('readline');

const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
async function ask(q) { return new Promise(r => rl.question(q, r)); }

const typeToCol = { '10': 2, '40 выезд/заезд': 3, '40 выезд': 4, '20': 5 };

function findFiles() {
  const results = [];
  const dirs = [
    path.join(__dirname, '..', 'отчёты_горничным_с_сортировкой', 'готовые_отчёты'),
    path.join(__dirname, '..', 'отчёты_горничным_без_сортировки', 'готовые_отчёты'),
    __dirname,
  ];
  for (const dir of dirs) {
    if (!fs.existsSync(dir)) continue;
    for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
      if (e.name.startsWith('~$')) continue;
      const full = path.join(dir, e.name);
      if (e.isDirectory()) {
        for (const s of fs.readdirSync(full, { withFileTypes: true })) {
          if (s.name.endsWith('.xlsx') && !s.name.startsWith('~$'))
            results.push({ file: path.join(full, s.name), mtime: fs.statSync(path.join(full, s.name)).mtimeMs });
        }
      } else if (e.name.endsWith('.xlsx')) {
        results.push({ file: full, mtime: fs.statSync(full).mtimeMs });
      }
    }
  }
  results.sort((a, b) => b.mtime - a.mtime);
  return results;
}

async function main() {
  console.log('=== РАЗБИВКА ПО ЦВЕТАМ ===\n');

  // --- Файл ---
  const inp = await ask('Путь к .xlsx (Enter — выбрать из списка): ');
  let inputFile;
  if (inp.trim()) {
    inputFile = inp.trim();
  } else {
    const files = findFiles();
    if (!files.length) { console.log('Файлы не найдены'); rl.close(); return; }
    files.slice(0, 15).forEach((f, i) => console.log(`  ${i+1}. ${path.relative(__dirname, f.file)}`));
    if (files.length > 15) console.log(`  ... ещё ${files.length - 15}`);
    const idx = parseInt(await ask('Номер: ')) || 1;
    inputFile = files[Math.min(idx - 1, files.length - 1)].file;
  }

  console.log(`\n📂 ${path.basename(inputFile)}`);

  // --- Читаем исходник ---
  const srcWb = new ExcelJS.Workbook();
  await srcWb.xlsx.readFile(inputFile);
  const ws = srcWb.getWorksheet('Уборки на сегодня');
  if (!ws) { console.log('Лист "Уборки на сегодня" не найден'); rl.close(); return; }

  // Собираем строки с цветами
  const allTasks = [];
  const colorGroups = {}; // argb -> Set<room>
  const colorOrder = [];

  for (let r = 4; r <= ws.rowCount; r++) {
    const roomVal = ws.getCell(r, 1).value;
    if (roomVal == null) continue;
    const room = parseInt(roomVal);
    if (isNaN(room)) continue;

    let type = null;
    let minutes = null;
    let fillArgb = null;

    for (const [t, col] of Object.entries(typeToCol)) {
      const cell = ws.getCell(r, col);
      if (!cell.value) continue;
      type = t;
      minutes = cell.value;
      fillArgb = cell.fill?.fgColor?.argb;
      break;
    }

    if (!type) continue; // пустая строка (112 из примера)

    allTasks.push({
      room,
      _cleaning: { type, minutes: typeof minutes === 'number' ? minutes : parseInt(minutes) || 0 },
      _comment: ws.getCell(r, 6).value || '',
      _guestCount: ws.getCell(r, 7).value || '',
    });

    if (fillArgb) {
      if (!colorGroups[fillArgb]) {
        colorGroups[fillArgb] = new Set();
        colorOrder.push(fillArgb);
      }
      colorGroups[fillArgb].add(room);
    }
  }

  const numHK = colorOrder.length;
  if (!numHK) { console.log('Нет закрашенных ячеек'); rl.close(); return; }
  console.log(`Найдено ${numHK} горничных (цветов), ${allTasks.filter(t => t._cleaning.type).length} уборок`);

  // Имена
  const hkNames = [];
  for (let i = 0; i < numHK; i++) {
    const name = (await ask(`Имя горничной ${i+1} (по умолч. Горничная ${i+1}): `)).trim();
    hkNames.push(name || `Горничная ${i+1}`);
  }

  // Дата из заголовка
  const titleDate = String(ws.getCell(1, 1).value || '').trim();
  const colWidths = [8, 10, 14, 10, 10, 30, 12];

  // --- Создаём НОВЫЙ файл ---
  const newWb = new ExcelJS.Workbook();

  for (let hk = 0; hk < numHK; hk++) {
    const hkName = hkNames[hk];
    const color = colorOrder[hk];
    const hkRoomSet = colorGroups[color];

    const wsHk = newWb.addWorksheet(hkName, { views: [{ state: 'frozen', ySplit: 3 }] });
    wsHk.pageSetup = { fitToPage: true, fitToWidth: 1, fitToHeight: 0, orientation: 'portrait' };
    wsHk.columns = colWidths.map(w => ({ width: w }));

    // Заголовок
    wsHk.addRow(['']);
    wsHk.addRow(['']);
    wsHk.mergeCells(1, 1, 2, 7);
    const titleCell = wsHk.getCell(1, 1);
    titleCell.value = `${titleDate} — ${hkName}`;
    titleCell.font = { size: 25, bold: true };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

    // Шапка
    const hRow = wsHk.addRow(['Номер', '10 минут', 'Выезд/Заезд', 'Выезд', '20 минут', 'Комментарий', 'Кол-во чел']);
    hRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF34495E' } };
    hRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    hRow.alignment = { horizontal: 'center', vertical: 'middle' };

    // Данные — все строки, цвет только её
    for (const t of allTasks) {
      const type = t._cleaning.type || '';
      const row = wsHk.addRow([
        t.room || '',
        type === '10' ? t._cleaning.minutes : '',
        type === '40 выезд/заезд' ? t._cleaning.minutes : '',
        type === '40 выезд' ? t._cleaning.minutes : '',
        type === '20' ? t._cleaning.minutes : '',
        t._comment || '',
        t._guestCount || '',
      ]);
      row.alignment = { horizontal: 'center', vertical: 'middle' };

      if (hkRoomSet.has(t.room)) {
        const col = typeToCol[type];
        if (col) {
          wsHk.getCell(row.number, col).fill = {
            type: 'pattern', pattern: 'solid', fgColor: { argb: color },
          };
          wsHk.getCell(row.number, col).font = { bold: true, color: { argb: 'FF000000' }, size: 11 };
        }
      }
    }
  }

  // --- Лист расходников ---
  const wsCons = newWb.addWorksheet('Расходники');
  wsCons.pageSetup = { fitToPage: true, fitToWidth: 1, fitToHeight: 0, orientation: 'portrait' };
  wsCons.columns = [{ width: 20 }, { width: 14 }, { width: 14 }, { width: 16 }];

  wsCons.addRow(['']);
  wsCons.addRow(['Расходники (наборов на номер):']);
  wsCons.getCell(2, 1).font = { bold: true, size: 12 };
  wsCons.addRow(['Горничная', 'На человек', 'Выезды ×2', 'Доп. гости']);

  const hRow = wsCons.getRow(3);
  for (let c = 1; c <= 4; c++) {
    hRow.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF34495E' } };
    hRow.getCell(c).font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  }

  for (let hk = 0; hk < numHK; hk++) {
    const hkRoomSet = colorGroups[colorOrder[hk]];
    let base = 0, extra = 0;
    for (const t of allTasks) {
      if (!hkRoomSet.has(t.room)) continue;
      const type = t._cleaning.type;
      if (type === '40 выезд' || type === '40 выезд/заезд') base += 2;
      if (type === '10' || type === '40 выезд/заезд') {
        const gc = t._guestCount;
        if (gc) {
          const total = String(gc).split('+').reduce((s, p) => s + (parseInt(p) || 0), 0);
          if (total > 2) extra += (total - 2);
        }
      }
    }
    const total = base + extra;
    const row = wsCons.addRow([hkNames[hk], total, base, extra > 0 ? '+' + extra : 0]);
    row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colorOrder[hk] } };
    row.getCell(1).font = { bold: true, size: 10 };
  }

  // --- Сохраняем ---
  const outDir = path.join(__dirname, 'готовые_отчёты');
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  const today = new Date();
  const dd = String(today.getDate()).padStart(2, '0');
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const yy = String(today.getFullYear()).slice(2);

  let ver = 1;
  while (fs.existsSync(path.join(outDir, `листы-${dd}.${mm}.${yy}-v${ver}.xlsx`))) ver++;
  const outFile = path.join(outDir, `листы-${dd}.${mm}.${yy}-v${ver}.xlsx`);

  await newWb.xlsx.writeFile(outFile);
  console.log(`\n✅ Создан: ${outFile}`);
  rl.close();
}

main().catch(err => { console.error(err); rl.close(); });
