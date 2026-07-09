const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const iconv = require('iconv-lite');
const fs = require('fs');
const path = require('path');

function fixEncoding(str) {
  if (typeof str !== 'string') return str;
  const buf = Buffer.from(str, 'latin1');
  return iconv.decode(buf, 'cp1251');
}

// ==============================
// АВТООПРЕДЕЛЕНИЕ ДАТЫ И ФАЙЛА
// ==============================
const TODAY = new Date();
const DD = String(TODAY.getDate()).padStart(2, '0');
const MM = String(TODAY.getMonth() + 1).padStart(2, '0');
const YY = String(TODAY.getFullYear()).slice(2);
const REPORT_DATE = TODAY;

const DATA_DIR = path.join(__dirname, 'отчёты_из_эдельвейса');
const OUT_DIR = path.join(__dirname, 'готовые_отчёты_уборка');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

function findInputFile() {
  const exact = path.join(DATA_DIR, `${DD}.${MM}.xls`);
  if (fs.existsSync(exact)) return exact;
  const files = fs.readdirSync(DATA_DIR).filter(f => f.endsWith('.xls') && !f.includes('~$'));
  if (files.length > 0) return path.join(DATA_DIR, files[0]);
  console.error(`❌ Файл ${DD}.${MM}.xls не найден в папке отчёты_из_эдельвейса`);
  process.exit(1);
}

const INPUT_FILE = findInputFile();
let ver = 1;
while (fs.existsSync(path.join(OUT_DIR, `график-уборок-${DD}.${MM}.${YY}-v${ver}.xlsx`))) ver++;
const OUTPUT_FILE = path.join(OUT_DIR, `график-уборок-${DD}.${MM}.${YY}-v${ver}.xlsx`);

const todayStr = `${DD}.${MM}.${TODAY.getFullYear()}`;

console.log(`📅 ${DD}.${MM}.${TODAY.getFullYear()}`);
console.log(`📥 ${path.basename(INPUT_FILE)} → 📤 ${path.basename(OUTPUT_FILE)}`);

const MAID_COLORS = [
  { name: 'Горничная 1', fill: 'E74C3C' },  // Красный
  { name: 'Горничная 2', fill: '3498DB' },  // Синий
  { name: 'Горничная 3', fill: '2ECC71' },  // Зелёный
];

const ALL_ROOMS = [
  101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 112, 113, 114, 115, 116,
  201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224, 225
];

// ==============================
// ПАРСИНГ
// ==============================

function parseDateRange(rangeStr) {
  // "29.06 (14:21) - 04.07 (12:00)"
  const m = String(rangeStr || '').match(/(\d{2})\.(\d{2})\s*\(.*?\)\s*-\s*(\d{2})\.(\d{2})/);
  if (!m) return { checkin: null, checkout: null };
  return {
    checkin: new Date(TODAY.getFullYear(), parseInt(m[2]) - 1, parseInt(m[1])),
    checkout: new Date(TODAY.getFullYear(), parseInt(m[4]) - 1, parseInt(m[3])),
  };
}

function parseEdelweiss(filePath) {
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets['Sheet1'];
  const data = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 });

  const sections = { arrivals: [], departures: [], staying: [] };
  let currentSection = null;

  for (const rawRow of data) {
    const firstCell = fixEncoding(String(rawRow[0] || '')).trim();
    if (firstCell.includes('ЗАЕЗДЫ')) { currentSection = 'arrivals'; continue; }
    if (firstCell.includes('ВЫЕЗДЫ')) { currentSection = 'departures'; continue; }
    if (firstCell.includes('ПРОЖИВАНИЯ')) { currentSection = 'staying'; continue; }
    if (!currentSection || firstCell.includes('Комната') || firstCell === '') continue;

    const roomNum = parseInt(rawRow[0]);
    if (isNaN(roomNum)) continue;

    const range = parseDateRange(fixEncoding(String(rawRow[7] || '')));
    const guestName = fixEncoding(String(rawRow[2] || '')).replace(/\s*\*+$/, '').trim();
    const notes = fixEncoding(String(rawRow[18] || ''));
    const status = fixEncoding(String(rawRow[19] || ''));

    sections[currentSection].push({
      room: roomNum,
      guest: guestName,
      type: fixEncoding(String(rawRow[1] || '')),
      notes: notes,
      nights: parseInt(rawRow[9]) || 0,
      adults: parseInt(rawRow[10]) || 0,
      checkin: range.checkin,
      checkout: range.checkout,
      status,
    });
  }

  return sections;
}

// ==============================
// ОПРЕДЕЛЕНИЕ УБОРОК
// ==============================

function getCleaning(room, sections, today) {
  const isArrival = sections.arrivals.some(r => r.room === room);
  const isDeparture = sections.departures.some(r => r.room === room);

  if (isArrival && isDeparture) {
    // Выезд + заезд
    return { type: '40 выезд/заезд', minutes: 40, col: 2 };
  }
  if (isDeparture) {
    return { type: '40 выезд', minutes: 40, col: 3 };
  }
  if (isArrival) {
    return { type: '10', minutes: 10, col: 1 };
  }

  // Проживание (текучка)
  const staying = sections.staying.filter(r => r.room === room);
  if (staying.length > 0) {
    return { type: '20', minutes: 20, col: 4 };
  }

  return { type: null, minutes: 0, col: -1 };
}

function makeComment(room, sections, cleaningType, today) {
  if (!cleaningType || !cleaningType.type) return '';

  if (cleaningType.type === '10' || cleaningType.type === '40 выезд/заезд') {
    const entry = sections.arrivals.find(r => r.room === room);
    if (!entry) return '1 кровать';
    const notes = entry.notes.toLowerCase();
    const parts = [];

    if (notes.includes('2 кроват') || notes.includes('2 раздельн') || notes.includes('2 кров')) parts.push('2 кровати');
    else parts.push('1 кровать');

    if (notes.includes('люльк')) parts.push('люлька');
    if (notes.includes('детская кроватк')) parts.push('детская кроватка');
    if (notes.includes('доп место') || notes.includes('доп.')) parts.push('доп. место');
    if (notes.includes('диван')) parts.push('диван');

    let comment = parts.join(' + ');
    // Добавляем особые отметки
    if (notes.includes('шампанск')) comment += '; шампанское (др)';

    return comment;
  }

  if (cleaningType.type === '40 выезд') {
    const dep = sections.departures.find(r => r.room === room);
    return dep ? `Выезд — ${dep.guest}` : '';
  }

  if (cleaningType.type === '20') {
    const entry = sections.staying.find(r => r.room === room);
    if (!entry) return '';

    const nightsStayed = Math.round((today - entry.checkin) / (1000 * 60 * 60 * 24));
    const remaining = entry.nights - nightsStayed;

    // Смена белья: гость живёт 2+ ночи, остаётся 2+, не слишком долгий заезд
    if (nightsStayed >= 2 && remaining >= 2 && nightsStayed < 7) {
      return 'смена белья';
    }
    return '';
  }

  return '';
}

function getGuestCount(room, sections) {
  for (const section of ['arrivals', 'staying']) {
    const entries = sections[section].filter(r => r.room === room);
    if (entries.length === 0) continue;

    let totalAdults = 0, totalChildren = 0;

    for (const entry of entries) {
      const notes = entry.notes.toLowerCase();
      let adults = 0, children = 0;

      const fullMatch = notes.match(/(\d+)\s*(?:человека?|человек|чел|взр)\s*\S*?\s*[+,]\s*(?:(\d+)\s*)?(?:ребен|реб|дет)/i);
      if (fullMatch) {
        adults = parseInt(fullMatch[1]);
        children = parseInt(fullMatch[2] || '1');
      } else {
        // "Гостей: X" или "X чел/человек/взр"
        const guestMatch = notes.match(/(?:гостей[:\s]+|^|\s)(\d+)\s*(?:чел|человек|взр)/i);
        const guestColon = notes.match(/гостей[:\s]+(\d+)/i);
        if (guestColon) {
          adults = parseInt(guestColon[1]);
        } else if (guestMatch) {
          adults = parseInt(guestMatch[1]);
        }
      }

      if (adults === 0 && children === 0) {
        adults = entry.adults;
      }

      totalAdults = Math.max(totalAdults, adults);
      totalChildren = Math.max(totalChildren, children);
    }

    if (totalAdults > 0 || totalChildren > 0) {
      return totalAdults + (totalChildren > 0 ? `+${totalChildren}реб` : '');
    }
  }
  return '';
}

// ==============================
// РАСПРЕДЕЛЕНИЕ
// ==============================

function getArea(room) {
  if (room >= 112 && room <= 116) return 98;
  return Math.floor(room / 100);
}

function distribute(tasks) {
  const activeTasks = tasks.filter(t => t._cleaning && t._cleaning.type);
  const maids = MAID_COLORS.map(m => ({ ...m, tasks: [], total: 0 }));

  // Группируем задачи по зонам
  const byArea = {};
  for (const t of activeTasks) {
    const a = getArea(t.room);
    if (!byArea[a]) byArea[a] = [];
    byArea[a].push(t);
  }

  // Зоны от большей к меньшей
  const areas = Object.entries(byArea)
    .map(([a, ts]) => ({ area: parseInt(a), tasks: ts, total: ts.reduce((s, t) => s + t._cleaning.minutes, 0) }))
    .sort((a, b) => b.total - a.total);

  const idealPerMaid = activeTasks.reduce((s, t) => s + t._cleaning.minutes, 0) / maids.length;

  for (const grp of areas) {
    const { tasks: areaTasks, total: areaTotal } = grp;
    const onFloor = maids.filter(m => m.tasks.some(t => getArea(t.room) === grp.area));
    const chosen = onFloor.length > 0
      ? onFloor.sort((a, b) => a.total - b.total)[0]
      : maids.sort((a, b) => a.total - b.total)[0];

    // Если зона большая (>1.3 идеала) — подключаем второго
    let secondChosen = null;
    if (areaTotal > idealPerMaid * 1.3) {
      secondChosen = maids.filter(m => m !== chosen).sort((a, b) => a.total - b.total)[0];
    }

    for (const task of areaTasks) {
      const type = task._cleaning.type;
      let pickMaid = chosen;
      if (secondChosen) {
        const cType = chosen.tasks.filter(t => t._cleaning.type === type).length;
        const sType = secondChosen.tasks.filter(t => t._cleaning.type === type).length;
        if (chosen.total > secondChosen.total + 40) pickMaid = secondChosen;
        else if (cType > sType + 1) pickMaid = secondChosen;
        else if (chosen.total > idealPerMaid && cType >= sType) pickMaid = secondChosen;
        else if (cType >= 3) pickMaid = secondChosen;
      }
      pickMaid.tasks.push(task);
      pickMaid.total += task._cleaning.minutes;
    }
  }

  // Балансировка
  for (let iter = 0; iter < 5; iter++) {
    const sorted = [...maids].sort((a, b) => b.total - a.total);
    const heavy = sorted[0], light = sorted[sorted.length - 1];
    if (heavy.total - light.total <= 20) break;
    let best = null, bestScore = Infinity;
    for (let i = 0; i < heavy.tasks.length; i++) {
      const task = heavy.tasks[i], type = task._cleaning.type;
      if (light.tasks.filter(t => t._cleaning.type === type).length >=
          heavy.tasks.filter(t => t._cleaning.type === type).length) continue;
      const nd = Math.abs((heavy.total - task._cleaning.minutes) - (light.total + task._cleaning.minutes));
      const sameArea = light.tasks.some(t => getArea(t.room) === getArea(task.room));
      if (nd + (sameArea ? 0 : 30) < bestScore) {
        bestScore = nd + (sameArea ? 0 : 30);
        best = { idx: i, task };
      }
    }
    if (best) {
      heavy.tasks.splice(best.idx, 1);
      light.tasks.push(best.task);
      heavy.total -= best.task._cleaning.minutes;
      light.total += best.task._cleaning.minutes;
    } else break;
  }
  return maids;
}

// ==============================
// ГЛАВНАЯ
// ==============================

async function main() {
  console.log('Читаю отчёт...');
  const sections = parseEdelweiss(INPUT_FILE);

  const arrivalRooms = [...new Set(sections.arrivals.map(r => r.room))];
  const depRooms = [...new Set(sections.departures.map(r => r.room))];
  const stayRooms = [...new Set(sections.staying.map(r => r.room))];
  console.log(`Заезды: ${arrivalRooms.length}, Выезды: ${depRooms.length}, Проживания: ${stayRooms.length}`);

  const tasks = ALL_ROOMS.map(room => ({
    room,
    _cleaning: getCleaning(room, sections, REPORT_DATE),
    _comment: '',
    _guestCount: '',
  }));

  for (const t of tasks) {
    t._comment = makeComment(t.room, sections, t._cleaning, REPORT_DATE);
    t._guestCount = getGuestCount(t.room, sections);
  }

  console.log('\n=== ГРАФИК ===');
  for (const t of tasks) {
    const type = t._cleaning.type || '-';
    const min = t._cleaning.minutes || '';
    const c = t._comment || '';
    const g = (t._cleaning.type === '10' || t._cleaning.type === '40 выезд/заезд') ? (t._guestCount || '') : '';
    console.log(
      String(t.room).padEnd(5),
      `| ${type}`.padEnd(18),
      `${min ? min + 'мин' : ''}`.padEnd(6),
      c ? `| ${c}` : '',
      g ? `| кол-во: ${g}` : ''
    );
  }

  const maids = distribute(tasks);

  // Назначим горничную каждой задаче
  const maidByRoom = {};
  for (const maid of maids) {
    for (const t of maid.tasks) {
      maidByRoom[t.room] = maid;
    }
  }
  for (const t of tasks) {
    t._maid = maidByRoom[t.room] || null;
  }

  // === EXCEL ===
  const wb = new ExcelJS.Workbook();

  // ---- ЛИСТ 1: Общий график ----
  const ws1 = wb.addWorksheet('Уборки на сегодня', {
    views: [{ state: 'frozen', ySplit: 1 }]
  });

  ws1.columns = [
    { header: 'Номер',       key: 'номер',    width: 8 },
    { header: '10 минут',    key: 'c10',      width: 10 },
    { header: 'Выезд/Заезд', key: 'c40vz',    width: 14 },
    { header: 'Выезд',       key: 'c40v',     width: 10 },
    { header: '20 минут',    key: 'c20',      width: 10 },
    { header: 'Комментарий', key: 'komm',     width: 40 },
    { header: 'Кол-во чел',  key: 'colvo',    width: 12 },
  ];

  // Шапка
  const hRow1 = ws1.getRow(1);
  hRow1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF34495E' } };
  hRow1.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  hRow1.alignment = { horizontal: 'center', vertical: 'middle' };

  // Сортируем по порядку номеров
  tasks.sort((a, b) => {
    const af = Math.floor(a.room / 100), bf = Math.floor(b.room / 100);
    return af !== bf ? af - bf : a.room - b.room;
  });

  // Маппинг типа уборки → колонка (1-based): Номер=1, 10мин=2, Выезд/Заезд=3, Выезд=4, 20мин=5
  const COL_MAP = { '10': 2, '40 выезд/заезд': 3, '40 выезд': 4, '20': 5 };

  tasks.forEach(t => {
    const type = t._cleaning.type || '';
    const maid = t._maid;
    const colorArgb = maid ? maid.fill : 'FFFFFF';

    const row = ws1.addRow({
      номер: t.room || '',
      c10: type === '10' ? t._cleaning.minutes : '',
      c40vz: type === '40 выезд/заезд' ? t._cleaning.minutes : '',
      c40v: type === '40 выезд' ? t._cleaning.minutes : '',
      c20: type === '20' ? t._cleaning.minutes : '',
      komm: t._comment || '',
      colvo: (type === '10' || type === '40 выезд/заезд') ? (t._guestCount || '') : '',
    });

    // Красим ячейку с цифрой уборки
    const colIdx = COL_MAP[type];
    if (colIdx) {
      const cell = row.getCell(colIdx);
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${colorArgb}` } };
      cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    }
  });

  
// ---- ЛИСТ: Карточки заездов (ключи) ----
  const wsK = wb.addWorksheet('Карточки заездов', {
    views: [{ state: 'frozen', ySplit: 1 }],
    pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true },
  });

  wsK.columns = [
    { width: 6 },   // Комната
    { width: 30 },  // ФИО гостя
    { width: 12 },  // Кол-во
    { width: 8 },   // Ночей
    { width: 14 },  // Заезд
    { width: 14 },  // Выезд
  ];

  const hK = wsK.addRow(['Комната', 'ФИО гостя', 'Кол-во', 'Ночей', 'Заезд', 'Выезд']);
  hK.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2980B9' } };
  hK.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  hK.alignment = { horizontal: 'center', vertical: 'middle' };

  // Собираем уникальные комнаты из заездов
  const arrivalMap = {};
  for (const entry of sections.arrivals) {
    if (!arrivalMap[entry.room]) {
      arrivalMap[entry.room] = { room: entry.room, guests: [], nights: entry.nights, notes: entry.notes };
    }
    if (entry.guest && !arrivalMap[entry.room].guests.includes(entry.guest)) {
      arrivalMap[entry.room].guests.push(entry.guest);
    }
    // Даты заезда/выезда
    if (entry.checkin) {
      const d = entry.checkin;
      arrivalMap[entry.room]._checkin = `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}`;
    }
    if (entry.checkout) {
      const d = entry.checkout;
      arrivalMap[entry.room]._checkout = `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}`;
    }
  }

  const arrivalRoomsSorted = Object.values(arrivalMap).sort((a, b) => a.room - b.room);

  for (const entry of arrivalRoomsSorted) {
    const notes = entry.notes || '';
    const guestCount = getGuestCount(entry.room, sections);
    const guestNames = entry.guests.map(g => g.replace(/\s*\*+$/, '').trim()).join(', ');
    // Даты уже сохранены в entry._checkin / entry._checkout

    wsK.addRow([
      entry.room,
      guestNames,
      guestCount,
      entry.nights || '',
      entry._checkin || '',
      entry._checkout || '',
    ]);
  }

  // ---- ЛИСТЫ КАЖДОЙ ГОРНИЧНОЙ (все номера, раскрашены только её ячейки) ----
  for (const maid of maids) {
    if (maid.tasks.length === 0) continue;

    const wsM = wb.addWorksheet(maid.name, {
      views: [{ state: 'frozen', ySplit: 1 }],
      pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true },
    });

    wsM.columns = [
      { width: 8 },   // Номер
      { width: 10 },  // 10 минут
      { width: 14 },  // Выезд/Заезд
      { width: 10 },  // Выезд
      { width: 10 },  // 20 минут
      { width: 40 },  // Комментарий
      { width: 12 },  // Кол-во чел
    ];

    // Дата (объединённая строка для подписи горничной)
    const dateRow = wsM.addRow([`ДАТА: ${DD}.${MM}.${YY}`, '', '', '', '', '', '']);
    dateRow.height = 28;
    dateRow.font = { size: 14, bold: true };
    dateRow.alignment = { horizontal: 'left', vertical: 'middle' };
    dateRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
    wsM.mergeCells('A1:G1');

    // Шапка
    const hRow = wsM.addRow(['Номер', '10 минут', 'Выезд/Заезд', 'Выезд', '20 минут', 'Комментарий', 'Кол-во чел']);
    hRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF34495E' } };
    hRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    hRow.alignment = { horizontal: 'center', vertical: 'middle' };

    const maidRoomSet = new Set(maid.tasks.map(t => t.room));

    for (const t of tasks) {
      const type = t._cleaning.type || '';
      const isHers = maidRoomSet.has(t.room);

      const row = wsM.addRow([
        t.room || '',
        type === '10' ? t._cleaning.minutes : '',
        type === '40 выезд/заезд' ? t._cleaning.minutes : '',
        type === '40 выезд' ? t._cleaning.minutes : '',
        type === '20' ? t._cleaning.minutes : '',
        t._comment || '',
        (type === '10' || type === '40 выезд/заезд') ? (t._guestCount || '') : '',
      ]);

      row.eachCell((cell, idx) => {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
        if (idx === 6) cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      });

      // Красим ячейку с цифрой, только если задача её
      if (type && isHers) {
        const colIdx = COL_MAP[type];
        if (colIdx) {
          const cell = row.getCell(colIdx);
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${maid.fill}` } };
          cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
        }
      }
    }

    // Итог
    const total = maid.tasks.reduce((s, t) => s + (t._cleaning.minutes || 0), 0);
    const itogRow = wsM.addRow(['ИТОГО', '', '', '', total, `${maid.tasks.length} задач`, '']);
    itogRow.font = { bold: true };

    // Строка загрузки
    const types = {};
    for (const t of maid.tasks) {
      const tp = t._cleaning.type;
      types[tp] = (types[tp] || 0) + 1;
    }
    const parts = Object.entries(types).map(([k, v]) => `${k}: ${v}`).join(', ');
    wsM.addRow([`Загрузка: ${maid.tasks.length} задач / ${total} мин`, '', '', '', '', parts, ''])
      .font = { italic: true, color: { argb: 'FF555555' }, size: 10 };
  }

    // ---- ЛИСТ 2: По горничным ----
  const ws2 = wb.addWorksheet('По горничным', {
    views: [{ state: 'frozen', ySplit: 1 }]
  });

  ws2.columns = [
    { header: 'Номер',       key: 'номер',     width: 8 },
    { header: '10 минут',    key: 'c10',       width: 10 },
    { header: 'Выезд/Заезд', key: 'c40vz',     width: 14 },
    { header: 'Выезд',       key: 'c40v',      width: 10 },
    { header: '20 минут',    key: 'c20',       width: 10 },
,
    { header: 'Мин',         key: 'min',       width: 6 },
    { header: 'Комментарий', key: 'komm',      width: 40 },
    { header: 'Кол-во чел',  key: 'colvo',     width: 12 },
  ];

  const hRow2 = ws2.getRow(1);
  hRow2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF34495E' } };
  hRow2.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  hRow2.alignment = { horizontal: 'center', vertical: 'middle' };

  const COL_MAP2 = { '10': 2, '40 выезд/заезд': 3, '40 выезд': 4, '20': 5 };

  for (const maid of maids) {
    if (maid.tasks.length === 0) continue;
    maid.tasks.sort((a, b) => a.room - b.room);

    for (const task of maid.tasks) {
      const type = task._cleaning.type || '';
      const colorArgb = `FF${maid.fill}`;

      const row = ws2.addRow({
        номер: task.room,
        c10: type === '10' ? task._cleaning.minutes : '',
        c40vz: type === '40 выезд/заезд' ? task._cleaning.minutes : '',
        c40v: type === '40 выезд' ? task._cleaning.minutes : '',
        c20: type === '20' ? task._cleaning.minutes : '',
        min: task._cleaning.minutes || '',
        komm: task._comment || '',
        colvo: (type === '10' || type === '40 выезд/заезд') ? (task._guestCount || '') : '',
      });

      // Красим ячейку с цифрой
      const colIdx = COL_MAP2[type];
      if (colIdx) {
        const cell = row.getCell(colIdx);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colorArgb } };
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 12 };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      }
    }

    // Итог по горничной
    const totalM = maid.tasks.reduce((s, t) => s + (t._cleaning.minutes || 0), 0);
    const itog = ws2.addRow(['', '', '', '', '', totalM, `${maid.tasks.length} задач`, '']);
    itog.font = { bold: true };
    ws2.addRow({}); // разделитель
  }

 

await wb.xlsx.writeFile(OUTPUT_FILE);
  console.log(`\nГотово: ${OUTPUT_FILE}`);

  console.log('\n=== ИТОГО ===');
  for (const maid of maids) {
    if (maid.tasks.length > 0) {
      const rooms = maid.tasks.map(t => t.room).join(', ');
      console.log(`${maid.name}: ${maid.tasks.length} задач, ${maid.total} мин — ${rooms}`);
    }
  }
}

main().catch(console.error);
