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
const STANDARD_CHECKOUT_HOUR = 12;

const DATA_DIR = path.join(__dirname, 'отчёты_из_эдельвейса');
const OUT_DIR = path.join(__dirname, 'готовые_отчёты');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

function findInputFile() {
  const todayDD = String(new Date().getDate()).padStart(2, '0');
  const todayMM = String(new Date().getMonth() + 1).padStart(2, '0');
  const exact = path.join(DATA_DIR, `${todayDD}.${todayMM}.xls`);
  if (fs.existsSync(exact)) return exact;
  const files = fs.readdirSync(DATA_DIR).filter(f => f.endsWith('.xls') && !f.includes('~$'));
  if (files.length > 0) return path.join(DATA_DIR, files[0]);
  console.error(`❌ Файл ${todayDD}.${todayMM}.xls не найден в папке отчёты_из_эдельвейса`);
  process.exit(1);
}

/** Извлечь дату из имени файла (04.07.xls → { dd: '04', mm: '07' }) */
function parseDateFromFilename(filePath) {
  const name = path.basename(filePath);
  const m = name.match(/^(\d{2})\.(\d{2})\.xls/);
  if (m) return { dd: m[1], mm: m[2] };
  // если имя не matches — используем сегодня
  const d = new Date();
  return {
    dd: String(d.getDate()).padStart(2, '0'),
    mm: String(d.getMonth() + 1).padStart(2, '0'),
  };
}

const INPUT_FILE = findInputFile();
const { dd: DD, mm: MM } = parseDateFromFilename(INPUT_FILE);

const CURRENT_YEAR = new Date().getFullYear();
const REPORT_DATE = new Date(CURRENT_YEAR, parseInt(MM) - 1, parseInt(DD));
const YY = String(CURRENT_YEAR).slice(2);

let ver = 1;
const dateFolder = `${DD}.${MM}.${YY}`;
const dateOutDir = path.join(OUT_DIR, dateFolder);
if (!fs.existsSync(dateOutDir)) fs.mkdirSync(dateOutDir, { recursive: true });
while (fs.existsSync(path.join(dateOutDir, `виды-уборок-${DD}.${MM}.${YY}-v${ver}.xlsx`))) ver++;
const OUTPUT_FILE = path.join(dateOutDir, `виды-уборок-${DD}.${MM}.${YY}-v${ver}.xlsx`);

console.log(`📅 ${DD}.${MM}.${REPORT_DATE.getFullYear()}`);
console.log(`📥 ${path.basename(INPUT_FILE)} → 📤 ${path.basename(OUTPUT_FILE)}`);

// ==============================
// ПАРСИНГ ОТЧЁТА
// ==============================

function parseDateRange(rangeStr) {
  const m = String(rangeStr || '').match(/(\d{2})\.(\d{2})\s*\((\d{2}):(\d{2})\)\s*-\s*(\d{2})\.(\d{2})\s*\((\d{2}):(\d{2})\)/);
  if (!m) return { checkin: null, checkout: null, checkoutHour: null, checkoutMinute: null };
  return {
    checkin: new Date(CURRENT_YEAR, parseInt(m[2]) - 1, parseInt(m[1])),
    checkout: new Date(CURRENT_YEAR, parseInt(m[6]) - 1, parseInt(m[5])),
    checkoutHour: parseInt(m[7]),
    checkoutMinute: parseInt(m[8]),
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
      notes,
      nights: parseInt(rawRow[9]) || 0,
      adults: parseInt(rawRow[10]) || 0,
      checkin: range.checkin,
      checkout: range.checkout,
      checkoutHour: range.checkoutHour,
      checkoutMinute: range.checkoutMinute,
      status,
    });
  }

  return sections;
}

// ==============================
// ОПРЕДЕЛЕНИЕ ТИПОВ УБОРОК
// ==============================

function getArea(room) {
  if (room >= 112 && room <= 116) return 1.5;
  return Math.floor(room / 100);
}

function getCleaning(room, sections) {
  const isArrival = sections.arrivals.some(r => r.room === room);
  const isDeparture = sections.departures.some(r => r.room === room);

  if (isArrival && isDeparture) return { type: '40 выезд/заезд', minutes: 40 };
  if (isDeparture) return { type: '40 выезд', minutes: 40 };
  if (isArrival) return { type: '10', minutes: 10 };

  const staying = sections.staying.filter(r => r.room === room);
  if (staying.length > 0) return { type: '20', minutes: 20 };

  return { type: null, minutes: 0 };
}

function getLateCheckoutStr(entry) {
  if (!entry || entry.checkoutHour == null) return '';
  if (entry.checkoutHour > STANDARD_CHECKOUT_HOUR) {
    const h = String(entry.checkoutHour).padStart(2, '0');
    const m = String(entry.checkoutMinute).padStart(2, '0');
    return `поздний выезд до ${h}:${m}`;
  }
  return '';
}

function makeComment(room, sections, cleaningType) {
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

    // доп. места если гостей больше 2
    const gc = getGuestCountSimple(room, sections);
    if (gc) {
      const total = gc.split('+').reduce((sum, p) => sum + (parseInt(p) || 0), 0);
      if (total > 2) parts.push('доп. места');
    }

    let comment = parts.join(' + ');
    if (notes.includes('шампанск')) comment += '; шампанское (др)';

    // для выезд/заезд — добавить информацию о позднем выезде
    if (cleaningType.type === '40 выезд/заезд') {
      const dep = sections.departures.find(r => r.room === room);
      const lateStr = getLateCheckoutStr(dep);
      if (lateStr) comment += `; ${lateStr}`;
    }

    return comment;
  }

  if (cleaningType.type === '40 выезд') {
    const dep = sections.departures.find(r => r.room === room);
    return getLateCheckoutStr(dep);
  }

  if (cleaningType.type === '20') {
    const entry = sections.staying.find(r => r.room === room);
    if (!entry) return '';
    const nightsStayed = Math.round((REPORT_DATE - entry.checkin) / (1000 * 60 * 60 * 24));
    const remaining = entry.nights - nightsStayed;
    if (nightsStayed >= 2 && nightsStayed % 2 === 0 && remaining >= 2) return 'смена белья';
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
        const guestColon = notes.match(/гостей[:\s]+(\d+)/i);
        const guestMatch = notes.match(/(?:^|[^\d])(\d+)\s*(?:чел|человек|взр)/i);
        if (guestColon) adults = parseInt(guestColon[1]);
        else if (guestMatch) adults = parseInt(guestMatch[1]);
      }

      if (adults === 0 && children === 0) adults = entry.adults;
      totalAdults = Math.max(totalAdults, adults);
      totalChildren = Math.max(totalChildren, children);
    }
    if (totalAdults > 0 || totalChildren > 0)
      return totalAdults + (totalChildren > 0 ? `+${totalChildren}реб` : '');
  }
  return '';
}

// ==============================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ КАРТОЧЕК
// ==============================

function getGuestCountSimple(room, sections) {
  return getGuestCount(room, sections);
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

  const ALL_ROOMS = [
    101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 112, 113, 114, 115, 116,
    201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 211, 212, 213, 214, 215, 216, 217, 218, 219, 220, 221, 222, 223, 224, 225
  ];

  const tasks = ALL_ROOMS.map(room => ({
    room,
    _cleaning: getCleaning(room, sections),
    _comment: '',
    _guestCount: '',
  }));

  for (const t of tasks) {
    t._comment = makeComment(t.room, sections, t._cleaning);
    t._guestCount = (t._cleaning.type === '10' || t._cleaning.type === '40 выезд/заезд') ? (getGuestCountSimple(t.room, sections) || '') : '';
  }

  console.log('\n=== ГРАФИК ===');
  for (const t of tasks) {
    const type = t._cleaning.type || '-';
    const min = t._cleaning.minutes || '';
    const c = t._comment || '';
    const g = t._guestCount || '';
    console.log(
      String(t.room).padEnd(5),
      `| ${type}`.padEnd(18),
      `${min ? min + 'мин' : ''}`.padEnd(6),
      c ? `| ${c}` : '',
      g ? `| кол-во: ${g}` : ''
    );
  }

  // === EXCEL ===
  const wb = new ExcelJS.Workbook();

  // ---- ЛИСТ 1: Уборки на сегодня ----
  const ws1 = wb.addWorksheet('Уборки на сегодня', { views: [{ state: 'frozen', ySplit: 3 }] });
  ws1.columns = [
    { width: 8 }, { width: 10 }, { width: 14 }, { width: 10 }, { width: 10 }, { width: 40 }, { width: 12 },
  ];

  // Заголовок (строчки 1-2 объединены)
  ws1.addRow(['']); // строка 1
  ws1.addRow(['']); // строка 2
  ws1.mergeCells(1, 1, 2, 7);
  const titleCell = ws1.getCell(1, 1);
  titleCell.value = `${DD}.${MM}.${YY}`;
  titleCell.font = { size: 25, bold: true };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

  const hRow1 = ws1.addRow(['Номер', '10 минут', 'Выезд/Заезд', 'Выезд', '20 минут', 'Комментарий', 'Кол-во чел']);
  hRow1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF34495E' } };
  hRow1.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
  hRow1.alignment = { horizontal: 'center', vertical: 'middle' };

  tasks.sort((a, b) => {
    const af = getArea(a.room), bf = getArea(b.room);
    return af !== bf ? af - bf : a.room - b.room;
  });

  for (const t of tasks) {
    const type = t._cleaning.type || '';
    const row = ws1.addRow([
      t.room || '',
      type === '10' ? t._cleaning.minutes : '',
      type === '40 выезд/заезд' ? t._cleaning.minutes : '',
      type === '40 выезд' ? t._cleaning.minutes : '',
      type === '20' ? t._cleaning.minutes : '',
      t._comment || '',
      t._guestCount || '',
    ]);
    row.alignment = { horizontal: 'center', vertical: 'middle' };
  }

  // ---- ЛИСТ 2: Карточки заездов ----
  const wsK = wb.addWorksheet('Карточки заездов', { views: [{ state: 'frozen', ySplit: 1 }] });
  wsK.columns = [{ width: 6 }, { width: 30 }, { width: 12 }, { width: 8 }, { width: 14 }, { width: 14 }];

  const hK = wsK.addRow(['Комната', 'ФИО гостя', 'Кол-во', 'Ночей', 'Заезд', 'Выезд']);
  hK.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2980B9' } };
  hK.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };

  const arrivalMap = {};
  for (const entry of sections.arrivals) {
    if (!arrivalMap[entry.room]) {
      arrivalMap[entry.room] = { room: entry.room, guests: [], nights: entry.nights, notes: entry.notes };
    }
    if (entry.guest && !arrivalMap[entry.room].guests.includes(entry.guest))
      arrivalMap[entry.room].guests.push(entry.guest);
    if (entry.checkin) {
      const d = entry.checkin;
      arrivalMap[entry.room]._checkin = `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}`;
    }
    if (entry.checkout) {
      const d = entry.checkout;
      arrivalMap[entry.room]._checkout = `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}`;
    }
  }

  const arrivalRoomsSorted = Object.values(arrivalMap).sort((a, b) => a.room - b.room);
  let totalGuests = 0;
  for (const entry of arrivalRoomsSorted) {
    const guestCount = getGuestCountSimple(entry.room, sections);
    const guestNames = entry.guests.map(g => g.replace(/\s*\*+$/, '').trim()).join(', ');
    const row = wsK.addRow([entry.room, guestNames, guestCount, entry.nights || '', entry._checkin || '', entry._checkout || '']);
    row.alignment = { horizontal: 'center', vertical: 'middle' };

    // подсчёт общего числа гостей
    if (guestCount) {
      totalGuests += guestCount.split('+').reduce((sum, p) => sum + (parseInt(p) || 0), 0);
    }
  }

  // итоговая строка с количеством карточек
  wsK.addRow([]);
  const totalRow = wsK.addRow([`Итого карточек: ${totalGuests}`]);
  totalRow.font = { bold: true, size: 12 };
  totalRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };

  // === РАСПРЕДЕЛЕНИЕ ПО ГОРНИЧНЫМ ===
  const readline = require('readline');
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  const numHK = await new Promise(resolve => {
    rl.question(`Сколько горничных? (по умолчанию 3): `, answer => {
      resolve(parseInt(answer) || 3);
    });
  });

  const hkNames = [];
  for (let i = 0; i < numHK; i++) {
    const name = await new Promise(resolve => {
      rl.question(`Имя горничной ${i + 1}: `, answer => {
        resolve(answer.trim() || `Горничная ${i + 1}`);
      });
    });
    hkNames.push(name);
  }
  rl.close();

  // Инициализация горничных
  const hkTasks = Array.from({ length: numHK }, () => []);
  const hkMins = new Array(numHK).fill(0);
  const hkVyzdyZaezd = new Array(numHK).fill(0);

  function assignToHK(hkIdx, task) {
    hkTasks[hkIdx].push(task);
    hkMins[hkIdx] += task._cleaning.minutes || 0;
    if (task._cleaning.type === '40 выезд/заезд') hkVyzdyZaezd[hkIdx]++;
  }

  const allMinutes = tasks.reduce((s, t) => s + (t._cleaning.minutes || 0), 0);
  const targetPerHK = allMinutes / numHK;

  console.log(`\n=== РАСПРЕДЕЛЕНИЕ (${numHK} горничных, цель ~${Math.round(targetPerHK)} мин/чел) ===`);

  // Шаг 1: зона 1.5 — горничной 1
  const zone15Tasks = tasks.filter(t => getArea(t.room) === 1.5);
  let mainTasks = tasks.filter(t => getArea(t.room) !== 1.5).sort((a, b) => a.room - b.room);

  if (zone15Tasks.length > 0 && numHK > 1) {
    for (const t of zone15Tasks.sort((a, b) => a.room - b.room)) assignToHK(0, t);
    console.log(`Зона 1.5 (корпус 112-116): горничная 1 (${hkMins[0]} мин)`);
  }

  // Шаг 2: сороковки — равномерно по счёту, потом компенсация двадцатками
  const allForties = tasks.filter(t =>
    t._cleaning.type === '40 выезд' || t._cleaning.type === '40 выезд/заезд'
  ).sort((a, b) => a.room - b.room);

  const totalForties = allForties.length;
  const base40 = Math.floor(totalForties / numHK);
  const rem40 = totalForties % numHK;

  // Сколько сороковок у каждой уже есть (из зоны 1.5)
  const fortyNeeded = [];
  for (let i = 0; i < numHK; i++) {
    const existing = hkTasks[i].filter(t =>
      t._cleaning.type === '40 выезд' || t._cleaning.type === '40 выезд/заезд'
    ).length;
    const target = base40 + (i < rem40 ? 1 : 0);
    fortyNeeded.push(Math.max(0, target - existing));
  }

  // Дополнительные сороковки из основного списка
  const extraForties = allForties.filter(t => !zone15Tasks.includes(t));

  for (let i = 0; i < numHK; i++) {
    for (let j = 0; j < fortyNeeded[i]; j++) {
      if (extraForties.length === 0) break;
      assignToHK(i, extraForties.shift());
    }
  }
  // Остатки — наименее загруженной
  while (extraForties.length > 0) {
    const least = hkMins.indexOf(Math.min(...hkMins));
    assignToHK(least, extraForties.shift());
  }

  // Считаем deficit: кому не хватает сороковок — компенсируем 2×20
  const fortyDebt = [];
  for (let i = 0; i < numHK; i++) {
    const has40 = hkTasks[i].filter(t =>
      t._cleaning.type === '40 выезд' || t._cleaning.type === '40 выезд/заезд'
    ).length;
    fortyDebt.push(Math.max(0, base40 - has40));
  }

  if (rem40 > 0) {
    console.log(`Сороковок: ${totalForties} (${base40}×${numHK} + ${rem40})`);
  } else {
    console.log(`Сороковок: ${totalForties} (по ${base40})`);
  }
  if (fortyDebt.some(d => d > 0)) {
    console.log(`Компенсация двадцатками: ${fortyDebt.map((d, i) => `Горн.${i+1} +${d*2}×20`).filter(s => !s.includes('+0')).join(', ')}`);
  }

  // Шаг 3: двадцатки и десятки — каждая добирает до целевых минут
  let twentyTenTasks = tasks.filter(t => {
    const type = t._cleaning.type;
    return type === '20' || type === '10';
  }).sort((a, b) => a.room - b.room);

  // Убираем те, что уже назначены (из зоны 1.5)
  const assignedRooms = new Set();
  for (let i = 0; i < numHK; i++) {
    for (const t of hkTasks[i]) assignedRooms.add(t.room);
  }
  twentyTenTasks = twentyTenTasks.filter(t => !assignedRooms.has(t.room));

  // Сначала компенсационные двадцатки для тех, кому не хватило сороковок
  for (let i = 0; i < numHK; i++) {
    for (let j = 0; j < fortyDebt[i] * 2; j++) {
      let bestIdx = 0;
      if (hkTasks[i].length > 0) {
        const myRooms = hkTasks[i].map(t => t.room);
        let bestDist = Infinity;
        for (let k = 0; k < twentyTenTasks.length; k++) {
          if (twentyTenTasks[k]._cleaning.type !== '20') continue;
          const dist = Math.min(...myRooms.map(r => Math.abs(r - twentyTenTasks[k].room)));
          if (dist < bestDist) { bestDist = dist; bestIdx = k; }
        }
      }
      if (twentyTenTasks.length > 0 && twentyTenTasks[bestIdx]._cleaning.type === '20') {
        assignToHK(i, twentyTenTasks.splice(bestIdx, 1)[0]);
      }
    }
  }

  // Остаток — блоками до целевых минут каждой горничной
  if (twentyTenTasks.length > 0) {
    // сколько каждой не хватает до общей цели
    const targets = [];
    for (let i = 0; i < numHK; i++) {
      targets.push(Math.max(0, targetPerHK - hkMins[i]));
    }

    let ptr = 0;
    for (let i = 0; i < numHK; i++) {
      let blockMin = 0;
      const isLast = (i === numHK - 1);
      while (ptr < twentyTenTasks.length) {
        const taskMin = twentyTenTasks[ptr]._cleaning.minutes || 0;
        if (!isLast && blockMin >= targets[i]) break;
        assignToHK(i, twentyTenTasks[ptr]);
        blockMin += taskMin;
        ptr++;
      }
    }
  }

  // Микро-корректировка: обмен задачами между богатой и бедной горничными
  for (let iter = 0; iter < 5; iter++) {
    let maxIdx = 0, minIdx = 0;
    for (let i = 0; i < numHK; i++) {
      if (hkMins[i] > hkMins[maxIdx]) maxIdx = i;
      if (hkMins[i] < hkMins[minIdx]) minIdx = i;
    }
    const diff = hkMins[maxIdx] - hkMins[minIdx];
    if (diff <= 10) break;

    // Ищем обмен: богатая отдаёт X, бедная отдаёт Y, чтобы diff уменьшился
    let bestSwap = null; // { richIdx, poorIdx, richTask, poorTask }
    let bestRoomDist = Infinity;

    for (let ri = 0; ri < hkTasks[maxIdx].length; ri++) {
      const rTask = hkTasks[maxIdx][ri];
      const rMin = rTask._cleaning.minutes || 0;
      for (let pi = 0; pi < hkTasks[minIdx].length; pi++) {
        const pTask = hkTasks[minIdx][pi];
        const pMin = pTask._cleaning.minutes || 0;

        // После обмена
        const newRich = hkMins[maxIdx] - rMin + pMin;
        const newPoor = hkMins[minIdx] + rMin - pMin;
        const newDiff = Math.abs(newRich - newPoor);

        if (newDiff < diff) {
          // Не ломаем счёт сороковок — только 20↔20, 20↔10, 10↔10
          const rIs40 = (rTask._cleaning.type === '40 выезд' || rTask._cleaning.type === '40 выезд/заезд');
          const pIs40 = (pTask._cleaning.type === '40 выезд' || pTask._cleaning.type === '40 выезд/заезд');
          if (rIs40 !== pIs40) continue;
          const roomDist = Math.abs(rTask.room - pTask.room);
          if (roomDist < bestRoomDist) {
            bestRoomDist = roomDist;
            bestSwap = { ri, pi, maxIdx, minIdx, rMin, pMin };
          }
        }
      }
    }

    if (bestSwap && bestRoomDist < 50) {
      const { ri, pi } = bestSwap;
      const richTask = hkTasks[maxIdx][ri];
      const poorTask = hkTasks[minIdx][pi];

      hkTasks[maxIdx][ri] = poorTask;
      hkTasks[minIdx][pi] = richTask;
      hkMins[maxIdx] = hkMins[maxIdx] - (richTask._cleaning.minutes || 0) + (poorTask._cleaning.minutes || 0);
      hkMins[minIdx] = hkMins[minIdx] - (poorTask._cleaning.minutes || 0) + (richTask._cleaning.minutes || 0);
    } else {
      // Обмен не нашёлся — пробуем перенести одну задачу (не сороковку)
      let bestMoveIdx = -1;
      let bestMoveDist = Infinity;
      for (let ri = 0; ri < hkTasks[maxIdx].length; ri++) {
        const t = hkTasks[maxIdx][ri];
        const tMin = t._cleaning.minutes || 0;
        if (tMin > diff) continue;
        const is40 = (t._cleaning.type === '40 выезд' || t._cleaning.type === '40 выезд/заезд');
        if (is40) continue;
        const dist = Math.min(...hkTasks[minIdx].map(pt => Math.abs(t.room - pt.room)));
        if (dist < bestMoveDist) { bestMoveDist = dist; bestMoveIdx = ri; }
      }
      if (bestMoveIdx !== -1 && bestMoveDist < 50) {
        const task = hkTasks[maxIdx].splice(bestMoveIdx, 1)[0];
        hkTasks[minIdx].push(task);
        hkMins[maxIdx] -= task._cleaning.minutes || 0;
        hkMins[minIdx] += task._cleaning.minutes || 0;
      } else break;
    }
  }

  // Проверка макс 2 выезд/заезд на горничную — при необходимости меняем
  for (let i = 0; i < numHK; i++) {
    let attempts = 0;
    while (hkVyzdyZaezd[i] > 2 && attempts < 10) {
      attempts++;
      // ищем выезд/заезд у этой горничной
      const swapIdx = hkTasks[i].findIndex(t => t._cleaning.type === '40 выезд/заезд');
      if (swapIdx === -1) break;
      // ищем другую горничную с < 2 выезд/заезд, предпочитая обмен на простой выезд
      let best = null;
      for (let j = 0; j < numHK; j++) {
        if (j === i || hkVyzdyZaezd[j] >= 2) continue;
        const their40 = hkTasks[j].findIndex(t => t._cleaning.type === '40 выезд');
        if (their40 !== -1) { best = { j, idx: their40 }; break; }
      }
      if (best) {
        const tmp = hkTasks[i][swapIdx];
        hkTasks[i][swapIdx] = hkTasks[best.j][best.idx];
        hkTasks[best.j][best.idx] = tmp;
        hkVyzdyZaezd[i]--;
        hkVyzdyZaezd[best.j]++;
      } else break;
    }
  }

  // --- Вывод ---
  if (zone15Tasks.length > 0) {
    console.log(`Зона 1.5 (корпус 112-116): горничная 1 (${hkMins[0]} мин)`);
  }
  for (let i = 0; i < numHK; i++) {
    const sorted = [...hkTasks[i]].sort((a, b) => a.room - b.room);
    const zones = [...new Set(sorted.map(t => getArea(t.room)))].sort().join(', ');
    console.log(`\nГорничная ${i + 1}: ${hkMins[i]} мин (зоны: ${zones})`);
    for (const t of sorted) {
      const type = t._cleaning.type || '-';
      const min = t._cleaning.minutes || '';
      const c = t._comment || '';
      const g = t._guestCount || '';
      console.log(`  ${t.room} | ${type} | ${min}мин${c ? ' | ' + c : ''}${g ? ' | кол-во: ' + g : ''}`);
    }
  }

  // --- Цветовая маркировка ячеек с цифрами уборки ---
  const HK_COLORS = [
    'FFE74C3C', // красный
    'FF3498DB', // синий
    'FF2ECC71', // зелёный
    'FF9B59B6', // фиолетовый (тёмный, виден на Ч/Б печати)
    'FFF39C12', // оранжевый
    'FF1ABC9C', // бирюзовый
  ];
  const HK_LABELS = ['Красный', 'Синий', 'Зелёный', 'Фиолетовый', 'Оранжевый', 'Бирюзовый'];

  // Сопоставляем номер комнаты → (индекс горничной, тип уборки)
  const roomInfo = {};
  for (let i = 0; i < numHK; i++) {
    for (const t of hkTasks[i]) {
      roomInfo[t.room] = { hkIdx: i, type: t._cleaning.type };
    }
  }

  // Какая колонка отвечает за какой тип уборки (1-based)
  const typeToCol = { '10': 2, '40 выезд/заезд': 3, '40 выезд': 4, '20': 5 };

  // Красим только ячейку с минутами (данные начинаются с row 4)
  for (let r = 4; r <= ws1.rowCount; r++) {
    const room = ws1.getCell(r, 1).value;
    if (room && roomInfo[room]) {
      const { hkIdx, type } = roomInfo[room];
      const col = typeToCol[type];
      if (col) {
        ws1.getCell(r, col).fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: HK_COLORS[hkIdx % HK_COLORS.length] },
        };
        ws1.getCell(r, col).font = { bold: true, color: { argb: 'FF000000' }, size: 11 };
      }
    }
  }

  // Легенда под таблицей
  ws1.addRow([]);
  ws1.addRow(['Распределение по горничным:']);
  for (let i = 0; i < numHK; i++) {
    // считаем разбивку по типам уборок
    const typeCounts = {};
    for (const t of hkTasks[i]) {
      const mins = t._cleaning.minutes || 0;
      if (mins > 0) {
        typeCounts[mins] = (typeCounts[mins] || 0) + 1;
      }
    }
    const breakdown = Object.entries(typeCounts)
      .sort((a, b) => Number(b[0]) - Number(a[0]))
      .map(([min, cnt]) => `${cnt}×${min}`)
      .join(', ');

    const legendRow = ws1.addRow([`${HK_LABELS[i]} — горничная ${i + 1} (${hkMins[i]} мин, ${breakdown})`]);
    legendRow.getCell(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: HK_COLORS[i % HK_COLORS.length] },
    };
    legendRow.getCell(1).font = { bold: true, color: { argb: 'FF000000' }, size: 10 };
  }

  // --- Индивидуальные листы для каждой горничной (все строки, но цвет только её) ---
  const hkSheetNames = ['Горничная 1', 'Горничная 2', 'Горничная 3', 'Горничная 4'];

  for (let hk = 0; hk < numHK; hk++) {
    const wsHk = wb.addWorksheet(hkNames[hk], { views: [{ state: 'frozen', ySplit: 3 }] });
    wsHk.columns = ws1.columns.map(c => ({ width: c.width }));

    // Заголовок с датой и именем
    wsHk.addRow(['']);
    wsHk.addRow(['']);
    wsHk.mergeCells(1, 1, 2, 7);
    const hkTitle = wsHk.getCell(1, 1);
    hkTitle.value = `${DD}.${MM}.${YY} — ${hkNames[hk]}`;
    hkTitle.font = { size: 25, bold: true };
    hkTitle.alignment = { horizontal: 'center', vertical: 'middle' };

    const hRow = wsHk.addRow(['Номер', '10 минут', 'Выезд/Заезд', 'Выезд', '20 минут', 'Комментарий', 'Кол-во чел']);
    hRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF34495E' } };
    hRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    hRow.alignment = { horizontal: 'center', vertical: 'middle' };

    const hkRoomSet = new Set(hkTasks[hk].map(t => t.room));
    const hkRoomInfo = {};
    for (const t of hkTasks[hk]) hkRoomInfo[t.room] = t._cleaning.type;

    const color = HK_COLORS[hk % HK_COLORS.length];

    for (const t of tasks) {
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

      // Цвет только если задача этой горничной
      if (hkRoomSet.has(t.room)) {
        const col = typeToCol[type];
        if (col) {
          wsHk.getCell(row.number, col).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: color },
          };
          wsHk.getCell(row.number, col).font = { bold: true, color: { argb: 'FF000000' }, size: 11 };
        }
      }
    }
  }

  await wb.xlsx.writeFile(OUTPUT_FILE);
  console.log(`\nГотово: ${OUTPUT_FILE}`);

  // перемещаем исходный файл в "старые"
  const archiveDir = path.join(DATA_DIR, 'старые');
  if (!fs.existsSync(archiveDir)) fs.mkdirSync(archiveDir, { recursive: true });
  const destPath = path.join(archiveDir, path.basename(INPUT_FILE));
  fs.renameSync(INPUT_FILE, destPath);
  console.log(`📦 Исходный файл перемещён в папку "старые"`);
}

main().catch(console.error);
