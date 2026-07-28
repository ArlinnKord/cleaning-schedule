const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const STANDARD_CHECKOUT_HOUR = 12;

const DATA_DIR = path.join(__dirname, 'отчёты_из_travelline');
const OUT_DIR = path.join(__dirname, 'готовые_отчёты');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

// ==============================
// ПОИСК ФАЙЛОВ
// ==============================

function findReportFile(prefix) {
  const files = fs.readdirSync(DATA_DIR)
    .filter(f => f.startsWith(prefix) && f.endsWith('.xlsx') && !f.includes('~$'))
    .sort()
    .reverse();
  if (files.length === 0) {
    console.error(`❌ Файл ${prefix}...xlsx не найден в папке отчёты_из_travelline`);
    process.exit(1);
  }
  return path.join(DATA_DIR, files[0]);
}

const ARRIVALS_FILE = findReportFile('ArrivalsReport');
const DEPARTURES_FILE = findReportFile('DeparturesReport');
const OCCUPIED_FILE = findReportFile('OccupiedRoomTypes');

console.log(`📥 Заезды: ${path.basename(ARRIVALS_FILE)}`);
console.log(`📥 Выезды: ${path.basename(DEPARTURES_FILE)}`);
console.log(`📥 Занятые: ${path.basename(OCCUPIED_FILE)}`);

const dateMatch = path.basename(ARRIVALS_FILE).match(/(\d{2})\.(\d{2})\.(\d{4})/);
const DD = dateMatch ? dateMatch[1] : String(new Date().getDate()).padStart(2, '0');
const MM = dateMatch ? dateMatch[2] : String(new Date().getMonth() + 1).padStart(2, '0');
const YYYY = dateMatch ? dateMatch[3] : String(new Date().getFullYear());
const YY = YYYY.slice(2);
const REPORT_DATE = new Date(parseInt(YYYY), parseInt(MM) - 1, parseInt(DD));

console.log(`📅 ${DD}.${MM}.${YYYY}`);

let ver = 1;
const dateFolder = `${DD}.${MM}.${YY}`;
const dateOutDir = path.join(OUT_DIR, dateFolder);
if (!fs.existsSync(dateOutDir)) fs.mkdirSync(dateOutDir, { recursive: true });
while (fs.existsSync(path.join(dateOutDir, `виды-уборок-${DD}.${MM}.${YY}-v${ver}.xlsx`))) ver++;
const OUTPUT_FILE = path.join(dateOutDir, `виды-уборок-${DD}.${MM}.${YY}-v${ver}.xlsx`);
console.log(`📤 ${path.basename(OUTPUT_FILE)}`);

// ==============================
// ПАРСИНГ ДАТЫ
// ==============================

function parseTLDate(str) {
  const m = String(str || '').match(/(\d{2})\.(\d{2})\.(\d{4}),\s*(\d{1,2}):(\d{2})/);
  if (!m) return { date: null, hour: null, minute: null };
  return {
    date: new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1])),
    hour: parseInt(m[4]),
    minute: parseInt(m[5]),
  };
}

// ==============================
// ПАРСИНГ ОТЧЁТОВ
// ==============================

function readSheet(filePath) {
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets['Report'];
  const data = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 });

  // найдём заголовок по № комнаты
  const headerRow = data[0];
  const roomCol = headerRow ? headerRow.indexOf('№ комнаты') : -1;
  if (roomCol === -1) {
    console.error(`❌ Колонка "№ комнаты" не найдена в ${path.basename(filePath)}`);
    process.exit(1);
  }

  return { data, roomCol };
}

function parseArrivals(filePath) {
  const { data, roomCol } = readSheet(filePath);
  const entries = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[roomCol]) continue;
    const room = parseInt(row[roomCol]);
    if (isNaN(room)) continue;

    const checkin = parseTLDate(row[6]);
    const checkout = parseTLDate(row[7]);
    const notes = [String(row[16] || ''), String(row[17] || '')].join(' ').trim();

    entries.push({
      room, guest: String(row[1] || '').replace(/\s*\*+$/, '').trim(),
      guestCount: parseInt(row[3]) || 0,
      checkin: checkin.date, checkout: checkout.date,
      checkoutHour: checkout.hour, checkoutMinute: checkout.minute,
      notes, status: String(row[11] || ''),
    });
  }
  return entries;
}

function parseDepartures(filePath) {
  const { data, roomCol } = readSheet(filePath);
  const entries = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[roomCol]) continue;
    const room = parseInt(row[roomCol]);
    if (isNaN(room)) continue;

    const checkin = parseTLDate(row[4]);
    const checkout = parseTLDate(row[5]);
    const notes = [String(row[13] || ''), String(row[14] || '')].join(' ').trim();

    entries.push({
      room, guest: String(row[1] || '').replace(/\s*\*+$/, '').trim(),
      guestCount: parseInt(row[3]) || 0,
      checkin: checkin.date, checkout: checkout.date,
      checkoutHour: checkout.hour, checkoutMinute: checkout.minute,
      notes, status: String(row[9] || ''),
    });
  }
  return entries;
}

function parseOccupied(filePath) {
  const { data, roomCol } = readSheet(filePath);
  const entries = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[roomCol]) continue;
    const room = parseInt(row[roomCol]);
    if (isNaN(room)) continue;

    const checkin = parseTLDate(row[3]);
    const checkout = parseTLDate(row[4]);
    const notes = [String(row[12] || ''), String(row[13] || '')].join(' ').trim();

    entries.push({
      room, guest: String(row[1] || '').replace(/\s*\*+$/, '').trim(),
      guestCount: parseInt(row[2]) || 0,
      checkin: checkin.date, checkout: checkout.date,
      checkoutHour: checkout.hour, checkoutMinute: checkout.minute,
      notes,
    });
  }
  return entries;
}

// ==============================
// СБОРКА СЕКЦИЙ
// ==============================

const arrivals = parseArrivals(ARRIVALS_FILE);
const departures = parseDepartures(DEPARTURES_FILE);
const staying = parseOccupied(OCCUPIED_FILE);

const sections = { arrivals, departures, staying };

const arrivalRooms = [...new Set(arrivals.map(r => r.room))];
const depRooms = [...new Set(departures.map(r => r.room))];
const stayRooms = [...new Set(staying.map(r => r.room))];
console.log(`Заезды: ${arrivalRooms.length} номеров, Выезды: ${depRooms.length} номеров, Проживания: ${stayRooms.length} номеров`);

// ==============================
// ОПРЕДЕЛЕНИЕ ТИПОВ УБОРОК
// ==============================

const ALL_ROOMS = [
  101,102,103,104,105,106,107,108,109,110,112,113,114,115,116,
  201,202,203,204,205,206,207,208,209,210,211,212,213,214,215,216,217,218,219,220,221,222,223,224,225
];

function getArea(room) {
  if (room >= 112 && room <= 116) return 1.5;
  return Math.floor(room / 100);
}

function getCleaning(room) {
  const isArrival = sections.arrivals.some(r => r.room === room);
  const isDeparture = sections.departures.some(r => r.room === room);
  if (isArrival && isDeparture) return { type: '40 выезд/заезд', minutes: 40 };
  if (isDeparture) return { type: '40 выезд', minutes: 40 };
  if (isArrival) return { type: '10', minutes: 10 };
  if (sections.staying.some(r => r.room === room)) return { type: '20', minutes: 20 };
  return { type: null, minutes: 0 };
}

function getLateCheckoutStr(entry) {
  if (!entry || entry.checkoutHour == null) return '';
  if (entry.checkoutHour > STANDARD_CHECKOUT_HOUR) {
    return `поздний выезд до ${String(entry.checkoutHour).padStart(2,'0')}:${String(entry.checkoutMinute).padStart(2,'0')}`;
  }
  return '';
}

function getGuestCount(room) {
  for (const entry of sections.arrivals)
    if (entry.room === room && entry.guestCount > 0) return String(entry.guestCount);
  for (const entry of sections.staying)
    if (entry.room === room && entry.guestCount > 0) return String(entry.guestCount);
  return '';
}

function makeComment(room, cleaningType) {
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
    if (notes.includes('шампанск')) parts.push('шампанское (др)');

    let comment = parts.join(' + ');

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
    const nightsStayed = Math.round((REPORT_DATE - entry.checkin) / (1000*60*60*24));
    let totalNights = 0;
    if (entry.checkout && entry.checkin)
      totalNights = Math.round((entry.checkout - entry.checkin) / (1000*60*60*24));
    const remaining = totalNights - nightsStayed;
    if (nightsStayed >= 2 && nightsStayed % 2 === 0 && remaining >= 2) return 'смена белья';
    return '';
  }
  return '';
}

// ==============================
// ЗАПРОС ДАННЫХ У ПОЛЬЗОВАТЕЛЯ
// ==============================

const readline = require('readline');
const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

async function ask(question, def) {
  return new Promise(resolve => {
    rl.question(question, answer => {
      resolve(answer.trim() || def);
    });
  });
}

// ==============================
// ГЛАВНАЯ
// ==============================

async function main() {
  const tasks = ALL_ROOMS.map(room => {
    const cleaning = getCleaning(room);
    return {
      room,
      _cleaning: cleaning,
      _comment: makeComment(room, cleaning),
      _guestCount: (cleaning.type === '10' || cleaning.type === '40 выезд/заезд') ? getGuestCount(room) : '',
    };
  });

  console.log('\n=== ГРАФИК ===');
  for (const t of tasks) {
    const type = t._cleaning.type || '-';
    const min = t._cleaning.minutes || '';
    console.log(
      String(t.room).padEnd(5),
      `| ${type}`.padEnd(18),
      min ? `${min}мин`.padEnd(6) : '',
      t._comment ? `| ${t._comment}` : '',
      t._guestCount ? `| кол-во: ${t._guestCount}` : ''
    );
  }

  const numHK = parseInt(await ask('Сколько горничных? (по умолчанию 3): ')) || 3;
  const hkNames = [];
  for (let i = 0; i < numHK; i++) {
    hkNames.push(await ask(`Имя горничной ${i + 1}: `) || `Горничная ${i + 1}`);
  }
  rl.close();

  // === РАСПРЕДЕЛЕНИЕ ===
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

  // Зона 1.5
  const zone15Tasks = tasks.filter(t => getArea(t.room) === 1.5);
  let mainTasks = tasks.filter(t => getArea(t.room) !== 1.5).sort((a, b) => a.room - b.room);
  if (zone15Tasks.length > 0 && numHK > 1) {
    for (const t of zone15Tasks.sort((a, b) => a.room - b.room)) assignToHK(0, t);
    console.log(`Зона 1.5 (корпус 112-116): горничная 1 (${hkMins[0]} мин)`);
  }

  // Сороковки
  const allForties = tasks.filter(t => t._cleaning.type === '40 выезд' || t._cleaning.type === '40 выезд/заезд')
    .sort((a, b) => a.room - b.room);
  const totalForties = allForties.length;
  const base40 = Math.floor(totalForties / numHK);
  const rem40 = totalForties % numHK;
  const fortyNeeded = [];
  for (let i = 0; i < numHK; i++) {
    const existing = hkTasks[i].filter(t => t._cleaning.type === '40 выезд' || t._cleaning.type === '40 выезд/заезд').length;
    fortyNeeded.push(Math.max(0, base40 + (i < rem40 ? 1 : 0) - existing));
  }
  const extraForties = allForties.filter(t => !zone15Tasks.includes(t));
  for (let i = 0; i < numHK; i++) {
    for (let j = 0; j < fortyNeeded[i]; j++) {
      if (extraForties.length === 0) break;
      assignToHK(i, extraForties.shift());
    }
  }
  while (extraForties.length > 0) {
    assignToHK(hkMins.indexOf(Math.min(...hkMins)), extraForties.shift());
  }

  const fortyDebt = [];
  for (let i = 0; i < numHK; i++) {
    const has40 = hkTasks[i].filter(t => t._cleaning.type === '40 выезд' || t._cleaning.type === '40 выезд/заезд').length;
    fortyDebt.push(Math.max(0, base40 - has40));
  }

  console.log(`Сороковок: ${totalForties} (${base40}×${numHK}${rem40 > 0 ? ' + ' + rem40 : ''})`);

  // Двадцатки и десятки
  let twentyTenTasks = tasks.filter(t => t._cleaning.type === '20' || t._cleaning.type === '10').sort((a, b) => a.room - b.room);
  const assignedRooms = new Set();
  for (let i = 0; i < numHK; i++)
    for (const t of hkTasks[i]) assignedRooms.add(t.room);
  twentyTenTasks = twentyTenTasks.filter(t => !assignedRooms.has(t.room));

  // Компенсация
  for (let i = 0; i < numHK; i++) {
    for (let j = 0; j < fortyDebt[i] * 2; j++) {
      let bestIdx = 0, bestDist = Infinity;
      if (hkTasks[i].length > 0) {
        const myRooms = hkTasks[i].map(t => t.room);
        for (let k = 0; k < twentyTenTasks.length; k++) {
          if (twentyTenTasks[k]._cleaning.type !== '20') continue;
          const dist = Math.min(...myRooms.map(r => Math.abs(r - twentyTenTasks[k].room)));
          if (dist < bestDist) { bestDist = dist; bestIdx = k; }
        }
      }
      if (twentyTenTasks.length > 0 && twentyTenTasks[bestIdx]._cleaning.type === '20')
        assignToHK(i, twentyTenTasks.splice(bestIdx, 1)[0]);
    }
  }

  // Остаток блоками
  if (twentyTenTasks.length > 0) {
    const targets = [];
    for (let i = 0; i < numHK; i++) targets.push(Math.max(0, targetPerHK - hkMins[i]));
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

  // Микро-корректировка (5 итераций)
  for (let iter = 0; iter < 5; iter++) {
    let maxIdx = 0, minIdx = 0;
    for (let i = 0; i < numHK; i++) {
      if (hkMins[i] > hkMins[maxIdx]) maxIdx = i;
      if (hkMins[i] < hkMins[minIdx]) minIdx = i;
    }
    const diff = hkMins[maxIdx] - hkMins[minIdx];
    if (diff <= 10) break;

    let bestSwap = null, bestRoomDist = Infinity;
    for (let ri = 0; ri < hkTasks[maxIdx].length; ri++) {
      const rTask = hkTasks[maxIdx][ri], rMin = rTask._cleaning.minutes || 0;
      for (let pi = 0; pi < hkTasks[minIdx].length; pi++) {
        const pTask = hkTasks[minIdx][pi], pMin = pTask._cleaning.minutes || 0;
        const newDiff = Math.abs((hkMins[maxIdx] - rMin + pMin) - (hkMins[minIdx] + rMin - pMin));
        if (newDiff < diff) {
          const rIs40 = (rTask._cleaning.type === '40 выезд' || rTask._cleaning.type === '40 выезд/заезд');
          const pIs40 = (pTask._cleaning.type === '40 выезд' || pTask._cleaning.type === '40 выезд/заезд');
          if (rIs40 !== pIs40) continue;
          const roomDist = Math.abs(rTask.room - pTask.room);
          if (roomDist < bestRoomDist) { bestRoomDist = roomDist; bestSwap = { ri, pi, maxIdx, minIdx }; }
        }
      }
    }

    if (bestSwap && bestRoomDist < 50) {
      const { ri, pi } = bestSwap;
      const richTask = hkTasks[maxIdx][ri], poorTask = hkTasks[minIdx][pi];
      hkTasks[maxIdx][ri] = poorTask; hkTasks[minIdx][pi] = richTask;
      hkMins[maxIdx] = hkMins[maxIdx] - (richTask._cleaning.minutes||0) + (poorTask._cleaning.minutes||0);
      hkMins[minIdx] = hkMins[minIdx] - (poorTask._cleaning.minutes||0) + (richTask._cleaning.minutes||0);
    } else {
      let bestMoveIdx = -1, bestDist = Infinity;
      for (let ri = 0; ri < hkTasks[maxIdx].length; ri++) {
        const t = hkTasks[maxIdx][ri], tMin = t._cleaning.minutes || 0;
        if (tMin > diff) continue;
        if (t._cleaning.type === '40 выезд' || t._cleaning.type === '40 выезд/заезд') continue;
        const dist = Math.min(...hkTasks[minIdx].map(pt => Math.abs(t.room - pt.room)));
        if (dist < bestDist) { bestDist = dist; bestMoveIdx = ri; }
      }
      if (bestMoveIdx !== -1 && bestDist < 50) {
        const task = hkTasks[maxIdx].splice(bestMoveIdx, 1)[0];
        hkTasks[minIdx].push(task);
        hkMins[maxIdx] -= task._cleaning.minutes||0;
        hkMins[minIdx] += task._cleaning.minutes||0;
      } else break;
    }
  }

  // Лимит 2 выезд/заезд на горничную
  for (let i = 0; i < numHK; i++) {
    let attempts = 0;
    while (hkVyzdyZaezd[i] > 2 && attempts < 10) {
      attempts++;
      const swapIdx = hkTasks[i].findIndex(t => t._cleaning.type === '40 выезд/заезд');
      if (swapIdx === -1) break;
      let best = null;
      for (let j = 0; j < numHK; j++) {
        if (j === i || hkVyzdyZaezd[j] >= 2) continue;
        const their40 = hkTasks[j].findIndex(t => t._cleaning.type === '40 выезд');
        if (their40 !== -1) { best = { j, idx: their40 }; break; }
      }
      if (best) {
        const tmp = hkTasks[i][swapIdx]; hkTasks[i][swapIdx] = hkTasks[best.j][best.idx]; hkTasks[best.j][best.idx] = tmp;
        hkVyzdyZaezd[i]--; hkVyzdyZaezd[best.j]++;
      } else break;
    }
  }

  // Консоль
  for (let i = 0; i < numHK; i++) {
    const sorted = [...hkTasks[i]].sort((a, b) => a.room - b.room);
    const zones = [...new Set(sorted.map(t => getArea(t.room)))].sort().join(', ');
    console.log(`\n${hkNames[i]}: ${hkMins[i]} мин (зоны: ${zones})`);
    for (const t of sorted) {
      console.log(`  ${t.room} | ${t._cleaning.type} | ${t._cleaning.minutes}мин${t._comment ? ' | ' + t._comment : ''}${t._guestCount ? ' | кол-во: ' + t._guestCount : ''}`);
    }
  }

  // ==============================
  // EXCEL
  // ==============================
  const wb = new ExcelJS.Workbook();

  // Цвета
  const HK_COLORS = ['FFE74C3C','FF3498DB','FF2ECC71','FF9B59B6','FFF39C12','FF1ABC9C'];
  const typeToCol = { '10': 2, '40 выезд/заезд': 3, '40 выезд': 4, '20': 5 };

  const roomInfo = {};
  for (let i = 0; i < numHK; i++)
    for (const t of hkTasks[i])
      roomInfo[t.room] = { hkIdx: i, type: t._cleaning.type };

  // --- Лист 1: Уборки на сегодня ---
  const ws1 = wb.addWorksheet('Уборки на сегодня', { views: [{ state: 'frozen', ySplit: 3 }] });
  ws1.columns = [{ width: 8 },{ width: 10 },{ width: 14 },{ width: 10 },{ width: 10 },{ width: 40 },{ width: 12 }];

  ws1.addRow(['']); ws1.addRow(['']);
  ws1.mergeCells(1, 1, 2, 7);
  const titleCell = ws1.getCell(1, 1);
  titleCell.value = `${DD}.${MM}.${YY}`;
  titleCell.font = { size: 25, bold: true };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

  const headerKeys = ['Номер','10 минут','Выезд/Заезд','Выезд','20 минут','Комментарий','Кол-во чел'];
  const hRow1 = ws1.addRow(headerKeys);
  hRow1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF34495E' } };
  hRow1.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  hRow1.alignment = { horizontal: 'center', vertical: 'middle' };

  tasks.sort((a, b) => getArea(a.room) !== getArea(b.room) ? getArea(a.room) - getArea(b.room) : a.room - b.room);
  for (const t of tasks) {
    const type = t._cleaning.type || '';
    const row = ws1.addRow([
      t.room,
      type === '10' ? t._cleaning.minutes : '',
      type === '40 выезд/заезд' ? t._cleaning.minutes : '',
      type === '40 выезд' ? t._cleaning.minutes : '',
      type === '20' ? t._cleaning.minutes : '',
      t._comment || '', t._guestCount || '',
    ]);
    row.alignment = { horizontal: 'center', vertical: 'middle' };
  }

  // Цвета в таблице
  for (let r = 4; r <= ws1.rowCount; r++) {
    const room = ws1.getCell(r, 1).value;
    if (room && roomInfo[room]) {
      const { hkIdx, type } = roomInfo[room];
      const col = typeToCol[type];
      if (col) {
        ws1.getCell(r, col).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HK_COLORS[hkIdx % HK_COLORS.length] } };
        ws1.getCell(r, col).font = { bold: true, color: { argb: 'FF000000' } };
      }
    }
  }

  // Легенда
  ws1.addRow([]);
  ws1.addRow(['Распределение по горничным:']);
  for (let i = 0; i < numHK; i++) {
    const typeCounts = {};
    for (const t of hkTasks[i]) {
      const mins = t._cleaning.minutes || 0;
      if (mins > 0) typeCounts[mins] = (typeCounts[mins] || 0) + 1;
    }
    const breakdown = Object.entries(typeCounts).sort((a,b) => Number(b[0]) - Number(a[0])).map(([m,c]) => `${c}×${m}`).join(', ');
    const row = ws1.addRow([`${hkNames[i]} (${hkMins[i]} мин, ${breakdown})`]);
    row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HK_COLORS[i % HK_COLORS.length] } };
    row.getCell(1).font = { bold: true, color: { argb: 'FF000000' }, size: 10 };
  }

  // Расходники
  console.log('\n=== РАСХОДНИКИ ===');
  const hkConsumables = new Array(numHK).fill(0).map(() => ({ base: 0, extra: 0, total: 0 }));
  for (let i = 0; i < numHK; i++) {
    for (const t of hkTasks[i]) {
      const type = t._cleaning.type;
      if (type === '40 выезд' || type === '40 выезд/заезд') hkConsumables[i].base += 2;
      if (type === '10' || type === '40 выезд/заезд') {
        const gc = t._guestCount;
        if (gc) {
          const total = String(gc).split('+').reduce((s, p) => s + (parseInt(p) || 0), 0);
          if (total > 2) hkConsumables[i].extra += (total - 2);
        }
      }
    }
    hkConsumables[i].total = hkConsumables[i].base + hkConsumables[i].extra;
    console.log(`${hkNames[i]}: по ${hkConsumables[i].total} шт каждого вида (${hkConsumables[i].base} — выезды ×2, +${hkConsumables[i].extra} — доп. гости)`);
  }

  ws1.addRow([]);
  ws1.addRow(['Расходники (наборов на номер):']);
  for (let i = 0; i < numHK; i++) {
    const c = hkConsumables[i];
    const row = ws1.addRow([`${hkNames[i]}: по ${c.total} шт каждого вида (выезды ×2: ${c.base}, доп. гости: +${c.extra})`]);
    row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HK_COLORS[i % HK_COLORS.length] } };
    row.getCell(1).font = { bold: true, color: { argb: 'FF000000' }, size: 10 };
  }

  // --- Лист 2: Карточки заездов ---
  const wsK = wb.addWorksheet('Карточки заездов', { views: [{ state: 'frozen', ySplit: 1 }] });
  wsK.columns = [{ width: 6 },{ width: 30 },{ width: 8 },{ width: 8 },{ width: 14 },{ width: 14 }];
  const hK = wsK.addRow(['Комната','ФИО гостя','Кол-во','Ночей','Заезд','Выезд']);
  hK.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2980B9' } };
  hK.font = { bold: true, color: { argb: 'FFFFFFFF' } };

  const arrivalMap = {};
  for (const entry of sections.arrivals) {
    if (!arrivalMap[entry.room]) arrivalMap[entry.room] = { room: entry.room, guests: [], nights: 0 };
    if (entry.guest && !arrivalMap[entry.room].guests.includes(entry.guest)) arrivalMap[entry.room].guests.push(entry.guest);
    if (entry.checkin) arrivalMap[entry.room]._checkin = `${String(entry.checkin.getDate()).padStart(2,'0')}.${String(entry.checkin.getMonth()+1).padStart(2,'0')}`;
    if (entry.checkout) arrivalMap[entry.room]._checkout = `${String(entry.checkout.getDate()).padStart(2,'0')}.${String(entry.checkout.getMonth()+1).padStart(2,'0')}`;
    if (entry.checkin && entry.checkout) {
      const nights = Math.round((entry.checkout - entry.checkin) / (1000*60*60*24));
      if (nights > 0) arrivalMap[entry.room].nights = nights;
    }
  }

  let totalGuests = 0;
  for (const entry of Object.values(arrivalMap).sort((a,b) => a.room - b.room)) {
    const gc = getGuestCount(entry.room);
    const names = entry.guests.join(', ');
    wsK.addRow([entry.room, names, gc, entry.nights || '', entry._checkin || '', entry._checkout || '']);
    if (gc) totalGuests += String(gc).split('+').reduce((s, p) => s + (parseInt(p) || 0), 0);
  }
  wsK.addRow([]);
  wsK.addRow([`Итого гостей: ${totalGuests}`]);

  // --- Индивидуальные листы ---
  for (let hk = 0; hk < numHK; hk++) {
    const wsHk = wb.addWorksheet(hkNames[hk], { views: [{ state: 'frozen', ySplit: 3 }] });
    wsHk.columns = ws1.columns.map(c => ({ width: c.width }));
    wsHk.addRow(['']); wsHk.addRow(['']);
    wsHk.mergeCells(1, 1, 2, 7);
    const hkTitle = wsHk.getCell(1, 1);
    hkTitle.value = `${DD}.${MM}.${YY} — ${hkNames[hk]}`;
    hkTitle.font = { size: 25, bold: true };
    hkTitle.alignment = { horizontal: 'center', vertical: 'middle' };

    const hRow = wsHk.addRow(headerKeys);
    hRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF34495E' } };
    hRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    hRow.alignment = { horizontal: 'center', vertical: 'middle' };

    const hkRoomSet = new Set(hkTasks[hk].map(t => t.room));
    const color = HK_COLORS[hk % HK_COLORS.length];
    for (const t of tasks) {
      const type = t._cleaning.type || '';
      const row = wsHk.addRow([
        t.room,
        type === '10' ? t._cleaning.minutes : '',
        type === '40 выезд/заезд' ? t._cleaning.minutes : '',
        type === '40 выезд' ? t._cleaning.minutes : '',
        type === '20' ? t._cleaning.minutes : '',
        t._comment || '', t._guestCount || '',
      ]);
      row.alignment = { horizontal: 'center', vertical: 'middle' };
      if (hkRoomSet.has(t.room)) {
        const col = typeToCol[type];
        if (col) {
          wsHk.getCell(row.number, col).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: color } };
          wsHk.getCell(row.number, col).font = { bold: true, color: { argb: 'FF000000' } };
        }
      }
    }
  }

  await wb.xlsx.writeFile(OUTPUT_FILE);
  console.log(`\n✅ Готово: ${OUTPUT_FILE}`);

  // Архив
  const archiveDir = path.join(DATA_DIR, 'старые');
  if (!fs.existsSync(archiveDir)) fs.mkdirSync(archiveDir, { recursive: true });
  for (const f of [ARRIVALS_FILE, DEPARTURES_FILE, OCCUPIED_FILE])
    fs.renameSync(f, path.join(archiveDir, path.basename(f)));
  console.log(`📦 Исходные файлы перемещены в "старые"`);
}

main().catch(console.error);
