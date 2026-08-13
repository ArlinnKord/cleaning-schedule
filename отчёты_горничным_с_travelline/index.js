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

function findFile(prefix) {
  const files = fs.readdirSync(DATA_DIR)
    .filter(f => f.startsWith(prefix) && f.endsWith('.xlsx') && !f.includes('~$'))
    .sort()
    .reverse();
  if (files.length === 0) return null;
  return path.join(DATA_DIR, files[0]);
}

function findReport(prefixes) {
  for (const p of prefixes) {
    const f = findFile(p);
    if (f) return f;
  }
  return null;
}

// Основной источник: русскоязычные отчёты «Заезды», «Выезды», «Проживания»
// (fallback — англоязычные ArrivalsReport / DeparturesReport / OccupiedRoomTypes)
// Плюс отчёт Housekeeping — это «шахматка»: по нему определяются типы уборок.
const ARRIVALS_FILE = findReport(['Заезды', 'ArrivalsReport']);
const DEPARTURES_FILE = findReport(['Выезды', 'DeparturesReport']);
const STAYING_FILE = findReport(['Проживания', 'OccupiedRoomTypes']);
const HOUSEKEEPING_FILE = findReport(['Housekeeping']);

if (!ARRIVALS_FILE || !DEPARTURES_FILE || !STAYING_FILE || !HOUSEKEEPING_FILE) {
  console.error('❌ Не найдены отчёты Заезды/Выезды/Проживания/Housekeeping в папке отчёты_из_travelline');
  process.exit(1);
}

console.log(`📥 Заезды: ${path.basename(ARRIVALS_FILE)}`);
console.log(`📥 Выезды: ${path.basename(DEPARTURES_FILE)}`);
console.log(`📥 Проживания: ${path.basename(STAYING_FILE)}`);
console.log(`📥 Шахматка (Housekeeping): ${path.basename(HOUSEKEEPING_FILE)}`);

// Дата из имени файла заездов (Заезды_01.08.2026_01.08.2026.xlsx)
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
// ПАРСИНГ ОТЧЁТОВ TRAVELLINE (ЗАЕЗДЫ / ВЫЕЗДЫ / ПРОЖИВАНИЯ)
// ==============================

function fmtDateStr(d) {
  return `${String(d.getDate()).padStart(2,'0')}.${String(d.getMonth()+1).padStart(2,'0')}.${d.getFullYear()}`;
}

function colIndex(header, names) {
  for (const n of names) {
    const i = header.indexOf(n);
    if (i !== -1) return i;
  }
  return -1;
}

/**
 * Читает один отчёт Travelline (Заезды/Выезды/Проживания или англ. аналог)
 * и возвращает записи по комнатам.
 * section: 'arrivals' | 'departures' | 'staying'
 */
function parseTLReport(filePath, todayDate, section) {
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets['Report'] || wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 });
  if (data.length < 2) return [];

  const h = data[0];
  const roomCol = colIndex(h, ['Номер комнаты', '№ комнаты']);
  const guestCol = colIndex(h, ['ФИО гостей', 'ФИО кириллицей', 'Гость', 'ФИО']);
  const countCol = colIndex(h, ['Количество гостей']);
  const commentCol = colIndex(h, ['Комментарий гостя']);
  const notesCol = colIndex(h, ['Заметки']);
  const nightsCol = colIndex(h, ['Количество ночей']);
  const checkinCol = colIndex(h, ['Время заезда', 'Заезд']);
  const checkoutCol = colIndex(h, ['Время выезда', 'Выезд']);
  const periodCol = colIndex(h, ['Период проживания']);
  if (roomCol === -1) return [];

  const todayStr = fmtDateStr(todayDate);
  const rooms = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[roomCol]) continue;
    const room = parseInt(row[roomCol]);
    if (isNaN(room)) continue;

    let checkinStr = '', checkoutStr = '';
    if (periodCol !== -1) {
      // Проживания: "31.07.2026, 14:04 - 02.08.2026, 12:00"
      const m = String(row[periodCol] || '').match(/(\d{2}\.\d{2}\.\d{4},\s*\d{1,2}:\d{2})\s*-\s*(\d{2}\.\d{2}\.\d{4},\s*\d{1,2}:\d{2})/);
      if (m) { checkinStr = m[1]; checkoutStr = m[2]; }
    }
    if (section === 'arrivals' && !checkinStr && checkinCol !== -1) checkinStr = String(row[checkinCol] || '');
    if (section === 'departures' && !checkoutStr && checkoutCol !== -1) checkoutStr = String(row[checkoutCol] || '');

    // Фильтр: заезды — только сегодняшние.
    // Выезды: отчёт выгружен на сегодня, но при «Задержке» в колонке
    // «Время выезда» остаётся ПЛАНОВАЯ дата (вчерашняя). Поэтому отбрасываем
    // только выезды, запланированные строго ПОСЛЕ сегодняшнего дня.
    if (section === 'arrivals' && !checkinStr.startsWith(todayStr)) continue;
    if (section === 'departures') {
      const co = parseTLDate(checkoutStr);
      if (co.date && co.date > todayDate) continue;
    }

    if (!rooms[room]) {
      rooms[room] = {
        room,
        guests: [],
        guestCount: 0,
        comment: '', notes: '',
        nights: 0,
        checkin: null, checkout: null,
        checkoutHour: null, checkoutMinute: null,
      };
    }
    const rec = rooms[room];

    const guest = String(row[guestCol] || '').trim();
    if (guest && !rec.guests.includes(guest)) rec.guests.push(guest);

    const cnt = parseInt(row[countCol]);
    if (!isNaN(cnt)) rec.guestCount += cnt;

    const cm = String(row[commentCol] || '').trim();
    if (cm) rec.comment = rec.comment ? rec.comment + '; ' + cm : cm;

    const nt = String(row[notesCol] || '').trim();
    if (nt) rec.notes = rec.notes ? rec.notes + '; ' + nt : nt;

    const nn = parseInt(row[nightsCol]);
    if (!isNaN(nn)) rec.nights = Math.max(rec.nights, nn);

    const ci = parseTLDate(checkinStr);
    const co = parseTLDate(checkoutStr);
    if (ci.date) rec.checkin = ci.date;
    if (co.date) { rec.checkout = co.date; rec.checkoutHour = co.hour; rec.checkoutMinute = co.minute; }
  }

  return Object.values(rooms).map(r => {
    let nights = r.nights;
    if (!nights && r.checkin && r.checkout)
      nights = Math.round((r.checkout - r.checkin) / (1000 * 60 * 60 * 24));
    return {
      room: r.room,
      guest: r.guests.join(', '),
      guestCount: r.guestCount,
      comment: r.comment,
      notes: r.notes,
      nights,
      checkin: r.checkin,
      checkout: r.checkout,
      checkoutHour: r.checkoutHour,
      checkoutMinute: r.checkoutMinute,
    };
  });
}

/**
 * Отчёт Housekeeping — «шахматка». По каждому номеру: занятость и плановые
 * даты «Номер будет освобожден» (выезд) / «Номер будет заселен» (заезд).
 * Эти даты — плановые, поэтому шахматка НЕ устаревает так быстро, как отчёт
 * «Выезды» (который показывает только уже подтверждённые события).
 */
function parseHousekeeping(filePath) {
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets['Report'] || wb.Sheets[wb.SheetNames[0]];
  const data = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 });
  if (data.length < 2) return null;
  const h = data[0];
  const roomCol = h.indexOf('Номер');
  const occCol = h.indexOf('Занятость');
  const releaseCol = h.indexOf('Номер будет освобожден');
  const occupyCol = h.indexOf('Номер будет заселен');
  if (roomCol === -1 || occCol === -1 || releaseCol === -1 || occupyCol === -1) return null;

  const map = new Map();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const room = parseInt(row[roomCol]);
    if (isNaN(room)) continue;
    const release = parseTLDate(String(row[releaseCol] || ''));
    const occupy = parseTLDate(String(row[occupyCol] || ''));
    map.set(room, {
      room,
      occupied: String(row[occCol] || '').startsWith('Занят'),
      releaseDate: release.date, releaseHour: release.hour, releaseMinute: release.minute,
      occupyDate: occupy.date, occupyHour: occupy.hour, occupyMinute: occupy.minute,
    });
  }
  return map;
}

// ==============================
// СБОРКА
// ==============================

// Детали (комментарии, гости, кол-во, ночи) из отчётов Заезды/Выезды/Проживания
const arrDetail = parseTLReport(ARRIVALS_FILE, REPORT_DATE, 'arrivals');
const depDetail = parseTLReport(DEPARTURES_FILE, REPORT_DATE, 'departures');
const stayDetail = parseTLReport(STAYING_FILE, REPORT_DATE, 'staying');

// Главный источник типов уборок — шахматка (Housekeeping).
// Отчёт «Выезды» может не успеть обновиться к запуску (устаревший снимок),
// а плановые «освобожден/заселен» из шахматки видны заранее.
const hkMap = parseHousekeeping(HOUSEKEEPING_FILE);
if (!hkMap || hkMap.size === 0) {
  console.error('❌ Не удалось прочитать шахматку (Housekeeping). Проверьте файл.');
  process.exit(1);
}
const isSameDay = d => d && d.getTime() === REPORT_DATE.getTime();

const hkDepRooms = [];
const hkArrRooms = [];
const hkStayRooms = [];
for (const info of hkMap.values()) {
  if (isSameDay(info.releaseDate)) hkDepRooms.push(info.room);
  if (isSameDay(info.occupyDate)) hkArrRooms.push(info.room);
  // Проживание = занят сегодня и не выезжает сегодня (номер, освобождённый
  // раньше даты отчёта, на дату отчёта уже пустой — уборка не нужна)
  const releasedBefore = info.releaseDate && info.releaseDate < REPORT_DATE;
  if (info.occupied && !isSameDay(info.releaseDate) && !releasedBefore) hkStayRooms.push(info.room);
}

const depRoomSet = new Set(hkDepRooms);
const arrivals = hkArrRooms.map(room => arrDetail.find(x => x.room === room) || {
  room, guest: '', guestCount: 0, comment: '', notes: '', nights: 0,
  checkin: null, checkout: null, checkoutHour: null, checkoutMinute: null,
});
const departures = hkDepRooms.map(room => {
  const d = depDetail.find(x => x.room === room);
  if (d) return d;
  const info = hkMap.get(room);
  return {
    room, guest: '', guestCount: 0, comment: '', notes: '', nights: 0,
    checkin: null, checkout: info ? info.releaseDate : null,
    checkoutHour: info ? info.releaseHour : null,
    checkoutMinute: info ? info.releaseMinute : null,
  };
});
const staying = hkStayRooms.map(room => stayDetail.find(x => x.room === room) || {
  room, guest: '', guestCount: 0, comment: '', notes: '', nights: 0,
  checkin: null, checkout: null, checkoutHour: null, checkoutMinute: null,
});

const arrivalRooms = [...new Set(arrivals.map(r => r.room))];
const depRoomsList = [...new Set(departures.map(r => r.room))];
const stayRooms = [...new Set(staying.map(r => r.room))];
console.log(`Заезды: ${arrivalRooms.length} номеров, Выезды: ${depRoomsList.length} номеров, Проживания: ${stayRooms.length} номеров`);

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
  const isArrival = arrivals.some(r => r.room === room);
  const isDeparture = departures.some(r => r.room === room);
  if (isArrival && isDeparture) return { type: '40 выезд/заезд', minutes: 40 };
  if (isDeparture) return { type: '40 выезд', minutes: 40 };
  if (isArrival) return { type: '10', minutes: 10 };
  if (staying.some(r => r.room === room)) return { type: '20', minutes: 20 };
  return { type: null, minutes: 0 };
}

/** Извлекает количество гостей из текста комментария: "4 чел, 2 взр +2 реб", "2 чел" и т.п. */
function parseGuestCountFromText(text) {
  const t = String(text || '').toLowerCase();
  // "N взр + M реб" / "N взрослых + M детей"
  const withKids = t.match(/(\d+)\s*(?:взр|взросл\w*)\s*\S{0,4}\s*[+,]\s*(\d+)\s*(?:ребен\w*|реб|дет\w*)/);
  if (withKids) return parseInt(withKids[1]) + parseInt(withKids[2]);
  // "N чел" / "N человек" / "N человека"
  const chel = t.match(/(?:^|[^\d])(\d+)\s*(?:чел|человек\w*)/);
  if (chel) return parseInt(chel[1]);
  // "гостей: N"
  const guestColon = t.match(/гостей[:\s]+(\d+)/);
  if (guestColon) return parseInt(guestColon[1]);
  return 0;
}

function getGuestCount(room) {
  const entry = arrivals.find(r => r.room === room) || staying.find(r => r.room === room);
  if (entry) {
    // Сначала парсим из комментария (в отчёте колонка "Количество гостей" часто = 1)
    const fromText = parseGuestCountFromText(`${entry.comment} ${entry.notes}`);
    if (fromText > 0) return String(fromText);
    if (entry.guestCount > 0) return String(entry.guestCount);
  }
  return '';
}

/**
 * Извлекает из комментария гостя информацию о кроватях и особых пожеланиях.
 * Примеры: "Double bed", "2 Single beds", "1 кровать", "2 раздельные кровати",
 * "доп место", "люлька", "ШАМПАНСКОЕ+ ОТКРЫТКУ"
 */
function parseBedComment(text) {
  const t = String(text || '').toLowerCase();
  const parts = [];

  // Тип кровати
  if (/(2\s*single|two\s*single|2\s*раздельн|2\s*отдельн|две\s*односпальн|2\s*кроват|раздельн|отдельн)/.test(t)) {
    parts.push('2 односпальные');
  } else if (/(double\s*bed|двуспальн|большая\s*двуспальн)/.test(t)) {
    parts.push('1 двуспальная');
  } else if (/(1\s*кроват|одна\s*кроват|one\s*bed|1\s*bed)/.test(t)) {
    parts.push('1 кровать');
  } else {
    parts.push('1 кровать'); // значение по умолчанию
  }

  if (/люльк/.test(t)) parts.push('люлька');
  if (/детская\s*кроватк/.test(t)) parts.push('детская кроватка');
  if (/доп[\s.]*мест|доп\.?\s*место/.test(t)) parts.push('доп. место');
  if (/диван/.test(t)) parts.push('диван');

  // Особый гость: открытка/шампанское
  if (/шампанск|открытк|день\s*рожден|блогер|блоггер/i.test(t)) parts.push('открытка/шампанское');

  return parts;
}

function makeComment(room, cleaningType) {
  if (!cleaningType || !cleaningType.type) return '';

  // Комментарий гостя для этой комнаты (из заездов или проживаний)
  const entry = arrivals.find(r => r.room === room) || staying.find(r => r.room === room);
  const guestText = `${entry ? entry.comment : ''} ${entry ? entry.notes : ''}`.trim();

  if (cleaningType.type === '10' || cleaningType.type === '40 выезд/заезд') {
    const parts = parseBedComment(guestText);

    // Доп. места если гостей больше 2
    const gc = getGuestCount(room);
    if (gc) {
      const total = String(gc).split('+').reduce((sum, p) => sum + (parseInt(p) || 0), 0);
      if (total > 2 && !parts.some(p => p.includes('доп'))) parts.push('доп. места');
    }

    if (cleaningType.type === '40 выезд/заезд') {
      const dep = departures.find(r => r.room === room);
      if (dep && dep.checkoutHour != null && dep.checkoutHour > STANDARD_CHECKOUT_HOUR) {
        parts.push(`поздний выезд до ${String(dep.checkoutHour).padStart(2,'0')}:${String(dep.checkoutMinute).padStart(2,'0')}`);
      }
    }

    return parts.join('; ');
  }

  if (cleaningType.type === '40 выезд') {
    const dep = departures.find(r => r.room === room);
    if (!dep || dep.checkoutHour == null) return '';
    if (dep.checkoutHour > STANDARD_CHECKOUT_HOUR) {
      return `поздний выезд до ${String(dep.checkoutHour).padStart(2,'0')}:${String(dep.checkoutMinute).padStart(2,'0')}`;
    }
    return '';
  }

  if (cleaningType.type === '20') {
    const entry20 = staying.find(r => r.room === room);
    const parts = [];
    if (entry20 && entry20.checkin) {
      const nightsStayed = Math.round((REPORT_DATE - entry20.checkin) / (1000*60*60*24));
      const remaining = entry20.nights - nightsStayed;
      if (nightsStayed >= 2 && nightsStayed % 2 === 0 && remaining >= 2) parts.push('смена белья');
    }
    // Особый гость для проживающих
    const special = parseBedComment(guestText).filter(p => p === 'открытка/шампанское');
    for (const sp of special) if (!parts.includes(sp)) parts.push(sp);
    return parts.join('; ');
  }
  return '';
}

// ==============================
// ЗАПРОС ДАННЫХ У ПОЛЬЗОВАТЕЛЯ
// ==============================

const readline = require('readline');
const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

// Queue-based ask that works with piped input (readline.question() breaks on pipes)
const _answerQueue = [];
const _waiters = [];
rl.on('line', line => {
  if (_waiters.length) {
    _waiters.shift()(line);
  } else {
    _answerQueue.push(line);
  }
});

async function ask(question, def) {
  process.stdout.write(question);
  return new Promise(resolve => {
    if (_answerQueue.length) {
      resolve(_answerQueue.shift().trim() || def);
    } else {
      _waiters.push(line => resolve(line.trim() || def));
    }
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

  // Сверка с шахматкой: отчёт «Выезды» — это снимок на момент скачивания,
  // он может устареть (гости выехали раньше/позже). Поэтому спрашиваем
  // про номера, где по шахматке выезд, а в отчёте их нет.
  const extraDepRaw = await ask('\nНомера на ВЫЕЗД по шахматке, которых нет в отчёте (через запятую, пусто — пропустить): ', '');
  const extraDepRooms = extraDepRaw.split(',').map(s => parseInt(s.trim())).filter(n => !isNaN(n));
  let extraDepCount = 0;
  for (const room of extraDepRooms) {
    const t = tasks.find(x => x.room === room);
    if (!t) continue;
    if (departures.some(r => r.room === room)) continue; // уже учтён как выезд
    const isArr = arrivals.some(r => r.room === room);
    t._cleaning = { type: isArr ? '40 выезд/заезд' : '40 выезд', minutes: 40 };
    t._comment = makeComment(room, t._cleaning);
    t._guestCount = (t._cleaning.type === '40 выезд/заезд') ? getGuestCount(room) : '';
    console.log(`  ➕ ${room} → ${t._cleaning.type}`);
    extraDepCount++;
  }
  if (extraDepCount === 0) console.log('  (доп. выездов не введено)');

  const numHK = parseInt(await ask('Сколько горничных? (по умолчанию 3): ')) || 3;
  const hkNames = [];
  const hkHours = [];
  for (let i = 0; i < numHK; i++) {
    const name = await ask(`Имя горничной ${i + 1}: `) || `Горничная ${i + 1}`;
    hkNames.push(name);
    const hours = parseFloat(await ask(`  ${name}: сколько часов работает? (по умолчанию 8): `)) || 8;
    hkHours.push(hours);
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

  const LARGE_ROOMS = new Set([107, 201, 208, 216, 217, 220, 221, 222, 224, 225]);
  // Светло-бежевая подсветка номера в столбце «Номер» — визуальный индикатор больших номеров, не влияет на распределение
  const LARGE_ROOM_COLOR = 'FFEFD9B3';

  // Все задачи по типам
  const allForties = tasks.filter(t => t._cleaning.type === '40 выезд' || t._cleaning.type === '40 выезд/заезд')
    .sort((a, b) => a.room - b.room);
  const allTwenties = tasks.filter(t => t._cleaning.type === '20').sort((a, b) => a.room - b.room);
  const allTens = tasks.filter(t => t._cleaning.type === '10').sort((a, b) => a.room - b.room);
  const totalForties = allForties.length;
  const totalTwenties = allTwenties.length;
  const totalTens = allTens.length;
  console.log(`40-минутных: ${totalForties}, 20-минутных: ${totalTwenties}, 10-минутных: ${totalTens}`);

  const allMinutes = tasks.reduce((s, t) => s + (t._cleaning.minutes || 0), 0);
  const totalHours = hkHours.reduce((s, h) => s + h, 0);
  const targetMins = hkHours.map(h => allMinutes * h / totalHours);
  console.log(`\n=== РАСПРЕДЕЛЕНИЕ (${numHK} горничных, всего ${allMinutes} мин) ===`);
  for (let i = 0; i < numHK; i++) {
    console.log(`  ${hkNames[i]}: ${hkHours[i]} ч → цель ~${Math.round(targetMins[i])} мин`);
  }

  // Сортируем горничных: от меньших часов к большим
  const hkByHours = hkHours.map((h, i) => i).sort((a, b) => hkHours[a] - hkHours[b]);

  // Выбор ближайшей задачи: приоритет — тот же этаж, потом расстояние
  function pickClosest(available, assignedRooms) {
    let best = 0, bestScore = Infinity;
    for (let k = 0; k < available.length; k++) {
      const aFloor = Math.floor(available[k].room / 100);
      const sameFloor = assignedRooms.length
        ? assignedRooms.some(r => Math.floor(r / 100) === aFloor) ? 0 : 1
        : 0;
      const dist = assignedRooms.length
        ? Math.min(...assignedRooms.map(r => Math.abs(r - available[k].room)))
        : 0;
      const score = sameFloor * 10000 + dist;
      if (score < bestScore) { bestScore = score; best = k; }
    }
    return best;
  }

  // Подсчёт больших номеров у горничной (по типу уборки)
  function countLarge(hkIdx, typeFilter) {
    return hkTasks[hkIdx].filter(t =>
      LARGE_ROOMS.has(t.room) && (!typeFilter || typeFilter === t._cleaning.type)
    ).length;
  }

  // Попытка обмена больших номеров между горничными
  function balanceLargeRooms(typeFilter, maxIter) {
    for (let iter = 0; iter < (maxIter || 20); iter++) {
      let maxIdx = -1, minIdx = -1;
      let maxCnt = -1, minCnt = Infinity;
      for (const i of hkByHours) {
        const c = countLarge(i, typeFilter);
        if (c > maxCnt) { maxCnt = c; maxIdx = i; }
        if (c < minCnt) { minCnt = c; minIdx = i; }
      }
      if (maxIdx === -1 || minIdx === -1 || maxCnt - minCnt <= 1) break;

      let swapped = false;
      for (let ri = 0; ri < hkTasks[maxIdx].length && !swapped; ri++) {
        const rTask = hkTasks[maxIdx][ri];
        if (!LARGE_ROOMS.has(rTask.room)) continue;
        if (typeFilter && rTask._cleaning.type !== typeFilter) continue;

        for (let pi = 0; pi < hkTasks[minIdx].length && !swapped; pi++) {
          const pTask = hkTasks[minIdx][pi];
          if (LARGE_ROOMS.has(pTask.room)) continue;
          if (typeFilter && pTask._cleaning.type !== typeFilter) continue;
          if (rTask._cleaning.minutes !== pTask._cleaning.minutes) continue;

          // Меняем
          hkTasks[maxIdx][ri] = pTask;
          hkTasks[minIdx][pi] = rTask;
          swapped = true;
        }
      }
      if (!swapped) break;
    }
  }

  // ============================================================
  // ФАЗА 1: 40-минутные уборки
  // Делим поровну. Остаток — тем, кто работает БОЛЬШЕ.
  // Короткий день получает base (без остатка).
  // ============================================================
  console.log('\n--- Фаза 1: 40-минутные уборки ---');

  // Отделяем выезд/заезд от простых выездов
  const vyezdZaezdTasks = allForties.filter(t => t._cleaning.type === '40 выезд/заезд')
    .sort((a, b) => a.room - b.room);
  const plainForties = allForties.filter(t => t._cleaning.type === '40 выезд')
    .sort((a, b) => a.room - b.room);

  // Сначала распределяем выезд/заезд: макс 2 на горничную
  const vzPool = [...vyezdZaezdTasks];
  for (let round = 0; round < 2 && vzPool.length > 0; round++) {
    for (const hkIdx of hkByHours.slice().reverse()) {
      // Обратный порядок: более загруженные первыми получают выезд/заезд
      if (vzPool.length === 0) break;
      if (hkVyzdyZaezd[hkIdx] >= 2) continue;
      const best = pickClosest(vzPool, hkTasks[hkIdx].map(t => t.room));
      assignToHK(hkIdx, vzPool.splice(best, 1)[0]);
    }
  }
  // Оставшиеся выезд/заезд — тем, у кого меньше 2
  while (vzPool.length > 0) {
    const hkIdx = hkByHours.reduce((best, i) =>
      hkVyzdyZaezd[i] < hkVyzdyZaezd[best] ? i : best, hkByHours[0]);
    const best = pickClosest(vzPool, hkTasks[hkIdx].map(t => t.room));
    assignToHK(hkIdx, vzPool.splice(best, 1)[0]);
  }

  // Считаем, сколько ещё 40-минутных нужно каждой горничной
  const totalFortiesAssigned = allForties.length;
  const base40 = Math.floor(totalFortiesAssigned / numHK);
  const rem40 = totalFortiesAssigned % numHK;
  const fortyTargets = new Array(numHK).fill(base40);
  for (let i = 0; i < rem40; i++) {
    // Остаток — тем, кто работает дольше
    fortyTargets[hkByHours[numHK - 1 - i]]++;
  }

  // Распределяем остальные 40-минутные (простые выезды)
  const plainPool = [...plainForties];
  for (const hkIdx of hkByHours) {
    const current40s = hkTasks[hkIdx].filter(t =>
      t._cleaning.type === '40 выезд' || t._cleaning.type === '40 выезд/заезд'
    ).length;
    const needed = fortyTargets[hkIdx] - current40s;
    for (let j = 0; j < needed && plainPool.length > 0; j++) {
      const best = pickClosest(plainPool, hkTasks[hkIdx].map(t => t.room));
      assignToHK(hkIdx, plainPool.splice(best, 1)[0]);
    }
  }
  while (plainPool.length > 0) {
    const hkIdx = hkByHours.reduce((best, i) => hkMins[i] < hkMins[best] ? i : best, hkByHours[0]);
    const best = pickClosest(plainPool, hkTasks[hkIdx].map(t => t.room));
    assignToHK(hkIdx, plainPool.splice(best, 1)[0]);
  }

  // Балансировка больших номеров среди 40-минутных уборок (только 40-минутные)
  balanceLargeRooms(null, 20);

  // Лимит 2 выезд/заезд на горничную (финальная проверка)
  for (let i = 0; i < numHK; i++) {
    let attempts = 0;
    while (hkVyzdyZaezd[i] > 2 && attempts < 10) {
      attempts++;
      const swapIdx = hkTasks[i].findIndex(t => t._cleaning.type === '40 выезд/заезд');
      if (swapIdx === -1) break;
      // Ищем, кому отдать: у кого < 2 выезд/заезд
      let bestJ = -1;
      for (let j = 0; j < numHK; j++) {
        if (j === i || hkVyzdyZaezd[j] >= 2) continue;
        const their40 = hkTasks[j].findIndex(t => t._cleaning.type === '40 выезд');
        if (their40 !== -1) { bestJ = j; break; }
      }
      if (bestJ !== -1) {
        const vzIdx = swapIdx;
        const p40Idx = hkTasks[bestJ].findIndex(t => t._cleaning.type === '40 выезд');
        const tmp = hkTasks[i][vzIdx];
        hkTasks[i][vzIdx] = hkTasks[bestJ][p40Idx];
        hkTasks[bestJ][p40Idx] = tmp;
        hkVyzdyZaezd[i]--;
        hkVyzdyZaezd[bestJ]++;
      } else break;
    }
  }

  console.log(`Сороковки: ${fortyTargets.map((c, i) => `${hkNames[i]}: ${c}`).join(', ')}`);

  // ============================================================
  // ФАЗА 1.5: Балансировка 20+20=40 (только если не превышает часы)
  // Если после 40-минутных у горничной дефицит ≥40 мин относительно
  // её пропорционального таргета — добавляем 2×20 из пула.
  // Короткий день (4ч) не получит лишнего сверх пропорции.
  // ============================================================
  console.log('\n--- Фаза 1.5: 20+20 балансировка ---');
  const used20as40 = new Set();

  for (const hkIdx of hkByHours) {
    const current40Min = hkTasks[hkIdx]
      .filter(t => t._cleaning.type === '40 выезд' || t._cleaning.type === '40 выезд/заезд')
      .reduce((s, t) => s + (t._cleaning.minutes || 0), 0);
    const deficit = targetMins[hkIdx] - current40Min;

    if (deficit >= 40) {
      const available20s = allTwenties.filter(t => !used20as40.has(t.room));
      if (available20s.length >= 2) {
        const pool = [...available20s];
        const firstIdx = pickClosest(pool, hkTasks[hkIdx].map(t => t.room));
        used20as40.add(pool[firstIdx].room);
        assignToHK(hkIdx, pool[firstIdx]);
        pool.splice(firstIdx, 1);
        if (pool.length > 0) {
          const secondIdx = pickClosest(pool, hkTasks[hkIdx].map(t => t.room));
          used20as40.add(pool[secondIdx].room);
          assignToHK(hkIdx, pool[secondIdx]);
        }
      }
    }
  }

  console.log(`Использовано 20-минутных для баланса: ${used20as40.size}`);

  // ============================================================
  // ФАЗА 2: Оставшиеся 20-минутные уборки
  // Делим поровну. Остаток — самой короткой.
  // Балансируем большие номера.
  // ============================================================
  console.log('\n--- Фаза 2: 20-минутные уборки ---');
  const remaining20s = allTwenties.filter(t => !used20as40.has(t.room));
  const totalRemaining20s = remaining20s.length;

  const base20 = Math.floor(totalRemaining20s / numHK);
  const rem20 = totalRemaining20s % numHK;
  const twentyTargets = new Array(numHK).fill(base20);
  for (let i = 0; i < rem20; i++) {
    // Остаток — тем, кто работает дольше (короткий день получает base)
    twentyTargets[hkByHours[numHK - 1 - i]]++;
  }

  const pool20 = [...remaining20s];
  for (const hkIdx of hkByHours) {
    const needed = twentyTargets[hkIdx];
    for (let j = 0; j < needed && pool20.length > 0; j++) {
      const best = pickClosest(pool20, hkTasks[hkIdx].map(t => t.room));
      assignToHK(hkIdx, pool20.splice(best, 1)[0]);
    }
  }
  while (pool20.length > 0) {
    const hkIdx = hkByHours.reduce((best, i) => hkMins[i] < hkMins[best] ? i : best, hkByHours[0]);
    const best = pickClosest(pool20, hkTasks[hkIdx].map(t => t.room));
    assignToHK(hkIdx, pool20.splice(best, 1)[0]);
  }

  // Балансировка больших номеров среди 20-минутных уборок
  balanceLargeRooms('20', 20);

  console.log(`Двадцатки: ${twentyTargets.map((c, i) => `${hkNames[i]}: ${c}`).join(', ')}`);

  // ============================================================
  // ФАЗА 3: 10-минутные уборки
  // Сначала добиваем хвостики 10+10=20 (кто отстаёт от таргета),
  // потом делим поровну.
  // ============================================================
  console.log('\n--- Фаза 3: 10-минутные уборки ---');

  const pool10 = [...allTens];

  // Для тех, кто сильно отстаёт от пропорционального таргета — 10+10=20
  for (const hkIdx of hkByHours) {
    if (pool10.length < 2) break;
    const diff = targetMins[hkIdx] - hkMins[hkIdx];
    if (diff >= 20) {
      const first = pickClosest(pool10, hkTasks[hkIdx].map(t => t.room));
      assignToHK(hkIdx, pool10.splice(first, 1)[0]);
      if (pool10.length > 0) {
        const second = pickClosest(pool10, hkTasks[hkIdx].map(t => t.room));
        assignToHK(hkIdx, pool10.splice(second, 1)[0]);
      }
    }
  }

  // Оставшиеся десятки — делим поровну
  const totalRemaining10s = pool10.length;
  const base10 = Math.floor(totalRemaining10s / numHK);
  const rem10 = totalRemaining10s % numHK;
  const tenTargets = new Array(numHK).fill(base10);
  for (let i = 0; i < rem10; i++) {
    tenTargets[hkByHours[numHK - 1 - i]]++;
  }

  for (const hkIdx of hkByHours) {
    const needed = tenTargets[hkIdx];
    for (let j = 0; j < needed && pool10.length > 0; j++) {
      const best = pickClosest(pool10, hkTasks[hkIdx].map(t => t.room));
      assignToHK(hkIdx, pool10.splice(best, 1)[0]);
    }
  }
  while (pool10.length > 0) {
    const hkIdx = hkByHours.reduce((best, i) => hkMins[i] < hkMins[best] ? i : best, hkByHours[0]);
    const best = pickClosest(pool10, hkTasks[hkIdx].map(t => t.room));
    assignToHK(hkIdx, pool10.splice(best, 1)[0]);
  }

  console.log(`Десятки: ${tenTargets.map((c, i) => `${hkNames[i]}: ${c}`).join(', ')}`);

  // Итоговый баланс по минутам
  console.log('\n=== ИТОГ ПО МИНУТАМ ===');
  for (let i = 0; i < numHK; i++) {
    const fortyCount = hkTasks[i].filter(t => t._cleaning.type === '40 выезд' || t._cleaning.type === '40 выезд/заезд').length;
    const twentyCount = hkTasks[i].filter(t => t._cleaning.type === '20').length;
    const tenCount = hkTasks[i].filter(t => t._cleaning.type === '10').length;
    const largeCount = hkTasks[i].filter(t => LARGE_ROOMS.has(t.room)).length;
    const vzCount = hkVyzdyZaezd[i];
    console.log(`  ${hkNames[i]}: ${hkMins[i]} мин (40: ${fortyCount}, 20: ${twentyCount}, 10: ${tenCount}, больших: ${largeCount}, выезд/заезд: ${vzCount})`);
  }

  // Общая балансировка больших номеров (ещё раз по всем типам)
  balanceLargeRooms(null, 10);

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

  const roomAssign = {};
  for (let i = 0; i < numHK; i++)
    for (const t of hkTasks[i])
      roomAssign[t.room] = { hkIdx: i, type: t._cleaning.type };

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
    if (room && LARGE_ROOMS.has(Number(room))) {
      ws1.getCell(r, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LARGE_ROOM_COLOR } };
    }
    if (room && roomAssign[room]) {
      const { hkIdx, type } = roomAssign[room];
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

  let totalGuests = 0;
  for (const entry of arrivals.sort((a,b) => a.room - b.room)) {
    const gc = getGuestCount(entry.room);
    const checkinStr = entry.checkin ? `${String(entry.checkin.getDate()).padStart(2,'0')}.${String(entry.checkin.getMonth()+1).padStart(2,'0')}` : '';
    const checkoutStr = entry.checkout ? `${String(entry.checkout.getDate()).padStart(2,'0')}.${String(entry.checkout.getMonth()+1).padStart(2,'0')}` : '';
    wsK.addRow([entry.room, entry.guest, gc, entry.nights || '', checkinStr, checkoutStr]);
    if (gc) totalGuests += parseInt(String(gc).split('+')[0]) || 0;
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
      if (LARGE_ROOMS.has(t.room)) {
        wsHk.getCell(row.number, 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LARGE_ROOM_COLOR } };
      }
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
  for (const f of [ARRIVALS_FILE, DEPARTURES_FILE, STAYING_FILE, HOUSEKEEPING_FILE].filter(f => f && fs.existsSync(f)))
    fs.renameSync(f, path.join(archiveDir, path.basename(f)));
  console.log(`📦 Исходные файлы перемещены в "старые"`);
}

main().catch(console.error);
