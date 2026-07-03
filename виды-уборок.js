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
while (fs.existsSync(path.join(OUT_DIR, `виды-уборок-${DD}.${MM}.${YY}-v${ver}.xlsx`))) ver++;
const OUTPUT_FILE = path.join(OUT_DIR, `виды-уборок-${DD}.${MM}.${YY}-v${ver}.xlsx`);

console.log(`📅 ${DD}.${MM}.${TODAY.getFullYear()}`);
console.log(`📥 ${path.basename(INPUT_FILE)} → 📤 ${path.basename(OUTPUT_FILE)}`);

// ==============================
// ПАРСИНГ ОТЧЁТА
// ==============================

function parseDateRange(rangeStr) {
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
      notes,
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
// ОПРЕДЕЛЕНИЕ ТИПОВ УБОРОК
// ==============================

function getArea(room) {
  if (room >= 112 && room <= 116) return 3;
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

    let comment = parts.join(' + ');
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
    const nightsStayed = Math.round((REPORT_DATE - entry.checkin) / (1000 * 60 * 60 * 24));
    const remaining = entry.nights - nightsStayed;
    if (nightsStayed >= 2 && remaining >= 2 && nightsStayed < 7) return 'смена белья';
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
  const ws1 = wb.addWorksheet('Уборки на сегодня', { views: [{ state: 'frozen', ySplit: 1 }] });
  ws1.columns = [
    { width: 8 }, { width: 10 }, { width: 14 }, { width: 10 }, { width: 10 }, { width: 40 }, { width: 12 },
  ];

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
    ws1.addRow([
      t.room || '',
      type === '10' ? t._cleaning.minutes : '',
      type === '40 выезд/заезд' ? t._cleaning.minutes : '',
      type === '40 выезд' ? t._cleaning.minutes : '',
      type === '20' ? t._cleaning.minutes : '',
      t._comment || '',
      t._guestCount || '',
    ]);
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
  for (const entry of arrivalRoomsSorted) {
    const guestCount = getGuestCountSimple(entry.room, sections);
    const guestNames = entry.guests.map(g => g.replace(/\s*\*+$/, '').trim()).join(', ');
    wsK.addRow([entry.room, guestNames, guestCount, entry.nights || '', entry._checkin || '', entry._checkout || '']);
  }

  await wb.xlsx.writeFile(OUTPUT_FILE);
  console.log(`\nГотово: ${OUTPUT_FILE}`);
}

main().catch(console.error);
