// src/parser/excelParser.ts
import * as XLSX from 'xlsx';
import {
  AmbiguityIssue,
  ClassType,
  DataCompleteness,
  GroupHierarchy,
  ParsedScheduleResult,
  ScheduleEvent,
} from '../types';

/** Normalizuje identyfikator grupy z komórki Excela */
export function normalizeGroupToken(raw: string): string {
  const cleaned = raw.trim().replace(/\s+/g, ' ');
  const mGrupa = cleaned.match(/^grupa\s*(\d+)$/i);
  if (mGrupa) return `grupa ${mGrupa[1]}`;
  const mSub = cleaned.replace(/\s+/g, '').match(/^(\d+)([a-cA-C])([12])?$/);
  if (mSub) {
    return `${mSub[1]}${mSub[2].toLowerCase()}${mSub[3] ?? ''}`;
  }
  return cleaned.toLowerCase();
}

/** Rozkłada identyfikator podgrupy 4-osobowej (np. "1c2") na hierarchię */
export function parseGroupHierarchy(sub4: string): GroupHierarchy {
  const norm = sub4.trim().toLowerCase().replace(/\s+/g, '');
  const m = norm.match(/^(\d+)([abc])([12])$/);
  if (!m) {
    throw new Error(`Nieprawidłowy identyfikator podgrupy 4-osobowej: ${sub4}`);
  }
  const deanGroup = parseInt(m[1], 10);
  const letter8 = m[2] as 'a' | 'b' | 'c';
  return {
    sub4: `${deanGroup}${letter8}${m[3]}`,
    sub8: `${deanGroup}${letter8}`,
    deanGroup,
    letter8,
  };
}

/** Sprawdza, czy wydarzenie dotyczy wybranej podgrupy 4-osobowej */
export function doesEventMatchStudent(
  event: ScheduleEvent,
  studentSub4: string,
  preferred12Half: 'a' | 'b' | 'both' = 'both'
): boolean {
  if (event.groups.includes('ALL')) return true;
  const h = parseGroupHierarchy(studentSub4);

  if (event.is12PersonSplit && event.split12Half) {
    // Ćwiczenia 12-osobowe: sprawdzamy najpierw numer grupy dziekańskiej
    const eventGroupNums = event.groups
      .map((g) => g.match(/^(\d+)/)?.[1])
      .filter(Boolean)
      .map(Number);
    if (!eventGroupNums.includes(h.deanGroup)) return false;

    if (h.letter8 === 'a') return event.split12Half === 'a';
    if (h.letter8 === 'b') return event.split12Half === 'b';
    // Student z podgrupy 'c' (np. 1c1, 1c2):
    if (preferred12Half === 'both') return true;
    return event.split12Half === preferred12Half;
  }

  return (
    event.groups.includes(h.sub4) ||
    event.groups.includes(h.sub8) ||
    event.groups.includes(`grupa ${h.deanGroup}`)
  );
}

/** Oblicza różnicę godzin zegarowych między HH:MM a HH:MM */
export function calculateDurationHours(
  startTime: string | null,
  endTime: string | null
): number | null {
  if (!startTime || !endTime) return null;
  const [sh, sm] = startTime.split(':').map(Number);
  const [eh, em] = endTime.split(':').map(Number);
  if (isNaN(sh) || isNaN(sm) || isNaN(eh) || isNaN(em)) return null;
  const diffMinutes = eh * 60 + em - (sh * 60 + sm);
  if (diffMinutes <= 0) return null;
  return Math.round((diffMinutes / 60) * 100) / 100;
}

/** Zamienia zapis godziny z Excela (np. "8.00", "15.15") na "08:00", "15:15" */
export function normalizeTimeStr(raw: string): string | null {
  const m = raw.trim().match(/^(\d{1,2})[.:](\d{2})$/);
  if (!m) return null;
  const hh = m[1].padStart(2, '0');
  const mm = m[2];
  return `${hh}:${mm}`;
}

/** Wyciąga przedział godzin (startTime, endTime) z tekstu np. "8.00 - 14.00" */
export function extractTimeRange(text: string): {
  startTime: string | null;
  endTime: string | null;
} {
  const m = text.match(/(\d{1,2}[.:]\d{2})\s*-\s*(\d{1,2}[.:]\d{2})/);
  if (!m) return { startTime: null, endTime: null };
  return {
    startTime: normalizeTimeStr(m[1]),
    endTime: normalizeTimeStr(m[2]),
  };
}

/** Parsuje zakres tygodnia z kolumny A, np. "05.10. - 09.10.2026" lub "30.11. - 04.12.2026" */
export function parseWeekRange(raw: string): {
  mondayIso: string;
  fridayIso: string;
  year: number;
} | null {
  const cleaned = raw.trim();
  const m = cleaned.match(
    /(\d{2})\.(\d{2})\.?\s*-\s*(\d{2})\.(\d{2})\.(\d{4})/
  );
  if (!m) return null;
  const d1 = parseInt(m[1], 10);
  const m1 = parseInt(m[2], 10);
  const d2 = parseInt(m[3], 10);
  const m2 = parseInt(m[4], 10);
  const year2 = parseInt(m[5], 10);
  const year1 = m1 > m2 ? year2 - 1 : year2;

  const mondayIso = `${year1}-${String(m1).padStart(2, '0')}-${String(d1).padStart(2, '0')}`;
  const fridayIso = `${year2}-${String(m2).padStart(2, '0')}-${String(d2).padStart(2, '0')}`;
  return { mondayIso, fridayIso, year: year2 };
}

/** Dodaje dni do daty ISO YYYY-MM-DD */
export function addDaysIso(isoDate: string, days: number): string {
  const [y, m, d] = isoDate.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d + days));
  const ry = dt.getUTCFullYear();
  const rm = String(dt.getUTCMonth() + 1).padStart(2, '0');
  const rd = String(dt.getUTCDate()).padStart(2, '0');
  return `${ry}-${rm}-${rd}`;
}

/** Zamienia krótką datę DD.MM. na YYYY-MM-DD w kontekście semestru zimowego 2026/2027 */
export function resolveShortDate(ddmm: string, defaultOctDecYear = 2026): string | null {
  const m = ddmm.trim().match(/(\d{2})\.(\d{2})\.?/);
  if (!m) return null;
  const day = m[1];
  const month = parseInt(m[2], 10);
  const year = month >= 9 ? defaultOctDecYear : defaultOctDecYear + 1;
  return `${year}-${String(month).padStart(2, '0')}-${day}`;
}

/** Pomocnik do odczytu wartości komórki z uwzględnieniem scalonych komórek */
function createSheetReader(ws: XLSX.WorkSheet) {
  const merges = ws['!merges'] || [];
  return function getVal(row1Based: number, col1Based: number): string {
    const r0 = row1Based - 1;
    const c0 = col1Based - 1;
    for (const rng of merges) {
      if (r0 >= rng.s.r && r0 <= rng.e.r && c0 >= rng.s.c && c0 <= rng.e.c) {
        const addr = XLSX.utils.encode_cell({ r: rng.s.r, c: rng.s.c });
        const cell = ws[addr];
        return cell && cell.v != null ? String(cell.v) : '';
      }
    }
    const addr = XLSX.utils.encode_cell({ r: r0, c: c0 });
    const cell = ws[addr];
    return cell && cell.v != null ? String(cell.v) : '';
  };
}

function getRawCellVal(ws: XLSX.WorkSheet, row1Based: number, col1Based: number): string {
  const addr = XLSX.utils.encode_cell({ r: row1Based - 1, c: col1Based - 1 });
  const cell = ws[addr];
  return cell && cell.v != null ? String(cell.v) : '';
}

/** Rozszyfrowuje skróty budynków WUM na pełne nazwy i adresy ze słownika w arkuszu */
function resolveBuildingCode(text: string): {
  facilityName: string | null;
  address: string | null;
  building: string | null;
  room: string | null;
} {
  const t = text.replace(/\s+/g, ' ').trim();
  const roomMatch = t.match(/(sala(?:\s+seminaryjna|\s+komputerowa)?\s+[^\s,]+|s\.\s*\d+\s*Pato|aula\s+[AB]|119\s*CBI)/i);
  const room = roomMatch ? roomMatch[1].trim() : null;

  if (/NZJ/i.test(t) || /Ciołka\s*27/i.test(t)) {
    return {
      facilityName: 'Zakład Rozwoju Pielęgniarstwa, Nauk Społecznych i Medycznych (NZJ)',
      address: 'ul. Ciołka 27, 01-145 Warszawa',
      building: 'Budynek WNoZ (ul. Ciołka 27)',
      room,
    };
  }
  if (/NZN/i.test(t) || /ZPK/i.test(t)) {
    return {
      facilityName: 'Zakład Pielęgniarstwa Klinicznego (NZN)',
      address: 'ul. Ciołka 27, 01-145 Warszawa',
      building: 'Budynek WNoZ (ul. Ciołka 27)',
      room,
    };
  }
  if (/NZA/i.test(t)) {
    return {
      facilityName: 'Zakład Podstaw Pielęgniarstwa (NZA)',
      address: 'ul. Ciołka 27, 01-145 Warszawa',
      building: 'Zakład Podstaw Pielęgniarstwa (NZA)',
      room,
    };
  }
  if (/CBI/i.test(t)) {
    return {
      facilityName: 'Centrum Biblioteczno-Informacyjne WUM (CBI)',
      address: 'ul. Żwirki i Wigury 63, 02-091 Warszawa',
      building: 'Centrum Biblioteczno-Informacyjne (CBI)',
      room: room || (t.includes('119') ? 'sala 119' : null),
    };
  }
  if (/\bCD\b/i.test(t) || /Trojdena\s*2a/i.test(t)) {
    return {
      facilityName: 'Centrum Dydaktyczne WUM (CD)',
      address: 'ul. Trojdena 2a, 02-109 Warszawa',
      building: 'Centrum Dydaktyczne (CD)',
      room,
    };
  }
  if (/Litewska/i.test(t) || /Pato/i.test(t)) {
    return {
      facilityName: 'Budynek Patomorfologii / Litewska',
      address: 'ul. Litewska 14/16, 00-575 Warszawa',
      building: t.includes('Pato') ? 'Budynek Patomorfologii' : null,
      room,
    };
  }
  if (/Pawińskiego\s*3c/i.test(t)) {
    return {
      facilityName: 'Katedra i Zakład Farmakologii Doświadczalnej i Klinicznej',
      address: 'ul. Pawińskiego 3c ( / ul. Banacha 1b), 02-097 Warszawa',
      building: 'Niski parter, ul. Pawińskiego 3c',
      room: 'sala nr 7',
    };
  }
  return { facilityName: null, address: null, building: null, room };
}

/** Parsuje wielką scaloną komórkę BG9:CT20 z przypisaniem sal do grup dla POZ seminaria */
export function parsePozSeminariaRooms(rawText: string): Record<string, Record<string, string>> {
  // Zwraca mapę: dayName -> { "grupa 1": "sala 210 w NZJ", ... }
  const result: Record<string, Record<string, string>> = {
    poniedziałek: {},
    wtorek: {},
    środa: {},
    czwartek: {},
    piątek: {},
  };
  if (!rawText) return result;

  const lines = rawText.split(/\r?\n/);
  let mode: 'top' | 'bottom' = 'top';

  for (const line of lines) {
    if (/wtorek\s*:/i.test(line) && /czwartek\s*:/i.test(line)) {
      mode = 'bottom';
      continue;
    }
    if (/poniedziałek/i.test(line) && /środa/i.test(line)) {
      mode = 'top';
      continue;
    }

    // Wyciągnij wszystkie wystąpienia "gr. X - opis" lub "gr. X sala opis"
    const matches = Array.from(
      line.matchAll(/gr\.\s*(\d+)\s*[-–]?\s*((?:sala|s\.)[^g]+?)(?=\s{2,}gr\.|\s*$)/gi)
    );
    if (matches.length === 0) continue;

    if (mode === 'top') {
      // Kolumny tekstowe w linii: 0 -> poniedziałek, 1 -> środa, 2 -> piątek
      const daysOrder = ['poniedziałek', 'środa', 'piątek'];
      matches.forEach((m, idx) => {
        const day = daysOrder[idx];
        if (day) {
          result[day][`grupa ${m[1]}`] = m[2].trim();
        }
      });
    } else {
      // Kolumny tekstowe w linii: 0 -> wtorek, 1 -> czwartek
      const daysOrder = ['wtorek', 'czwartek'];
      matches.forEach((m, idx) => {
        const day = daysOrder[idx];
        if (day) {
          result[day][`grupa ${m[1]}`] = m[2].trim();
        }
      });
    }
  }
  return result;
}

const DAY_OFFSETS: Record<string, number> = {
  poniedziałek: 0,
  wtorek: 1,
  środa: 2,
  czwartek: 3,
  piątek: 4,
};

/** Główna funkcja parsująca skoroszyt Excela */
export function parseScheduleWorkbook(wb: XLSX.WorkBook): ParsedScheduleResult {
  const events: ScheduleEvent[] = [];
  const wsPlan = wb.Sheets['PLAN ZAJĘĆ'];
  const wsWyklady = wb.Sheets['WYKŁADY'];

  if (!wsPlan) {
    throw new Error('Brak arkusza "PLAN ZAJĘĆ" w pliku Excel.');
  }

  const getPlanVal = createSheetReader(wsPlan);

  // Słownik sal POZ seminaria z komórki BG9 (col 59, row 9)
  const pozRoomsRaw = getPlanVal(9, 59);
  const pozRoomsMap = parsePozSeminariaRooms(pozRoomsRaw);

  // 1. Odczyt tygodni z wierszy 7..21, kolumna A (1)
  const weeks: { row: number; mondayIso: string; fridayIso: string }[] = [];
  for (let r = 7; r <= 21; r++) {
    const rawRange = getRawCellVal(wsPlan, r, 1);
    const parsed = parseWeekRange(rawRange);
    if (parsed) {
      weeks.push({ row: r, mondayIso: parsed.mondayIso, fridayIso: parsed.fridayIso });
    }
  }

  const semesterStart = weeks.length > 0 ? weeks[0].mondayIso : null;
  const semesterEnd = weeks.length > 0 ? weeks[weeks.length - 1].fridayIso : null;

  // Słownik klinik CHIRURGII (kolumny 2..9) na podstawie wierszy 24..32
  const chirurgiaClinics: Record<number, { instructor: string; facility: string; address: string }> = {
    2: {
      instructor: 'Prof. M. Grąt',
      facility: 'Katedra Chirurgii Ogólnej, Transplantacyjnej i Wątroby (1WB1)',
      address: 'ul. Banacha 1, 02-097 Warszawa',
    },
    3: {
      instructor: 'Prof. M. Grąt',
      facility: 'Katedra Chirurgii Ogólnej, Transplantacyjnej i Wątroby (1WB1)',
      address: 'ul. Banacha 1, 02-097 Warszawa',
    },
    4: {
      instructor: 'Prof. M. Słodkowski',
      facility: 'Katedra i Klinika Chirurgii Ogólnej, Gastroenterologicznej i Onkologicznej (1W8)',
      address: 'ul. Banacha 1, 02-097 Warszawa',
    },
    5: {
      instructor: 'Prof. Z. Gałązka',
      facility: 'Katedra i Klinika Chirurgii Ogólnej, Naczyniowej, Endokrynologicznej i Transplantacyjnej (1W14)',
      address: 'ul. Banacha 1a, 02-097 Warszawa',
    },
    6: {
      instructor: 'Prof. Z. Gałązka',
      facility: 'Katedra i Klinika Chirurgii Ogólnej, Naczyniowej, Endokrynologicznej i Transplantacyjnej (1W14)',
      address: 'ul. Banacha 1a, 02-097 Warszawa',
    },
    7: {
      instructor: 'Prof. dr hab. M. Kosieradzki',
      facility: 'Katedra i Klinika Chirurgii Ogólnej i Transplantacyjnej (1W23)',
      address: 'ul. Nowogrodzka 59, 02-006 Warszawa',
    },
    8: {
      instructor: 'Prof. dr hab. M. Kosieradzki',
      facility: 'Katedra i Klinika Chirurgii Ogólnej i Transplantacyjnej (1W23)',
      address: 'ul. Nowogrodzka 59, 02-006 Warszawa',
    },
    9: {
      instructor: 'Prof. dr hab. M. Kosieradzki',
      facility: 'Katedra i Klinika Chirurgii Ogólnej i Transplantacyjnej (1W23)',
      address: 'ul. Nowogrodzka 59, 02-006 Warszawa',
    },
  };

  // Słownik klinik INTERNY (kolumny 11..17) na podstawie wierszy 33..44
  const internaClinics: Record<
    number,
    { instructor: string | null; facility: string; address: string; startTime: string; notes: string[] }
  > = {
    11: {
      instructor: 'Prof. K. Mucha',
      facility: 'Katedra i Klinika Transplantologii, Immunologii, Nefrologii i Chorób Wewnętrznych (1W22)',
      address: 'ul. Nowogrodzka 59, 02-006 Warszawa',
      startTime: '07:30',
      notes: ['Zajęcia u Prof. K. Muchy zaczynają się od 7:30'],
    },
    12: {
      instructor: 'Prof. K. Mucha',
      facility: 'Katedra i Klinika Transplantologii, Immunologii, Nefrologii i Chorób Wewnętrznych (1W22)',
      address: 'ul. Nowogrodzka 59, 02-006 Warszawa',
      startTime: '07:30',
      notes: ['Zajęcia u Prof. K. Muchy zaczynają się od 7:30'],
    },
    13: {
      instructor: 'Prof. K. Mucha',
      facility: 'Katedra i Klinika Transplantologii, Immunologii, Nefrologii i Chorób Wewnętrznych (1W22)',
      address: 'ul. Nowogrodzka 59, 02-006 Warszawa',
      startTime: '07:30',
      notes: ['Zajęcia u Prof. K. Muchy zaczynają się od 7:30'],
    },
    14: {
      instructor: null,
      facility: 'Klinika Gastroenterologii WNoZ',
      address: 'ul. Kondratowicza 8, 03-242 Warszawa',
      startTime: '08:00',
      notes: [],
    },
    15: {
      instructor: 'Prof. dr hab. Ł. Czyżewski',
      facility: 'Klinika Chorób Wewnętrznych i Kardiologii (Zakład Pielęgniarstwa Geriatrycznego NZE)',
      address: 'ul. Lindleya 4 (jednostka macierzysta: ul. Oczki 4), 02-005 Warszawa',
      startTime: '07:30',
      notes: [
        'Zajęcia rozpoczynają się od 7:30!',
        'Realizacja w Klinice Chorób Wewnętrznych i Kardiologii, ul. Lindleya 4 (w wierszu 37 podano również siedzibę Zakładu: ul. Oczki 4).',
      ],
    },
    16: {
      instructor: 'Prof. dr hab. Z. Gaciong',
      facility: 'Katedra i Klinika Chorób Wewnętrznych, Nadciśnienia Tętniczego i Angiologii',
      address: 'ul. Banacha 1a, 02-097 Warszawa',
      startTime: '08:00',
      notes: [],
    },
    17: {
      instructor: null,
      facility: 'Klinika Chorób Wewnętrznych WNoZ (NZW)',
      address: 'ul. Grenadierów 51/59, 04-073 Warszawa',
      startTime: '08:00',
      notes: [],
    },
  };

  // Słownik klinik PEDIATRII (kolumny 27..40) na podstawie wierszy 45..59
  const pediatriaClinics: Record<
    number,
    { instructor: string | null; facility: string; address: string; notes: string[] }
  > = {
    27: {
      instructor: 'Prof. dr hab. A. Banaszkiewicz',
      facility: 'Klinika Gastroenterologii i Żywienia Dzieci (1W33)',
      address: 'ul. Żwirki i Wigury 63A, 02-091 Warszawa',
      notes: ['Pierwsze spotkanie z Asystentem o godz. 8:00 przed wejściem do Kliniki / przy punkcie informacyjnym w holu głównym DSK.'],
    },
    28: {
      instructor: 'lek. med. Katarzyna Bogusz-Gogol',
      facility: 'Klinika Ortopedii i Traumatologii Dziecięcej',
      address: 'ul. Żwirki i Wigury 63A, 02-091 Warszawa',
      notes: ['Pierwsze spotkanie z Asystentem o godz. 8:00 przed wejściem do Kliniki / przy punkcie informacyjnym w holu głównym DSK.'],
    },
    29: {
      instructor: 'Prof. M. Kulus',
      facility: 'Klinika Pneumonologii i Alergologii Wieku Dziecięcego (1W34)',
      address: 'ul. Żwirki i Wigury 63A, 02-091 Warszawa',
      notes: ['Pierwsze spotkanie z Asystentem o godz. 8:00 przed wejściem do Kliniki / przy punkcie informacyjnym w holu głównym DSK.'],
    },
    30: {
      instructor: 'Dr hab. E. Kuchar',
      facility: 'Klinika Pediatrii z Oddziałem Obserwacyjnym (2W9)',
      address: 'ul. Żwirki i Wigury 63A, 02-091 Warszawa',
      notes: ['Pierwsze spotkanie z Asystentem o godz. 8:00 przed wejściem do Kliniki / przy punkcie informacyjnym w holu głównym DSK.'],
    },
    31: {
      instructor: null,
      facility: 'Szpital Dziecięcy (Niekłańska 1) — Oddział Pediatrii',
      address: 'ul. Niekłańska 4/24, 03-924 Warszawa',
      notes: ['Pierwsze spotkanie z koordynatorem o godz. 8:00 przy wejściu głównym.'],
    },
    32: {
      instructor: null,
      facility: 'Szpital Dziecięcy (Niekłańska 2) — Pododdział Patologii Noworodka i Niemowlęcia',
      address: 'ul. Niekłańska 4/24, 03-924 Warszawa',
      notes: ['Pierwsze spotkanie z koordynatorem o godz. 8:00 przy wejściu głównym.'],
    },
    33: {
      instructor: null,
      facility: 'Szpital Dziecięcy (Niekłańska 3) — Oddział Otolaryngologii',
      address: 'ul. Niekłańska 4/24, 03-924 Warszawa',
      notes: ['Pierwsze spotkanie z koordynatorem o godz. 8:00 przy wejściu głównym.'],
    },
    34: {
      instructor: null,
      facility: 'Szpital Dziecięcy (Niekłańska 4) — Oddział Neurochirurgii',
      address: 'ul. Niekłańska 4/24, 03-924 Warszawa',
      notes: ['Pierwsze spotkanie z koordynatorem o godz. 8:00 przy wejściu głównym.'],
    },
    35: {
      instructor: 'Dr hab. L. Zawadzka-Głos',
      facility: 'Klinika Otolaryngologii Dziecięcej (WLS16)',
      address: 'ul. Żwirki i Wigury 63A, 02-091 Warszawa',
      notes: ['Pierwsze spotkanie z Asystentem o godz. 8:00 przed wejściem do Kliniki / przy punkcie informacyjnym w holu głównym DSK.'],
    },
    36: {
      instructor: 'Dr hab. K. Szymańska',
      facility: 'Klinika Neurologii Dziecięcej i Chorób Rzadkich (WLS13)',
      address: 'ul. Żwirki i Wigury 63A, 02-091 Warszawa',
      notes: ['Pierwsze spotkanie z Asystentem o godz. 8:00 przed wejściem do Kliniki / przy punkcie informacyjnym w holu głównym DSK.'],
    },
    37: {
      instructor: 'Dr hab. P. Sosnowska-Sienkiewicz',
      facility: 'Klinika Chirurgii Dziecięcej (1W7)',
      address: 'ul. Żwirki i Wigury 63A, 02-091 Warszawa',
      notes: ['Pierwsze spotkanie z Asystentem o godz. 8:00 przed wejściem do Kliniki / przy punkcie informacyjnym w holu głównym DSK.'],
    },
    38: {
      instructor: 'Dr hab. R. Pietrzak',
      facility: 'Klinika Kardiologii Wieku Dziecięcego i Pediatrii Ogólnej (2M6)',
      address: 'ul. Żwirki i Wigury 63A, 02-091 Warszawa',
      notes: ['Pierwsze spotkanie z Asystentem o godz. 8:00 przed wejściem do Kliniki / przy punkcie informacyjnym w holu głównym DSK.'],
    },
    39: {
      instructor: 'Prof. dr hab. M. Mizerska-Wasiak',
      facility: 'Katedra i Klinika Pediatrii i Nefrologii (1WI)',
      address: 'ul. Żwirki i Wigury 63A, 02-091 Warszawa',
      notes: ['Pierwsze spotkanie z Asystentem o godz. 8:00 przed wejściem do Kliniki / przy punkcie informacyjnym w holu głównym DSK.'],
    },
    40: {
      instructor: 'Prof. dr hab. P. Łaguna',
      facility: 'Katedra i Klinika Onkologii, Hematologii Dziecięcej, Transplantologii Klinicznej i Chorób Rzadkich (1WG)',
      address: 'ul. Żwirki i Wigury 63A, 02-091 Warszawa',
      notes: ['Pierwsze spotkanie z Asystentem o godz. 8:00 przed wejściem do Kliniki / przy punkcie informacyjnym w holu głównym DSK.'],
    },
  };

  // Słownik POZ zajęcia praktyczne (kolumny 47..52) na podstawie wierszy 62..64
  const pozPraktClinics: Record<number, { instructor: string | null; facility: string; address: string }> = {
    47: {
      instructor: 'Dr Mariola Pietrzak',
      facility: 'SZPZLO Warszawa-Praga Północ',
      address: 'ul. Jagiellońska 34, Warszawa',
    },
    48: {
      instructor: 'Mgr Barbara Knoff',
      facility: 'SZPZLO Warszawa-Praga Północ',
      address: 'ul. Otwocka 1, Warszawa',
    },
    49: {
      instructor: 'Dr Anna Idzik',
      facility: 'SZPZLO Warszawa-Ochota',
      address: 'ul. Skarżyńskiego 1, Warszawa',
    },
    50: {
      instructor: 'Dr Beata Dziedzic',
      facility: 'SZPZLO Warszawa-Ochota',
      address: 'ul. Sanocka 6, Warszawa',
    },
    51: {
      instructor: null,
      facility: 'Przychodnia lekarska',
      address: 'ul. Jadźwingów 9, Warszawa',
    },
    52: {
      instructor: null,
      facility: 'Przychodnia POZ',
      address: 'ul. Soczi 1, Warszawa',
    },
  };

  // Iteracja po wszystkich 15 tygodniach (wiersze 7..21)
  for (const wk of weeks) {
    const r = wk.row;

    // --- A. CHIRURGIA I BLOK OPERACYJNY (kol. 2..9) + CSM (kol. 10) ---
    const csmChirRaw = getRawCellVal(wsPlan, r, 10).trim();
    const csmChirDate = csmChirRaw ? resolveShortDate(csmChirRaw) : null;
    const chirGroupsInRow: string[] = [];

    for (let c = 2; c <= 9; c++) {
      const rawVal = getRawCellVal(wsPlan, r, c).trim();
      if (!rawVal) continue;
      const grp = normalizeGroupToken(rawVal);
      chirGroupsInRow.push(grp);
      const clinic = chirurgiaClinics[c];

      for (let d = 0; d < 5; d++) {
        const dateIso = addDaysIso(wk.mondayIso, d);
        const colLetter = XLSX.utils.encode_col(c - 1);
        events.push({
          id: `CHIR-PRAKT-${r}-${c}-${dateIso}-${grp}`,
          subject: 'CHIRURGIA I BLOK OPERACYJNY',
          type: 'praktyczne',
          date: dateIso,
          startTime: '08:00',
          endTime: '14:00',
          durationHours: 6,
          groups: [grp],
          instructor: `${clinic.instructor} (Zakład NZS — p.o. kier. dr Marta Hreńczuk)`,
          facilityName: clinic.facility,
          address: clinic.address,
          building: null,
          room: null,
          notes:
            dateIso === csmChirDate
              ? [
                  `Uwaga: W tym tygodniu w kolumnie J wpisano zajęcia w Centrum Symulacji Medycznych (ul. Pawińskiego 3a, godz. 8:00–14:00) dla grup odbywających blok chirurgii.`,
                ]
              : [],
          completeness: 'complete',
          completenessNote: null,
          sourceRef: `PLAN ZAJĘĆ!${colLetter}${r}`,
        });
      }
    }

    if (csmChirDate && chirGroupsInRow.length > 0) {
      events.push({
        id: `CHIR-CSM-${r}-J-${csmChirDate}`,
        subject: 'CHIRURGIA I BLOK OPERACYJNY — Centrum Symulacji Medycznych',
        type: 'praktyczne',
        date: csmChirDate,
        startTime: '08:00',
        endTime: '14:00',
        durationHours: 6,
        groups: [...chirGroupsInRow],
        instructor: 'Zakład NZS (dr Marta Hreńczuk)',
        facilityName: 'Centrum Symulacji Medycznych WUM (CSM)',
        address: 'ul. Pawińskiego 3a, 02-106 Warszawa',
        building: 'Centrum Symulacji Medycznych',
        room: null,
        notes: [
          'Zajęcia w CSM odbywają się w wyznaczony wtorek w ramach bloku Chirurgii (w miejsce dyżuru klinicznego tego dnia).',
        ],
        completeness: 'ambiguous',
        completenessNote:
          'W kolumnie J podano datę wtorku w CSM bez wskazania, która z podgrup z wiersza ma CSM w 1., a która w 2. tygodniu bloku.',
        sourceRef: `PLAN ZAJĘĆ!J${r}`,
      });
    }

    // --- B. INTERNA ZAJĘCIA PRAKTYCZNE (kol. 11..17), PROF. R. STEC (kol. 18..25) i CSM (kol. 26) ---
    // Najpierw sprawdźmy, w które dni grupy idą do Prof. R. Steca (kol. 18..25)
    // R,S (18,19) -> poniedziałek (0); T,U (20,21) -> wtorek (1); V,W (22,23) -> czwartek (3); X,Y (24,25) -> piątek (4)
    const stecDaysByGroup: Record<string, Set<number>> = {};
    const stecColDayMap: Record<number, number> = {
      18: 0,
      19: 0,
      20: 1,
      21: 1,
      22: 3,
      23: 3,
      24: 4,
      25: 4,
    };

    for (let c = 18; c <= 25; c++) {
      const rawVal = getRawCellVal(wsPlan, r, c).trim();
      if (!rawVal) continue;
      const grp = normalizeGroupToken(rawVal);
      const dayOffset = stecColDayMap[c];
      if (!stecDaysByGroup[grp]) stecDaysByGroup[grp] = new Set();
      stecDaysByGroup[grp].add(dayOffset);

      const dateIso = addDaysIso(wk.mondayIso, dayOffset);
      const colLetter = XLSX.utils.encode_col(c - 1);
      events.push({
        id: `INT-STEC-${r}-${c}-${dateIso}-${grp}`,
        subject: 'INTERNA (Klinika Onkologii)',
        type: 'praktyczne',
        date: dateIso,
        startTime: '08:00',
        endTime: '14:00',
        durationHours: 6,
        groups: [grp],
        instructor: 'Prof. R. Stec',
        facilityName: 'Klinika Onkologii (2W11)',
        address: 'ul. Banacha 1a, 02-097 Warszawa',
        building: null,
        room: null,
        notes: ['Zajęcia u Prof. R. Steca realizowane w wyznaczony dzień tygodnia w ramach bloku Interny.'],
        completeness: 'complete',
        completenessNote: null,
        sourceRef: `PLAN ZAJĘĆ!${colLetter}${r}`,
      });
    }

    const internaGroupsInRow: string[] = [];
    for (let c = 11; c <= 17; c++) {
      const rawVal = getRawCellVal(wsPlan, r, c).trim();
      if (!rawVal) continue;
      const grp = normalizeGroupToken(rawVal);
      internaGroupsInRow.push(grp);
      const clinic = internaClinics[c];

      for (let d = 0; d < 5; d++) {
        // Pomijamy dni, w których grupa ma zajęcia u Prof. R. Steca (zgodnie z nagłówkiem K3:Q3)
        if (stecDaysByGroup[grp]?.has(d)) continue;

        const dateIso = addDaysIso(wk.mondayIso, d);
        const colLetter = XLSX.utils.encode_col(c - 1);
        const dur = calculateDurationHours(clinic.startTime, '14:00');
        events.push({
          id: `INT-PRAKT-${r}-${c}-${dateIso}-${grp}`,
          subject: 'INTERNA',
          type: 'praktyczne',
          date: dateIso,
          startTime: clinic.startTime,
          endTime: '14:00',
          durationHours: dur,
          groups: [grp],
          instructor: clinic.instructor,
          facilityName: clinic.facility,
          address: clinic.address,
          building: null,
          room: null,
          notes: [...clinic.notes],
          completeness: 'complete',
          completenessNote: null,
          sourceRef: `PLAN ZAJĘĆ!${colLetter}${r}`,
        });
      }
    }

    // Kolumna Z (26) — INTERNA w CSM
    const csmIntRaw = getRawCellVal(wsPlan, r, 26).trim();
    if (csmIntRaw && internaGroupsInRow.length > 0) {
      if (csmIntRaw.includes('10a') && csmIntRaw.includes('13.10')) {
        // Specjalny wpis w Z8: "gr. 10a - 13.10.\npozostałe grupy 14.10."
        events.push({
          id: `INT-CSM-${r}-Z-2026-10-13-10a`,
          subject: 'INTERNA — Centrum Symulacji Medycznych',
          type: 'praktyczne',
          date: '2026-10-13',
          startTime: '08:00',
          endTime: '14:00',
          durationHours: 6,
          groups: ['10a'],
          instructor: null,
          facilityName: 'Centrum Symulacji Medycznych WUM (CSM)',
          address: 'ul. Pawińskiego 3a, 02-106 Warszawa',
          building: 'Centrum Symulacji Medycznych',
          room: null,
          notes: ['Termin wyznaczony indywidualnie dla gr. 10a w komórce Z8 (13.10.2026).'],
          completeness: 'complete',
          completenessNote: null,
          sourceRef: `PLAN ZAJĘĆ!Z${r}`,
        });
        const otherGroups = internaGroupsInRow.filter((g) => g !== '10a');
        if (otherGroups.length > 0) {
          events.push({
            id: `INT-CSM-${r}-Z-2026-10-14-others`,
            subject: 'INTERNA — Centrum Symulacji Medycznych',
            type: 'praktyczne',
            date: '2026-10-14',
            startTime: '08:00',
            endTime: '14:00',
            durationHours: 6,
            groups: otherGroups,
            instructor: null,
            facilityName: 'Centrum Symulacji Medycznych WUM (CSM)',
            address: 'ul. Pawińskiego 3a, 02-106 Warszawa',
            building: 'Centrum Symulacji Medycznych',
            room: null,
            notes: ['Zajęcia w CSM dla pozostałych grup z bloku Interny (14.10.2026).'],
            completeness: 'ambiguous',
            completenessNote:
              'W kolumnie Z podano datę zajęć w CSM dla grup z bloku Interny bez rozbicia na 1. i 2. tydzień bloku.',
            sourceRef: `PLAN ZAJĘĆ!Z${r}`,
          });
        }
      } else {
        const csmIntDate = resolveShortDate(csmIntRaw);
        if (csmIntDate) {
          events.push({
            id: `INT-CSM-${r}-Z-${csmIntDate}`,
            subject: 'INTERNA — Centrum Symulacji Medycznych',
            type: 'praktyczne',
            date: csmIntDate,
            startTime: '08:00',
            endTime: '14:00',
            durationHours: 6,
            groups: [...internaGroupsInRow],
            instructor: null,
            facilityName: 'Centrum Symulacji Medycznych WUM (CSM)',
            address: 'ul. Pawińskiego 3a, 02-106 Warszawa',
            building: 'Centrum Symulacji Medycznych',
            room: null,
            notes: ['Zajęcia w CSM w ramach bloku Interny (w miejsce dyżuru klinicznego tego dnia).'],
            completeness: 'ambiguous',
            completenessNote:
              'W kolumnie Z podano datę CSM bez wskazania, która z podgrup z wiersza ma CSM w 1., a która w 2. tygodniu bloku.',
            sourceRef: `PLAN ZAJĘĆ!Z${r}`,
          });
        }
      }
    }

    // --- C. PEDIATRIA ZAJĘCIA PRAKTYCZNE (kol. 27..40) ---
    for (let c = 27; c <= 40; c++) {
      const rawVal = getRawCellVal(wsPlan, r, c).trim();
      if (!rawVal) continue;
      const grp = normalizeGroupToken(rawVal);
      const clinic = pediatriaClinics[c];

      for (let d = 0; d < 5; d++) {
        const dateIso = addDaysIso(wk.mondayIso, d);
        const colLetter = XLSX.utils.encode_col(c - 1);
        events.push({
          id: `PED-PRAKT-${r}-${c}-${dateIso}-${grp}`,
          subject: 'PEDIATRIA',
          type: 'praktyczne',
          date: dateIso,
          startTime: '08:00',
          endTime: '14:00',
          durationHours: 6,
          groups: [grp],
          instructor: clinic.instructor,
          facilityName: clinic.facility,
          address: clinic.address,
          building: null,
          room: null,
          notes: d === 0 ? [...clinic.notes] : [],
          completeness: 'complete',
          completenessNote: null,
          sourceRef: `PLAN ZAJĘĆ!${colLetter}${r}`,
        });
      }
    }

    // --- D. PODSTAWY REHABILITACJI ĆWICZENIA (kol. 41..46, grupy 12-osobowe) ---
    const rehabColConfig: Record<number, { dayOffset: number; start: string; end: string }> = {
      41: { dayOffset: 0, start: '08:15', end: '11:15' },
      42: { dayOffset: 0, start: '11:15', end: '14:15' },
      43: { dayOffset: 1, start: '08:00', end: '14:00' },
      44: { dayOffset: 2, start: '08:00', end: '14:00' },
      45: { dayOffset: 3, start: '08:00', end: '14:00' },
      46: { dayOffset: 4, start: '08:00', end: '14:00' },
    };

    for (let c = 41; c <= 46; c++) {
      const rawVal = getRawCellVal(wsPlan, r, c).trim();
      if (!rawVal) continue;
      const cfg = rehabColConfig[c];
      const lines = rawVal.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
      const grp = normalizeGroupToken(lines[0]);
      const half = grp.endsWith('b') ? 'b' : 'a';

      let dateIso = addDaysIso(wk.mondayIso, cfg.dayOffset);
      let startTime: string | null = cfg.start;
      let endTime: string | null = cfg.end;
      let room = 'sala 101';
      const extraNotes: string[] = [];

      // Sprawdź, czy w komórce nadpisano datę/godziny/salę (jak w wierszu 7 dla 10a/10b)
      if (lines.length > 1) {
        const fullCellText = lines.slice(1).join(' ');
        const customDateMatch = fullCellText.match(/(\d{2}\.\d{2}\.)/);
        if (customDateMatch) {
          const resolved = resolveShortDate(customDateMatch[1]);
          if (resolved) {
            if (resolved !== dateIso) {
              extraNotes.push(`Uwaga: W komórce wpisano niestandardową datę ${customDateMatch[1]} (zamiast domyślnego dnia z wiersza).`);
            }
            dateIso = resolved;
          }
        }
        const customTime = extractTimeRange(fullCellText);
        if (customTime.startTime && customTime.endTime) {
          startTime = customTime.startTime;
          endTime = customTime.endTime;
        }
        const roomM = fullCellText.match(/sala\s+\d+[^,]*/i);
        if (roomM) room = roomM[0].trim();
      }

      const colLetter = XLSX.utils.encode_col(c - 1);
      events.push({
        id: `REHAB-CW-${r}-${c}-${dateIso}-${grp}`,
        subject: 'PODSTAWY REHABILITACJI',
        type: 'cwiczenia',
        date: dateIso,
        startTime,
        endTime,
        durationHours: calculateDurationHours(startTime, endTime),
        groups: [grp],
        is12PersonSplit: true,
        split12Half: half,
        instructor: 'Prof. dr hab. B. Czarkowska-Pączek',
        facilityName: 'Zakład Pielęgniarstwa Klinicznego (NZN)',
        address: 'ul. Ciołka 27, 01-145 Warszawa',
        building: 'Budynek WNoZ (ul. Ciołka 27)',
        room,
        notes: extraNotes,
        completeness: 'complete',
        completenessNote: null,
        sourceRef: `PLAN ZAJĘĆ!${colLetter}${r}`,
      });
    }

    // --- E. POZ ZAJĘCIA PRAKTYCZNE (kol. 47..52, grupy 4-osobowe) ---
    for (let c = 47; c <= 52; c++) {
      const rawVal = getRawCellVal(wsPlan, r, c).trim();
      if (!rawVal) continue;
      const grp = normalizeGroupToken(rawVal);
      const clinic = pozPraktClinics[c];

      for (let d = 0; d < 5; d++) {
        const dateIso = addDaysIso(wk.mondayIso, d);
        const colLetter = XLSX.utils.encode_col(c - 1);
        events.push({
          id: `POZ-PRAKT-${r}-${c}-${dateIso}-${grp}`,
          subject: 'POZ (Podstawowa Opieka Zdrowotna)',
          type: 'praktyczne',
          date: dateIso,
          startTime: null,
          endTime: null,
          durationHours: null,
          groups: [grp],
          instructor: clinic.instructor,
          facilityName: clinic.facility,
          address: clinic.address,
          building: null,
          room: null,
          notes: [
            'Pierwsze spotkanie — szkolenie RODO i BHP od godz. 8:30 (adresy jednostek szkoleniowych na stronie Zakładu Rozwoju Pielęgniarstwa, ul. Ciołka 27).',
          ],
          completeness: 'partial',
          completenessNote:
            'Brak dokładnych godzin codziennych praktyk POZ w pliku Excel (podano jedynie łączny wymiar 40h oraz godz. 8:30 dla szkolenia BHP/RODO).',
          sourceRef: `PLAN ZAJĘĆ!${colLetter}${r}`,
        });
      }
    }

    // --- F. PROMOCJA ZDROWIA ZAJĘCIA PRAKTYCZNE (kol. 53..55, grupy 8-osobowe) ---
    const promPraktConfig: Record<number, { start: string; end: string }> = {
      53: { start: '08:00', end: '11:00' },
      54: { start: '11:15', end: '14:15' },
      55: { start: '08:00', end: '11:00' },
    };

    for (let c = 53; c <= 55; c++) {
      const rawVal = getRawCellVal(wsPlan, r, c).trim();
      if (!rawVal) continue;
      const grp = normalizeGroupToken(rawVal);
      const cfg = promPraktConfig[c];

      for (let d = 0; d < 5; d++) {
        const dateIso = addDaysIso(wk.mondayIso, d);
        const colLetter = XLSX.utils.encode_col(c - 1);

        let facilityName = 'Zakład Podstaw Pielęgniarstwa (NZA)';
        let address: string | null = 'ul. Ciołka 27, 01-145 Warszawa';
        let building: string | null = 'Budynek WNoZ (NZA)';
        let room: string | null = null;
        const notes: string[] = [];

        if (c === 53 || c === 54) {
          if (d === 0) {
            room = 'sala 202 w NZA';
          } else if (d === 4) {
            facilityName = 'Centrum Dydaktyczne WUM (CD)';
            address = 'ul. Trojdena 2a, 02-109 Warszawa';
            building = 'Centrum Dydaktyczne (CD)';
            room = 'sala 221 w CD';
          } else {
            room = 'pon.: s. 202 (NZA) / pt.: s. 221 (CD)';
            notes.push('W nagłówku kolumny wskazano sale dla poniedziałku (s. 202 NZA) i piątku (s. 221 CD).');
          }
        } else if (c === 55) {
          if (d === 0) {
            facilityName = 'Centrum Dydaktyczne WUM (CD)';
            address = 'ul. Trojdena 2a, 02-109 Warszawa';
            building = 'Centrum Dydaktyczne (CD)';
            room = 'sala 126 w CD';
          } else if (d === 4) {
            if (dateIso <= '2026-10-30') {
              facilityName = 'Zakład Rozwoju Pielęgniarstwa (NZJ)';
              address = 'ul. Ciołka 27, 01-145 Warszawa';
              building = 'Budynek WNoZ (NZJ)';
              room = 'sala 210 w NZJ';
            } else {
              facilityName = 'Budynek WUM przy ul. Litewskiej';
              address = 'ul. Litewska 14/16, 00-575 Warszawa';
              building = null;
              room = 'sala 104';
            }
          } else {
            room = 'pon.: s. 126 (CD) / pt.: s. 210 (NZJ) lub s. 104 (Litewska 14/16)';
          }
        }

        events.push({
          id: `PROM-PRAKT-${r}-${c}-${dateIso}-${grp}`,
          subject: 'PROMOCJA ZDROWIA',
          type: 'praktyczne',
          date: dateIso,
          startTime: cfg.start,
          endTime: cfg.end,
          durationHours: 3,
          groups: [grp],
          instructor: 'Dr hab. E. Krzych-Fałta',
          facilityName,
          address,
          building,
          room,
          notes,
          completeness: 'complete',
          completenessNote: null,
          sourceRef: `PLAN ZAJĘĆ!${colLetter}${r}`,
        });
      }
    }

    // --- G. PROMOCJA ZDROWIA SEMINARIA (kol. 56..57) ---
    for (const c of [56, 57]) {
      const rawVal = getRawCellVal(wsPlan, r, c).trim();
      if (!rawVal) continue;
      const grp = normalizeGroupToken(rawVal);
      const dayOffset = c === 56 ? 2 : 4; // 56 -> środa, 57 -> piątek
      const dateIso = addDaysIso(wk.mondayIso, dayOffset);
      const colLetter = XLSX.utils.encode_col(c - 1);

      events.push({
        id: `PROM-SEM-${r}-${c}-${dateIso}-${grp}`,
        subject: 'PROMOCJA ZDROWIA',
        type: 'seminarium',
        date: dateIso,
        startTime: '15:00',
        endTime: '18:45',
        durationHours: 3.75,
        groups: [grp],
        instructor: 'Prof. dr hab. M. Gujski',
        facilityName: 'Zakład Zdrowia Publicznego (NZB) — CBI',
        address: 'ul. Żwirki i Wigury 63, 02-091 Warszawa',
        building: 'Centrum Biblioteczno-Informacyjne (CBI)',
        room: 'sala 119 CBI',
        notes: [],
        completeness: 'complete',
        completenessNote: null,
        sourceRef: `PLAN ZAJĘĆ!${colLetter}${r}`,
      });
    }

    // --- H. POZ SEMINARIA (kol. 58..99, wiersz 7) ---
    if (r === 7) {
      for (let c = 58; c <= 99; c++) {
        const rawVal = getRawCellVal(wsPlan, r, c).trim();
        if (!rawVal) continue;
        const grp = normalizeGroupToken(rawVal);
        const dayHeader = getPlanVal(4, c).trim().toLowerCase();
        const timeHeader = getPlanVal(5, c).trim();
        const dayOffset = DAY_OFFSETS[dayHeader] ?? 0;
        const dateIso = addDaysIso(wk.mondayIso, dayOffset);
        const { startTime, endTime } = extractTimeRange(timeHeader);

        const roomDesc = pozRoomsMap[dayHeader]?.[grp] || '';
        const loc = roomDesc
          ? resolveBuildingCode(roomDesc)
          : {
              facilityName: 'Zakład Rozwoju Pielęgniarstwa, Nauk Społecznych i Medycznych (NZJ)',
              address: 'ul. Ciołka 27, 01-145 Warszawa',
              building: 'Budynek WNoZ (ul. Ciołka 27)',
              room: null,
            };

        const colLetter = XLSX.utils.encode_col(c - 1);
        events.push({
          id: `POZ-SEM-${r}-${c}-${dateIso}-${grp}`,
          subject: 'POZ (Podstawowa Opieka Zdrowotna)',
          type: 'seminarium',
          date: dateIso,
          startTime,
          endTime,
          durationHours: calculateDurationHours(startTime, endTime),
          groups: [grp],
          instructor: 'Dr hab. T. Kryczka',
          facilityName: loc.facilityName || 'Zakład Rozwoju Pielęgniarstwa (NZJ)',
          address: loc.address || 'ul. Ciołka 27, 01-145 Warszawa',
          building: loc.building,
          room: roomDesc || loc.room,
          notes: [],
          completeness: 'complete',
          completenessNote: null,
          sourceRef: `PLAN ZAJĘĆ!${colLetter}${r}`,
        });
      }
    }

    // --- I. POZ ĆWICZENIA (kol. 100..102, grupy 12-osobowe) ---
    const pozCwConfig: Record<
      number,
      { dayOffset: number; start: string; end: string; room: string; building: string; facility: string }
    > = {
      100: {
        dayOffset: 2,
        start: '08:00',
        end: '11:45',
        room: 'sala 210 w NZJ',
        building: 'Budynek WNoZ (NZJ)',
        facility: 'Zakład Rozwoju Pielęgniarstwa, Nauk Społecznych i Medycznych (NZJ)',
      },
      101: {
        dayOffset: 2,
        start: '12:00',
        end: '15:45',
        room: 'sala 210 w NZJ',
        building: 'Budynek WNoZ (NZJ)',
        facility: 'Zakład Rozwoju Pielęgniarstwa, Nauk Społecznych i Medycznych (NZJ)',
      },
      102: {
        dayOffset: 4,
        start: '15:00',
        end: '18:45',
        room: 'sala 102 w NZN',
        building: 'Budynek WNoZ (NZN)',
        facility: 'Zakład Pielęgniarstwa Klinicznego (NZN)',
      },
    };

    for (let c = 100; c <= 102; c++) {
      const rawVal = getRawCellVal(wsPlan, r, c).trim();
      if (!rawVal) continue;
      const grp = normalizeGroupToken(rawVal);
      const half = grp.endsWith('b') ? 'b' : 'a';
      const cfg = pozCwConfig[c];
      const dateIso = addDaysIso(wk.mondayIso, cfg.dayOffset);
      const colLetter = XLSX.utils.encode_col(c - 1);

      events.push({
        id: `POZ-CW-${r}-${c}-${dateIso}-${grp}`,
        subject: 'POZ (Podstawowa Opieka Zdrowotna)',
        type: 'cwiczenia',
        date: dateIso,
        startTime: cfg.start,
        endTime: cfg.end,
        durationHours: 3.75,
        groups: [grp],
        is12PersonSplit: true,
        split12Half: half,
        instructor: 'Dr hab. T. Kryczka',
        facilityName: cfg.facility,
        address: 'ul. Ciołka 27, 01-145 Warszawa',
        building: cfg.building,
        room: cfg.room,
        notes: [],
        completeness: 'complete',
        completenessNote: null,
        sourceRef: `PLAN ZAJĘĆ!${colLetter}${r}`,
      });
    }

    // --- J. FARMAKOLOGIA SEMINARIA (kol. 103..105) ---
    const farmConfig: Record<number, { dayOffset: number; defaultRoom: string }> = {
      103: { dayOffset: 0, defaultRoom: 'sala 203 w CD' },
      104: { dayOffset: 1, defaultRoom: 'sala nr 7 niski parter, ul. Pawińskiego 3c' },
      105: { dayOffset: 3, defaultRoom: 'sala komputerowa 234 w CD' },
    };

    for (let c = 103; c <= 105; c++) {
      const rawVal = getRawCellVal(wsPlan, r, c).trim();
      if (!rawVal) continue;
      const lines = rawVal.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
      const grp = normalizeGroupToken(lines[0]);
      const cfg = farmConfig[c];
      const dateIso = addDaysIso(wk.mondayIso, cfg.dayOffset);
      const extraText = lines.slice(1).join(' ');
      const roomText = extraText || cfg.defaultRoom;
      const loc = resolveBuildingCode(roomText);
      const colLetter = XLSX.utils.encode_col(c - 1);

      events.push({
        id: `FARM-SEM-${r}-${c}-${dateIso}-${grp}`,
        subject: 'FARMAKOLOGIA',
        type: 'seminarium',
        date: dateIso,
        startTime: '10:15',
        endTime: '14:00',
        durationHours: 3.75,
        groups: [grp],
        instructor: 'Prof. dr hab. D. Mirowska-Guzel',
        facilityName: loc.facilityName || 'Katedra i Zakład Farmakologii Doświadczalnej i Klinicznej (1M9)',
        address: loc.address || 'ul. Banacha 1b, 02-097 Warszawa',
        building: loc.building,
        room: extraText ? `${cfg.defaultRoom} -> ZMIANA: ${extraText}` : cfg.defaultRoom,
        notes: extraText ? [extraText] : [],
        completeness: 'complete',
        completenessNote: null,
        sourceRef: `PLAN ZAJĘĆ!${colLetter}${r}`,
      });
    }

    // --- K. INTERNA SEMINARIA — DR HAB. T. KRYCZKA (kol. 106..110) ---
    const intSemKryczkaConfig: Record<
      number,
      { dayOffset: number; start: string; end: string; defaultRoom: string }
    > = {
      106: { dayOffset: 1, start: '15:00', end: '18:45', defaultRoom: 'sala 210 w NZJ' },
      107: { dayOffset: 2, start: '15:45', end: '19:30', defaultRoom: 'sala 210 w NZJ' },
      108: { dayOffset: 3, start: '16:45', end: '20:30', defaultRoom: 'aula A w CD' },
      109: { dayOffset: 3, start: '15:00', end: '18:45', defaultRoom: 'sala 210 w NZJ' },
      110: { dayOffset: 4, start: '15:15', end: '19:00', defaultRoom: 'sala 125 w CBI' },
    };

    for (let c = 106; c <= 110; c++) {
      const rawVal = getRawCellVal(wsPlan, r, c).trim();
      if (!rawVal) continue;
      const lines = rawVal.split(/\r?\n/).map((l) => l.trim()).filter(Boolean);
      const grp = normalizeGroupToken(lines[0]);
      const cfg = intSemKryczkaConfig[c];
      const dateIso = addDaysIso(wk.mondayIso, cfg.dayOffset);
      const extraText = lines.slice(1).join(' ');
      const roomText = extraText || cfg.defaultRoom;
      const loc = resolveBuildingCode(roomText);
      const colLetter = XLSX.utils.encode_col(c - 1);

      events.push({
        id: `INT-SEM-KRY-${r}-${c}-${dateIso}-${grp}`,
        subject: 'INTERNA',
        type: 'seminarium',
        date: dateIso,
        startTime: cfg.start,
        endTime: cfg.end,
        durationHours: calculateDurationHours(cfg.start, cfg.end),
        groups: [grp],
        instructor: 'Dr hab. T. Kryczka',
        facilityName: loc.facilityName || 'Zakład Rozwoju Pielęgniarstwa (NZJ)',
        address: loc.address || 'ul. Ciołka 27, 01-145 Warszawa',
        building: loc.building,
        room: extraText ? extraText : cfg.defaultRoom,
        notes: extraText ? [`Zmiana sali w planie: ${extraText}`] : [],
        completeness: 'complete',
        completenessNote: null,
        sourceRef: `PLAN ZAJĘĆ!${colLetter}${r}`,
      });
    }

    // --- L. PEDIATRIA NZYN SEMINARIA (kol. 111..112) ---
    for (const c of [111, 112]) {
      const rawVal = getRawCellVal(wsPlan, r, c).trim();
      if (!rawVal) continue;
      const grp = normalizeGroupToken(rawVal);
      const dayOffset = c === 111 ? 2 : 4;
      const start = c === 111 ? '15:30' : '15:15';
      const end = c === 111 ? '19:15' : '19:00';
      const dateIso = addDaysIso(wk.mondayIso, dayOffset);
      const colLetter = XLSX.utils.encode_col(c - 1);

      events.push({
        id: `PED-SEM-${r}-${c}-${dateIso}-${grp}`,
        subject: 'PEDIATRIA (NZYN)',
        type: 'seminarium',
        date: dateIso,
        startTime: start,
        endTime: end,
        durationHours: 3.75,
        groups: [grp],
        instructor: null,
        facilityName: 'Klinika Neonatologii i Chorób Rzadkich (DSK)',
        address: 'ul. Żwirki i Wigury 63A, 02-091 Warszawa',
        building: 'Szpital Pediatryczny WUM',
        room: 'sala seminaryjna w Klinice Neonatologii i Chorób Rzadkich',
        notes: [],
        completeness: 'complete',
        completenessNote: null,
        sourceRef: `PLAN ZAJĘĆ!${colLetter}${r}`,
      });
    }

    // --- M. INTERNA SEMINARIA — PROF. J. WYZGAŁ (kol. 113..114) ---
    for (const c of [113, 114]) {
      const rawVal = getRawCellVal(wsPlan, r, c).trim();
      if (!rawVal) continue;
      const grp = normalizeGroupToken(rawVal);
      const dayOffset = c === 113 ? 2 : 4;
      const dateIso = addDaysIso(wk.mondayIso, dayOffset);
      const colLetter = XLSX.utils.encode_col(c - 1);

      events.push({
        id: `INT-SEM-WYZ-${r}-${c}-${dateIso}-${grp}`,
        subject: 'INTERNA',
        type: 'seminarium',
        date: dateIso,
        startTime: '15:15',
        endTime: '19:00',
        durationHours: 3.75,
        groups: [grp],
        instructor: 'Prof. J. Wyzgał',
        facilityName: 'Zakład Podstaw Pielęgniarstwa (NZM)',
        address: 'ul. Nowogrodzka 59, 02-006 Warszawa',
        building: 'Pawilon nr VIII, I piętro (boczne wejście do laboratorium)',
        room: 'sala seminaryjna I p.',
        notes: ['Boczne wejście do laboratorium, pawilon nr VIII, I piętro.'],
        completeness: 'complete',
        completenessNote: null,
        sourceRef: `PLAN ZAJĘĆ!${colLetter}${r}`,
      });
    }

    // --- N. CHIRURGIA SEMINARIA — DR M. HREŃCZUK (kol. 115..116) ---
    for (const c of [115, 116]) {
      const rawVal = getRawCellVal(wsPlan, r, c).trim();
      if (!rawVal) continue;
      const grp = normalizeGroupToken(rawVal);
      const dayOffset = c === 115 ? 2 : 4;
      const dateIso = addDaysIso(wk.mondayIso, dayOffset);
      const colLetter = XLSX.utils.encode_col(c - 1);

      events.push({
        id: `CHIR-SEM-${r}-${c}-${dateIso}-${grp}`,
        subject: 'CHIRURGIA',
        type: 'seminarium',
        date: dateIso,
        startTime: '15:15',
        endTime: '19:00',
        durationHours: 3.75,
        groups: [grp],
        instructor: 'Dr M. Hreńczuk',
        facilityName: 'Zakład Pielęgniarstwa Chirurgicznego, Transplantacyjnego i Leczenia Pozaustrojowego (NZS)',
        address: 'ul. Nowogrodzka 59, 02-006 Warszawa',
        building: 'Pawilon XI D1',
        room: 'sala seminaryjna 128 w NZS',
        notes: [],
        completeness: 'complete',
        completenessNote: null,
        sourceRef: `PLAN ZAJĘĆ!${colLetter}${r}`,
      });
    }
  }

  // 2. Parsowanie arkusza "WYKŁADY"
  const elearningNotes: string[] = [];
  if (wsWyklady) {
    const note1 = getRawCellVal(wsWyklady, 2, 1).trim();
    const note2 = getRawCellVal(wsWyklady, 20, 1).trim();
    if (note1) elearningNotes.push(note1);
    if (note2) elearningNotes.push(note2);

    // Wiersze 4..17 (Wtorki — Aula B) oraz 21..35 (Czwartki — Aula A)
    for (let r = 4; r <= 35; r++) {
      if (r >= 18 && r <= 20) continue;
      const rawDate = getRawCellVal(wsWyklady, r, 1).trim();
      if (!rawDate) continue;
      const dateIso = resolveShortDate(rawDate);
      if (!dateIso) continue;

      // Uwaga: w wierszu 21 (08.10.) treść wykładu wpisano w kolumnie C (3) zamiast B (2)!
      const rawB = getRawCellVal(wsWyklady, r, 2).trim();
      const rawC = getRawCellVal(wsWyklady, r, 3).trim();
      const content = rawB || rawC;
      if (!content || /nie ma wykładów/i.test(content)) continue;

      const { startTime, endTime } = extractTimeRange(content);
      const isTeams = /TEAMS/i.test(content);
      const isTuesdayBlock = r <= 17;

      // Wyciągnij przedmiot (wielkie litery na początku) i prowadzącego
      const beforeTime = content.split(/\d{1,2}[.:]\d{2}/)[0].trim();
      const subjMatch = beforeTime.match(/^([A-ZŚĆŻŹŁÓĘĄ\s]+?)\s+((?:prof\.|dr)\s+.+)$/i);
      const subject = subjMatch ? subjMatch[1].trim() : beforeTime;
      const instructor = subjMatch ? subjMatch[2].trim() : null;

      events.push({
        id: `WYKLAD-${r}-${dateIso}`,
        subject: `${subject} (Wykład)`,
        type: 'wyklad',
        date: dateIso,
        startTime,
        endTime,
        durationHours: calculateDurationHours(startTime, endTime),
        groups: ['ALL'],
        instructor,
        facilityName: isTeams ? 'Zajęcia zdalne (MS Teams)' : 'Centrum Dydaktyczne WUM (CD)',
        address: isTeams ? 'Online — platforma MS Teams' : 'ul. Trojdena 2a, 02-109 Warszawa',
        building: isTeams ? null : 'Centrum Dydaktyczne',
        room: isTeams ? 'MS Teams' : isTuesdayBlock ? 'Aula B' : 'Aula A',
        notes: isTeams ? ['Wykład realizowany zdalnie na platformie MS Teams.'] : [],
        completeness: 'complete',
        completenessNote: null,
        sourceRef: `WYKŁADY!${rawB ? 'B' : 'C'}${r}`,
      });
    }
  }

  // Lista wszystkich 78 podgrup 4-osobowych (1a1..13c2)
  const availableSub4Groups: string[] = [];
  for (let g = 1; g <= 13; g++) {
    for (const l of ['a', 'b', 'c']) {
      for (const n of ['1', '2']) {
        availableSub4Groups.push(`${g}${l}${n}`);
      }
    }
  }

  const ambiguities: AmbiguityIssue[] = [
    {
      id: 'AMB-12OS',
      title: 'Podział grup 24-osobowych na 12-osobowe (Rehabilitacja ćw. i POZ ćw.)',
      description:
        'W ćwiczeniach 12-osobowych grupy dzielą się na dwie połowy (np. 1a i 1b), podczas gdy na zajęciach klinicznych 8-osobowych dzielą się na trzy części (1a, 1b, 1c). Plik Excel nie precyzuje, którzy studenci z podgrupy "c" należą do połowy "a", a którzy do "b".',
      sourceCells: 'PLAN ZAJĘĆ!AO2:AT21 oraz CV2:CX21',
      resolution:
        'Dla studentów podgrup "c" (np. 1c1, 1c2) aplikacja pozwala wybrać w ustawieniach przypisaną połowę 12-osobową (A lub B) albo wyświetla obie z odpowiednim oznaczeniem.',
    },
    {
      id: 'AMB-POZ-HOURS',
      title: 'Brak godzin trwania praktyk POZ',
      description:
        'Dla przedmiotu "POZ zajęcia praktyczne" (40 godz.) podano jedynie godzinę rozpoczęcia szkolenia BHP i RODO pierwszego dnia (od godz. 8:30), bez codziennych godzin od–do w przychodniach.',
      sourceCells: 'PLAN ZAJĘĆ!AU2:AZ21',
      resolution:
        'Wydarzenia mają jawnie oznaczony brak godzin w planie. Są liczone jako wydarzenia w semestrze, ale nie dodają zmyślonej liczby godzin do statystyk.',
    },
    {
      id: 'AMB-CSM',
      title: 'Przypisanie dni w Centrum Symulacji Medycznych (Chirurgia i Interna)',
      description:
        'W kolumnach J (Chirurgia CSM) i Z (Interna CSM) wpisano daty zamiast nazw grup. Ponieważ bloki trwają po 2 tygodnie, z samej siatki (poza gr. 10a w Z8) nie wynika, które podgrupy z danego wiersza idą do CSM w pierwszym, a które w drugim tygodniu bloku.',
      sourceCells: 'PLAN ZAJĘĆ!J8:J21 oraz Z8:Z21',
      resolution:
        'Wydarzenia CSM są przypisywane do grup z danego wiersza z ostrzeżeniem o konieczności weryfikacji tygodnia w wewnętrznym grafiku CSM.',
    },
    {
      id: 'AMB-REHAB-NOV6',
      title: 'Nietypowa data 06.11. w pierwszym wierszu Podstaw Rehabilitacji (AT7)',
      description:
        'W wierszu 7 (tydzień 05.10–09.10.2026) w komórce AT7 wpisano dla grupy 10b tekst: "10b / 06.11. - 8.00 - 14.00 / sala 101 ZPK".',
      sourceCells: 'PLAN ZAJĘĆ!AT7',
      resolution:
        'Parser odczytuje rzeczywistą datę wpisaną w komórce (06.11.2026) i umieszcza zajęcia we właściwym dniu kalendarzowym z adnotacją.',
    },
  ];

  return {
    events,
    availableSub4Groups,
    semesterStart,
    semesterEnd,
    elearningNotes,
    ambiguities,
  };
}