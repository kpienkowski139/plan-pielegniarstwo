// src/utils/timeAndStats.ts
import { ScheduleEvent } from '../types';

// Jawne i spójne progi używane w całej aplikacji (Porównywarka + Wskaźnik ciężkości dnia)
export const SCHEDULE_THRESHOLDS = {
  EARLY_MORNING_HOUR: 8, // przed 08:00 (czyli np. 07:30 u Prof. Muchy / Dr. Czyżewskiego)
  EARLY_START_STANDARD: '08:00', // wczesny start <= 08:00
  LONG_DAY_HOURS: 8, // łączny czas zajęć w ciągu dnia >= 8h zegarowych
  LONG_BREAK_MINUTES: 60, // okienko między zajęciami > 60 minut
};

/** Zwraca aktualną datę YYYY-MM-DD oraz czas HH:MM w strefie Europe/Warsaw */
export function getWarsawNow(now: Date = new Date()): {
  dateIso: string;
  timeHm: string;
  timestampMs: number;
} {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Europe/Warsaw',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false,
  }).formatToParts(now);

  const map: Record<string, string> = {};
  for (const p of parts) map[p.type] = p.value;

  const h = map.hour === '24' ? '00' : map.hour;
  const dateIso = `${map.year}-${map.month}-${map.day}`;
  const timeHm = `${h}:${map.minute}`;
  return { dateIso, timeHm, timestampMs: now.getTime() };
}

export function hmToMinutes(hm: string): number {
  const [h, m] = hm.split(':').map(Number);
  return h * 60 + m;
}

export type EventTimeStatus = 'past' | 'ongoing' | 'upcoming';

export function getEventStatus(
  ev: ScheduleEvent,
  warsawDate: string,
  warsawTimeHm: string
): EventTimeStatus {
  if (ev.date < warsawDate) return 'past';
  if (ev.date > warsawDate) return 'upcoming';

  // Ten sam dzień
  if (!ev.startTime || !ev.endTime) {
    // Jeśli brak godzin (np. POZ praktyczne), w danym dniu traktujemy jako trwające/dzisiejsze do godz. 15:00
    return hmToMinutes(warsawTimeHm) < 15 * 60 ? 'ongoing' : 'past';
  }
  const nowMin = hmToMinutes(warsawTimeHm);
  const startMin = hmToMinutes(ev.startTime);
  const endMin = hmToMinutes(ev.endTime);

  if (nowMin >= endMin) return 'past';
  if (nowMin >= startMin && nowMin < endMin) return 'ongoing';
  return 'upcoming';
}

/** Formatuje czas pozostały do rozpoczęcia zajęć */
export function formatTimeUntil(
  targetDateIso: string,
  targetStartHm: string | null,
  warsawDateIso: string,
  warsawTimeHm: string
): string | null {
  if (!targetStartHm) return null;
  const [y1, m1, d1] = warsawDateIso.split('-').map(Number);
  const [y2, m2, d2] = targetDateIso.split('-').map(Number);
  const daysDiff = Math.round(
    (Date.UTC(y2, m2 - 1, d2) - Date.UTC(y1, m1 - 1, d1)) / (1000 * 60 * 60 * 24)
  );
  const totalMinutes =
    daysDiff * 24 * 60 + (hmToMinutes(targetStartHm) - hmToMinutes(warsawTimeHm));
  if (totalMinutes <= 0) return null;

  if (totalMinutes < 60) {
    return `za ${totalMinutes} min`;
  }
  const hours = Math.floor(totalMinutes / 60);
  const mins = totalMinutes % 60;
  if (daysDiff === 0) {
    return mins > 0 ? `za ${hours} godz. ${mins} min` : `za ${hours} godz.`;
  }
  if (daysDiff === 1 && hours < 24) {
    return `jutro (za ${hours} godz. ${mins} min)`;
  }
  return `za ${daysDiff} dni (${targetDateIso})`;
}

/** Oblicza postęp semestru na podstawie dat początkowej i końcowej */
export function calculateSemesterProgress(
  semesterStart: string | null,
  semesterEnd: string | null,
  currentDateIso: string
): {
  canCalculate: boolean;
  percent: number;
  daysRemaining: number;
  statusText: string;
  humorComment: string;
} {
  if (!semesterStart || !semesterEnd) {
    return {
      canCalculate: false,
      percent: 0,
      daysRemaining: 0,
      statusText: 'Brak jednoznacznych dat granic semestru w pliku.',
      humorComment: 'Nie można obliczyć postępu semestru bez dat granicznych.',
    };
  }

  const [sy, sm, sd] = semesterStart.split('-').map(Number);
  const [ey, em, ed] = semesterEnd.split('-').map(Number);
  const [cy, cm, cd] = currentDateIso.split('-').map(Number);

  const startMs = Date.UTC(sy, sm - 1, sd);
  const endMs = Date.UTC(ey, em - 1, ed);
  const currMs = Date.UTC(cy, cm - 1, cd);

  if (endMs <= startMs) {
    return {
      canCalculate: false,
      percent: 0,
      daysRemaining: 0,
      statusText: 'Nieprawidłowy zakres dat semestru.',
      humorComment: '',
    };
  }

  if (currMs < startMs) {
    const daysToStart = Math.ceil((startMs - currMs) / (1000 * 60 * 60 * 24));
    return {
      canCalculate: true,
      percent: 0,
      daysRemaining: Math.ceil((endMs - startMs) / (1000 * 60 * 60 * 24)),
      statusText: `Przed rozpoczęciem semestru (start za ${daysToStart} dni)`,
      humorComment: 'Semestr w 0%. Ciesz się wolnością, póki stetoskop leży w szufladzie!',
    };
  }

  if (currMs > endMs) {
    return {
      canCalculate: true,
      percent: 100,
      daysRemaining: 0,
      statusText: 'Semestr zakończony!',
      humorComment: '100% za Tobą! Przetrwałeś dyżury, czas na zasłużony odpoczynek.',
    };
  }

  const rawPercent = ((currMs - startMs) / (endMs - startMs)) * 100;
  const percent = Math.min(100, Math.max(0, Math.round(rawPercent)));
  const daysRemaining = Math.max(0, Math.ceil((endMs - currMs) / (1000 * 60 * 60 * 24)));

  let humorComment = '';
  if (percent < 20) {
    humorComment = `Semestr za Tobą w ${percent}%. Dopiero się rozkręca — kawa jeszcze działa!`;
  } else if (percent < 45) {
    humorComment = `Semestr za Tobą w ${percent}%. Znasz już rozkład autobusów pod każdy szpital w Warszawie.`;
  } else if (percent < 60) {
    humorComment = `Połowa za Tobą (${percent}%). Widać już światełko w tunelu (i to nie jest karetka)!`;
  } else if (percent < 85) {
    humorComment = `Już ${percent}% za Tobą! Jeszcze kilka dyżurów i wychodzimy na prostą.`;
  } else {
    humorComment = `Semestr prawie za nami (${percent}%). Trzymaj się, meta już tuż-tuż!`;
  }

  return {
    canCalculate: true,
    percent,
    daysRemaining,
    statusText: `Pozostało ${daysRemaining} dni do końca zajęć`,
    humorComment,
  };
}

/** Wskaźnik ciężkości dnia ("Spokojnie" | "Trzymaj się" | "Tryb przetrwania") */
export function evaluateDayDifficulty(dayEvents: ScheduleEvent[]): {
  level: 'Spokojnie' | 'Trzymaj się' | 'Tryb przetrwania';
  badgeColor: 'green' | 'amber' | 'red';
  totalHours: number;
  blocksCount: number;
  longBreaksCount: number;
  earliestStart: string | null;
  latestEnd: string | null;
  message: string;
} {
  // Nie liczymy wydarzeń wielokrotnie ani opcjonalnych duplikatów CSM o tej samej porze
  const uniqueEvents = deduplicateSameSlotEvents(dayEvents);

  if (uniqueEvents.length === 0) {
    return {
      level: 'Spokojnie',
      badgeColor: 'green',
      totalHours: 0,
      blocksCount: 0,
      longBreaksCount: 0,
      earliestStart: null,
      latestEnd: null,
      message: 'Dzień bez zajęć w planie! Regeneruj siły.',
    };
  }

  const timed = uniqueEvents
    .filter((e) => e.startTime && e.endTime)
    .sort((a, b) => hmToMinutes(a.startTime!) - hmToMinutes(b.startTime!));

  const totalHours = uniqueEvents.reduce((acc, e) => acc + (e.durationHours ?? 0), 0);
  const blocksCount = uniqueEvents.length;
  const earliestStart = timed.length > 0 ? timed[0].startTime : null;
  const latestEnd = timed.length > 0 ? timed[timed.length - 1].endTime : null;

  let longBreaksCount = 0;
  for (let i = 1; i < timed.length; i++) {
    const prevEnd = hmToMinutes(timed[i - 1].endTime!);
    const currStart = hmToMinutes(timed[i].startTime!);
    if (currStart - prevEnd > SCHEDULE_THRESHOLDS.LONG_BREAK_MINUTES) {
      longBreaksCount++;
    }
  }

  const startsVeryEarly = earliestStart ? hmToMinutes(earliestStart) < 8 * 60 : false;
  const endsLate = latestEnd ? hmToMinutes(latestEnd) >= 19 * 60 : false;

  if (totalHours >= 8.5 || (totalHours >= 6 && endsLate) || (startsVeryEarly && blocksCount >= 2)) {
    return {
      level: 'Tryb przetrwania',
      badgeColor: 'red',
      totalHours: Math.round(totalHours * 100) / 100,
      blocksCount,
      longBreaksCount,
      earliestStart,
      latestEnd,
      message: 'Weź termos z kawą i kanapki — dzisiaj pełny maraton od rana do wieczora!',
    };
  }

  if (totalHours >= 5 || blocksCount >= 2 || startsVeryEarly || longBreaksCount >= 1) {
    return {
      level: 'Trzymaj się',
      badgeColor: 'amber',
      totalHours: Math.round(totalHours * 100) / 100,
      blocksCount,
      longBreaksCount,
      earliestStart,
      latestEnd,
      message: 'Solidny dzień dyżurowy, ale wieczór jeszcze należy do Ciebie.',
    };
  }

  return {
    level: 'Spokojnie',
    badgeColor: 'green',
    totalHours: Math.round(totalHours * 100) / 100,
    blocksCount,
    longBreaksCount,
    earliestStart,
    latestEnd,
    message: 'Lekki rozkład godzin — idealny dzień na spokojną kawę.',
  };
}

/** Zapobiega podwójnemu liczeniu godzin, gdy w tym samym slocie 8:00-14:00 wisi klinika oraz adnotacja CSM */
export function deduplicateSameSlotEvents(events: ScheduleEvent[]): ScheduleEvent[] {
  const result: ScheduleEvent[] = [];
  for (const ev of events) {
    if (ev.completeness === 'ambiguous' && ev.subject.includes('Centrum Symulacji Medycznych')) {
      const hasMainClinicSameSlot = events.some(
        (other) =>
          other.id !== ev.id &&
          other.date === ev.date &&
          other.startTime === ev.startTime &&
          other.endTime === ev.endTime
      );
      if (hasMainClinicSameSlot) continue;
    }
    result.push(ev);
  }
  return result;
}

export interface GroupComparisonMetrics {
  groupId: string;
  totalKnownHours: number;
  unknownHoursEventsCount: number;
  daysWithClasses: number;
  earlyMorningDaysCount: number; // start <= 08:00 (w tym 07:30)
  veryEarly730DaysCount: number; // start < 08:00 (07:30)
  longDaysCount: number;         // >= 8h zajęć w dniu
  longBreaksCount: number;       // przerwy > 60 min
}

export function computeGroupComparisonMetrics(
  groupId: string,
  events: ScheduleEvent[]
): GroupComparisonMetrics {
  const cleanEvents = deduplicateSameSlotEvents(events);
  const byDate: Record<string, ScheduleEvent[]> = {};
  let totalKnownHours = 0;
  let unknownHoursEventsCount = 0;

  for (const ev of cleanEvents) {
    if (!byDate[ev.date]) byDate[ev.date] = [];
    byDate[ev.date].push(ev);
    if (ev.durationHours != null) {
      totalKnownHours += ev.durationHours;
    } else {
      unknownHoursEventsCount++;
    }
  }

  let earlyMorningDaysCount = 0;
  let veryEarly730DaysCount = 0;
  let longDaysCount = 0;
  let longBreaksCount = 0;

  for (const date of Object.keys(byDate)) {
    const dayEvs = byDate[date];
    const timed = dayEvs
      .filter((e) => e.startTime && e.endTime)
      .sort((a, b) => hmToMinutes(a.startTime!) - hmToMinutes(b.startTime!));

    const dayHours = dayEvs.reduce((acc, e) => acc + (e.durationHours ?? 0), 0);
    if (dayHours >= SCHEDULE_THRESHOLDS.LONG_DAY_HOURS) {
      longDaysCount++;
    }

    if (timed.length > 0) {
      const firstMin = hmToMinutes(timed[0].startTime!);
      if (firstMin <= 8 * 60) earlyMorningDaysCount++;
      if (firstMin < 8 * 60) veryEarly730DaysCount++;
    }

    for (let i = 1; i < timed.length; i++) {
      const prevEnd = hmToMinutes(timed[i - 1].endTime!);
      const nextStart = hmToMinutes(timed[i].startTime!);
      if (nextStart - prevEnd > SCHEDULE_THRESHOLDS.LONG_BREAK_MINUTES) {
        longBreaksCount++;
      }
    }
  }

  return {
    groupId,
    totalKnownHours: Math.round(totalKnownHours * 100) / 100,
    unknownHoursEventsCount,
    daysWithClasses: Object.keys(byDate).length,
    earlyMorningDaysCount,
    veryEarly730DaysCount,
    longDaysCount,
    longBreaksCount,
  };
}