// src/App.tsx
import React, { useEffect, useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import {
  doesEventMatchStudent,
  parseGroupHierarchy,
  parseScheduleWorkbook,
} from './parser/excelParser';
import { ParsedScheduleResult, ScheduleEvent } from './types';
import {
  calculateSemesterProgress,
  computeGroupComparisonMetrics,
  deduplicateSameSlotEvents,
  evaluateDayDifficulty,
  formatTimeUntil,
  getEventStatus,
  getWarsawNow,
  hmToMinutes,
  SCHEDULE_THRESHOLDS,
} from './utils/timeAndStats';
import './App.css';

const STORAGE_KEYS = {
  GROUP: 'wum_nursing_selected_group_v1',
  HALF12: 'wum_nursing_selected_12half_v1',
  CACHED_SCHEDULE: 'wum_nursing_cached_schedule_v1',
  LAST_FETCH_TIME: 'wum_nursing_last_fetch_v1',
};

const POLISH_DAYS = [
  'Niedziela',
  'Poniedziałek',
  'Wtorek',
  'Środa',
  'Czwartek',
  'Piątek',
  'Sobota',
];

function formatDateReadable(isoDate: string): string {
  const [y, m, d] = isoDate.split('-').map(Number);
  const dt = new Date(Date.UTC(y, m - 1, d));
  const dayName = POLISH_DAYS[dt.getUTCDay()];
  return `${dayName}, ${String(d).padStart(2, '0')}.${String(m).padStart(2, '0')}.${y}`;
}

export default function App() {
  const [scheduleData, setScheduleData] = useState<ParsedScheduleResult | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [fetchError, setFetchError] = useState<string | null>(null);
  const [usingCachedCopy, setUsingCachedCopy] = useState<boolean>(false);
  const [lastFetchedAt, setLastFetchedAt] = useState<string | null>(
    localStorage.getItem(STORAGE_KEYS.LAST_FETCH_TIME)
  );

  // Wybrana podgrupa 4-osobowa (np. "1a1")
  const [selectedGroup, setSelectedGroup] = useState<string | null>(
    localStorage.getItem(STORAGE_KEYS.GROUP)
  );
  // Dla podgrup "c" (np. 1c1) — preferowana połowa 12-osobowa ('a' | 'b' | 'both')
  const [preferred12Half, setPreferred12Half] = useState<'a' | 'b' | 'both'>(
    (localStorage.getItem(STORAGE_KEYS.HALF12) as 'a' | 'b' | 'both') || 'both'
  );

  // Zegar czasu rzeczywistego (strefa Europe/Warsaw), odświeżany co 30 sekund bez przeładowania strony
  const [warsawClock, setWarsawClock] = useState(() => getWarsawNow());

  // Widok pełnego planu
  const [viewMode, setViewMode] = useState<'upcoming' | 'day' | 'week'>('upcoming');
  const [selectedDateFilter, setSelectedDateFilter] = useState<string>(() => getWarsawNow().dateIso);
  const [selectedWeekIndex, setSelectedWeekIndex] = useState<number>(0);
  const [typeFilter, setTypeFilter] = useState<'all' | 'praktyczne' | 'wyklad'>('all');

  // Porównywarka grup
  const [compareGroup, setCompareGroup] = useState<string>('2a1');
  const [compareScope, setCompareScope] = useState<'semester' | 'week'>('semester');

  // Małe elementy do poklikania
  const [badnessComment, setBadnessComment] = useState<string | null>(null);
  const [showSettingsModal, setShowSettingsModal] = useState<boolean>(false);

  useEffect(() => {
    const timer = setInterval(() => {
      setWarsawClock(getWarsawNow());
    }, 30_000);
    return () => clearInterval(timer);
  }, []);

  const loadExcelSchedule = async () => {
    setLoading(true);
    setFetchError(null);
    try {
      const url = `${import.meta.env.BASE_URL}plan.xlsx?t=${Date.now()}`;
      const response = await fetch(url);
      if (!response.ok) {
        throw new Error(`Błąd HTTP ${response.status} przy pobieraniu plan.xlsx`);
      }
      const buffer = await response.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const parsed = parseScheduleWorkbook(workbook);

      setScheduleData(parsed);
      setUsingCachedCopy(false);

      const nowStamp = new Intl.DateTimeFormat('pl-PL', {
        timeZone: 'Europe/Warsaw',
        dateStyle: 'short',
        timeStyle: 'medium',
      }).format(new Date());

      setLastFetchedAt(nowStamp);
      localStorage.setItem(STORAGE_KEYS.LAST_FETCH_TIME, nowStamp);
      localStorage.setItem(STORAGE_KEYS.CACHED_SCHEDULE, JSON.stringify(parsed));
    } catch (err: any) {
      const cached = localStorage.getItem(STORAGE_KEYS.CACHED_SCHEDULE);
      if (cached) {
        try {
          setScheduleData(JSON.parse(cached));
          setUsingCachedCopy(true);
          setFetchError(
            'Nie udało się pobrać najnowszej wersji pliku plan.xlsx z serwera. Wyświetlam zapisaną wcześniej kopię z pamięci przeglądarki.'
          );
        } catch {
          setFetchError(
            'Nie udało się pobrać pliku plan.xlsx, a zapisana kopia jest uszkodzona.'
          );
        }
      } else {
        setFetchError(
          `Nie można pobrać pliku plan.xlsx (${err?.message || 'Błąd sieci'}). Sprawdź połączenie z internetem i spróbuj ponownie.`
        );
      }
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadExcelSchedule();
  }, []);

  const handleSelectGroup = (grp: string) => {
    setSelectedGroup(grp);
    localStorage.setItem(STORAGE_KEYS.GROUP, grp);
    setShowSettingsModal(false);
  };

  const handleSelect12Half = (half: 'a' | 'b' | 'both') => {
    setPreferred12Half(half);
    localStorage.setItem(STORAGE_KEYS.HALF12, half);
  };

  // Wydarzenia przefiltrowane dla wybranej grupy studenta, posortowane chronologicznie
  const studentEvents = useMemo(() => {
    if (!scheduleData || !selectedGroup) return [];
    return scheduleData.events
      .filter((ev) => doesEventMatchStudent(ev, selectedGroup, preferred12Half))
      .sort((a, b) => {
        if (a.date !== b.date) return a.date.localeCompare(b.date);
        const aStart = a.startTime ? hmToMinutes(a.startTime) : 480;
        const bStart = b.startTime ? hmToMinutes(b.startTime) : 480;
        return aStart - bStart;
      });
  }, [scheduleData, selectedGroup, preferred12Half]);

  // Wydarzenia bez duplikatów slotów CSM (do liczników godzin i statystyk)
  const deduplicatedStudentEvents = useMemo(
    () => deduplicateSameSlotEvents(studentEvents),
    [studentEvents]
  );

  // Obliczenia dla kart "TERAZ" i "NASTĘPNE ZAJĘCIA"
  const { nowEvent, nextEvent, todayRemainingCount } = useMemo(() => {
    const ongoing: ScheduleEvent[] = [];
    const upcoming: ScheduleEvent[] = [];

    for (const ev of deduplicatedStudentEvents) {
      const st = getEventStatus(ev, warsawClock.dateIso, warsawClock.timeHm);
      if (st === 'ongoing') ongoing.push(ev);
      else if (st === 'upcoming') upcoming.push(ev);
    }

    const todayRemaining = upcoming.filter((e) => e.date === warsawClock.dateIso).length;

    return {
      nowEvent: ongoing.length > 0 ? ongoing[0] : null,
      nextEvent: upcoming.length > 0 ? upcoming[0] : null,
      todayRemainingCount: todayRemaining,
    };
  }, [deduplicatedStudentEvents, warsawClock]);

  // Lista tygodni do przeglądarki tygodniowej
  const weeksList = useMemo(() => {
    if (!scheduleData) return [];
    const uniqueMondays = Array.from(
      new Set(
        scheduleData.events.map((e) => {
          const [y, m, d] = e.date.split('-').map(Number);
          const dt = new Date(Date.UTC(y, m - 1, d));
          const day = dt.getUTCDay() || 7;
          dt.setUTCDate(dt.getUTCDate() - (day - 1));
          return dt.toISOString().slice(0, 10);
        })
      )
    ).sort();
    return uniqueMondays;
  }, [scheduleData]);

  // Statystyki licznika zajęć
  const stats = useMemo(() => {
    let pastCount = 0;
    let upcomingCount = 0;
    let remainingHours = 0;
    let totalKnownHours = 0;
    let lecturesCount = 0;
    let practicalsCount = 0;

    for (const ev of deduplicatedStudentEvents) {
      const st = getEventStatus(ev, warsawClock.dateIso, warsawClock.timeHm);
      if (st === 'past') {
        pastCount++;
      } else {
        upcomingCount++;
        if (ev.durationHours != null) {
          remainingHours += ev.durationHours;
        }
      }
      if (ev.durationHours != null) {
        totalKnownHours += ev.durationHours;
      }
      if (ev.type === 'wyklad') lecturesCount++;
      else practicalsCount++;
    }

    return {
      totalCount: deduplicatedStudentEvents.length,
      pastCount,
      upcomingCount,
      remainingHours: Math.round(remainingHours * 100) / 100,
      totalKnownHours: Math.round(totalKnownHours * 100) / 100,
      lecturesCount,
      practicalsCount,
    };
  }, [deduplicatedStudentEvents, warsawClock]);

  // Wskaźnik ciężkości wybranego dnia (domyślnie dzisiaj lub dzień następnych zajęć)
  const targetDayForDifficulty = useMemo(() => {
    const todayEvs = deduplicatedStudentEvents.filter((e) => e.date === warsawClock.dateIso);
    if (todayEvs.length > 0) return warsawClock.dateIso;
    if (nextEvent) return nextEvent.date;
    return warsawClock.dateIso;
  }, [deduplicatedStudentEvents, warsawClock.dateIso, nextEvent]);

  const dayDifficulty = useMemo(() => {
    const evs = deduplicatedStudentEvents.filter((e) => e.date === targetDayForDifficulty);
    return evaluateDayDifficulty(evs);
  }, [deduplicatedStudentEvents, targetDayForDifficulty]);

  // Ciekawostki do poklikania
  const funFacts = useMemo(() => {
    const byDate: Record<string, ScheduleEvent[]> = {};
    for (const ev of deduplicatedStudentEvents) {
      if (!byDate[ev.date]) byDate[ev.date] = [];
      byDate[ev.date].push(ev);
    }

    let hardestDayDate: string | null = null;
    let hardestDayHours = 0;
    let longestSpanDate: string | null = null;
    let longestSpanMinutes = 0;
    let earliestEvent: ScheduleEvent | null = null;

    for (const [dt, evs] of Object.entries(byDate)) {
      const hrs = evs.reduce((acc, e) => acc + (e.durationHours ?? 0), 0);
      if (hrs > hardestDayHours) {
        hardestDayHours = hrs;
        hardestDayDate = dt;
      }

      const timed = evs
        .filter((e) => e.startTime && e.endTime)
        .sort((a, b) => hmToMinutes(a.startTime!) - hmToMinutes(b.startTime!));

      if (timed.length > 0) {
        const span =
          hmToMinutes(timed[timed.length - 1].endTime!) - hmToMinutes(timed[0].startTime!);
        if (span > longestSpanMinutes) {
          longestSpanMinutes = span;
          longestSpanDate = dt;
        }
        if (
          !earliestEvent ||
          hmToMinutes(timed[0].startTime!) < hmToMinutes(earliestEvent.startTime!)
        ) {
          earliestEvent = timed[0];
        }
      }
    }

    // Bieżący tydzień
    const [cy, cm, cd] = warsawClock.dateIso.split('-').map(Number);
    const currDt = new Date(Date.UTC(cy, cm - 1, cd));
    const dayOfWeek = currDt.getUTCDay() || 7;
    currDt.setUTCDate(currDt.getUTCDate() - (dayOfWeek - 1));
    const weekStartIso = currDt.toISOString().slice(0, 10);
    currDt.setUTCDate(currDt.getUTCDate() + 6);
    const weekEndIso = currDt.toISOString().slice(0, 10);

    const currentWeekEvents = deduplicatedStudentEvents.filter(
      (e) => e.date >= weekStartIso && e.date <= weekEndIso
    );
    const currentWeekHours = currentWeekEvents.reduce(
      (acc, e) => acc + (e.durationHours ?? 0),
      0
    );

    return {
      hardestDayDate,
      hardestDayHours: Math.round(hardestDayHours * 100) / 100,
      longestSpanDate,
      longestSpanHours: Math.round((longestSpanMinutes / 60) * 10) / 10,
      earliestEvent,
      currentWeekCount: currentWeekEvents.length,
      currentWeekHours: Math.round(currentWeekHours * 100) / 100,
    };
  }, [deduplicatedStudentEvents, warsawClock.dateIso]);

  // Porównywarka grup
  const comparisonResult = useMemo(() => {
    if (!scheduleData || !selectedGroup || !compareGroup) return null;
    const filterByScope = (evs: ScheduleEvent[]) => {
      if (compareScope === 'semester') return evs;
      const wkMonday = weeksList[selectedWeekIndex];
      if (!wkMonday) return evs;
      const [y, m, d] = wkMonday.split('-').map(Number);
      const endDt = new Date(Date.UTC(y, m - 1, d + 6)).toISOString().slice(0, 10);
      return evs.filter((e) => e.date >= wkMonday && e.date <= endDt);
    };

    const myEvs = filterByScope(
      scheduleData.events.filter((e) => doesEventMatchStudent(e, selectedGroup, preferred12Half))
    );
    const otherEvs = filterByScope(
      scheduleData.events.filter((e) => doesEventMatchStudent(e, compareGroup, 'both'))
    );

    const m1 = computeGroupComparisonMetrics(selectedGroup, myEvs);
    const m2 = computeGroupComparisonMetrics(compareGroup, otherEvs);

    let verdict = '';
    const hasMissingHoursDiff = m1.unknownHoursEventsCount !== m2.unknownHoursEventsCount;

    if (hasMissingHoursDiff) {
      verdict =
        'Uwaga: Jedna z grup ma w wybranym okresie praktyki POZ bez podanych godzin zegarowych w pliku Excel. Nie przyznajemy zwycięstwa w godzinach — brak pełnych danych do rzetelnego werdyktu!';
    } else if (Math.abs(m1.totalKnownHours - m2.totalKnownHours) >= 3) {
      const diff = Math.round(Math.abs(m1.totalKnownHours - m2.totalKnownHours) * 10) / 10;
      if (m1.totalKnownHours > m2.totalKnownHours) {
        verdict = `Twoja grupa (${selectedGroup}) ma o ${diff} godz. więcej zajęć. Składamy szczere kondolencje!`;
      } else {
        verdict = `Grupa ${compareGroup} ma o ${diff} godz. więcej zajęć. Tym razem wygrywasz życie!`;
      }
    } else if (m1.veryEarly730DaysCount !== m2.veryEarly730DaysCount) {
      if (m1.veryEarly730DaysCount > m2.veryEarly730DaysCount) {
        verdict = `Liczba godzin podobna, ale Twoja grupa częściej wstaje na 7:30 (${m1.veryEarly730DaysCount} vs ${m2.veryEarly730DaysCount} dni). Budzik płacze.`;
      } else {
        verdict = `Liczba godzin podobna, ale grupa ${compareGroup} ma więcej poranków o 7:30 (${m2.veryEarly730DaysCount} vs ${m1.veryEarly730DaysCount} dni). Śpisz dłużej!`;
      }
    } else if (m1.longBreaksCount !== m2.longBreaksCount) {
      if (m1.longBreaksCount > m2.longBreaksCount) {
        verdict = `Godziny identyczne, ale Twoja grupa ma więcej długich okienek (${m1.longBreaksCount} vs ${m2.longBreaksCount}). Więcej czasu na kawę!`;
      } else {
        verdict = `Druga grupa (${compareGroup}) ma więcej okienek (${m2.longBreaksCount} vs ${m1.longBreaksCount}). Więcej czasu na kawę na Trojdena.`;
      }
    } else {
      verdict = 'Wasze plany są niemal bliźniacze. Cierpicie solidarnie w tym samym wymiarze!';
    }

    return { m1, m2, verdict, hasMissingHoursDiff };
  }, [
    scheduleData,
    selectedGroup,
    compareGroup,
    compareScope,
    weeksList,
    selectedWeekIndex,
    preferred12Half,
  ]);

  const semesterProgress = useMemo(() => {
    return calculateSemesterProgress(
      scheduleData?.semesterStart ?? null,
      scheduleData?.semesterEnd ?? null,
      warsawClock.dateIso
    );
  }, [scheduleData, warsawClock.dateIso]);

  const generateBadnessComment = () => {
    const comments = [
      `W całym semestrze czekają Cię ${stats.totalCount} wydarzenia dydaktyczne (${stats.totalKnownHours} znanych godzin zegarowych + praktyki POZ). Dasz radę!`,
      funFacts.earliestEvent
        ? `Najwcześniejsza pobudka: zajęcia od ${funFacts.earliestEvent.startTime} (${funFacts.earliestEvent.subject}). Nie zapomnij nastawić dwóch budzików!`
        : 'Wszystkie zajęcia zaczynają się o ludzkiej porze!',
      `Pozostało jeszcze ${stats.upcomingCount} wydarzeń do końca semestru. Każdy odhaczony dyżur przybliża Cię do wakacji!`,
      `W tym tygodniu masz zaplanowanych ${funFacts.currentWeekCount} zajęć (${funFacts.currentWeekHours} godz. zegarowych).`,
    ];
    const randomIdx = Math.floor(Math.random() * comments.length);
    setBadnessComment(comments[randomIdx]);
  };

  // Ekran ładowania
  if (loading && !scheduleData) {
    return (
      <div className="app-container center-screen" role="status" aria-live="polite">
        <div className="card">
          <h1>Wczytywanie planu zajęć…</h1>
          <p>Pobieranie i analiza pliku <code>plan.xlsx</code> dla II roku Pielęgniarstwa WUM.</p>
        </div>
      </div>
    );
  }

  // Krytyczny błąd braku danych
  if (!scheduleData) {
    return (
      <div className="app-container center-screen" role="alert">
        <div className="card error-card">
          <h1>Błąd wczytywania planu</h1>
          <p>{fetchError || 'Nie udało się odczytać pliku plan.xlsx.'}</p>
          <button className="btn-primary" onClick={loadExcelSchedule}>
            Spróbuj ponownie wczytać plan
          </button>
        </div>
      </div>
    );
  }

  // Ekran pierwszego uruchomienia: "Do której grupy należysz?"
  if (!selectedGroup || showSettingsModal) {
    const groupedByDean: Record<number, string[]> = {};
    for (const g of scheduleData.availableSub4Groups) {
      const h = parseGroupHierarchy(g);
      if (!groupedByDean[h.deanGroup]) groupedByDean[h.deanGroup] = [];
      groupedByDean[h.deanGroup].push(g);
    }

    const currentHierarchy = selectedGroup ? parseGroupHierarchy(selectedGroup) : null;

    return (
      <div className="app-container">
        <div className="card onboarding-card">
          <h1>Do której grupy należysz?</h1>
          <p className="subtitle">
            Wybierz swoją podgrupę 4-osobową (np. <strong>1a1</strong>). Aplikacja automatycznie
            dopasuje Twoje zajęcia 4-osobowe (Pediatria, POZ), 8-osobowe (np. <strong>1a</strong> —
            Chirurgia, Interna, Promocja Zdrowia), grupę dziekańską (np. <strong>grupa 1</strong> —
            seminaria) oraz wykłady dla całego roku.
          </p>

          {currentHierarchy && currentHierarchy.letter8 === 'c' && (
            <div className="warning-box" role="region" aria-label="Ustawienie grupy 12-osobowej">
              <strong>Należysz do podgrupy „{currentHierarchy.sub8}” (człon C):</strong>
              <p>
                Na ćwiczeniach 12-osobowych (<em>Podstawy Rehabilitacji</em> oraz <em>POZ ćw.</em>)
                grupa dziekańska dzieli się tylko na dwie połowy: <strong>{currentHierarchy.deanGroup}a</strong> i{' '}
                <strong>{currentHierarchy.deanGroup}b</strong>. Wybierz, do której połowy 12-osobowej
                zostałeś dopisany:
              </p>
              <div className="button-row">
                <button
                  className={preferred12Half === 'both' ? 'btn-primary' : 'btn-secondary'}
                  onClick={() => handleSelect12Half('both')}
                >
                  Pokaż obie połowy ({currentHierarchy.deanGroup}a i {currentHierarchy.deanGroup}b)
                </button>
                <button
                  className={preferred12Half === 'a' ? 'btn-primary' : 'btn-secondary'}
                  onClick={() => handleSelect12Half('a')}
                >
                  Połowa {currentHierarchy.deanGroup}a (12-os.)
                </button>
                <button
                  className={preferred12Half === 'b' ? 'btn-primary' : 'btn-secondary'}
                  onClick={() => handleSelect12Half('b')}
                >
                  Połowa {currentHierarchy.deanGroup}b (12-os.)
                </button>
              </div>
            </div>
          )}

          <div className="dean-groups-grid">
            {Object.entries(groupedByDean).map(([deanNum, subGroups]) => (
              <div key={deanNum} className="dean-group-box">
                <h2>Grupa dziekańska {deanNum}</h2>
                <div className="subgroups-buttons">
                  {subGroups.map((sg) => (
                    <button
                      key={sg}
                      className={selectedGroup === sg ? 'btn-group active' : 'btn-group'}
                      onClick={() => handleSelectGroup(sg)}
                    >
                      {sg}
                    </button>
                  ))}
                </div>
              </div>
            ))}
          </div>

          {selectedGroup && (
            <div className="modal-footer">
              <button className="btn-primary" onClick={() => setShowSettingsModal(false)}>
                Gotowe — wróć do planu ({selectedGroup})
              </button>
            </div>
          )}
        </div>
      </div>
    );
  }

  const hierarchy = parseGroupHierarchy(selectedGroup);

  // Filtrowanie listy zajęć do sekcji "Pełny plan zajęć"
  const displayedEvents = studentEvents.filter((ev) => {
    if (typeFilter === 'wyklad' && ev.type !== 'wyklad') return false;
    if (typeFilter === 'praktyczne' && ev.type === 'wyklad') return false;

    if (viewMode === 'upcoming') {
      return getEventStatus(ev, warsawClock.dateIso, warsawClock.timeHm) !== 'past';
    }
    if (viewMode === 'day') {
      return ev.date === selectedDateFilter;
    }
    if (viewMode === 'week') {
      const wkMonday = weeksList[selectedWeekIndex];
      if (!wkMonday) return true;
      const [y, m, d] = wkMonday.split('-').map(Number);
      const wkSunday = new Date(Date.UTC(y, m - 1, d + 6)).toISOString().slice(0, 10);
      return ev.date >= wkMonday && ev.date <= wkSunday;
    }
    return true;
  });

  return (
    <div className="app-container">
      {/* Nagłówek aplikacji */}
      <header className="app-header">
        <div className="header-left">
          <span className="app-badge">WUM • II rok Pielęgniarstwo</span>
          <h1>Twój Plan Zajęć</h1>
          <div className="sync-status">
            <span>
              Strefa: <strong>Europe/Warsaw ({warsawClock.timeHm})</strong>
            </span>
            {lastFetchedAt && (
              <span>
                • Ostatnio pobrano: <strong>{lastFetchedAt}</strong>
                {usingCachedCopy ? ' (kopia offline)' : ' (aktualny plik)'}
              </span>
            )}
          </div>
        </div>

        <div className="header-right">
          <div className="current-group-pill" aria-label="Aktualnie wybrana grupa">
            Grupa: <strong>{hierarchy.sub4}</strong>{' '}
            <span className="pill-sub">
              ({hierarchy.sub8} / gr. {hierarchy.deanGroup})
            </span>
          </div>
          <button
            className="btn-secondary"
            onClick={() => setShowSettingsModal(true)}
            aria-label="Zmień grupę w ustawieniach"
          >
            Zmień grupę
          </button>
          <button
            className="btn-secondary"
            onClick={loadExcelSchedule}
            aria-label="Odśwież plik Excel"
          >
            Odśwież plan
          </button>
        </div>
      </header>

      {/* Komunikaty błędów odświeżania / trybu offline */}
      {fetchError && (
        <div className="alert-banner" role="alert">
          <span>⚠️ {fetchError}</span>
          <button className="btn-small" onClick={loadExcelSchedule}>
            Ponów pobieranie
          </button>
        </div>
      )}

      {/* GŁÓWNY EKRAN: Karty "TERAZ" i "NASTĘPNE ZAJĘCIA" */}
      <section className="hero-cards" aria-label="Status bieżących i najbliższych zajęć">
        {/* KARTA "TERAZ" */}
        <div className={`hero-card ${nowEvent ? 'card-now-active' : 'card-now-idle'}`}>
          <div className="card-header-row">
            <span className="status-tag tag-now">TERAZ</span>
            <span className="status-subtext">
              {nowEvent ? 'Zajęcia aktualnie trwają' : 'Brak trwających zajęć'}
            </span>
          </div>

          {nowEvent ? (
            <div className="hero-content">
              <h2>{nowEvent.subject}</h2>
              <p className="hero-time">
                📅 {formatDateReadable(nowEvent.date)} • ⏰{' '}
                <strong>
                  {nowEvent.startTime && nowEvent.endTime
                    ? `${nowEvent.startTime} – ${nowEvent.endTime}`
                    : 'Brak dokładnych godzin w planie'}
                </strong>
              </p>
              <div className="location-highlight">
                <p>
                  🏥 <strong>Placówka:</strong>{' '}
                  {nowEvent.facilityName || 'Brak nazwy placówki w planie'}
                </p>
                <p>
                  📍 <strong>Adres:</strong> {nowEvent.address || 'Brak adresu w planie'}
                </p>
                {nowEvent.building && (
                  <p>
                    🏢 <strong>Budynek:</strong> {nowEvent.building}
                  </p>
                )}
                <p>
                  🚪 <strong>Sala:</strong> {nowEvent.room || 'Brak numeru sali w planie'}
                </p>
                {nowEvent.instructor && (
                  <p>
                    👩‍⚕️ <strong>Prowadzący:</strong> {nowEvent.instructor}
                  </p>
                )}
              </div>
              {nowEvent.notes.map((n, i) => (
                <p key={i} className="event-note">
                  ℹ️ {n}
                </p>
              ))}
            </div>
          ) : (
            <div className="hero-empty">
              {todayRemainingCount === 0 ? (
                <>
                  <h2>Na dziś koniec. Możesz odpocząć! ☕</h2>
                  <p>W dniu dzisiejszym ({formatDateReadable(warsawClock.dateIso)}) nie masz już żadnych zajęć.</p>
                </>
              ) : (
                <>
                  <h2>Przerwa przed kolejnymi zajęciami</h2>
                  <p>Aktualnie żadne zajęcia nie trwają, ale dzisiaj czekają Cię jeszcze {todayRemainingCount} zajęcia.</p>
                </>
              )}
            </div>
          )}
        </div>

        {/* KARTA "NASTĘPNE ZAJĘCIA" */}
        <div className="hero-card card-next">
          <div className="card-header-row">
            <span className="status-tag tag-next">NASTĘPNE ZAJĘCIA</span>
            {nextEvent && (
              <span className="countdown-badge">
                {formatTimeUntil(
                  nextEvent.date,
                  nextEvent.startTime,
                  warsawClock.dateIso,
                  warsawClock.timeHm
                ) || `Dzień: ${nextEvent.date}`}
              </span>
            )}
          </div>

          {nextEvent ? (
            <div className="hero-content">
              <h2>{nextEvent.subject}</h2>
              <p className="hero-time">
                📅 {formatDateReadable(nextEvent.date)} • ⏰{' '}
                <strong>
                  {nextEvent.startTime && nextEvent.endTime
                    ? `${nextEvent.startTime} – ${nextEvent.endTime}`
                    : 'Godziny nieokreślone w planie'}
                </strong>
              </p>
              <div className="location-highlight">
                <p>
                  🏥 <strong>Placówka:</strong>{' '}
                  {nextEvent.facilityName || 'Brak nazwy placówki w planie'}
                </p>
                <p>
                  📍 <strong>Adres:</strong> {nextEvent.address || 'Brak adresu w planie'}
                </p>
                {nextEvent.building && (
                  <p>
                    🏢 <strong>Budynek:</strong> {nextEvent.building}
                  </p>
                )}
                <p>
                  🚪 <strong>Sala:</strong> {nextEvent.room || 'Brak numeru sali w planie'}
                </p>
                {nextEvent.instructor && (
                  <p>
                    👩‍⚕️ <strong>Prowadzący:</strong> {nextEvent.instructor}
                  </p>
                )}
              </div>
              {nextEvent.completeness !== 'complete' && nextEvent.completenessNote && (
                <p className="completeness-warning">⚠️ {nextEvent.completenessNote}</p>
              )}
            </div>
          ) : (
            <div className="hero-empty">
              <h2>Brak przyszłych wydarzeń w planie</h2>
              <p>Wszystkie zaplanowane w pliku Excel zajęcia dla Twojej grupy już się odbyły.</p>
            </div>
          )}
        </div>
      </section>

      {/* WSKAŹNIK CIĘŻKOŚCI DNIA + POSTĘP SEMESTRU */}
      <section className="status-strip-grid">
        {/* Wskaźnik ciężkości dnia */}
        <div className="card compact-card">
          <div className="card-header-row">
            <h3>Jak ciężki będzie ten dzień? ({targetDayForDifficulty})</h3>
            <span className={`difficulty-badge diff-${dayDifficulty.badgeColor}`}>
              {dayDifficulty.level}
            </span>
          </div>
          <p className="humor-text">{dayDifficulty.message}</p>
          <div className="mini-metrics">
            <span>⏱️ Godzin: <strong>{dayDifficulty.totalHours}h</strong></span>
            <span>📦 Bloków: <strong>{dayDifficulty.blocksCount}</strong></span>
            <span>
              🌅 Start: <strong>{dayDifficulty.earliestStart || '—'}</strong> / Koniec:{' '}
              <strong>{dayDifficulty.latestEnd || '—'}</strong>
            </span>
            <span>☕ Długie przerwy (&gt;60m): <strong>{dayDifficulty.longBreaksCount}</strong></span>
          </div>
          <p className="disclaimer-tiny">
            * Umowna statystyka rozrywkowa oparta wyłącznie na rozkładzie godzin w planie.
          </p>
        </div>

        {/* Postęp semestru */}
        <div className="card compact-card">
          <div className="card-header-row">
            <h3>Postęp semestru zimowego</h3>
            {semesterProgress.canCalculate && (
              <span className="progress-pct">{semesterProgress.percent}%</span>
            )}
          </div>
          {semesterProgress.canCalculate ? (
            <>
              <div
                className="progress-bar-track"
                role="progressbar"
                aria-valuenow={semesterProgress.percent}
                aria-valuemin={0}
                aria-valuemax={100}
              >
                <div
                  className="progress-bar-fill"
                  style={{ width: `${semesterProgress.percent}%` }}
                />
              </div>
              <p className="humor-text">{semesterProgress.humorComment}</p>
              <div className="mini-metrics">
                <span>🟢 Początek: <strong>{scheduleData.semesterStart}</strong></span>
                <span>🏁 Koniec: <strong>{scheduleData.semesterEnd}</strong></span>
                <span>⏳ Pozostało: <strong>{semesterProgress.daysRemaining} dni</strong></span>
              </div>
            </>
          ) : (
            <p>{semesterProgress.statusText}</p>
          )}
        </div>
      </section>

      {/* PEŁNY PLAN ZAJĘĆ */}
      <section className="card schedule-section" aria-label="Pełny plan zajęć">
        <div className="schedule-controls">
          <h2>Pełny plan zajęć ({hierarchy.sub4})</h2>

          <div className="tabs-row" role="tablist">
            <button
              role="tab"
              aria-selected={viewMode === 'upcoming'}
              className={viewMode === 'upcoming' ? 'tab-btn active' : 'tab-btn'}
              onClick={() => setViewMode('upcoming')}
            >
              Najbliższe zajęcia
            </button>
            <button
              role="tab"
              aria-selected={viewMode === 'day'}
              className={viewMode === 'day' ? 'tab-btn active' : 'tab-btn'}
              onClick={() => setViewMode('day')}
            >
              Według dnia
            </button>
            <button
              role="tab"
              aria-selected={viewMode === 'week'}
              className={viewMode === 'week' ? 'tab-btn active' : 'tab-btn'}
              onClick={() => setViewMode('week')}
            >
              Według tygodnia
            </button>
          </div>

          <div className="subfilters-row">
            {viewMode === 'day' && (
              <label className="filter-label">
                Wybierz dzień:{' '}
                <input
                  type="date"
                  value={selectedDateFilter}
                  min={scheduleData.semesterStart || undefined}
                  max={scheduleData.semesterEnd || undefined}
                  onChange={(e) => setSelectedDateFilter(e.target.value)}
                />
              </label>
            )}

            {viewMode === 'week' && (
              <label className="filter-label">
                Wybierz tydzień:{' '}
                <select
                  value={selectedWeekIndex}
                  onChange={(e) => setSelectedWeekIndex(Number(e.target.value))}
                >
                  {weeksList.map((monday, idx) => (
                    <option key={monday} value={idx}>
                      Tydzień {idx + 1} (od {monday})
                    </option>
                  ))}
                </select>
              </label>
            )}

            <div className="type-filter-pills">
              <button
                className={typeFilter === 'all' ? 'pill-btn active' : 'pill-btn'}
                onClick={() => setTypeFilter('all')}
              >
                Wszystkie
              </button>
              <button
                className={typeFilter === 'praktyczne' ? 'pill-btn active' : 'pill-btn'}
                onClick={() => setTypeFilter('praktyczne')}
              >
                Praktyczne / Ćw. / Sem.
              </button>
              <button
                className={typeFilter === 'wyklad' ? 'pill-btn active' : 'pill-btn'}
                onClick={() => setTypeFilter('wyklad')}
              >
                Tylko wykłady
              </button>
            </div>
          </div>
        </div>

        {displayedEvents.length === 0 ? (
          <div className="empty-list-box">
            Brak zajęć spełniających wybrane kryteria w tym widoku.
          </div>
        ) : (
          <div className="events-list">
            {displayedEvents.map((ev) => {
              const status = getEventStatus(ev, warsawClock.dateIso, warsawClock.timeHm);
              return (
                <article
                  key={ev.id}
                  className={`event-card status-${status} type-${ev.type}`}
                >
                  <div className="event-card-top">
                    <div className="badges-group">
                      <span className={`type-badge badge-${ev.type}`}>
                        {ev.type === 'wyklad'
                          ? 'WYKŁAD (Wspólny)'
                          : ev.type === 'seminarium'
                          ? 'SEMINARIUM'
                          : ev.type === 'cwiczenia'
                          ? 'ĆWICZENIA'
                          : 'ZAJĘCIA PRAKTYCZNE'}
                      </span>
                      <span className={`state-badge state-${status}`}>
                        {status === 'past'
                          ? 'Zakończone'
                          : status === 'ongoing'
                          ? 'Trwa teraz'
                          : 'Przyszłe'}
                      </span>
                      {ev.is12PersonSplit && hierarchy.letter8 === 'c' && (
                        <span className="split-badge">
                          Podział 12-os. (połowa {ev.split12Half?.toUpperCase()})
                        </span>
                      )}
                    </div>
                    <span className="source-ref" title="Komórka źródłowa w pliku Excel">
                      {ev.sourceRef}
                    </span>
                  </div>

                  <h3>{ev.subject}</h3>

                  <div className="event-datetime">
                    <span>📅 <strong>{formatDateReadable(ev.date)}</strong></span>
                    <span>
                      ⏰{' '}
                      <strong>
                        {ev.startTime && ev.endTime
                          ? `${ev.startTime} – ${ev.endTime} (${ev.durationHours}h)`
                          : 'Brak godzin w planie'}
                      </strong>
                    </span>
                  </div>

                  <div className="event-location-grid">
                    <div>
                      <strong>Placówka:</strong>{' '}
                      {ev.facilityName || <em className="missing-field">Brak danych w planie</em>}
                    </div>
                    <div>
                      <strong>Adres:</strong>{' '}
                      {ev.address || <em className="missing-field">Brak adresu w planie</em>}
                    </div>
                    {ev.building && (
                      <div>
                        <strong>Budynek / Pawilon:</strong> {ev.building}
                      </div>
                    )}
                    <div>
                      <strong>Sala:</strong>{' '}
                      {ev.room || <em className="missing-field">Brak numeru sali w planie</em>}
                    </div>
                    <div>
                      <strong>Prowadzący:</strong>{' '}
                      {ev.instructor || <em className="missing-field">Brak w planie</em>}
                    </div>
                  </div>

                  {ev.notes.length > 0 && (
                    <ul className="event-notes-list">
                      {ev.notes.map((note, idx) => (
                        <li key={idx}>ℹ️ {note}</li>
                      ))}
                    </ul>
                  )}

                  {ev.completeness !== 'complete' && ev.completenessNote && (
                    <div className="event-warning-bar">⚠️ {ev.completenessNote}</div>
                  )}
                </article>
              );
            })}
          </div>
        )}
      </section>

      {/* DODATKOWE STATYSTYKI, LICZNIK ZAJĘĆ I PORÓWNYWARKA GRUP */}
      <details className="card collapsible-section" open>
        <summary className="section-summary">
          📊 Licznik zajęć i ciekawostki z Twojego planu
        </summary>
        <div className="section-body">
          <div className="stats-counters-grid">
            <div className="stat-box">
              <span className="stat-num">{stats.totalCount}</span>
              <span className="stat-label">Wszystkich wydarzeń w semestrze</span>
            </div>
            <div className="stat-box">
              <span className="stat-num">{stats.pastCount}</span>
              <span className="stat-label">Zajęć już za Tobą</span>
            </div>
            <div className="stat-box">
              <span className="stat-num">{stats.upcomingCount}</span>
              <span className="stat-label">Zajęć przed Tobą</span>
            </div>
            <div className="stat-box">
              <span className="stat-num">{stats.remainingHours}h</span>
              <span className="stat-label">Pozostałych znanych godzin zegarowych</span>
            </div>
            <div className="stat-box">
              <span className="stat-num">{stats.practicalsCount} / {stats.lecturesCount}</span>
              <span className="stat-label">Praktyczne+Sem.+Ćw. / Wykłady</span>
            </div>
          </div>

          {/* Małe elementy do poklikania */}
          <div className="fun-widgets-grid">
            <div className="fun-card">
              <h4>🏋️ Najcięższy dzień w semestrze</h4>
              <p>
                {funFacts.hardestDayDate
                  ? `${formatDateReadable(funFacts.hardestDayDate)} (${funFacts.hardestDayHours} godz. zajęć)`
                  : 'Brak danych'}
              </p>
            </div>
            <div className="fun-card">
              <h4>🌅 Najwcześniejsze rozpoczęcie</h4>
              <p>
                {funFacts.earliestEvent
                  ? `Godz. ${funFacts.earliestEvent.startTime} (${funFacts.earliestEvent.subject})`
                  : 'Brak danych'}
              </p>
            </div>
            <div className="fun-card">
              <h4>📏 Najdłuższy dzień (od startu do końca)</h4>
              <p>
                {funFacts.longestSpanDate
                  ? `${formatDateReadable(funFacts.longestSpanDate)} (${funFacts.longestSpanHours}h rozpiętości)`
                  : 'Brak danych'}
              </p>
            </div>
            <div className="fun-card">
              <h4>📅 Bieżący tydzień</h4>
              <p>
                Liczba wydarzeń: <strong>{funFacts.currentWeekCount}</strong> (łącznie{' '}
                <strong>{funFacts.currentWeekHours}h</strong>)
              </p>
            </div>
          </div>

          <div className="badness-box">
            <button className="btn-primary" onClick={generateBadnessComment}>
              🎲 Jak bardzo jest źle? (Kliknij po diagnozę)
            </button>
            {badnessComment && <p className="badness-result">{badnessComment}</p>}
          </div>
        </div>
      </details>

      {/* PORÓWNYWARKA GRUP: "KTO MA GORZEJ?" */}
      <details className="card collapsible-section">
        <summary className="section-summary">
          ⚖️ Porównywarka grup: „Kto ma gorzej?”
        </summary>
        <div className="section-body">
          <div className="comparator-controls">
            <label>
              Porównaj swoją grupę (<strong>{selectedGroup}</strong>) z grupą:{' '}
              <select
                value={compareGroup}
                onChange={(e) => setCompareGroup(e.target.value)}
              >
                {scheduleData.availableSub4Groups
                  .filter((g) => g !== selectedGroup)
                  .map((g) => (
                    <option key={g} value={g}>
                      {g}
                    </option>
                  ))}
              </select>
            </label>

            <label>
              Okres porównania:{' '}
              <select
                value={compareScope}
                onChange={(e) => setCompareScope(e.target.value as 'semester' | 'week')}
              >
                <option value="semester">Cały semestr zimowy</option>
                <option value="week">Wybrany tydzień ({weeksList[selectedWeekIndex]})</option>
              </select>
            </label>
          </div>

          <p className="thresholds-info">
            Jawne progi porównania: <strong>Wczesny poranek</strong> = start &le;{' '}
            {SCHEDULE_THRESHOLDS.EARLY_START_STANDARD} (w tym dyżury o 07:30) •{' '}
            <strong>Długi dzień</strong> = min. {SCHEDULE_THRESHOLDS.LONG_DAY_HOURS}h zajęć •{' '}
            <strong>Długa przerwa (okienko)</strong> = przerwa &gt;{' '}
            {SCHEDULE_THRESHOLDS.LONG_BREAK_MINUTES} min.
          </p>

          {comparisonResult && (
            <>
              <div className="comparison-table-wrapper">
                <table className="comparison-table">
                  <thead>
                    <tr>
                      <th>Kryterium</th>
                      <th>Twoja grupa ({comparisonResult.m1.groupId})</th>
                      <th>Grupa {comparisonResult.m2.groupId}</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td>Łączna liczba znanych godzin zegarowych</td>
                      <td>
                        <strong>{comparisonResult.m1.totalKnownHours}h</strong>
                        {comparisonResult.m1.unknownHoursEventsCount > 0 &&
                          ` (+${comparisonResult.m1.unknownHoursEventsCount} dni POZ bez podanych godz.)`}
                      </td>
                      <td>
                        <strong>{comparisonResult.m2.totalKnownHours}h</strong>
                        {comparisonResult.m2.unknownHoursEventsCount > 0 &&
                          ` (+${comparisonResult.m2.unknownHoursEventsCount} dni POZ bez podanych godz.)`}
                      </td>
                    </tr>
                    <tr>
                      <td>Liczba dni z zajęciami</td>
                      <td>{comparisonResult.m1.daysWithClasses} dni</td>
                      <td>{comparisonResult.m2.daysWithClasses} dni</td>
                    </tr>
                    <tr>
                      <td>Dni rozpoczynające się o 7:30 / &le; 8:00</td>
                      <td>
                        {comparisonResult.m1.veryEarly730DaysCount} dni o 7:30 (łącznie{' '}
                        {comparisonResult.m1.earlyMorningDaysCount} &le; 8:00)
                      </td>
                      <td>
                        {comparisonResult.m2.veryEarly730DaysCount} dni o 7:30 (łącznie{' '}
                        {comparisonResult.m2.earlyMorningDaysCount} &le; 8:00)
                      </td>
                    </tr>
                    <tr>
                      <td>Liczba długich dni (&ge; {SCHEDULE_THRESHOLDS.LONG_DAY_HOURS}h)</td>
                      <td>{comparisonResult.m1.longDaysCount} dni</td>
                      <td>{comparisonResult.m2.longDaysCount} dni</td>
                    </tr>
                    <tr>
                      <td>Liczba okienek (&gt; {SCHEDULE_THRESHOLDS.LONG_BREAK_MINUTES} min)</td>
                      <td>{comparisonResult.m1.longBreaksCount}</td>
                      <td>{comparisonResult.m2.longBreaksCount}</td>
                    </tr>
                  </tbody>
                </table>
              </div>

              <div className="verdict-banner">
                <strong>Werdykt:</strong> {comparisonResult.verdict}
              </div>
            </>
          )}
        </div>
      </details>

      {/* RAPORT Z NIEJEDNOZNACZNOŚCI I E-LEARNINGU */}
      <details className="card collapsible-section">
        <summary className="section-summary">
          🔍 Informacje o e-learningu oraz raport z niejednoznaczności w pliku Excel
        </summary>
        <div className="section-body">
          <h4>Wykłady e-learningowe (bez stałych godzin w kalendarzu):</h4>
          <ul>
            {scheduleData.elearningNotes.map((n, i) => (
              <li key={i}>{n}</li>
            ))}
          </ul>

          <h4>Wykryte niejednoznaczności w pliku źródłowym plan.xlsx:</h4>
          <div className="ambiguities-list">
            {scheduleData.ambiguities.map((amb) => (
              <div key={amb.id} className="ambiguity-item">
                <h5>{amb.title}</h5>
                <p><strong>Komórki w Excelu:</strong> <code>{amb.sourceCells}</code></p>
                <p>{amb.description}</p>
                <p className="amb-res"><strong>Sposób obsługi w aplikacji:</strong> {amb.resolution}</p>
              </div>
            ))}
          </div>
        </div>
      </details>
    </div>
  );
}
