// src/types.ts

export type ClassType =
  | 'praktyczne'
  | 'cwiczenia'
  | 'seminarium'
  | 'wyklad';

export type DataCompleteness = 'complete' | 'partial' | 'ambiguous';

export interface GroupHierarchy {
  /** Pełny identyfikator podgrupy 4-osobowej, np. "1a1" */
  sub4: string;
  /** Identyfikator podgrupy 8-osobowej, np. "1a" */
  sub8: string;
  /** Numer grupy dziekańskiej (seminaryjnej), np. 1 ("grupa 1") */
  deanGroup: number;
  /** Człon literowy 8-osobowy: 'a' | 'b' | 'c' */
  letter8: 'a' | 'b' | 'c';
}

export interface ScheduleEvent {
  /** Unikalny, deterministyczny identyfikator wydarzenia (zapobiega duplikatom) */
  id: string;
  /** Nazwa przedmiotu, np. "CHIRURGIA I BLOK OPERACYJNY" */
  subject: string;
  /** Typ zajęć: praktyczne | cwiczenia | seminarium | wyklad */
  type: ClassType;
  /** Data wydarzenia w formacie ISO YYYY-MM-DD (strefa Europe/Warsaw) */
  date: string;
  /** Godzina rozpoczęcia HH:MM lub null, jeśli brak w planie (np. POZ praktyczne) */
  startTime: string | null;
  /** Godzina zakończenia HH:MM lub null, jeśli brak w planie */
  endTime: string | null;
  /** Czas trwania w godzinach zegarowych (np. 6.0) lub null, gdy godziny nieznane */
  durationHours: number | null;
  /** Lista znormalizowanych oznaczeń grup z Excela, których dotyczy wydarzenie (lub ['ALL'] dla wykładów) */
  groups: string[];
  /** Czy wydarzenie dotyczy podziału 12-osobowego (gdzie podgrupa 'c' dzieli się między 'a' i 'b') */
  is12PersonSplit?: boolean;
  /** Wskazana połowa 12-osobowa ('a' | 'b'), jeśli dotyczy ćwiczeń 12-osobowych */
  split12Half?: 'a' | 'b';
  /** Prowadzący zajęcia lub null, jeśli brak w źródle */
  instructor: string | null;
  /** Nazwa placówki / szpitala / katedry / zakładu lub null */
  facilityName: string | null;
  /** Dokładny adres (ulica, numer, kod pocztowy, miasto) lub null */
  address: string | null;
  /** Budynek / pawilon / skrzydło (np. "Centrum Dydaktyczne (CD)", "Pawilon XI D1") lub null */
  building: string | null;
  /** Numer sali (np. "sala 101", "Aula B") lub null */
  room: string | null;
  /** Dodatkowe wskazówki z planu (np. miejsce pierwszego spotkania, zmiana sali, TEAMS) */
  notes: string[];
  /** Oznaczenie kompletności danych */
  completeness: DataCompleteness;
  /** Powód niepełnych lub niejednoznacznych danych wyświetlany użytkownikowi */
  completenessNote: string | null;
  /** Komórka źródłowa w pliku Excel (do celów weryfikacji i audytu, np. "PLAN ZAJĘĆ!B8") */
  sourceRef: string;
}

export interface AmbiguityIssue {
  id: string;
  title: string;
  description: string;
  sourceCells: string;
  resolution: string;
}

export interface ParsedScheduleResult {
  events: ScheduleEvent[];
  availableSub4Groups: string[]; // ['1a1', '1a2', ..., '13c2']
  semesterStart: string | null;  // '2026-10-05'
  semesterEnd: string | null;    // '2027-01-29'
  elearningNotes: string[];
  ambiguities: AmbiguityIssue[];
}