// src/parser/__tests__/excelParser.test.ts
import { describe, expect, it } from 'vitest';
import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';
import {
  doesEventMatchStudent,
  normalizeGroupToken,
  parseScheduleWorkbook,
} from '../excelParser';
import {
  calculateSemesterProgress,
  computeGroupComparisonMetrics,
  deduplicateSameSlotEvents,
  getEventStatus,
} from '../../utils/timeAndStats';

describe('Parser planu zajęć pielęgniarstwa (plan.xlsx)', () => {
  const excelPath = path.resolve(process.cwd(), 'public/plan.xlsx');
  const fileBuffer = fs.readFileSync(excelPath);
  const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
  const parsed = parseScheduleWorkbook(workbook);

  it('poprawnie normalizuje niejednolite oznaczenia grup z Excela', () => {
    expect(normalizeGroupToken(' 10 a')).toBe('10a');
    expect(normalizeGroupToken('5c 1')).toBe('5c1');
    expect(normalizeGroupToken('10B')).toBe('10b');
    expect(normalizeGroupToken('grupa 12')).toBe('grupa 12');
  });

  it('wykrywa dokładne daty początku i końca semestru z kolumny A', () => {
    expect(parsed.semesterStart).toBe('2026-10-05');
    expect(parsed.semesterEnd).toBe('2027-01-29');
  });

  it('odczytuje wszystkie 78 podgrup 4-osobowych (1a1..13c2)', () => {
    expect(parsed.availableSub4Groups.length).toBe(78);
    expect(parsed.availableSub4Groups).toContain('1a1');
    expect(parsed.availableSub4Groups).toContain('13c2');
  });

  it('poprawnie odczytuje salę POZ seminarium ze scalonej komórki BG9:CT20', () => {
    const pozSemGr11 = parsed.events.find(
      (e) =>
        e.subject.startsWith('POZ') &&
        e.type === 'seminarium' &&
        e.groups.includes('grupa 11') &&
        e.date === '2026-10-05'
    );
    expect(pozSemGr11).toBeDefined();
    expect(pozSemGr11?.address).toContain('Litewska 14/16');
    expect(pozSemGr11?.room).toContain('101 Pato');
  });

  it('nie wymyśla godzin dla POZ zajęcia praktyczne (startTime i durationHours są null)', () => {
    const pozPrakt = parsed.events.filter(
      (e) => e.subject.startsWith('POZ') && e.type === 'praktyczne'
    );
    expect(pozPrakt.length).toBeGreaterThan(0);
    expect(pozPrakt[0].startTime).toBeNull();
    expect(pozPrakt[0].durationHours).toBeNull();
    expect(pozPrakt[0].completeness).toBe('partial');
  });

  it('zastępuje klinikę Interny zajęciami u Prof. R. Steca we wskazane dni tygodnia (brak duplikatów)', () => {
    // Grupa 1a w tygodniu 02.11 - 06.11.2026 (R11) ma w poniedziałek i wtorek zajęcia u Prof. Steca
    const evs1aMon = parsed.events.filter(
      (e) =>
        e.date === '2026-11-02' &&
        doesEventMatchStudent(e, '1a1') &&
        e.type === 'praktyczne'
    );
    expect(evs1aMon.length).toBe(1);
    expect(evs1aMon[0].instructor).toContain('Prof. R. Stec');
    expect(evs1aMon[0].address).toContain('Banacha 1a');
  });

  it('uwzględnia nadpisaną datę 06.11.2026 w komórce AT7 (Podstawy Rehabilitacji dla 10b)', () => {
    const rehab10bNov6 = parsed.events.find(
      (e) =>
        e.subject === 'PODSTAWY REHABILITACJI' &&
        e.date === '2026-11-06' &&
        e.groups.includes('10b')
    );
    expect(rehab10bNov6).toBeDefined();
    expect(rehab10bNov6?.startTime).toBe('08:00');
  });

  it('przypisuje wykłady z arkusza WYKŁADY jako wspólne (ALL), w tym czwartek 08.10 z kolumny C', () => {
    const wykladOct8 = parsed.events.find(
      (e) => e.type === 'wyklad' && e.date === '2026-10-08'
    );
    expect(wykladOct8).toBeDefined();
    expect(wykladOct8?.startTime).toBe('16:30');
    expect(wykladOct8?.endTime).toBe('18:00');
    expect(doesEventMatchStudent(wykladOct8!, '7b2')).toBe(true);
  });

  it('poprawnie wylicza status TERAZ / NASTĘPNE oraz postęp semestru', () => {
    const sampleEv = parsed.events.find(
      (e) => e.date === '2026-10-13' && e.type === 'wyklad'
    )!;
    expect(getEventStatus(sampleEv, '2026-10-13', '14:30')).toBe('upcoming');
    expect(getEventStatus(sampleEv, '2026-10-13', '15:30')).toBe('ongoing');
    expect(getEventStatus(sampleEv, '2026-10-13', '18:00')).toBe('past');

    const progBefore = calculateSemesterProgress('2026-10-05', '2027-01-29', '2026-09-20');
    expect(progBefore.percent).toBe(0);

    const progAfter = calculateSemesterProgress('2026-10-05', '2027-01-29', '2027-02-10');
    expect(progAfter.percent).toBe(100);
  });

  it('nie duplikuje godzin w statystykach dla grup', () => {
    const evs1a1 = parsed.events.filter((e) => doesEventMatchStudent(e, '1a1', 'a'));
    const dedup = deduplicateSameSlotEvents(evs1a1);
    const metrics = computeGroupComparisonMetrics('1a1', dedup);
    expect(metrics.totalKnownHours).toBeGreaterThan(200);
    expect(metrics.unknownHoursEventsCount).toBe(10); // 2 tygodnie (10 dni) POZ praktyczne bez podanych godzin
  });
});