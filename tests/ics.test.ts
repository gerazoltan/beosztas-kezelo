import { describe, expect, it } from 'vitest';
import type { CalendarEvent } from '../src/domain/types';
import { buildIcs, escapeIcsText, foldIcsLine, icsFileName, stableUid } from '../src/services/ics';
import { interpretSchedule } from '../src/services/shifts';

function calendarEvent(id: string, summary: 'OMSZ' | 'KMR' = 'OMSZ'): CalendarEvent {
  return {
    id,
    summary,
    shiftType: summary === 'KMR' ? 'KMR' : 'Nappalos 06–18',
    shiftTime: { start: '2026-08-10T06:00:00', end: '2026-08-10T18:00:00' },
    calendarTime: { start: '2026-08-10T06:00:00', end: '2026-08-10T18:00:00' },
    timeZone: 'Europe/Budapest',
  };
}

describe('ICS-generátor', () => {
  it('Europe/Budapest időzónás, nem egész napos időpontokat ír', () => {
    const content = buildIcs([calendarEvent('a')], new Date('2026-01-01T00:00:00Z'));
    expect(content).toContain('BEGIN:VTIMEZONE\r\nTZID:Europe/Budapest');
    expect(content).toContain('DTSTART;TZID=Europe/Budapest:20260810T060000');
    expect(content).toContain('DTEND;TZID=Europe/Budapest:20260810T180000');
    expect(content).toContain('RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU');
    expect(content.endsWith('\r\n')).toBe(true);
  });

  it('több eseményt és egyedi, stabil UID-ket készít', () => {
    const first = calendarEvent('a');
    const second = {
      ...calendarEvent('b', 'KMR'),
      shiftTime: { start: '2026-08-11T05:00:00', end: '2026-08-12T01:00:00' },
      calendarTime: { start: '2026-08-11T05:00:00', end: '2026-08-12T01:00:00' },
    };
    const content = buildIcs([first, second]);
    expect(content.match(/BEGIN:VEVENT/g)).toHaveLength(2);
    expect(stableUid(first)).toBe(stableUid(first));
    expect(stableUid(first)).not.toBe(stableUid(second));
  });

  it('a 17–7 szolgálatot 07:00-tól másnap 06:59-ig exportálja', () => {
    const diagnostic = {
      address: 'G5',
      rawValue: '17',
      displayedText: '17',
      isMerged: false,
      italic: false,
      bold: false,
    };
    const result = interpretSchedule(
      [
        {
          date: { year: 2026, month: 8, day: 3 },
          group: { day: 3, startColumn: 7, endColumn: 8, valid: true },
          kind: 'single',
          marker: '17',
          normalizedMarker: '17',
          diagnostics: [diagnostic],
          selectedDiagnostic: diagnostic,
        },
        {
          date: { year: 2026, month: 8, day: 4 },
          group: { day: 4, startColumn: 9, endColumn: 10, valid: true },
          kind: 'single',
          marker: '7',
          normalizedMarker: '7',
          diagnostics: [{ ...diagnostic, address: 'I5', rawValue: '7', displayedText: '7' }],
          selectedDiagnostic: {
            ...diagnostic,
            address: 'I5',
            rawValue: '7',
            displayedText: '7',
          },
        },
      ],
      { legend: { blue12: [], green12: [] } },
    );

    const content = buildIcs(result.events);
    expect(content).toContain('DTSTART;TZID=Europe/Budapest:20260803T070000');
    expect(content).toContain('DTEND;TZID=Europe/Budapest:20260804T065900');
    expect(content).not.toContain('DTEND;TZID=Europe/Budapest:20260804T070000');
  });

  it('a fehér 12 szolgálatot 07:00–19:00 között exportálja', () => {
    const diagnostic = {
      address: 'C5',
      rawValue: '12',
      displayedText: '12',
      isMerged: false,
      hasVisibleFill: false,
      italic: false,
      bold: false,
    };
    const result = interpretSchedule(
      [
        {
          date: { year: 2026, month: 8, day: 1 },
          group: { day: 1, startColumn: 3, endColumn: 4, valid: true },
          kind: 'single',
          marker: '12',
          normalizedMarker: '12',
          diagnostics: [diagnostic],
          selectedDiagnostic: diagnostic,
        },
      ],
      { legend: { blue12: [], green12: [] } },
    );

    const content = buildIcs(result.events);
    expect(content).toContain('DTSTART;TZID=Europe/Budapest:20260801T070000');
    expect(content).toContain('DTEND;TZID=Europe/Budapest:20260801T190000');
  });

  it('escape-eli a különleges karaktereket és UTF-8 bájthosszon hajtogat', () => {
    expect(escapeIcsText('árvíz, sor; \\\núj')).toBe('árvíz\\, sor\\; \\\\\\núj');
    const folded = foldIcsLine(`DESCRIPTION:${'árvíztűrő '.repeat(12)}`);
    expect(folded).toContain('\r\n ');
    for (const line of folded.split('\r\n')) {
      expect(new TextEncoder().encode(line).length).toBeLessThanOrEqual(75);
    }
  });

  it('normalizált magyar fájlnevet készít', () => {
    expect(icsFileName('Teszt  Elek', 2026, 8)).toBe('teszt-elek-2026-augusztus.ics');
  });

  it('a bizonytalan bejegyzésből nem kerül VEVENT az ICS-be', () => {
    const diagnostic = {
      address: 'C5',
      rawValue: '12',
      displayedText: '12',
      isMerged: false,
      fillColor: '#F4CCCC',
      hasVisibleFill: true,
      italic: false,
      bold: false,
    };
    const result = interpretSchedule(
      [
        {
          date: { year: 2026, month: 8, day: 1 },
          group: { day: 1, startColumn: 3, endColumn: 4, valid: true },
          kind: 'single',
          marker: '12',
          normalizedMarker: '12',
          diagnostics: [diagnostic],
          selectedDiagnostic: diagnostic,
        },
      ],
      { legend: { blue12: [], green12: [] } },
    );
    expect(buildIcs(result.events)).not.toContain('BEGIN:VEVENT');
  });
});
