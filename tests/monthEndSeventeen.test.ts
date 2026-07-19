import { describe, expect, it, vi } from 'vitest';
import type { CellDiagnostic, DayEntry } from '../src/domain/types';
import type { EmployeeScheduleEntries } from '../src/excel/dayEntries';
import { buildDailyServicePatterns } from '../src/services/dailyServiceInference';
import { GoogleCalendarClient } from '../src/services/googleCalendar';
import { buildIcs } from '../src/services/ics';
import { interpretSchedule } from '../src/services/shifts';

function entry(
  row: number,
  year: number,
  month: number,
  day: number,
  marker: string,
  fontColor = '#000000',
): DayEntry {
  const diagnostic: CellDiagnostic = {
    address: `C${row}`,
    rawValue: marker,
    displayedText: marker,
    isMerged: false,
    positionInDayGroup: 1,
    fontColor,
    underline: false,
    italic: false,
    bold: false,
  };
  return {
    date: { year, month, day },
    group: { day, startColumn: 3, endColumn: 4, valid: true },
    kind: marker === '' ? 'empty' : 'single',
    marker,
    normalizedMarker: marker,
    diagnostics: marker === '' ? [] : [diagnostic],
    selectedDiagnostic: marker === '' ? undefined : diagnostic,
  };
}

function response(body: unknown): Response {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
  });
}

describe('hónap utolsó napi 17 feltételezett lezárása', () => {
  it.each([
    {
      name: '31 napos hónap fekete 17',
      date: [2026, 8, 31] as const,
      color: '#000000',
      category: 'Parti szolgálat',
      end: '2026-09-01',
    },
    {
      name: '30 napos hónap piros 17',
      date: [2026, 9, 30] as const,
      color: '#FF0000',
      category: 'Esetszolgálat',
      end: '2026-10-01',
    },
    {
      name: 'február 28 nem szökőévben',
      date: [2026, 2, 28] as const,
      color: '#000000',
      category: 'Parti szolgálat',
      end: '2026-03-01',
    },
    {
      name: 'február 29 szökőévben',
      date: [2024, 2, 29] as const,
      color: '#FF0000',
      category: 'Esetszolgálat',
      end: '2024-03-01',
    },
    {
      name: 'december 31 évváltással',
      date: [2026, 12, 31] as const,
      color: '#000000',
      category: 'Parti szolgálat',
      end: '2027-01-01',
    },
  ])('$name esetén exportálható 24 órás eseményt készít', ({
    date,
    color,
    category,
    end,
  }) => {
    const [year, month, day] = date;
    const result = interpretSchedule([
      entry(5, year, month, day, '17', color),
    ], {});

    expect(result.events).toHaveLength(1);
    expect(result.events[0]).toMatchObject({
      shiftType: '24 órás szolgálat',
      serviceCategory: category,
      shiftTime: {
        start: `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}T07:00:00`,
        end: `${end}T07:00:00`,
      },
      calendarTime: {
        end: `${end}T06:59:00`,
      },
    });
    expect(result.rows[0]).toMatchObject({
      status: 'Exportálható',
      technicalNote:
        'A hónap utolsó napi 17 jelölése a következő hónap első napján feltételezett 7-tel lett lezárva.',
      serviceResolution: {
        assumedBoundaryPairing: true,
        pairingSource: 'assumed',
        pairingCell: 'feltételezett következő havi 7',
      },
    });
  });

  it('hónap közepén pár nélküli 17 továbbra is hibás párosítás', () => {
    const result = interpretSchedule([
      entry(5, 2026, 8, 15, '17'),
    ], {});

    expect(result.events).toHaveLength(0);
    expect(result.rows[0]?.status).toBe('Hibás párosítás');
  });

  it('tényleges következő havi 7 mellett egyetlen, ténylegesen párosított eseményt készít', () => {
    const current = entry(5, 2026, 8, 31, '17');
    const next = entry(5, 2026, 9, 1, '7');
    const result = interpretSchedule([current], { next });

    expect(result.events).toHaveLength(1);
    expect(result.rows).toHaveLength(1);
    expect(result.rows[0]).toMatchObject({
      pairingReferences: [{ direction: 'next', address: 'C5' }],
      serviceResolution: {
        assumedBoundaryPairing: false,
        pairingSource: 'actual',
        pairingCell: 'C5',
      },
    });
  });

  it('üres következő havi első cellánál feltételezett lezárást használ', () => {
    const result = interpretSchedule(
      [entry(5, 2026, 8, 31, '17')],
      { next: entry(5, 2026, 9, 1, '') },
    );

    expect(result.events).toHaveLength(1);
    expect(result.events[0]?.calendarTime.end).toBe('2026-09-01T06:59:00');
    expect(result.rows[0]?.serviceResolution?.assumedBoundaryPairing).toBe(true);
  });

  it('nem üres, de nem 7 következő havi cellánál nem feltételez lezárást', () => {
    const result = interpretSchedule(
      [entry(5, 2026, 8, 31, '17')],
      { next: entry(5, 2026, 9, 1, 'x') },
    );

    expect(result.events).toHaveLength(0);
    expect(result.rows[0]?.status).toBe('Hibás párosítás');
  });

  it('piros 17 mellett az egyértelmű hónapvégi zöld 17-et Parti szolgálattá állítja helyre', () => {
    const emergency: EmployeeScheduleEntries = {
      current: [entry(5, 2026, 8, 31, '17', '#FF0000')],
    };
    const green: EmployeeScheduleEntries = {
      current: [entry(7, 2026, 8, 31, '17', '#008000')],
    };
    const patterns = buildDailyServicePatterns([emergency, green]);
    const result = interpretSchedule(green.current, {
      dailyServicePatterns: patterns,
    });

    expect(patterns.get('2026-08-31')?.seventeenCorrection?.target).toBe('party');
    expect(result.events[0]).toMatchObject({
      serviceCategory: 'Parti szolgálat',
      calendarTime: { end: '2026-09-01T06:59:00' },
    });
  });

  it('napi támpont nélküli hónapvégi zöld 17 bizonytalan marad', () => {
    const green: EmployeeScheduleEntries = {
      current: [entry(7, 2026, 8, 31, '17', '#008000')],
    };
    const patterns = buildDailyServicePatterns([green]);
    const result = interpretSchedule(green.current, {
      dailyServicePatterns: patterns,
    });

    expect(patterns.get('2026-08-31')?.seventeenCorrection).toBeUndefined();
    expect(result.events).toHaveLength(0);
    expect(result.rows[0]?.status).toBe('Bizonytalan');
  });

  it('a hónapvégi 17 ICS- és Google-eseménye is 06:59-kor zárul', async () => {
    const calendarEvent = interpretSchedule([
      entry(5, 2026, 8, 31, '17'),
    ], {}).events[0];
    if (!calendarEvent) throw new Error('Hiányzó hónapvégi 17 tesztesemény.');

    expect(buildIcs([calendarEvent])).toContain(
      'DTEND;TZID=Europe/Budapest:20260901T065900',
    );
    let body: unknown;
    const fetcher = vi.fn<typeof fetch>().mockImplementation((_input, init) => {
      if (typeof init?.body !== 'string') throw new Error('Hiányzó request body.');
      body = JSON.parse(init.body) as unknown;
      return Promise.resolve(response({ id: 'month-end-17', colorId: '10' }));
    });
    await new GoogleCalendarClient('token', fetcher).insertEvent('primary', calendarEvent);
    expect(body).toMatchObject({
      end: { dateTime: '2026-09-01T06:59:00' },
    });
  });
});
