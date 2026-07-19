import { describe, expect, it } from 'vitest';
import type { CellDiagnostic, DayEntry } from '../src/domain/types';
import { GoogleCalendarClient } from '../src/services/googleCalendar';
import { buildIcs } from '../src/services/ics';
import { interpretSchedule } from '../src/services/shifts';

function entry(
  year: number,
  month: number,
  day: number,
  marker: string,
  style: Partial<CellDiagnostic> = {},
): DayEntry {
  const diagnostic: CellDiagnostic = {
    address: 'C5',
    rawValue: marker,
    displayedText: marker,
    isMerged: false,
    positionInDayGroup: 1,
    fontColor: '#000000',
    underline: false,
    italic: false,
    bold: false,
    ...style,
  };
  return {
    date: { year, month, day },
    group: { day, startColumn: 3, endColumn: 4, valid: true },
    kind: 'single',
    marker,
    normalizedMarker: marker,
    diagnostics: [diagnostic],
    selectedDiagnostic: diagnostic,
  };
}

function response(body: unknown): Response {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
  });
}

describe('hónap utolsó napi 5 feltételezett lezárása', () => {
  it.each([
    {
      name: '31 napos hónap fekete 5',
      date: [2026, 8, 31] as const,
      fontColor: '#000000',
      fillColor: '#C6EFCE',
      category: 'Parti szolgálat',
      endDate: '2026-09-01',
    },
    {
      name: '30 napos hónap piros 5',
      date: [2026, 9, 30] as const,
      fontColor: '#FF0000',
      fillColor: '#FFF2CC',
      category: 'Esetszolgálat',
      endDate: '2026-10-01',
    },
    {
      name: 'február 28 nem szökőévben',
      date: [2026, 2, 28] as const,
      fontColor: '#000000',
      category: 'Parti szolgálat',
      endDate: '2026-03-01',
    },
    {
      name: 'február 29 szökőévben',
      date: [2024, 2, 29] as const,
      fontColor: '#FF0000',
      category: 'Esetszolgálat',
      endDate: '2024-03-01',
    },
    {
      name: 'december 31 évváltással',
      date: [2026, 12, 31] as const,
      fontColor: '#000000',
      category: 'Parti szolgálat',
      endDate: '2027-01-01',
    },
  ])('$name esetén exportálható éjszakai eseményt készít', ({
    date,
    fontColor,
    fillColor,
    category,
    endDate,
  }) => {
    const [year, month, day] = date;
    const result = interpretSchedule([
      entry(year, month, day, '5', { fontColor, fillColor }),
    ], {});

    expect(result.events).toHaveLength(1);
    expect(result.events[0]).toMatchObject({
      shiftType: 'Éjszakai szolgálat',
      serviceCategory: category,
      shiftTime: {
        start: `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}T19:00:00`,
        end: `${endDate}T07:00:00`,
      },
      calendarTime: {
        end: `${endDate}T06:59:00`,
      },
    });
    expect(result.rows[0]).toMatchObject({
      status: 'Exportálható',
      note: 'A hónap utolsó napi 5 jelölése a következő hónap első napján feltételezett 7-tel lett lezárva.',
      technicalNote:
        'A hónap utolsó napi 5 jelölése a következő hónap első napján feltételezett 7-tel lett lezárva.',
    });
  });

  it('hónap közepén a pár nélküli 5 továbbra is hibás párosítás', () => {
    const result = interpretSchedule([entry(2026, 8, 15, '5')], {});

    expect(result.events).toHaveLength(0);
    expect(result.rows[0]?.status).toBe('Hibás párosítás');
  });

  it('tényleges következő havi 7 mellett pontosan egy eseményt készít', () => {
    const current = entry(2026, 8, 31, '5');
    const next = entry(2026, 9, 1, '7', { fontColor: '#008000' });
    const result = interpretSchedule([current], { next });

    expect(result.events).toHaveLength(1);
    expect(result.rows).toHaveLength(1);
    expect(result.events[0]?.calendarTime).toEqual({
      start: '2026-08-31T19:00:00',
      end: '2026-09-01T06:59:00',
    });
    expect(result.rows[0]?.pairingReferences).toEqual([
      { direction: 'next', address: 'C5' },
    ]);
  });

  it('a hónapvégi 5 akkor is feltételezett 7-tel zárul, ha a következő havi adat nem 7', () => {
    const current = entry(2026, 8, 31, '5');
    const next = entry(2026, 9, 1, 'x');
    const result = interpretSchedule([current], { next });

    expect(result.events).toHaveLength(1);
    expect(result.events[0]?.calendarTime.end).toBe('2026-09-01T06:59:00');
    expect(result.rows[0]?.pairingReferences).toBeUndefined();
    expect(result.rows[0]?.note).toContain('feltételezett 7-tel');
  });

  it('az ICS és a Google request body is pontosan 06:59-kor zárul', async () => {
    const result = interpretSchedule([entry(2026, 8, 31, '5')], {});
    const calendarEvent = result.events[0];
    if (!calendarEvent) throw new Error('Hiányzó hónapvégi tesztesemény.');

    const ics = buildIcs([calendarEvent]);
    expect(ics).toContain('DTSTART;TZID=Europe/Budapest:20260831T190000');
    expect(ics).toContain('DTEND;TZID=Europe/Budapest:20260901T065900');
    expect(ics).not.toContain('DTEND;TZID=Europe/Budapest:20260901T070000');

    let requestBody: unknown;
    const fetcher: typeof fetch = (_input, init) => {
      if (typeof init?.body !== 'string') throw new Error('Hiányzó Google request body.');
      requestBody = JSON.parse(init.body) as unknown;
      return Promise.resolve(response({ id: 'created-month-end-five', colorId: '10' }));
    };
    await new GoogleCalendarClient('token', fetcher).insertEvent('primary', calendarEvent);
    expect(requestBody).toMatchObject({
      summary: 'OMSZ',
      start: {
        dateTime: '2026-08-31T19:00:00',
        timeZone: 'Europe/Budapest',
      },
      end: {
        dateTime: '2026-09-01T06:59:00',
        timeZone: 'Europe/Budapest',
      },
      colorId: '10',
    });
  });
});
