import { describe, expect, it, vi } from 'vitest';
import type { CellDiagnostic, DayEntry } from '../src/domain/types';
import type { EmployeeScheduleEntries } from '../src/excel/dayEntries';
import { buildDailyServicePatterns } from '../src/services/dailyServiceInference';
import { GoogleCalendarClient } from '../src/services/googleCalendar';
import { buildIcs } from '../src/services/ics';
import { interpretSchedule } from '../src/services/shifts';

function entry(
  row: number,
  day: number,
  marker: string,
  style: Partial<CellDiagnostic> = {},
): DayEntry {
  const column = day === 1 ? 'C' : 'E';
  const diagnostic: CellDiagnostic = {
    address: `${column}${row}`,
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
    date: { year: 2026, month: 8, day },
    group: {
      day,
      startColumn: day === 1 ? 3 : 5,
      endColumn: day === 1 ? 4 : 6,
      valid: true,
    },
    kind: marker === '' ? 'empty' : 'single',
    marker,
    normalizedMarker: marker,
    diagnostics: [diagnostic],
    selectedDiagnostic: diagnostic,
  };
}

function schedule(...current: DayEntry[]): EmployeeScheduleEntries {
  return { current };
}

function service17(row: number, fontColor: string, withClosing = true): EmployeeScheduleEntries {
  return schedule(
    entry(row, 1, '17', { fontColor }),
    ...(withClosing ? [entry(row, 2, '7', { fontColor: '#000000' })] : []),
  );
}

function interpretWithDailyContext(
  selected: EmployeeScheduleEntries,
  all: EmployeeScheduleEntries[],
) {
  const dailyServicePatterns = buildDailyServicePatterns(all);
  return {
    pattern: dailyServicePatterns.get('2026-08-01'),
    result: interpretSchedule(selected.current, {
      previous: selected.previous,
      next: selected.next,
      dailyServicePatterns,
    }),
  };
}

function response(body: unknown): Response {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
  });
}

describe('számos szolgálati formázások célzott korrekciója', () => {
  it('a közvetlen kék 12 ugyanúgy 6-os kocsi, mint a következtetett kék 12', () => {
    const result = interpretSchedule([
      entry(5, 1, '12', { fontColor: '#0000FF', fillColor: '#FFF2CC' }),
    ], {});

    expect(result.events[0]).toMatchObject({
      serviceCategory: '6-os kocsi',
      shiftTime: {
        start: '2026-08-01T06:00:00',
        end: '2026-08-01T18:00:00',
      },
    });
  });

  it('a közvetlen kék 12 ICS- és Google-leírása 6-os kocsit tartalmaz', async () => {
    const calendarEvent = interpretSchedule([
      entry(5, 1, '12', { fontColor: '#0000FF' }),
    ], {}).events[0];
    if (!calendarEvent) throw new Error('Hiányzó kék 12 tesztesemény.');

    expect(buildIcs([calendarEvent])).toContain('Szolgálati jelleg: 6-os kocsi');

    let body: unknown;
    const fetcher = vi.fn<typeof fetch>().mockImplementation((_input, init) => {
      if (typeof init?.body !== 'string') throw new Error('Hiányzó request body.');
      body = JSON.parse(init.body) as unknown;
      return Promise.resolve(response({ id: 'blue-12', colorId: '10' }));
    });
    await new GoogleCalendarClient('token', fetcher).insertEvent('primary', calendarEvent);
    expect(body).toMatchObject({ description: 'Szolgálati jelleg: 6-os kocsi' });
  });

  it.each([
    ['aláhúzott fehér háttéren', true, '#FFFFFF'],
    ['nem aláhúzott hétvégi zöld háttéren', false, '#C6EFCE'],
    ['nem aláhúzott ünnepnapi sárga háttéren', false, '#FFF2CC'],
  ] as const)('a zöld 12 $name is 10-es kocsi', (_name, underline, fillColor) => {
    const result = interpretSchedule([
      entry(5, 1, '12', { fontColor: '#008000', underline, fillColor }),
    ], {});

    expect(result.events[0]).toMatchObject({
      serviceCategory: '10-es kocsi',
      shiftTime: {
        start: '2026-08-01T10:00:00',
        end: '2026-08-01T22:00:00',
      },
    });
    expect(result.rows[0]?.status).toBe('Exportálható');
    if (!underline) {
      expect(result.rows[0]?.technicalNote).toBe(
        'Zöld 12 felismerve 10-es kocsiként. Az aláhúzás hiányzik, de a zöld betűszín alapján a szolgálat egyértelmű.',
      );
    }
  });

  it('piros és zöld 17 mellett a zöld jelölést Parti szolgálattá állítja helyre', () => {
    const emergency = service17(5, '#FF0000');
    const green = service17(7, '#008000');
    const { pattern, result } = interpretWithDailyContext(green, [emergency, green]);

    expect(pattern?.seventeenCorrection).toMatchObject({
      candidateAddress: 'C7',
      target: 'party',
    });
    expect(result.events[0]).toMatchObject({
      serviceCategory: 'Parti szolgálat',
      shiftTime: {
        start: '2026-08-01T07:00:00',
        end: '2026-08-02T07:00:00',
      },
      calendarTime: { end: '2026-08-02T06:59:00' },
    });
    expect(result.rows[0]?.technicalNote).toBe(
      'A zöld 17 formázási hibaként lett felismerve. Az adott napon Esetszolgálat már szerepelt, Parti szolgálat viszont hiányzott, ezért a jelölés Parti 24 órás szolgálatként lett értelmezve.',
    );
  });

  it('fekete és zöld 17 mellett a zöld jelölést Esetszolgálattá állítja helyre', () => {
    const party = service17(5, '#000000');
    const green = service17(7, '#006600');
    const { pattern, result } = interpretWithDailyContext(green, [party, green]);

    expect(pattern?.seventeenCorrection?.target).toBe('emergency');
    expect(result.events[0]?.serviceCategory).toBe('Esetszolgálat');
    expect(result.rows[0]?.serviceResolution).toMatchObject({
      originalServiceCategory: 'Nem meghatározható',
      finalServiceCategory: 'Esetszolgálat',
      formattingCorrectionApplied: true,
      dailyInferenceApplied: true,
    });
  });

  it.each([
    {
      name: 'csak egy zöld 17 van',
      schedules: (green: EmployeeScheduleEntries) => [green],
    },
    {
      name: 'fekete, piros és zöld 17 is van',
      schedules: (green: EmployeeScheduleEntries) => [
        service17(5, '#000000'),
        service17(6, '#FF0000'),
        green,
      ],
    },
    {
      name: 'két zöld 17 van',
      schedules: (green: EmployeeScheduleEntries) => [
        green,
        service17(8, '#006600'),
        service17(5, '#FF0000'),
      ],
    },
  ])('$name esetén nincs önkényes zöld 17 korrekció', ({ schedules }) => {
    const green = service17(7, '#008000');
    const { pattern, result } = interpretWithDailyContext(green, schedules(green));

    expect(pattern?.seventeenCorrection).toBeUndefined();
    expect(result.events).toHaveLength(0);
    expect(result.rows[0]?.status).toBe('Bizonytalan');
    expect(result.rows[0]?.note).toContain('nem következtethető biztosan');
  });

  it('párosítás nélküli zöld 17-ből napi összefüggés mellett sem készít eseményt', () => {
    const green = service17(7, '#008000', false);
    const { pattern, result } = interpretWithDailyContext(green, [
      service17(5, '#FF0000'),
      green,
    ]);

    expect(pattern?.seventeenCorrection).toBeUndefined();
    expect(result.events).toHaveLength(0);
    expect(result.rows[0]?.status).toBe('Hibás párosítás');
  });

  it('a helyreállított zöld 17 ICS- és Google-vége 06:59', async () => {
    const green = service17(7, '#008000');
    const { result } = interpretWithDailyContext(green, [
      service17(5, '#FF0000'),
      green,
    ]);
    const calendarEvent = result.events[0];
    if (!calendarEvent) throw new Error('Hiányzó helyreállított zöld 17 esemény.');

    expect(buildIcs([calendarEvent])).toContain(
      'DTEND;TZID=Europe/Budapest:20260802T065900',
    );
    let body: unknown;
    const fetcher = vi.fn<typeof fetch>().mockImplementation((_input, init) => {
      if (typeof init?.body !== 'string') throw new Error('Hiányzó request body.');
      body = JSON.parse(init.body) as unknown;
      return Promise.resolve(response({ id: 'green-17', colorId: '10' }));
    });
    await new GoogleCalendarClient('token', fetcher).insertEvent('primary', calendarEvent);
    expect(body).toMatchObject({
      end: { dateTime: '2026-08-02T06:59:00' },
    });
  });
});
