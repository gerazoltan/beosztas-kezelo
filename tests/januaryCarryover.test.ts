import { describe, expect, it, vi } from 'vitest';
import type { CellDiagnostic, DayEntry } from '../src/domain/types';
import { GoogleCalendarClient } from '../src/services/googleCalendar';
import { buildIcs } from '../src/services/ics';
import { interpretSchedule } from '../src/services/shifts';

function entry(
  year: number,
  month: number,
  day: number,
  marker: string,
  fontColor = '#000000',
): DayEntry {
  const diagnostic: CellDiagnostic = {
    address: 'C5',
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

describe('január 1-jei előző havi áthúzódás', () => {
  it.each([
    ['#000000', 'Parti szolgálat'],
    ['#FF0000', 'Esetszolgálat'],
  ] as const)(
    'előző decemberi lap nélkül a %s 7 részleges %s eseményt készít',
    (fontColor, serviceCategory) => {
      const result = interpretSchedule([
        entry(2026, 1, 1, '7', fontColor),
      ], {});

      expect(result.events).toHaveLength(1);
      expect(result.events[0]).toMatchObject({
        shiftType: 'Előző hónapról áthúzódó szolgálat',
        serviceCategory,
        shiftTime: {
          start: '2026-01-01T00:00:00',
          end: '2026-01-01T06:59:00',
        },
        calendarTime: {
          start: '2026-01-01T00:00:00',
          end: '2026-01-01T06:59:00',
        },
        specialKind: 'previous-month-carryover-partial',
      });
      expect(result.rows[0]).toMatchObject({
        status: 'Exportálható',
        technicalNote:
          'A január 1-jei 7 az előző év decemberéről áthúzódó szolgálat lezáró jelöléseként lett felismerve.',
      });
    },
  );

  it('ismeretlen színű január 1-jei 7 bizonytalan és nem exportálható', () => {
    const result = interpretSchedule([
      entry(2026, 1, 1, '7', '#800080'),
    ], {});

    expect(result.events).toHaveLength(0);
    expect(result.rows[0]?.status).toBe('Bizonytalan');
  });

  it.each([
    {
      marker: '17',
      shiftType: '24 órás szolgálat',
      start: '2025-12-31T07:00:00',
    },
    {
      marker: '5',
      shiftType: 'Éjszakai szolgálat',
      start: '2025-12-31T19:00:00',
    },
  ])(
    'elérhető december 31-i $marker mellett teljes $shiftType eseményt és nem részlegeset készít',
    ({ marker, shiftType, start }) => {
      const result = interpretSchedule(
        [entry(2026, 1, 1, '7')],
        { previous: entry(2025, 12, 31, marker) },
      );

      expect(result.events).toHaveLength(1);
      expect(result.events[0]).toMatchObject({
        shiftType,
        calendarTime: { start },
        specialKind: undefined,
      });
    },
  );

  it('meglévő december 31-i teljes OMSZ-eseménynél nem küld beszúrást', async () => {
    const partial = interpretSchedule([
      entry(2026, 1, 1, '7'),
    ], {}).events[0];
    if (!partial) throw new Error('Hiányzó részleges tesztesemény.');
    const fetcher = vi.fn<typeof fetch>().mockResolvedValue(
      response({
        items: [
          {
            summary: 'OMSZ',
            start: { dateTime: '2025-12-31T07:00:00+01:00' },
            end: { dateTime: '2026-01-01T07:00:00+01:00' },
          },
        ],
      }),
    );

    const results = await new GoogleCalendarClient('token', fetcher).addEvents(
      'primary',
      [partial],
    );

    expect(fetcher).toHaveBeenCalledOnce();
    expect(results[0]).toMatchObject({
      status: 'Már szerepel a naptárban',
      message:
        'Már szerepel a naptárban az előző hónapról áthúzódó teljes szolgálat.',
      technicalDetails: 'Átfedő előző havi teljes esemény található: igen.',
    });
    expect(fetcher.mock.calls.some((call) => call[1]?.method === 'POST')).toBe(false);
  });

  it('06:59-kor záródó decemberi teljes eseményt is átfedésként talál meg', async () => {
    const partial = interpretSchedule([
      entry(2026, 1, 1, '7'),
    ], {}).events[0];
    if (!partial) throw new Error('Hiányzó részleges tesztesemény.');
    const fetcher = vi.fn<typeof fetch>().mockResolvedValue(
      response({
        items: [
          {
            summary: 'OMSZ',
            start: { dateTime: '2025-12-31T19:00:00+01:00' },
            end: { dateTime: '2026-01-01T06:59:00+01:00' },
          },
        ],
      }),
    );

    await expect(
      new GoogleCalendarClient('token', fetcher).hasPreviousMonthCarryoverOverlap(
        'primary',
        partial,
      ),
    ).resolves.toBe(true);
  });

  it('előző havi átfedés hiányában létrehozza a 00:00–06:59 részleges eseményt', async () => {
    const partial = interpretSchedule([
      entry(2026, 1, 1, '7', '#FF0000'),
    ], {}).events[0];
    if (!partial) throw new Error('Hiányzó részleges tesztesemény.');
    let requestBody: unknown;
    const fetcher = vi
      .fn<typeof fetch>()
      .mockResolvedValueOnce(response({ items: [] }))
      .mockResolvedValueOnce(response({ items: [] }))
      .mockImplementationOnce((_input, init) => {
        if (typeof init?.body !== 'string') throw new Error('Hiányzó request body.');
        requestBody = JSON.parse(init.body) as unknown;
        return Promise.resolve(response({ id: 'partial', colorId: '10' }));
      });

    const results = await new GoogleCalendarClient('token', fetcher).addEvents(
      'primary',
      [partial],
    );

    expect(results[0]).toMatchObject({
      status: 'Létrehozva',
      technicalDetails: 'Átfedő előző havi teljes esemény található: nem.',
    });
    expect(requestBody).toMatchObject({
      summary: 'OMSZ',
      start: { dateTime: '2026-01-01T00:00:00' },
      end: { dateTime: '2026-01-01T06:59:00' },
    });
  });

  it('az ICS tartalmazza a részleges eseményt és az importálási figyelmeztetést', () => {
    const partial = interpretSchedule([
      entry(2026, 1, 1, '7'),
    ], {}).events[0];
    if (!partial) throw new Error('Hiányzó részleges tesztesemény.');

    const ics = buildIcs([partial]);
    const unfoldedIcs = ics.replace(/\r\n /g, '');
    expect(ics).toContain('DTSTART;TZID=Europe/Budapest:20260101T000000');
    expect(ics).toContain('DTEND;TZID=Europe/Budapest:20260101T065900');
    expect(unfoldedIcs).toContain('Előző hónapról áthúzódó részleges szolgálat.');
    expect(unfoldedIcs).toContain('teljes december 31-i esemény nem szerepel-e már');
  });
});
