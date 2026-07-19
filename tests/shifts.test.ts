import { describe, expect, it } from 'vitest';
import type { CellDiagnostic, DayEntry, LocalDate } from '../src/domain/types';
import { interpretSchedule } from '../src/services/shifts';
import { normalizeMarker } from '../src/utils/normalize';

function entry(
  day: number,
  marker: string,
  options: {
    month?: number;
    year?: number;
    style?: Partial<CellDiagnostic>;
    kind?: DayEntry['kind'];
  } = {},
): DayEntry {
  const date: LocalDate = { year: options.year ?? 2026, month: options.month ?? 8, day };
  const diagnostic: CellDiagnostic = {
    address: `C${day + 4}`,
    rawValue: marker,
    displayedText: marker,
    isMerged: false,
    positionInDayGroup: 1,
    fontColor: '#000000',
    underline: false,
    italic: false,
    bold: false,
    ...options.style,
  };
  return {
    date,
    group: { day, startColumn: 3, endColumn: 4, valid: true },
    kind: options.kind ?? (marker ? 'single' : 'empty'),
    marker,
    normalizedMarker: normalizeMarker(marker),
    diagnostics: [diagnostic],
    selectedDiagnostic: diagnostic,
  };
}

function sevenFiveEntry(
  day: number,
  secondStyle: Partial<CellDiagnostic> = {},
): DayEntry {
  const first: CellDiagnostic = {
    address: `G${day + 4}`,
    rawValue: '7',
    displayedText: '7',
    isMerged: false,
    positionInDayGroup: 1,
    fontColor: '#000000',
    underline: false,
    italic: false,
    bold: false,
  };
  const second: CellDiagnostic = {
    address: `H${day + 4}`,
    rawValue: '5',
    displayedText: '5',
    isMerged: false,
    positionInDayGroup: 2,
    fontColor: '#000000',
    underline: false,
    italic: false,
    bold: false,
    ...secondStyle,
  };
  return {
    date: { year: 2026, month: 8, day },
    group: { day, startColumn: 7, endColumn: 8, valid: true },
    kind: 'double',
    marker: '7 / 5',
    normalizedMarker: '',
    diagnostics: [first, second],
  };
}

describe('szolgálatértelmező', () => {
  it.each(['#FFFFFF', '#C6EFCE', '#FFF2CC'])(
    'fekete 12 %s háttéren is Parti szolgálat 07:00–19:00',
    (fillColor) => {
      const result = interpretSchedule([
        entry(1, '12', { style: { fontColor: '#000000', fillColor } }),
      ], {});
      expect(result.events[0]).toMatchObject({
        summary: 'OMSZ',
        shiftType: 'Nappalos 07–19',
        serviceCategory: 'Parti szolgálat',
        shiftTime: { start: '2026-08-01T07:00:00', end: '2026-08-01T19:00:00' },
        calendarTime: { start: '2026-08-01T07:00:00', end: '2026-08-01T19:00:00' },
      });
    },
  );

  it('piros 12 a háttértől függetlenül Esetszolgálat 07:00–19:00', () => {
    const result = interpretSchedule([
      entry(1, '12', { style: { fontColor: '#FF0000', fillColor: '#C6EFCE' } }),
    ], {});
    expect(result.events[0]).toMatchObject({
      shiftType: 'Nappalos 07–19',
      serviceCategory: 'Esetszolgálat',
      shiftTime: { start: '2026-08-01T07:00:00', end: '2026-08-01T19:00:00' },
    });
  });

  it('kék 12 a háttértől függetlenül 06:00–18:00 szolgálat', () => {
    const result = interpretSchedule([
      entry(1, '12', { style: { fontColor: '#0000FF', fillColor: '#FFF2CC' } }),
    ], {});
    expect(result.events[0]).toMatchObject({
      shiftType: 'Nappalos 06–18',
      serviceCategory: 'Nappalos 06–18',
      shiftTime: { start: '2026-08-01T06:00:00', end: '2026-08-01T18:00:00' },
    });
  });

  it('zöld és aláhúzott 12-ből 10-es kocsi 10:00–22:00 szolgálatot készít', () => {
    const result = interpretSchedule([
      entry(1, '12', {
        style: { fontColor: '#008000', underline: true, fillColor: '#FFF2CC' },
      }),
    ], {});
    expect(result.events[0]).toMatchObject({
      shiftType: 'Nappalos 10–22',
      serviceCategory: '10-es kocsi',
      shiftTime: { start: '2026-08-01T10:00:00', end: '2026-08-01T22:00:00' },
    });
  });

  it('zöld háttér fekete 12-vel Parti marad, nem lesz 10-es kocsi', () => {
    const result = interpretSchedule([
      entry(1, '12', { style: { fontColor: '#000000', fillColor: '#00FF00' } }),
    ], {});
    expect(result.events[0]?.serviceCategory).toBe('Parti szolgálat');
  });

  it('zöld, de nem aláhúzott 12 bizonytalan', () => {
    const result = interpretSchedule([
      entry(1, '12', { style: { fontColor: '#008000', underline: false } }),
    ], {});
    expect(result.events).toHaveLength(0);
    expect(result.rows[0]?.status).toBe('Bizonytalan');
  });

  it.each([
    ['#FF0000', 'Esetszolgálat'],
    ['#000000', 'Parti szolgálat'],
  ] as const)('a %s betűszínű 17 szolgálati jellegét a kezdő cella adja', (fontColor, category) => {
    const result = interpretSchedule([
      entry(1, '17', { style: { fontColor } }),
      entry(2, '7', { style: { fontColor: '#008000' } }),
    ], {});
    expect(result.events[0]).toMatchObject({
      shiftType: '24 órás szolgálat',
      serviceCategory: category,
      shiftTime: { start: '2026-08-01T07:00:00', end: '2026-08-02T07:00:00' },
      calendarTime: { start: '2026-08-01T07:00:00', end: '2026-08-02T06:59:00' },
    });
  });

  it.each([
    ['#FF0000', 'Esetszolgálat'],
    ['#000000', 'Parti szolgálat'],
  ] as const)('a %s betűszínű 5 szolgálati jellegét a kezdő cella adja', (fontColor, category) => {
    const result = interpretSchedule([
      entry(1, '5', { style: { fontColor } }),
      entry(2, '7', { style: { fontColor: '#0000FF' } }),
    ], {});
    expect(result.events[0]).toMatchObject({
      shiftType: 'Éjszakai szolgálat',
      serviceCategory: category,
      shiftTime: { start: '2026-08-01T19:00:00', end: '2026-08-02T07:00:00' },
    });
  });

  it('az 5 → 7/5 → 7 mintából két folytonos éjszakai eseményt készít', () => {
    const result = interpretSchedule([
      entry(1, '5', { style: { fontColor: '#000000' } }),
      sevenFiveEntry(2, { fontColor: '#FF0000' }),
      entry(3, '7', { style: { fontColor: '#008000' } }),
    ], {});

    expect(result.events).toHaveLength(2);
    expect(result.events).toEqual([
      expect.objectContaining({
        serviceCategory: 'Parti szolgálat',
        calendarTime: {
          start: '2026-08-01T19:00:00',
          end: '2026-08-02T07:00:00',
        },
      }),
      expect.objectContaining({
        serviceCategory: 'Esetszolgálat',
        calendarTime: {
          start: '2026-08-02T19:00:00',
          end: '2026-08-03T07:00:00',
        },
      }),
    ]);
    expect(
      result.events.some(
        (item) =>
          item.calendarTime.start.endsWith('T00:00:00') ||
          item.calendarTime.end.endsWith('T24:00:00'),
      ),
    ).toBe(false);
  });

  it('a 7/5 nap két külön logikai sort kap hibás párosítás nélkül', () => {
    const result = interpretSchedule([
      entry(1, '5'),
      sevenFiveEntry(2),
      entry(3, '7'),
    ], {});
    const rows = result.rows.filter((row) => row.date.day === 2);

    expect(rows.map((row) => row.marker)).toEqual(['7', '5']);
    expect(rows.map((row) => row.status)).toEqual(['Felismerve', 'Exportálható']);
    expect(rows.some((row) => row.status === 'Hibás párosítás')).toBe(false);
    expect(rows[0]?.pairingReferences).toEqual([{ direction: 'previous', address: 'C5' }]);
    expect(rows[1]?.pairingReferences).toEqual([{ direction: 'next', address: 'C7' }]);
  });

  it('KMR-t másnap 01:00-ig értelmez', () => {
    const result = interpretSchedule([entry(31, 'KMR')], {});
    expect(result.events[0]).toMatchObject({
      summary: 'KMR',
      serviceCategory: 'KMR',
      shiftTime: { start: '2026-08-31T05:00:00', end: '2026-09-01T01:00:00' },
    });
  });

  it.each(['7', '17', '5'])('a párosítatlan %s hibás párosítás', (marker) => {
    const result = interpretSchedule([entry(3, marker)], {});
    expect(result.events).toHaveLength(0);
    expect(result.rows[0]?.status).toBe('Hibás párosítás');
  });

  it('kezeli a 17–7 hónapváltást', () => {
    const current = entry(1, '7', { month: 9 });
    const previous = entry(31, '17', { month: 8 });
    const result = interpretSchedule([current], { previous });
    expect(result.events[0]).toMatchObject({
      serviceCategory: 'Parti szolgálat',
      shiftTime: { start: '2026-08-31T07:00:00', end: '2026-09-01T07:00:00' },
      calendarTime: { start: '2026-08-31T07:00:00', end: '2026-09-01T06:59:00' },
    });
  });

  it('kezeli az 5–7 hónapváltást a hónap végéről', () => {
    const current = entry(30, '5', { month: 9 });
    const next = entry(1, '7', { month: 10 });
    const result = interpretSchedule([current], { next });
    expect(result.events).toHaveLength(1);
    expect(result.events[0]).toMatchObject({
      serviceCategory: 'Parti szolgálat',
      shiftTime: { start: '2026-09-30T19:00:00', end: '2026-10-01T07:00:00' },
      calendarTime: { start: '2026-09-30T19:00:00', end: '2026-10-01T06:59:00' },
    });
  });

  it.each(['x', 'sz.', 'ÁP', 'TK', 'szabadság', 'betegállomány', 'pihenőnap'])(
    '%s nem szolgálat',
    (marker) => {
      const result = interpretSchedule([entry(1, marker)], {});
      expect(result.events).toHaveLength(0);
      expect(result.rows[0]?.status).toBe('Kizárva');
    },
  );
});
