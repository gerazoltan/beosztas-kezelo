import { describe, expect, it } from 'vitest';
import type { DayEntry, LocalDate } from '../src/domain/types';
import { interpretSchedule } from '../src/services/shifts';
import { normalizeMarker } from '../src/utils/normalize';

const legend = { blue12: [], green12: [] };

function entry(
  day: number,
  marker: string,
  options: {
    month?: number;
    year?: number;
    style?: Partial<DayEntry['selectedDiagnostic']>;
    kind?: DayEntry['kind'];
  } = {},
): DayEntry {
  const date: LocalDate = { year: options.year ?? 2026, month: options.month ?? 8, day };
  const diagnostic = {
    address: `C${day + 4}`,
    rawValue: marker,
    displayedText: marker,
    isMerged: false,
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

describe('szolgálatértelmező', () => {
  it('kék 12-ből 06:00–18:00 szolgálatot készít', () => {
    const result = interpretSchedule([entry(1, '12', { style: { fillColor: '#C5D9F1' } })], {
      legend,
    });
    expect(result.events[0]).toMatchObject({
      shiftType: 'Nappalos 06–18',
      shiftTime: { start: '2026-08-01T06:00:00', end: '2026-08-01T18:00:00' },
      calendarTime: { start: '2026-08-01T06:00:00', end: '2026-08-01T18:00:00' },
    });
  });

  it('zöld, dőlt 12-ből 10:00–22:00 szolgálatot készít', () => {
    const result = interpretSchedule(
      [entry(1, '12', { style: { fontColor: '#008000', italic: true } })],
      { legend },
    );
    expect(result.events[0]).toMatchObject({
      shiftType: 'Nappalos 10–22',
      shiftTime: { start: '2026-08-01T10:00:00', end: '2026-08-01T22:00:00' },
      calendarTime: { start: '2026-08-01T10:00:00', end: '2026-08-01T22:00:00' },
    });
  });

  it('kitöltés nélküli 12-ből 07:00–19:00 szolgálatot készít', () => {
    const result = interpretSchedule([entry(1, '12')], { legend });
    expect(result.events[0]).toMatchObject({
      summary: 'OMSZ',
      shiftType: 'Nappalos 07–19',
      shiftTime: { start: '2026-08-01T07:00:00', end: '2026-08-01T19:00:00' },
      calendarTime: { start: '2026-08-01T07:00:00', end: '2026-08-01T19:00:00' },
    });
    expect(result.rows[0]).toMatchObject({
      status: 'Exportálható',
      note: 'Fehér vagy kitöltés nélküli 12 felismerve.',
      diagnostics: [expect.objectContaining({ fillCategory: 'noFill' })],
    });
  });

  it('más, nem támogatott színű 12 bizonytalan és nem exportálható', () => {
    const result = interpretSchedule(
      [entry(1, '12', { style: { fillColor: '#F4CCCC', hasVisibleFill: true } })],
      { legend },
    );
    expect(result.events).toHaveLength(0);
    expect(result.rows[0]).toMatchObject({
      status: 'Bizonytalan',
      diagnostics: [expect.objectContaining({ fillCategory: 'unsupported' })],
    });
  });

  it.each(['7', '5', '17', 'x', ''])(
    'a fehér vagy kitöltés nélküli %s nem válik önálló 07:00–19:00 szolgálattá',
    (marker) => {
      const result = interpretSchedule([entry(1, marker, { style: { hasVisibleFill: false } })], {
        legend,
      });
      expect(result.events).toHaveLength(0);
      expect(result.rows.some((row) => row.shiftType === 'Nappalos 07–19')).toBe(false);
    },
  );

  it('17–7 esetén a listában 24 órát, a naptárban 06:59-es befejezést használ', () => {
    const result = interpretSchedule([entry(1, '17'), entry(2, '7')], { legend });
    expect(result.events).toHaveLength(1);
    expect(result.events[0]).toMatchObject({
      shiftType: '24 órás szolgálat',
      shiftTime: { start: '2026-08-01T07:00:00', end: '2026-08-02T07:00:00' },
      calendarTime: { start: '2026-08-01T07:00:00', end: '2026-08-02T06:59:00' },
    });
    expect(result.rows.filter((row) => row.event)).toHaveLength(1);
    expect(result.rows.find((row) => row.event)?.technicalNote).toBe(
      'A naptáresemény befejezése 06:59 a jobb naptári elkülönítés érdekében.',
    );
  });

  it('az 5–7 éjszakai szolgálat listában és naptárban is másnap 07:00-ig tart', () => {
    const result = interpretSchedule([entry(1, '5'), entry(2, '7')], { legend });
    expect(result.events).toHaveLength(1);
    expect(result.events[0]).toMatchObject({
      shiftType: 'Éjszakai szolgálat',
      shiftTime: { start: '2026-08-01T19:00:00', end: '2026-08-02T07:00:00' },
      calendarTime: { start: '2026-08-01T19:00:00', end: '2026-08-02T07:00:00' },
    });
  });

  it('KMR-t másnap 01:00-ig értelmez', () => {
    const result = interpretSchedule([entry(31, 'KMR')], { legend });
    expect(result.events[0]).toMatchObject({
      summary: 'KMR',
      shiftTime: { start: '2026-08-31T05:00:00', end: '2026-09-01T01:00:00' },
      calendarTime: { start: '2026-08-31T05:00:00', end: '2026-09-01T01:00:00' },
    });
  });

  it.each(['7', '17', '5'])('a párosítatlan %s hibás párosítás', (marker) => {
    const result = interpretSchedule([entry(3, marker)], { legend });
    expect(result.events).toHaveLength(0);
    expect(result.rows[0]?.status).toBe('Hibás párosítás');
  });

  it('kezeli a 17–7 hónapváltást', () => {
    const current = entry(1, '7', { month: 9 });
    const previous = entry(31, '17', { month: 8 });
    const result = interpretSchedule([current], { legend, previous });
    expect(result.events[0]).toMatchObject({
      shiftTime: { start: '2026-08-31T07:00:00', end: '2026-09-01T07:00:00' },
      calendarTime: { start: '2026-08-31T07:00:00', end: '2026-09-01T06:59:00' },
    });
  });

  it('kezeli az 5–7 hónapváltást a hónap végéről', () => {
    const current = entry(30, '5', { month: 9 });
    const next = entry(1, '7', { month: 10 });
    const result = interpretSchedule([current], { legend, next });
    expect(result.events[0]).toMatchObject({
      shiftTime: { start: '2026-09-30T19:00:00', end: '2026-10-01T07:00:00' },
      calendarTime: { start: '2026-09-30T19:00:00', end: '2026-10-01T07:00:00' },
    });
  });

  it('hiányzó szomszéd hónapnál nem exportál', () => {
    expect(interpretSchedule([entry(31, '17')], { legend }).events).toHaveLength(0);
  });

  it.each(['x', 'sz.', 'ÁP', 'TK', 'szabadság', 'betegállomány', 'pihenőnap'])(
    '%s nem szolgálat',
    (marker) => {
      const result = interpretSchedule([entry(1, marker)], { legend });
      expect(result.events).toHaveLength(0);
      expect(result.rows[0]?.status).toBe('Kizárva');
    },
  );
});
