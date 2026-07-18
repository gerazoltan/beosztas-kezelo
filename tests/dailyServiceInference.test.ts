import { describe, expect, it } from 'vitest';
import type { CellDiagnostic, DayEntry } from '../src/domain/types';
import type { EmployeeScheduleEntries } from '../src/excel/dayEntries';
import { buildDailyServicePatterns } from '../src/services/dailyServiceInference';
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
    kind: 'single',
    marker,
    normalizedMarker: marker,
    diagnostics: [diagnostic],
    selectedDiagnostic: diagnostic,
  };
}

function schedule(...current: DayEntry[]): EmployeeScheduleEntries {
  return { current };
}

function partyTwentyFourHour(row = 5, fontColor = '#000000'): EmployeeScheduleEntries {
  return schedule(
    entry(row, 1, '17', { fontColor }),
    entry(row, 2, '7', { fontColor: '#008000' }),
  );
}

function twelve(
  row: number,
  kind: 'black' | 'blue' | 'tenCar',
  fillColor?: string,
): EmployeeScheduleEntries {
  const style =
    kind === 'blue'
      ? { fontColor: '#0000FF', underline: false, fillColor }
      : kind === 'tenCar'
        ? { fontColor: '#008000', underline: true, fillColor }
        : { fontColor: '#000000', underline: false, fillColor };
  return schedule(entry(row, 1, '12', style));
}

function interpretedBlackTwelve(
  black: EmployeeScheduleEntries,
  schedules: EmployeeScheduleEntries[],
) {
  const dailyServicePatterns = buildDailyServicePatterns(schedules);
  return {
    pattern: dailyServicePatterns.get('2026-08-01'),
    result: interpretSchedule(black.current, { dailyServicePatterns }),
  };
}

describe('teljes munkalapos napi szolgálati következtetés', () => {
  it('Parti 24 óra + kék 12 + egy fekete 12 esetén 10-es kocsit következtet', () => {
    const black = twelve(7, 'black');
    const { pattern, result } = interpretedBlackTwelve(black, [
      partyTwentyFourHour(),
      twelve(6, 'blue'),
      black,
    ]);

    expect(pattern).toMatchObject({
      partyTwentyFourHourCount: 1,
      blueTwelveCount: 1,
      tenCarTwelveCount: 0,
      blackTwelveCandidateCount: 1,
      correction: { candidateAddress: 'C7', target: 'tenCar' },
    });
    expect(result.events[0]).toMatchObject({
      shiftType: 'Nappalos 10–22',
      serviceCategory: '10-es kocsi',
      calendarTime: {
        start: '2026-08-01T10:00:00',
        end: '2026-08-01T22:00:00',
      },
      inference: {
        source: 'daily-service-pattern',
        target: 'tenCar',
        originalServiceCategory: 'Parti szolgálat',
      },
    });
    expect(result.rows[0]).toMatchObject({
      status: 'Exportálható',
      dailyInference: {
        partyTwentyFourHourPresent: true,
        blueTwelvePresent: true,
        tenCarTwelvePresent: false,
        blackTwelveCandidateCount: 1,
        correctionApplied: true,
        originalShiftType: 'Nappalos 07–19',
        finalShiftType: 'Nappalos 10–22',
      },
    });
  });

  it('Parti 24 óra + zöld-aláhúzott 12 + egy fekete 12 esetén 6-os kocsit következtet', () => {
    const black = twelve(7, 'black');
    const { pattern, result } = interpretedBlackTwelve(black, [
      partyTwentyFourHour(),
      twelve(6, 'tenCar'),
      black,
    ]);

    expect(pattern?.correction).toMatchObject({ candidateAddress: 'C7', target: 'blue' });
    expect(result.events[0]).toMatchObject({
      shiftType: 'Nappalos 06–18',
      serviceCategory: '6-os kocsi',
      calendarTime: {
        start: '2026-08-01T06:00:00',
        end: '2026-08-01T18:00:00',
      },
      inference: { target: 'blue' },
    });
  });

  it.each([
    {
      name: 'mindkét nappali kocsitípus jelen van',
      schedules: (black: EmployeeScheduleEntries) => [
        partyTwentyFourHour(),
        twelve(6, 'blue'),
        twelve(8, 'tenCar'),
        black,
      ],
    },
    {
      name: 'egyik nappali kocsitípus sincs jelen',
      schedules: (black: EmployeeScheduleEntries) => [partyTwentyFourHour(), black],
    },
    {
      name: 'csak piros 24 órás Esetszolgálat van',
      schedules: (black: EmployeeScheduleEntries) => [
        partyTwentyFourHour(5, '#FF0000'),
        twelve(6, 'blue'),
        black,
      ],
    },
    {
      name: 'két fekete 12 jelölt van',
      schedules: (black: EmployeeScheduleEntries) => [
        partyTwentyFourHour(),
        twelve(6, 'blue'),
        black,
        twelve(8, 'black'),
      ],
    },
    {
      name: 'nincs 24 órás Parti szolgálat',
      schedules: (black: EmployeeScheduleEntries) => [twelve(6, 'blue'), black],
    },
    {
      name: 'két 24 órás Parti szolgálat verseng',
      schedules: (black: EmployeeScheduleEntries) => [
        partyTwentyFourHour(5),
        partyTwentyFourHour(9),
        twelve(6, 'blue'),
        black,
      ],
    },
    {
      name: 'másik hibás szolgálati marker teszi bizonytalanná a napot',
      schedules: (black: EmployeeScheduleEntries) => [
        partyTwentyFourHour(),
        twelve(6, 'blue'),
        black,
        schedule(entry(9, 1, '17')),
      ],
    },
  ])('$name esetén nem korrigál automatikusan', ({ schedules }) => {
    const black = twelve(7, 'black');
    const { pattern, result } = interpretedBlackTwelve(black, schedules(black));

    expect(pattern?.correction).toBeUndefined();
    expect(result.events[0]).toMatchObject({
      shiftType: 'Nappalos 07–19',
      serviceCategory: 'Parti szolgálat',
      calendarTime: {
        start: '2026-08-01T07:00:00',
        end: '2026-08-01T19:00:00',
      },
      inference: undefined,
    });
  });

  it('a hétvégi vagy ünnepnapi háttér nem befolyásolja a következtetést', () => {
    const black = twelve(7, 'black', '#FFF2CC');
    const { result } = interpretedBlackTwelve(black, [
      partyTwentyFourHour(),
      twelve(6, 'blue', '#C6EFCE'),
      black,
    ]);

    expect(result.events[0]).toMatchObject({
      serviceCategory: '10-es kocsi',
      calendarTime: {
        start: '2026-08-01T10:00:00',
        end: '2026-08-01T22:00:00',
      },
    });
  });
});
