import type {
  DailyServicePattern,
  LocalDate,
  ReviewRow,
} from '../domain/types';
import type { EmployeeScheduleEntries } from '../excel/dayEntries';
import { isGreen } from '../excel/colors';
import { normalizeMarker } from '../utils/normalize';
import { localDateKey } from './dates';
import { classifyTwelve, interpretSchedule } from './shifts';

interface DailyAccumulator {
  date: LocalDate;
  partyTwentyFourHourCount: number;
  emergencyTwentyFourHourCount: number;
  blueTwelveAddresses: Set<string>;
  tenCarTwelveAddresses: Set<string>;
  blackTwelveCandidateAddresses: Set<string>;
  greenSeventeenCandidateAddresses: Set<string>;
  conflictingSeventeenCount: number;
  conflictingServiceMarkerCount: number;
}

function containsServiceMarker(row: ReviewRow): boolean {
  return row.diagnostics.some((diagnostic) =>
    ['12', '17', '5', '7'].includes(normalizeMarker(diagnostic.displayedText)),
  );
}

function twelveAddress(row: ReviewRow): string | undefined {
  return row.diagnostics.find(
    (diagnostic) => normalizeMarker(diagnostic.displayedText) === '12',
  )?.address;
}

function markerAddress(row: ReviewRow, marker: string): string | undefined {
  return row.diagnostics.find(
    (diagnostic) => normalizeMarker(diagnostic.displayedText) === marker,
  )?.address;
}

export function buildDailyServicePatterns(
  schedules: EmployeeScheduleEntries[],
): ReadonlyMap<string, DailyServicePattern> {
  const daily = new Map<string, DailyAccumulator>();
  for (const schedule of schedules) {
    for (const entry of schedule.current) {
      if (!entry.group.valid) continue;
      const key = localDateKey(entry.date);
      if (!daily.has(key)) {
        daily.set(key, {
          date: entry.date,
          partyTwentyFourHourCount: 0,
          emergencyTwentyFourHourCount: 0,
          blueTwelveAddresses: new Set(),
          tenCarTwelveAddresses: new Set(),
          blackTwelveCandidateAddresses: new Set(),
          greenSeventeenCandidateAddresses: new Set(),
          conflictingSeventeenCount: 0,
          conflictingServiceMarkerCount: 0,
        });
      }
    }
  }

  for (const schedule of schedules) {
    const result = interpretSchedule(schedule.current, {
      previous: schedule.previous,
      next: schedule.next,
    });
    for (const row of result.rows) {
      const accumulator = daily.get(localDateKey(row.date));
      if (!accumulator) continue;

      if (
        row.status === 'Exportálható' &&
        row.shiftType === '24 órás szolgálat' &&
        row.event?.shiftTime.start.slice(0, 10) === localDateKey(row.date)
      ) {
        if (row.serviceCategory === 'Parti szolgálat') {
          accumulator.partyTwentyFourHourCount += 1;
        } else if (row.serviceCategory === 'Esetszolgálat') {
          accumulator.emergencyTwentyFourHourCount += 1;
        }
      }

      if (normalizeMarker(row.marker) === '12') {
        const address = twelveAddress(row);
        const diagnostic = address
          ? row.diagnostics.find((item) => item.address === address)
          : undefined;
        const kind = classifyTwelve(diagnostic);
        if (address && row.status === 'Exportálható') {
          if (kind === 'blue') accumulator.blueTwelveAddresses.add(address);
          else if (kind === 'tenCar') accumulator.tenCarTwelveAddresses.add(address);
          else if (kind === 'party') accumulator.blackTwelveCandidateAddresses.add(address);
        }
      }

      if (normalizeMarker(row.marker) === '17') {
        const address = markerAddress(row, '17');
        const diagnostic = address
          ? row.diagnostics.find((item) => item.address === address)
          : undefined;
        if (
          address &&
          row.status === 'Bizonytalan' &&
          isGreen(diagnostic?.fontColor)
        ) {
          accumulator.greenSeventeenCandidateAddresses.add(address);
        }
        if (row.status === 'Bizonytalan' || row.status === 'Hibás párosítás') {
          accumulator.conflictingSeventeenCount += 1;
        }
      }

      if (
        (row.status === 'Bizonytalan' || row.status === 'Hibás párosítás') &&
        containsServiceMarker(row)
      ) {
        accumulator.conflictingServiceMarkerCount += 1;
      }
    }
  }

  return new Map(
    [...daily.entries()].map(([key, accumulator]) => {
      const blueTwelveCount = accumulator.blueTwelveAddresses.size;
      const tenCarTwelveCount = accumulator.tenCarTwelveAddresses.size;
      const blackTwelveCandidateCount = accumulator.blackTwelveCandidateAddresses.size;
      const greenSeventeenCandidateCount =
        accumulator.greenSeventeenCandidateAddresses.size;
      const candidateAddress = [...accumulator.blackTwelveCandidateAddresses][0];
      const safeBase =
        accumulator.partyTwentyFourHourCount === 1 &&
        accumulator.conflictingServiceMarkerCount === 0 &&
        blackTwelveCandidateCount === 1 &&
        candidateAddress !== undefined;
      const correction =
        safeBase && blueTwelveCount === 1 && tenCarTwelveCount === 0
          ? {
              candidateAddress,
              target: 'tenCar' as const,
              explanation:
                'A napi 24 órás Parti szolgálat és a 6-os kocsi mellett hiányzott a 10-es kocsi. A fekete 12 ezért elírt zöld 12-ként, 10:00–22:00 szolgálatként lett felismerve.',
            }
          : safeBase && blueTwelveCount === 0 && tenCarTwelveCount === 1
            ? {
                candidateAddress,
                target: 'blue' as const,
                explanation:
                  'A napi 24 órás Parti szolgálat és a 10-es kocsi mellett hiányzott a 6-os kocsi. A fekete 12 ezért elírt kék 12-ként, 06:00–18:00 szolgálatként lett felismerve.',
              }
            : undefined;
      const greenSeventeenAddress =
        [...accumulator.greenSeventeenCandidateAddresses][0];
      const seventeenCorrection =
        greenSeventeenCandidateCount === 1 &&
        accumulator.conflictingSeventeenCount === 1 &&
        greenSeventeenAddress !== undefined &&
        accumulator.emergencyTwentyFourHourCount === 1 &&
        accumulator.partyTwentyFourHourCount === 0
          ? {
              candidateAddress: greenSeventeenAddress,
              target: 'party' as const,
              explanation:
                'A zöld 17 formázási hibaként lett felismerve. Az adott napon Esetszolgálat már szerepelt, Parti szolgálat viszont hiányzott, ezért a jelölés Parti 24 órás szolgálatként lett értelmezve.',
            }
          : greenSeventeenCandidateCount === 1 &&
              accumulator.conflictingSeventeenCount === 1 &&
              greenSeventeenAddress !== undefined &&
              accumulator.partyTwentyFourHourCount === 1 &&
              accumulator.emergencyTwentyFourHourCount === 0
            ? {
                candidateAddress: greenSeventeenAddress,
                target: 'emergency' as const,
                explanation:
                  'A zöld 17 formázási hibaként lett felismerve. Az adott napon Parti szolgálat már szerepelt, Esetszolgálat viszont hiányzott, ezért a jelölés Esetszolgálatként lett értelmezve.',
              }
            : undefined;

      return [
        key,
        {
          date: accumulator.date,
          partyTwentyFourHourCount: accumulator.partyTwentyFourHourCount,
          emergencyTwentyFourHourCount: accumulator.emergencyTwentyFourHourCount,
          blueTwelveCount,
          tenCarTwelveCount,
          blackTwelveCandidateCount,
          greenSeventeenCandidateCount,
          conflictingSeventeenCount: accumulator.conflictingSeventeenCount,
          conflictingServiceMarkerCount: accumulator.conflictingServiceMarkerCount,
          correction,
          seventeenCorrection,
        },
      ];
    }),
  );
}
