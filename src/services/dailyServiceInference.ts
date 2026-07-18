import type {
  DailyServicePattern,
  LocalDate,
  ReviewRow,
} from '../domain/types';
import type { EmployeeScheduleEntries } from '../excel/dayEntries';
import { normalizeMarker } from '../utils/normalize';
import { localDateKey } from './dates';
import { classifyTwelve, interpretSchedule } from './shifts';

interface DailyAccumulator {
  date: LocalDate;
  partyTwentyFourHourCount: number;
  blueTwelveAddresses: Set<string>;
  tenCarTwelveAddresses: Set<string>;
  blackTwelveCandidateAddresses: Set<string>;
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
          blueTwelveAddresses: new Set(),
          tenCarTwelveAddresses: new Set(),
          blackTwelveCandidateAddresses: new Set(),
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
        row.serviceCategory === 'Parti szolgálat' &&
        row.event?.shiftTime.start.slice(0, 10) === localDateKey(row.date)
      ) {
        accumulator.partyTwentyFourHourCount += 1;
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

      return [
        key,
        {
          date: accumulator.date,
          partyTwentyFourHourCount: accumulator.partyTwentyFourHourCount,
          blueTwelveCount,
          tenCarTwelveCount,
          blackTwelveCandidateCount,
          conflictingServiceMarkerCount: accumulator.conflictingServiceMarkerCount,
          correction,
        },
      ];
    }),
  );
}
