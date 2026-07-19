import type {
  CalendarEvent,
  CellDiagnostic,
  DailyInferenceTechnicalDetails,
  DailyServicePattern,
  DayEntry,
  EventTimeRange,
  LegendStyles,
  ResolvedStyle,
  ReviewRow,
  ScheduleResult,
  ServiceCategory,
  ServiceInference,
  ShiftType,
} from '../domain/types';
import { isBlack, isBlue, isGreen, isRed } from '../excel/colors';
import { normalizeMarker } from '../utils/normalize';
import { addDays, daysInMonth, localDateKey, localDateTime } from './dates';

const NON_SERVICE_MARKERS = new Set([
  '',
  'x',
  'sz',
  'áp',
  'ap',
  'tk',
  'szabadság',
  'szabadsag',
  'betegállomány',
  'betegallomany',
  'beteg',
  'táppénz',
  'tappenz',
  'pihenőnap',
  'pihenonap',
  'pihenő',
  'piheno',
]);

function stableHash(value: string): string {
  let hash = 0x811c9dc5;
  for (const character of value) {
    hash ^= character.codePointAt(0) ?? 0;
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, '0');
}

function event(
  summary: 'OMSZ' | 'KMR',
  shiftType: ShiftType,
  serviceCategory: ServiceCategory,
  shiftTime: EventTimeRange,
  calendarTime: EventTimeRange = shiftTime,
  inference?: ServiceInference,
): CalendarEvent {
  return {
    id: stableHash(
      `${summary}|${shiftType}|${serviceCategory}|${calendarTime.start}|${calendarTime.end}`,
    ),
    summary,
    shiftType,
    serviceCategory,
    shiftTime,
    calendarTime,
    timeZone: 'Europe/Budapest',
    inference,
  };
}

function timeRange(start: string, end: string): EventTimeRange {
  return { start, end };
}

function hasDefaultBlackFont(style: ResolvedStyle | undefined): boolean {
  if (!style) return false;
  return (
    isBlack(style.fontColor) ||
    (style.fontColor === undefined && style.fontColorRaw === undefined)
  );
}

type StartingServiceKind = 'party' | 'emergency' | 'unknown';

function classifyStartingService(style: ResolvedStyle | undefined): StartingServiceKind {
  if (!style) return 'unknown';
  if (isRed(style.fontColor)) return 'emergency';
  if (hasDefaultBlackFont(style)) return 'party';
  return 'unknown';
}

export type TwelveKind = 'party' | 'emergency' | 'blue' | 'tenCar' | 'unknown';

export function classifyTwelve(style: ResolvedStyle | undefined): TwelveKind {
  if (!style) return 'unknown';
  if (style.underline && isGreen(style.fontColor)) return 'tenCar';
  if (isBlue(style.fontColor)) return 'blue';
  if (isRed(style.fontColor)) return 'emergency';
  if (hasDefaultBlackFont(style)) return 'party';
  return 'unknown';
}

interface RowMetadata {
  timeRule: string;
  pairingReferences?: ReviewRow['pairingReferences'];
  dailyInference?: DailyInferenceTechnicalDetails;
  technicalNote?: string;
}

function rowForEvent(
  entry: DayEntry,
  created: CalendarEvent,
  note: string,
  metadata: RowMetadata,
): ReviewRow {
  return {
    id: `${localDateKey(entry.date)}-${created.id}`,
    date: entry.date,
    marker: entry.marker,
    shiftType: created.shiftType,
    serviceCategory: created.serviceCategory,
    summary: created.summary,
    status: 'Exportálható',
    note,
    timeRule: metadata.timeRule,
    pairingReferences: metadata.pairingReferences,
    dailyInference: metadata.dailyInference,
    technicalNote:
      metadata.technicalNote ??
      (created.shiftType === '24 órás szolgálat'
        ? 'A naptáresemény befejezése 06:59 a jobb naptári elkülönítés érdekében.'
        : undefined),
    diagnostics: entry.diagnostics,
    event: created,
  };
}

function issueRow(
  entry: DayEntry,
  status: ReviewRow['status'],
  note: string,
  metadata?: Partial<RowMetadata>,
): ReviewRow {
  return {
    id: `${localDateKey(entry.date)}-${status}-${
      entry.selectedDiagnostic?.address ?? entry.group.startColumn
    }`,
    date: entry.date,
    marker: entry.marker,
    status,
    note,
    timeRule: metadata?.timeRule,
    pairingReferences: metadata?.pairingReferences,
    diagnostics: entry.diagnostics,
  };
}

function recognizedPairingRow(
  entry: DayEntry,
  paired: CalendarEvent,
  note: string,
  metadata: RowMetadata,
): ReviewRow {
  return {
    id: `${localDateKey(entry.date)}-Felismerve-${
      entry.selectedDiagnostic?.address ?? entry.group.startColumn
    }`,
    date: entry.date,
    marker: entry.marker,
    shiftType: paired.shiftType,
    serviceCategory: paired.serviceCategory,
    summary: paired.summary,
    status: 'Felismerve',
    note,
    timeRule: metadata.timeRule,
    pairingReferences: metadata.pairingReferences,
    diagnostics: entry.diagnostics,
  };
}

function isConsecutive(first: DayEntry, second: DayEntry): boolean {
  return localDateKey(addDays(first.date, 1)) === localDateKey(second.date);
}

function isLastCalendarDay(entry: DayEntry): boolean {
  return entry.date.day === daysInMonth(entry.date.year, entry.date.month);
}

interface MarkerOccurrence {
  entry: DayEntry;
  diagnostic: CellDiagnostic;
  normalizedMarker: string;
}

function markerOccurrences(entry: DayEntry): MarkerOccurrence[] {
  return entry.diagnostics
    .filter((diagnostic) => diagnostic.displayedText.trim() !== '')
    .sort((first, second) => first.positionInDayGroup - second.positionInDayGroup)
    .map((diagnostic) => ({
      entry,
      diagnostic,
      normalizedMarker: normalizeMarker(diagnostic.displayedText),
    }));
}

function occurrenceEntry(occurrence: MarkerOccurrence): DayEntry {
  const marker = occurrence.diagnostic.displayedText.trim();
  return {
    ...occurrence.entry,
    kind: 'single',
    marker,
    normalizedMarker: occurrence.normalizedMarker,
    selectedDiagnostic: occurrence.diagnostic,
  };
}

function findMarker(entry: DayEntry, marker: string): MarkerOccurrence | undefined {
  return markerOccurrences(entry).find((occurrence) => occurrence.normalizedMarker === marker);
}

function findStartingMarker(entry: DayEntry): MarkerOccurrence | undefined {
  return markerOccurrences(entry)
    .filter((occurrence) => ['17', '5'].includes(occurrence.normalizedMarker))
    .at(-1);
}

interface PairedShift {
  event?: CalendarEvent;
  timeRule: string;
}

function pairedShift(
  startOccurrence: MarkerOccurrence,
  endOccurrence: MarkerOccurrence,
): PairedShift | undefined {
  const startEntry = startOccurrence.entry;
  const endEntry = endOccurrence.entry;
  const marker = startOccurrence.normalizedMarker;
  if (
    endOccurrence.normalizedMarker !== '7' ||
    !['17', '5'].includes(marker) ||
    !isConsecutive(startEntry, endEntry)
  ) {
    return undefined;
  }

  const serviceKind = classifyStartingService(startOccurrence.diagnostic);
  const serviceCategory =
    serviceKind === 'party'
      ? ('Parti szolgálat' as const)
      : serviceKind === 'emergency'
        ? ('Esetszolgálat' as const)
        : undefined;

  if (marker === '17') {
    const timeRule = '17 + következő napi 7 → 07:00–másnap 06:59';
    return {
      event: serviceCategory
        ? event(
            'OMSZ',
            '24 órás szolgálat',
            serviceCategory,
            timeRange(
              localDateTime(startEntry.date, '07:00'),
              localDateTime(endEntry.date, '07:00'),
            ),
            timeRange(
              localDateTime(startEntry.date, '07:00'),
              localDateTime(endEntry.date, '06:59'),
            ),
          )
        : undefined,
      timeRule,
    };
  }

  const monthEnd = isLastCalendarDay(startEntry);
  const timeRule = monthEnd
    ? 'Hónap utolsó napi 5 + következő havi 7 → 19:00–másnap 06:59'
    : '5 + következő napi 7 → 19:00–másnap 07:00';
  const shiftTime = timeRange(
    localDateTime(startEntry.date, '19:00'),
    localDateTime(endEntry.date, '07:00'),
  );
  return {
    event: serviceCategory
      ? event(
          'OMSZ',
          'Éjszakai szolgálat',
          serviceCategory,
          shiftTime,
          monthEnd
            ? timeRange(
                localDateTime(startEntry.date, '19:00'),
                localDateTime(endEntry.date, '06:59'),
              )
            : shiftTime,
        )
      : undefined,
    timeRule,
  };
}

function assumedMonthEndFiveShift(startOccurrence: MarkerOccurrence): PairedShift | undefined {
  if (
    startOccurrence.normalizedMarker !== '5' ||
    !isLastCalendarDay(startOccurrence.entry)
  ) {
    return undefined;
  }
  const serviceKind = classifyStartingService(startOccurrence.diagnostic);
  const serviceCategory =
    serviceKind === 'party'
      ? ('Parti szolgálat' as const)
      : serviceKind === 'emergency'
        ? ('Esetszolgálat' as const)
        : undefined;
  const endDate = addDays(startOccurrence.entry.date, 1);
  return {
    event: serviceCategory
      ? event(
          'OMSZ',
          'Éjszakai szolgálat',
          serviceCategory,
          timeRange(
            localDateTime(startOccurrence.entry.date, '19:00'),
            localDateTime(endDate, '07:00'),
          ),
          timeRange(
            localDateTime(startOccurrence.entry.date, '19:00'),
            localDateTime(endDate, '06:59'),
          ),
        )
      : undefined,
    timeRule: 'Hónap utolsó napi 5 + feltételezett következő havi 7 → 19:00–másnap 06:59',
  };
}

export interface InterpretOptions {
  legend?: LegendStyles;
  previous?: DayEntry;
  next?: DayEntry;
  dailyServicePatterns?: ReadonlyMap<string, DailyServicePattern>;
}

export function interpretSchedule(
  entries: DayEntry[],
  { previous, next, dailyServicePatterns }: InterpretOptions,
): ScheduleResult {
  const rows: ReviewRow[] = [];
  const events: CalendarEvent[] = [];

  const processStartingOccurrence = (occurrence: MarkerOccurrence, index: number): void => {
    const entry = occurrenceEntry(occurrence);
    const marker = occurrence.normalizedMarker;
    const following = entries[index + 1] ?? next;
    const closing = following ? findMarker(following, '7') : undefined;
    const assumedClosing = closing ? undefined : assumedMonthEndFiveShift(occurrence);
    const paired = closing ? pairedShift(occurrence, closing) : assumedClosing;

    if (!paired) {
      rows.push(
        issueRow(
          entry,
          'Hibás párosítás',
          following
            ? `A ${marker} jelölést nem követi a következő naptári napon 7.`
            : 'A következő havi munkalap hiányzik; a párosítás nem ellenőrizhető.',
        ),
      );
      return;
    }
    if (!paired.event) {
      rows.push(
        issueRow(
          entry,
          'Bizonytalan',
          `A kezdő ${marker} betűszíne nem sorolható Parti vagy Esetszolgálathoz.`,
          {
            timeRule: paired.timeRule,
            pairingReferences: closing
              ? [{ direction: 'next', address: closing.diagnostic.address }]
              : undefined,
          },
        ),
      );
      return;
    }

    events.push(paired.event);
    rows.push(
      rowForEvent(
        entry,
        paired.event,
        occurrence.entry.kind === 'double' && marker === '5'
          ? 'Az 5 új éjszakai szolgálatot indít; az adott napon 19:00–24:00 szolgálati szakasz.'
          : assumedClosing
            ? 'A hónap utolsó napi 5 jelölése a következő hónap első napján feltételezett 7-tel lett lezárva.'
            : `${marker}–7 pár egyetlen szolgálatként felismerve.`,
        {
          timeRule: paired.timeRule,
          pairingReferences: closing
            ? [{ direction: 'next', address: closing.diagnostic.address }]
            : undefined,
          technicalNote: assumedClosing
            ? 'A hónap utolsó napi 5 jelölése a következő hónap első napján feltételezett 7-tel lett lezárva.'
            : undefined,
        },
      ),
    );
  };

  const processClosingSeven = (occurrence: MarkerOccurrence, index: number): void => {
    const entry = occurrenceEntry(occurrence);
    const preceding = index > 0 ? entries[index - 1] : previous;
    const start = preceding ? findStartingMarker(preceding) : undefined;
    const paired = start ? pairedShift(start, occurrence) : undefined;

    if (!paired || !start) {
      rows.push(
        issueRow(
          entry,
          'Hibás párosítás',
          preceding
            ? 'A 7 előtt nincs 17 vagy 5 a megelőző naptári napon.'
            : 'Az előző havi munkalap hiányzik; a 7 nem párosítható.',
        ),
      );
      return;
    }
    if (!paired.event) {
      rows.push(
        issueRow(
          entry,
          'Bizonytalan',
          `A kezdő ${start.normalizedMarker} betűszíne nem sorolható Parti vagy Esetszolgálathoz.`,
          {
            timeRule: paired.timeRule,
            pairingReferences: [{ direction: 'previous', address: start.diagnostic.address }],
          },
        ),
      );
      return;
    }

    const metadata = {
      timeRule: paired.timeRule,
      pairingReferences: [
        { direction: 'previous' as const, address: start.diagnostic.address },
      ],
    };
    if (index === 0) {
      events.push(paired.event);
      rows.push(
        rowForEvent(
          entry,
          paired.event,
          'Az előző hónap utolsó napi jelölésével párosítva.',
          metadata,
        ),
      );
    } else {
      rows.push(
        recognizedPairingRow(
          entry,
          paired.event,
          occurrence.entry.kind === 'double'
            ? 'A 7 az előző napi szolgálatot zárja; az adott napon 00:00–07:00 szolgálati szakasz.'
            : 'Az előző napi szolgálat befejező 7 jelölése.',
          metadata,
        ),
      );
    }
  };

  entries.forEach((sourceEntry, index) => {
    if (sourceEntry.kind === 'empty') return;
    if (sourceEntry.kind === 'invalid-date') {
      rows.push(
        issueRow(
          sourceEntry,
          'Hibás párosítás',
          'Érvénytelen naptári nap; nem exportálható.',
        ),
      );
      return;
    }

    const occurrences = markerOccurrences(sourceEntry);
    const isSevenFiveDay =
      sourceEntry.kind === 'double' &&
      occurrences.length === 2 &&
      occurrences[0]?.normalizedMarker === '7' &&
      occurrences[1]?.normalizedMarker === '5';
    if (sourceEntry.kind === 'double' && !isSevenFiveDay) {
      rows.push(
        issueRow(
          sourceEntry,
          'Hibás párosítás',
          'A napi cellacsoport két értéke nem támogatott kombináció.',
        ),
      );
      return;
    }

    for (const occurrence of occurrences) {
      const entry = occurrenceEntry(occurrence);
      const marker = occurrence.normalizedMarker;

      if (marker === '12') {
        const twelveKind = classifyTwelve(occurrence.diagnostic);
        const dailyPattern = dailyServicePatterns?.get(localDateKey(entry.date));
        const inferredCorrection =
          twelveKind === 'party' &&
          dailyPattern?.correction?.candidateAddress === occurrence.diagnostic.address
            ? dailyPattern.correction
            : undefined;
        const twelveConfig =
          inferredCorrection?.target === 'tenCar'
            ? {
                shiftType: 'Nappalos 10–22' as const,
                serviceCategory: '10-es kocsi' as const,
                start: '10:00',
                end: '22:00',
                note: inferredCorrection.explanation,
              }
            : inferredCorrection?.target === 'blue'
              ? {
                  shiftType: 'Nappalos 06–18' as const,
                  serviceCategory: '6-os kocsi' as const,
                  start: '06:00',
                  end: '18:00',
                  note: inferredCorrection.explanation,
                }
              : twelveKind === 'tenCar'
                ? {
                    shiftType: 'Nappalos 10–22' as const,
                    serviceCategory: '10-es kocsi' as const,
                    start: '10:00',
                    end: '22:00',
                    note: 'Zöld és aláhúzott 12 felismerve: 10-es kocsi.',
                  }
                : twelveKind === 'blue'
                  ? {
                      shiftType: 'Nappalos 06–18' as const,
                      serviceCategory: 'Nappalos 06–18' as const,
                      start: '06:00',
                      end: '18:00',
                      note: 'Kék 12 felismerve.',
                    }
                  : twelveKind === 'emergency'
                    ? {
                        shiftType: 'Nappalos 07–19' as const,
                        serviceCategory: 'Esetszolgálat' as const,
                        start: '07:00',
                        end: '19:00',
                        note: 'Piros 12 felismerve: Esetszolgálat.',
                      }
                    : twelveKind === 'party'
                      ? {
                          shiftType: 'Nappalos 07–19' as const,
                          serviceCategory: 'Parti szolgálat' as const,
                          start: '07:00',
                          end: '19:00',
                          note: 'Fekete 12 felismerve: Parti szolgálat.',
                        }
                      : undefined;

        if (!twelveConfig) {
          rows.push(
            issueRow(
              entry,
              'Bizonytalan',
              'A 12 betűszíne és aláhúzása nem sorolható megbízhatóan támogatott szolgálathoz.',
            ),
          );
          continue;
        }

        const resolvedTime = timeRange(
          localDateTime(entry.date, twelveConfig.start),
          localDateTime(entry.date, twelveConfig.end),
        );
        const inference: ServiceInference | undefined = inferredCorrection
          ? {
              source: 'daily-service-pattern',
              target: inferredCorrection.target,
              explanation: inferredCorrection.explanation,
              originalServiceCategory: 'Parti szolgálat',
              originalShiftType: 'Nappalos 07–19',
            }
          : undefined;
        const created = event(
          'OMSZ',
          twelveConfig.shiftType,
          twelveConfig.serviceCategory,
          resolvedTime,
          resolvedTime,
          inference,
        );
        events.push(created);
        rows.push(
          rowForEvent(entry, created, twelveConfig.note, {
            timeRule: inferredCorrection
              ? `12 + napi szolgálati összesítés → ${twelveConfig.start}–${twelveConfig.end}`
              : `12 → ${twelveConfig.start}–${twelveConfig.end}`,
            dailyInference:
              inferredCorrection && dailyPattern
                ? {
                    partyTwentyFourHourPresent:
                      dailyPattern.partyTwentyFourHourCount === 1,
                    blueTwelvePresent: dailyPattern.blueTwelveCount > 0,
                    tenCarTwelvePresent: dailyPattern.tenCarTwelveCount > 0,
                    blackTwelveCandidateCount: dailyPattern.blackTwelveCandidateCount,
                    correctionApplied: true,
                    originalServiceCategory: 'Parti szolgálat',
                    originalShiftType: 'Nappalos 07–19',
                    finalServiceCategory: twelveConfig.serviceCategory,
                    finalShiftType: twelveConfig.shiftType,
                    finalTime: resolvedTime,
                  }
                : undefined,
          }),
        );
        continue;
      }

      if (marker === 'kmr') {
        const created = event(
          'KMR',
          'KMR',
          'KMR',
          timeRange(
            localDateTime(entry.date, '05:00'),
            localDateTime(addDays(entry.date, 1), '01:00'),
          ),
        );
        events.push(created);
        rows.push(
          rowForEvent(
            entry,
            created,
            'Napi KMR-bejegyzés a naptári nap cellacsoportjában.',
            { timeRule: 'KMR → 05:00–másnap 01:00' },
          ),
        );
        continue;
      }

      if (marker === '17' || marker === '5') {
        processStartingOccurrence(occurrence, index);
        continue;
      }

      if (marker === '7') {
        processClosingSeven(occurrence, index);
        continue;
      }

      if (NON_SERVICE_MARKERS.has(marker)) {
        rows.push(issueRow(entry, 'Kizárva', 'Távolléti vagy nem szolgálati jelölés.'));
        continue;
      }

      rows.push(issueRow(entry, 'Bizonytalan', 'Ismeretlen napi jelölés; nem exportálható.'));
    }
  });

  const uncertain = rows.filter((row) => row.status === 'Bizonytalan').length;
  const invalid = rows.filter((row) => row.status === 'Hibás párosítás').length;
  return {
    rows: rows.sort((first, second) =>
      localDateKey(first.date).localeCompare(localDateKey(second.date)),
    ),
    events,
    summary: {
      recognized: events.length,
      omsz: events.filter((item) => item.summary === 'OMSZ').length,
      kmr: events.filter((item) => item.summary === 'KMR').length,
      uncertain,
      invalid,
      exportable: events.length,
    },
  };
}
