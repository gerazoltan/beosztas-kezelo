import type {
  CalendarEvent,
  DayEntry,
  EventTimeRange,
  FillCategory,
  LegendStyles,
  ResolvedStyle,
  ReviewRow,
  ScheduleResult,
  ShiftType,
} from '../domain/types';
import { addDays, localDateKey, localDateTime } from './dates';
import { isBlue, isGreen, isWhite } from '../excel/colors';

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
  shiftTime: EventTimeRange,
  calendarTime: EventTimeRange = shiftTime,
): CalendarEvent {
  return {
    id: stableHash(`${summary}|${shiftType}|${calendarTime.start}|${calendarTime.end}`),
    summary,
    shiftType,
    shiftTime,
    calendarTime,
    timeZone: 'Europe/Budapest',
  };
}

function timeRange(start: string, end: string): EventTimeRange {
  return { start, end };
}

function sameStyle(first: ResolvedStyle, second: ResolvedStyle): boolean {
  return (
    first.fillColor === second.fillColor &&
    first.hasVisibleFill === second.hasVisibleFill &&
    first.fontColor === second.fontColor &&
    first.italic === second.italic &&
    first.bold === second.bold
  );
}

export type TwelveKind = 'blue' | 'white' | 'green' | 'unknown';

export function classifyTwelve(style: ResolvedStyle | undefined, legend: LegendStyles): TwelveKind {
  if (!style) return 'unknown';
  if (legend.green12.some((reference) => sameStyle(reference, style))) return 'green';
  if (legend.blue12.some((reference) => sameStyle(reference, style))) return 'blue';
  if (style.italic && isGreen(style.fontColor)) return 'green';
  if (!style.italic && isBlue(style.fillColor)) return 'blue';
  const hasVisibleFill = style.hasVisibleFill ?? style.fillColor !== undefined;
  if (!hasVisibleFill || isWhite(style.fillColor)) return 'white';
  return 'unknown';
}

function fillCategoryForTwelve(kind: TwelveKind, style: ResolvedStyle | undefined): FillCategory {
  if (kind === 'blue' || kind === 'green') return kind;
  if (kind === 'white') return style?.hasVisibleFill ? 'white' : 'noFill';
  return 'unsupported';
}

function withTwelveFillCategory(entry: DayEntry, kind: TwelveKind): DayEntry {
  const selected = entry.selectedDiagnostic;
  if (!selected) return entry;
  const selectedDiagnostic = {
    ...selected,
    fillCategory: fillCategoryForTwelve(kind, selected),
  };
  return {
    ...entry,
    selectedDiagnostic,
    diagnostics: entry.diagnostics.map((item) =>
      item.address === selected.address ? selectedDiagnostic : item,
    ),
  };
}

function rowForEvent(entry: DayEntry, created: CalendarEvent, note: string): ReviewRow {
  return {
    id: `${localDateKey(entry.date)}-${created.id}`,
    date: entry.date,
    marker: entry.marker,
    shiftType: created.shiftType,
    summary: created.summary,
    status: 'Exportálható',
    note,
    technicalNote:
      created.shiftType === '24 órás szolgálat'
        ? 'A naptáresemény befejezése 06:59 a jobb naptári elkülönítés érdekében.'
        : undefined,
    diagnostics: entry.diagnostics,
    event: created,
  };
}

function issueRow(entry: DayEntry, status: ReviewRow['status'], note: string): ReviewRow {
  return {
    id: `${localDateKey(entry.date)}-${status}-${entry.group.startColumn}`,
    date: entry.date,
    marker: entry.marker,
    status,
    note,
    diagnostics: entry.diagnostics,
  };
}

function isConsecutive(first: DayEntry, second: DayEntry): boolean {
  return localDateKey(addDays(first.date, 1)) === localDateKey(second.date);
}

function pairedEvent(startEntry: DayEntry, endEntry: DayEntry): CalendarEvent | undefined {
  if (endEntry.normalizedMarker !== '7' || !isConsecutive(startEntry, endEntry)) return undefined;
  if (startEntry.normalizedMarker === '17') {
    return event(
      'OMSZ',
      '24 órás szolgálat',
      timeRange(localDateTime(startEntry.date, '07:00'), localDateTime(endEntry.date, '07:00')),
      timeRange(localDateTime(startEntry.date, '07:00'), localDateTime(endEntry.date, '06:59')),
    );
  }
  if (startEntry.normalizedMarker === '5') {
    return event(
      'OMSZ',
      'Éjszakai szolgálat',
      timeRange(localDateTime(startEntry.date, '19:00'), localDateTime(endEntry.date, '07:00')),
    );
  }
  return undefined;
}

export interface InterpretOptions {
  legend: LegendStyles;
  previous?: DayEntry;
  next?: DayEntry;
}

export function interpretSchedule(
  entries: DayEntry[],
  { legend, previous, next }: InterpretOptions,
): ScheduleResult {
  const rows: ReviewRow[] = [];
  const events: CalendarEvent[] = [];
  const consumed = new Set<number>();

  entries.forEach((entry, index) => {
    if (consumed.has(index) || entry.kind === 'empty') return;
    if (entry.kind === 'invalid-date') {
      rows.push(issueRow(entry, 'Hibás párosítás', 'Érvénytelen naptári nap; nem exportálható.'));
      return;
    }
    if (entry.kind === 'double') {
      rows.push(
        issueRow(entry, 'Hibás párosítás', 'A napi cellacsoport mindkét oldalán van érték.'),
      );
      return;
    }

    const marker = entry.normalizedMarker;
    if (marker === '12') {
      const twelveKind = classifyTwelve(entry.selectedDiagnostic, legend);
      const categorizedEntry = withTwelveFillCategory(entry, twelveKind);
      if (twelveKind === 'unknown') {
        rows.push(
          issueRow(
            categorizedEntry,
            'Bizonytalan',
            'A 12 látható formázása nem sorolható megbízhatóan a támogatott szolgálatokhoz.',
          ),
        );
        return;
      }
      const twelveConfig =
        twelveKind === 'green'
          ? {
              shiftType: 'Nappalos 10–22' as const,
              start: '10:00',
              end: '22:00',
              note: 'Zöld 12 felismerve.',
            }
          : twelveKind === 'white'
            ? {
                shiftType: 'Nappalos 07–19' as const,
                start: '07:00',
                end: '19:00',
                note: 'Fehér vagy kitöltés nélküli 12 felismerve.',
              }
            : {
                shiftType: 'Nappalos 06–18' as const,
                start: '06:00',
                end: '18:00',
                note: 'Kék 12 felismerve.',
              };
      const created = event(
        'OMSZ',
        twelveConfig.shiftType,
        timeRange(
          localDateTime(entry.date, twelveConfig.start),
          localDateTime(entry.date, twelveConfig.end),
        ),
      );
      events.push(created);
      rows.push(rowForEvent(categorizedEntry, created, twelveConfig.note));
      return;
    }

    if (marker === 'kmr') {
      const created = event(
        'KMR',
        'KMR',
        timeRange(
          localDateTime(entry.date, '05:00'),
          localDateTime(addDays(entry.date, 1), '01:00'),
        ),
      );
      events.push(created);
      rows.push(rowForEvent(entry, created, 'Napi KMR-bejegyzés a naptári nap cellacsoportjában.'));
      return;
    }

    if (marker === '17' || marker === '5') {
      const following = entries[index + 1] ?? next;
      const created = following ? pairedEvent(entry, following) : undefined;
      if (!created) {
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
      events.push(created);
      rows.push(rowForEvent(entry, created, `${marker}–7 pár egyetlen szolgálatként felismerve.`));
      if (following && entries[index + 1] === following) {
        consumed.add(index + 1);
        rows.push(
          issueRow(following, 'Felismerve', 'Az előző napi szolgálat befejező 7 jelölése.'),
        );
      }
      return;
    }

    if (marker === '7') {
      const preceding = index > 0 ? entries[index - 1] : previous;
      const created = preceding ? pairedEvent(preceding, entry) : undefined;
      if (created && index === 0) {
        events.push(created);
        rows.push(rowForEvent(entry, created, 'Az előző hónap utolsó napi jelölésével párosítva.'));
      } else {
        rows.push(
          issueRow(
            entry,
            created ? 'Felismerve' : 'Hibás párosítás',
            created
              ? 'Az előző napi szolgálat befejező 7 jelölése.'
              : preceding
                ? 'A 7 előtt nincs 17 vagy 5 a megelőző naptári napon.'
                : 'Az előző havi munkalap hiányzik; a 7 nem párosítható.',
          ),
        );
      }
      return;
    }

    if (NON_SERVICE_MARKERS.has(marker)) {
      rows.push(issueRow(entry, 'Kizárva', 'Távolléti vagy nem szolgálati jelölés.'));
      return;
    }

    rows.push(issueRow(entry, 'Bizonytalan', 'Ismeretlen napi jelölés; nem exportálható.'));
  });

  const uncertain = rows.filter((row) => row.status === 'Bizonytalan').length;
  const invalid = rows.filter((row) => row.status === 'Hibás párosítás').length;
  return {
    rows: rows.sort((a, b) => localDateKey(a.date).localeCompare(localDateKey(b.date))),
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
