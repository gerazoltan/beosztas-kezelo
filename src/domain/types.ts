import type ExcelJS from 'exceljs';

export const HUNGARIAN_MONTHS = [
  'január',
  'február',
  'március',
  'április',
  'május',
  'június',
  'július',
  'augusztus',
  'szeptember',
  'október',
  'november',
  'december',
] as const;

export type FillCategory = 'noFill' | 'white' | 'blue' | 'green' | 'unsupported';

export interface ResolvedStyle {
  styleId?: number;
  fillType?: string;
  fillPatternType?: string;
  fillForegroundRaw?: string;
  fillBackgroundRaw?: string;
  fillColor?: string;
  hasVisibleFill?: boolean;
  fillCategory?: FillCategory;
  fontColorRaw?: string;
  fontColor?: string;
  underline: boolean;
  italic: boolean;
  bold: boolean;
}

export interface CellDiagnostic extends ResolvedStyle {
  address: string;
  rawValue: string;
  displayedText: string;
  isMerged: boolean;
  mergeMaster?: string;
  positionInDayGroup: number;
}

export interface DayGroup {
  day: number;
  startColumn: number;
  endColumn: number;
  valid: boolean;
  validationMessage?: string;
}

export interface EmployeeRef {
  name: string;
  normalizedName: string;
  rows: number[];
}

export interface LegendStyles {
  blue12: ResolvedStyle[];
  green12: ResolvedStyle[];
}

export interface MonthSheet {
  sheetName: string;
  month: number;
  year: number;
  headerRow: number;
  nameColumn: number;
  dayGroups: DayGroup[];
  employees: EmployeeRef[];
  warnings: string[];
  legendStyles: LegendStyles;
}

export interface OoxmlMetadata {
  themeColors: string[];
  styleIds: Map<string, number>;
}

export interface WorkbookSession {
  fileName: string;
  workbook: ExcelJS.Workbook;
  months: MonthSheet[];
  ooxml: OoxmlMetadata;
  warnings: string[];
}

export type EntryKind = 'empty' | 'single' | 'double' | 'invalid-date';

export interface DayEntry {
  date: LocalDate;
  group: DayGroup;
  kind: EntryKind;
  marker: string;
  normalizedMarker: string;
  diagnostics: CellDiagnostic[];
  selectedDiagnostic?: CellDiagnostic;
}

export interface LocalDate {
  year: number;
  month: number;
  day: number;
}

export type ShiftType =
  | 'Nappalos 06–18'
  | 'Nappalos 07–19'
  | 'Nappalos 10–22'
  | '24 órás szolgálat'
  | 'Éjszakai szolgálat'
  | 'Előző hónapról áthúzódó szolgálat'
  | 'KMR';

export type ServiceCategory =
  | 'Parti szolgálat'
  | 'Esetszolgálat'
  | '6-os kocsi'
  | '10-es kocsi'
  | 'Nappalos 06–18'
  | 'KMR';

export type InferredTwelveKind = 'blue' | 'tenCar';

export interface DailyServicePattern {
  date: LocalDate;
  partyTwentyFourHourCount: number;
  emergencyTwentyFourHourCount: number;
  blueTwelveCount: number;
  tenCarTwelveCount: number;
  blackTwelveCandidateCount: number;
  greenSeventeenCandidateCount: number;
  conflictingSeventeenCount: number;
  conflictingServiceMarkerCount: number;
  correction?: {
    candidateAddress: string;
    target: InferredTwelveKind;
    explanation: string;
  };
  seventeenCorrection?: {
    candidateAddress: string;
    target: 'party' | 'emergency';
    explanation: string;
  };
}

export interface ServiceInference {
  source: 'daily-service-pattern';
  target: InferredTwelveKind | 'party' | 'emergency';
  explanation: string;
  originalServiceCategory?: ServiceCategory | 'Nem meghatározható';
  originalShiftType?: ShiftType;
}

export interface EventTimeRange {
  start: string;
  end: string;
}

export interface DailyInferenceTechnicalDetails {
  partyTwentyFourHourPresent: boolean;
  blueTwelvePresent: boolean;
  tenCarTwelvePresent: boolean;
  blackTwelveCandidateCount: number;
  correctionApplied: boolean;
  originalServiceCategory: 'Parti szolgálat';
  originalShiftType: 'Nappalos 07–19';
  finalServiceCategory: ServiceCategory;
  finalShiftType: ShiftType;
  finalTime: EventTimeRange;
}

export interface CalendarEvent {
  id: string;
  summary: 'OMSZ' | 'KMR';
  shiftType: ShiftType;
  serviceCategory: ServiceCategory;
  shiftTime: EventTimeRange;
  calendarTime: EventTimeRange;
  timeZone: 'Europe/Budapest';
  inference?: ServiceInference;
  specialKind?: 'previous-month-carryover-partial';
}

export type ReviewStatus =
  | 'Exportálható'
  | 'Felismerve'
  | 'Kizárva'
  | 'Bizonytalan'
  | 'Hibás párosítás'
  | 'Létrehozás folyamatban'
  | 'Már szerepel a naptárban'
  | 'Létrehozva'
  | 'Sikertelen';

export interface GoogleEventState {
  status: Extract<
    ReviewStatus,
    'Létrehozás folyamatban' | 'Már szerepel a naptárban' | 'Létrehozva' | 'Sikertelen'
  >;
  message: string;
  technicalDetails?: string;
}

export interface ServiceResolutionTechnicalDetails {
  originalServiceCategory: ServiceCategory | 'Nem meghatározható';
  finalServiceCategory?: ServiceCategory;
  formattingCorrectionApplied: boolean;
  dailyInferenceApplied: boolean;
  assumedBoundaryPairing: boolean;
  pairingSource?: 'actual' | 'assumed' | 'previous-month-carryover';
  pairingCell?: string;
  finalShiftTime?: EventTimeRange;
  finalCalendarTime?: EventTimeRange;
}

export interface ReviewRow {
  id: string;
  date: LocalDate;
  marker: string;
  shiftType?: ShiftType;
  serviceCategory?: ServiceCategory;
  summary?: 'OMSZ' | 'KMR';
  status: ReviewStatus;
  note: string;
  timeRule?: string;
  pairingReferences?: Array<{
    direction: 'previous' | 'next';
    address: string;
  }>;
  dailyInference?: DailyInferenceTechnicalDetails;
  serviceResolution?: ServiceResolutionTechnicalDetails;
  technicalNote?: string;
  diagnostics: CellDiagnostic[];
  event?: CalendarEvent;
}

export interface ScheduleResult {
  rows: ReviewRow[];
  events: CalendarEvent[];
  summary: {
    recognized: number;
    omsz: number;
    kmr: number;
    uncertain: number;
    invalid: number;
    exportable: number;
  };
}
