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

export interface ResolvedStyle {
  styleId?: number;
  fillColor?: string;
  fontColor?: string;
  italic: boolean;
  bold: boolean;
}

export interface CellDiagnostic extends ResolvedStyle {
  address: string;
  rawValue: string;
  displayedText: string;
  isMerged: boolean;
  mergeMaster?: string;
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
  'Nappalos 06–18' | 'Nappalos 10–22' | '24 órás szolgálat' | 'Éjszakai szolgálat' | 'KMR';

export interface EventTimeRange {
  start: string;
  end: string;
}

export interface CalendarEvent {
  id: string;
  summary: 'OMSZ' | 'KMR';
  shiftType: ShiftType;
  shiftTime: EventTimeRange;
  calendarTime: EventTimeRange;
  timeZone: 'Europe/Budapest';
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

export interface ReviewRow {
  id: string;
  date: LocalDate;
  marker: string;
  shiftType?: ShiftType;
  summary?: 'OMSZ' | 'KMR';
  status: ReviewStatus;
  note: string;
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
