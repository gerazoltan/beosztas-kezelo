import type ExcelJS from 'exceljs';
import type {
  CellDiagnostic,
  DayEntry,
  EmployeeRef,
  MonthSheet,
  WorkbookSession,
} from '../domain/types';
import { AppError } from '../domain/errors';
import { localDateKey } from '../services/dates';
import { normalizeMarker } from '../utils/normalize';
import { resolveCellStyle } from './colors';
import { displayedCellText } from './cellValues';

function rawValue(cell: ExcelJS.Cell): string {
  const value = cell.value;
  if (value === null || value === undefined) return '';
  if (typeof value === 'object') {
    if ('result' in value && value.result !== undefined) {
      const result = value.result;
      return typeof result === 'string' || typeof result === 'number' || typeof result === 'boolean'
        ? String(result)
        : (JSON.stringify(result) ?? '');
    }
    if ('richText' in value) return value.richText.map((part) => part.text).join('');
    return JSON.stringify(value);
  }
  return String(value);
}

function diagnostic(
  session: WorkbookSession,
  worksheet: ExcelJS.Worksheet,
  cell: ExcelJS.Cell,
  positionInDayGroup: number,
): CellDiagnostic {
  const master = cell.isMerged ? cell.master : cell;
  const address = master.address;
  return {
    address,
    rawValue: rawValue(master),
    displayedText: displayedCellText(master),
    isMerged: cell.isMerged,
    mergeMaster: cell.isMerged ? address : undefined,
    positionInDayGroup,
    ...resolveCellStyle(
      master,
      session.ooxml.themeColors,
      session.ooxml.styleIds.get(`${worksheet.name}!${address}`),
    ),
  };
}

function findEmployee(
  month: MonthSheet,
  normalizedName: string,
  rowOverride?: number,
): EmployeeRef {
  const employee = month.employees.find((item) => item.normalizedName === normalizedName);
  if (!employee) throw new AppError('EMPLOYEE_NOT_FOUND');
  if (rowOverride !== undefined) {
    if (!employee.rows.includes(rowOverride)) throw new AppError('EMPLOYEE_NOT_FOUND');
    return { ...employee, rows: [rowOverride] };
  }
  if (employee.rows.length !== 1) {
    throw new AppError(
      'EMPLOYEE_DUPLICATE',
      `${month.sheetName}: ${employee.rows.join(', ')}. sor`,
    );
  }
  return employee;
}

export function readMonthEntries(
  session: WorkbookSession,
  month: MonthSheet,
  normalizedName: string,
  rowOverride?: number,
): DayEntry[] {
  const worksheet = session.workbook.getWorksheet(month.sheetName);
  if (!worksheet) throw new AppError('NO_MONTH_SHEET', month.sheetName);
  const employee = findEmployee(month, normalizedName, rowOverride);
  const rowNumber = employee.rows[0] as number;
  return month.dayGroups.map((group) => {
    const allDiagnostics = new Map<string, CellDiagnostic>();
    for (let column = group.startColumn; column <= group.endColumn; column += 1) {
      const cell = worksheet.getCell(rowNumber, column);
      const item = diagnostic(session, worksheet, cell, column - group.startColumn + 1);
      if (!allDiagnostics.has(item.address)) allDiagnostics.set(item.address, item);
    }
    const diagnostics = [...allDiagnostics.values()];
    const nonEmpty = diagnostics.filter((item) => item.displayedText.trim() !== '');
    const selected = nonEmpty.length === 1 ? nonEmpty[0] : undefined;
    return {
      date: { year: month.year, month: month.month, day: group.day },
      group,
      kind: !group.valid
        ? 'invalid-date'
        : nonEmpty.length === 0
          ? 'empty'
          : nonEmpty.length === 1
            ? 'single'
            : 'double',
      marker:
        selected?.displayedText.trim() ?? nonEmpty.map((item) => item.displayedText).join(' / '),
      normalizedMarker: normalizeMarker(selected?.displayedText ?? ''),
      diagnostics,
      selectedDiagnostic: selected,
    };
  });
}

export interface EmployeeScheduleEntries {
  current: DayEntry[];
  previous?: DayEntry;
  next?: DayEntry;
}

export interface WorksheetEmployeeScheduleEntries extends EmployeeScheduleEntries {
  normalizedName: string;
  row: number;
}

function adjacentMonth(
  months: MonthSheet[],
  year: number,
  month: number,
  offset: number,
): MonthSheet | undefined {
  const date = new Date(Date.UTC(year, month - 1 + offset, 1));
  return months.find(
    (item) => item.year === date.getUTCFullYear() && item.month === date.getUTCMonth() + 1,
  );
}

function safelyReadBoundary(
  session: WorkbookSession,
  month: MonthSheet | undefined,
  normalizedName: string,
  edge: 'first' | 'last',
): DayEntry | undefined {
  if (!month) return undefined;
  try {
    const entries = readMonthEntries(session, month, normalizedName).filter(
      (entry) => entry.group.valid,
    );
    return edge === 'first' ? entries[0] : entries.at(-1);
  } catch {
    return undefined;
  }
}

export function readEmployeeScheduleEntries(
  session: WorkbookSession,
  month: MonthSheet,
  normalizedName: string,
  rowOverride?: number,
): EmployeeScheduleEntries {
  const current = readMonthEntries(session, month, normalizedName, rowOverride);
  const previousMonth = adjacentMonth(session.months, month.year, month.month, -1);
  const nextMonth = adjacentMonth(session.months, month.year, month.month, 1);
  const previous = safelyReadBoundary(session, previousMonth, normalizedName, 'last');
  const next = safelyReadBoundary(session, nextMonth, normalizedName, 'first');
  if (previous && localDateKey(previous.date) >= localDateKey(current[0]?.date ?? previous.date))
    return { current, next };
  return { current, previous, next };
}

export function readWorksheetScheduleEntries(
  session: WorkbookSession,
  month: MonthSheet,
): WorksheetEmployeeScheduleEntries[] {
  return month.employees.flatMap((employee) =>
    employee.rows.map((row) => ({
      normalizedName: employee.normalizedName,
      row,
      ...readEmployeeScheduleEntries(session, month, employee.normalizedName, row),
    })),
  );
}
