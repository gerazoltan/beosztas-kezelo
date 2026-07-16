import ExcelJS from 'exceljs';
import type {
  DayGroup,
  EmployeeRef,
  LegendStyles,
  MonthSheet,
  ResolvedStyle,
  WorkbookSession,
} from '../domain/types';
import { HUNGARIAN_MONTHS } from '../domain/types';
import { AppError, toAppError } from '../domain/errors';
import { daysInMonth } from '../services/dates';
import { monthFromText, normalizeLookup, normalizeWhitespace } from '../utils/normalize';
import { resolveCellStyle } from './colors';
import { displayedCellText } from './cellValues';
import { extractOoxmlMetadata } from './ooxml';

interface HeaderCandidate {
  row: number;
  nameColumn: number;
  days: Array<{ day: number; column: number }>;
}

function cellText(cell: ExcelJS.Cell): string {
  return normalizeWhitespace(displayedCellText(cell));
}

function meaningfulCells(row: ExcelJS.Row): ExcelJS.Cell[] {
  const cells: ExcelJS.Cell[] = [];
  row.eachCell({ includeEmpty: false }, (cell) => {
    if (cellText(cell) !== '') cells.push(cell);
  });
  return cells;
}

function findHeader(worksheet: ExcelJS.Worksheet): HeaderCandidate {
  const candidates: HeaderCandidate[] = [];
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    const cells = meaningfulCells(row);
    const names = cells.filter((cell) => normalizeLookup(cellText(cell)) === 'név');
    if (names.length > 1)
      throw new AppError('NAME_HEADER_MULTIPLE', `${worksheet.name}, ${rowNumber}. sor`);
    if (names.length !== 1) return;
    const nameColumn = names[0]?.col;
    if (typeof nameColumn !== 'number') return;
    const days = cells
      .map((cell) => ({ day: Number(cellText(cell)), column: Number(cell.col) }))
      .filter(
        (item) =>
          item.column > nameColumn && Number.isInteger(item.day) && item.day >= 1 && item.day <= 31,
      )
      .sort((a, b) => a.column - b.column);
    const consecutive: typeof days = [];
    for (const item of days) {
      if (consecutive.length === 0) {
        if (item.day === 1) consecutive.push(item);
        continue;
      }
      const lastDay = consecutive.at(-1)?.day ?? 0;
      if (item.day === lastDay) continue;
      if (item.day === lastDay + 1) consecutive.push(item);
      else break;
    }
    if (consecutive.length >= 1) {
      candidates.push({ row: rowNumber, nameColumn, days: consecutive });
    }
  });
  if (candidates.length === 0) throw new AppError('NAME_HEADER_MISSING', worksheet.name);
  if (candidates.length > 1) throw new AppError('NAME_HEADER_MULTIPLE', worksheet.name);
  return candidates[0] as HeaderCandidate;
}

function monthAndYearFromHeader(
  worksheet: ExcelJS.Worksheet,
  headerRow: number,
  sheetMonth: number,
): { month: number; year: number } {
  const headerTexts: string[] = [];
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber > headerRow) return;
    for (const cell of meaningfulCells(row)) headerTexts.push(cellText(cell));
  });
  const monthValues = new Set(
    headerTexts.map(monthFromText).filter((value): value is number => value !== undefined),
  );
  if (monthValues.size !== 1 || !monthValues.has(sheetMonth)) {
    throw new AppError('AMBIGUOUS_MONTH', worksheet.name);
  }
  const years = new Set<number>();
  for (const text of headerTexts) {
    for (const match of text.matchAll(/\b(20\d{2})\b/g)) years.add(Number(match[1]));
  }
  if (years.size !== 1) throw new AppError('AMBIGUOUS_YEAR', worksheet.name);
  return { month: sheetMonth, year: [...years][0] as number };
}

function createDayGroups(
  candidate: HeaderCandidate,
  year: number,
  month: number,
): { groups: DayGroup[]; warnings: string[] } {
  const widths = candidate.days.slice(0, -1).map((item, index) => {
    const next = candidate.days[index + 1];
    return next ? next.column - item.column : 2;
  });
  const usualWidth =
    widths.length > 0
      ? Math.max(1, Math.round(widths.reduce((a, b) => a + b, 0) / widths.length))
      : 2;
  const maximumDay = daysInMonth(year, month);
  const warnings: string[] = [];
  const foundDays = new Set(candidate.days.map((item) => item.day));
  const missingDays = Array.from({ length: maximumDay }, (_, index) => index + 1).filter(
    (day) => !foundDays.has(day),
  );
  if (missingDays.length > 0) {
    warnings.push(`Hiányzó napfejlécek: ${missingDays.join(', ')}.`);
  }
  const groups = candidate.days.map((item, index) => {
    const next = candidate.days[index + 1];
    const endColumn = next ? next.column - 1 : item.column + usualWidth - 1;
    const valid = item.day <= maximumDay;
    if (!valid)
      warnings.push(
        `${item.day}. nap nem létezik ${year}. ${HUNGARIAN_MONTHS[month - 1]} hónapban.`,
      );
    if (endColumn - item.column + 1 !== 2) {
      warnings.push(`${item.day}. naphoz ${endColumn - item.column + 1} fizikai oszlop tartozik.`);
    }
    return {
      day: item.day,
      startColumn: item.column,
      endColumn,
      valid,
      validationMessage: valid ? undefined : 'Érvénytelen naptári nap; nem exportálható.',
    };
  });
  return { groups, warnings };
}

function readEmployees(
  worksheet: ExcelJS.Worksheet,
  headerRow: number,
  nameColumn: number,
): EmployeeRef[] {
  const byName = new Map<string, EmployeeRef>();
  const lastMeaningfulRow = Math.min(worksheet.rowCount, 2000);
  for (let rowNumber = headerRow + 1; rowNumber <= lastMeaningfulRow; rowNumber += 1) {
    const name = normalizeWhitespace(displayedCellText(worksheet.getCell(rowNumber, nameColumn)));
    const lookup = normalizeLookup(name);
    if (lookup === 'összesen') break;
    if (!name || /^(név|ápoló|gkv|kmr)$/iu.test(lookup)) continue;
    const existing = byName.get(lookup);
    if (existing) existing.rows.push(rowNumber);
    else byName.set(lookup, { name, normalizedName: lookup, rows: [rowNumber] });
  }
  return [...byName.values()].sort((a, b) => a.name.localeCompare(b.name, 'hu-HU'));
}

function styleEquals(first: ResolvedStyle, second: ResolvedStyle): boolean {
  return (
    first.fillColor === second.fillColor &&
    first.fontColor === second.fontColor &&
    first.italic === second.italic &&
    first.bold === second.bold
  );
}

function readLegendStyles(
  worksheet: ExcelJS.Worksheet,
  themeColors: string[],
  styleIds: Map<string, number>,
): LegendStyles {
  const result: LegendStyles = { blue12: [], green12: [] };
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (normalizeLookup(cellText(cell)) !== '12') return;
      const nearby: string[] = [];
      for (let offset = -4; offset <= 4; offset += 1) {
        const column = Number(cell.col) + offset;
        if (offset !== 0 && column >= 1)
          nearby.push(normalizeLookup(cellText(row.getCell(column))));
      }
      const label = nearby.find((text) => text.includes('nappalos')) ?? '';
      const style = resolveCellStyle(
        cell,
        themeColors,
        styleIds.get(`${worksheet.name}!${cell.address}`),
      );
      const destination = /06\D*18|6\D*18/u.test(label)
        ? result.blue12
        : /10\D*22/u.test(label)
          ? result.green12
          : undefined;
      if (destination && !destination.some((entry) => styleEquals(entry, style)))
        destination.push(style);
    });
  });
  return result;
}

function parseMonthSheet(
  worksheet: ExcelJS.Worksheet,
  sheetMonth: number,
  themeColors: string[],
  styleIds: Map<string, number>,
): MonthSheet {
  const header = findHeader(worksheet);
  const { month, year } = monthAndYearFromHeader(worksheet, header.row, sheetMonth);
  const { groups, warnings } = createDayGroups(header, year, month);
  const employees = readEmployees(worksheet, header.row, header.nameColumn);
  if (employees.length === 0) throw new AppError('NO_EMPLOYEE', worksheet.name);
  return {
    sheetName: worksheet.name,
    month,
    year,
    headerRow: header.row,
    nameColumn: header.nameColumn,
    dayGroups: groups,
    employees,
    warnings,
    legendStyles: readLegendStyles(worksheet, themeColors, styleIds),
  };
}

export function validateExcelFileName(fileName: string): void {
  if (!/\.(xlsx|xls)$/iu.test(fileName)) throw new AppError('INVALID_FILE_TYPE');
  if (/\.xls$/iu.test(fileName)) throw new AppError('LEGACY_XLS');
}

export async function parseWorkbook(
  buffer: ArrayBuffer,
  fileName: string,
): Promise<WorkbookSession> {
  validateExcelFileName(fileName);
  const signature = new Uint8Array(buffer, 0, Math.min(buffer.byteLength, 8));
  if ([0xd0, 0xcf, 0x11, 0xe0].every((value, index) => signature[index] === value)) {
    throw new AppError(
      'PROTECTED_WORKBOOK',
      'OLE Compound File fejléc észlelve .xlsx kiterjesztéssel.',
    );
  }
  try {
    const [ooxml, workbook] = await Promise.all([
      extractOoxmlMetadata(buffer.slice(0)),
      (async () => {
        const instance = new ExcelJS.Workbook();
        await instance.xlsx.load(buffer.slice(0));
        return instance;
      })(),
    ]);
    const months: MonthSheet[] = [];
    const monthErrors: AppError[] = [];
    const warnings: string[] = [];
    for (const worksheet of workbook.worksheets) {
      const normalizedName = normalizeLookup(worksheet.name);
      if (/^munka\s*\d*$/iu.test(normalizedName)) continue;
      const sheetMonth = monthFromText(normalizedName);
      if (!sheetMonth) continue;
      try {
        months.push(parseMonthSheet(worksheet, sheetMonth, ooxml.themeColors, ooxml.styleIds));
      } catch (error) {
        const appError = error instanceof AppError ? error : toAppError(error);
        monthErrors.push(appError);
        warnings.push(
          `${worksheet.name}: ${appError.message}${appError.technicalDetails ? ` (${appError.technicalDetails})` : ''}`,
        );
      }
    }
    if (months.length === 0) {
      const firstError = monthErrors[0];
      if (firstError && monthErrors.every((item) => item.code === firstError.code)) {
        throw new AppError(firstError.code, warnings.join(' | '));
      }
      throw new AppError('NO_MONTH_SHEET', warnings.join(' | '));
    }
    months.sort((a, b) => a.year - b.year || a.month - b.month);
    return { fileName, workbook, months, ooxml, warnings };
  } catch (error) {
    throw toAppError(error);
  }
}

export interface DefaultMonthSelection {
  month: MonthSheet;
  usedFallback: boolean;
}

export function chooseDefaultMonth(months: MonthSheet[], now = new Date()): DefaultMonthSelection {
  const next = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const preferred = months.find(
    (month) => month.year === next.getFullYear() && month.month === next.getMonth() + 1,
  );
  const viable = months.filter(
    (month) => month.employees.length > 0 && month.dayGroups.some((group) => group.valid),
  );
  const selected = preferred ?? viable[0];
  if (!selected) throw new AppError('NO_MONTH_SHEET');
  return { month: selected, usedFallback: preferred === undefined };
}
