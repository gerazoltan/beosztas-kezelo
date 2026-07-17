import { HUNGARIAN_MONTHS, type MonthSheet } from '../domain/types';

type MonthIdentity = Pick<MonthSheet, 'year' | 'month' | 'sheetName'>;

export function monthOptionValue(month: MonthIdentity): string {
  return `${month.year}-${month.month}-${month.sheetName}`;
}

export function monthOptionLabel(month: Pick<MonthIdentity, 'year' | 'month'>): string {
  const monthName = HUNGARIAN_MONTHS[month.month - 1];
  if (!monthName) throw new RangeError(`Érvénytelen hónapszám: ${month.month}`);
  return `${month.year}. ${monthName}`;
}
