import type { LocalDate } from '../domain/types';

export const BUDAPEST_TIME_ZONE = 'Europe/Budapest' as const;

export function daysInMonth(year: number, month: number): number {
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

export function addDays(date: LocalDate, amount: number): LocalDate {
  const value = new Date(Date.UTC(date.year, date.month - 1, date.day + amount));
  return {
    year: value.getUTCFullYear(),
    month: value.getUTCMonth() + 1,
    day: value.getUTCDate(),
  };
}

export function localDateKey(date: LocalDate): string {
  return `${String(date.year).padStart(4, '0')}-${String(date.month).padStart(2, '0')}-${String(date.day).padStart(2, '0')}`;
}

export function localDateTime(date: LocalDate, time: string): string {
  return `${localDateKey(date)}T${time}:00`;
}

export function formatHungarianDate(date: LocalDate): string {
  return new Intl.DateTimeFormat('hu-HU', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).format(new Date(Date.UTC(date.year, date.month - 1, date.day, 12)));
}

export function weekdayHungarian(date: LocalDate): string {
  return new Intl.DateTimeFormat('hu-HU', { weekday: 'long', timeZone: 'UTC' }).format(
    new Date(Date.UTC(date.year, date.month - 1, date.day, 12)),
  );
}

function partsInZone(instant: Date, timeZone: string): number[] {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hourCycle: 'h23',
  }).formatToParts(instant);
  const pick = (type: Intl.DateTimeFormatPartTypes): number =>
    Number(parts.find((part) => part.type === type)?.value ?? 0);
  return [pick('year'), pick('month'), pick('day'), pick('hour'), pick('minute'), pick('second')];
}

export function zonedLocalToInstant(local: string, timeZone = BUDAPEST_TIME_ZONE): Date {
  const match = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?$/.exec(local);
  if (!match) throw new Error(`Érvénytelen helyi dátum-idő: ${local}`);
  const desired = match.slice(1).map(Number);
  while (desired.length < 6) desired.push(0);
  const utcGuess = Date.UTC(
    desired[0] ?? 0,
    (desired[1] ?? 1) - 1,
    desired[2] ?? 1,
    desired[3] ?? 0,
    desired[4] ?? 0,
    desired[5] ?? 0,
  );
  let result = utcGuess;
  for (let attempt = 0; attempt < 3; attempt += 1) {
    const actual = partsInZone(new Date(result), timeZone);
    const actualAsUtc = Date.UTC(
      actual[0] ?? 0,
      (actual[1] ?? 1) - 1,
      actual[2] ?? 1,
      actual[3] ?? 0,
      actual[4] ?? 0,
      actual[5] ?? 0,
    );
    result += utcGuess - actualAsUtc;
  }
  return new Date(result);
}

export function instantToLocal(iso: string, timeZone = BUDAPEST_TIME_ZONE): string {
  const [year, month, day, hour, minute, second] = partsInZone(new Date(iso), timeZone);
  return `${String(year).padStart(4, '0')}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}T${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}:${String(second).padStart(2, '0')}`;
}
