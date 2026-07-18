import type { CalendarEvent } from '../domain/types';
import { HUNGARIAN_MONTHS } from '../domain/types';
import { safeFileStem } from '../utils/normalize';
import { calendarEventDescription } from './calendarEventDescription';

const CRLF = '\r\n';

export function escapeIcsText(value: string): string {
  return value
    .replace(/\\/g, '\\\\')
    .replace(/\r?\n/g, '\\n')
    .replace(/,/g, '\\,')
    .replace(/;/g, '\\;');
}

export function formatIcsLocal(value: string): string {
  return value.replace(/[-:]/g, '');
}

export function stableUid(item: CalendarEvent): string {
  let hash = 0x811c9dc5;
  const value = `${item.summary}|${item.calendarTime.start}|${item.calendarTime.end}|${item.shiftType}|${item.serviceCategory}`;
  for (const character of value) {
    hash ^= character.codePointAt(0) ?? 0;
    hash = Math.imul(hash, 0x01000193);
  }
  return `${(hash >>> 0).toString(16)}-${item.id}@beosztas-kezelo`;
}

function utf8Length(value: string): number {
  return new TextEncoder().encode(value).length;
}

export function foldIcsLine(line: string): string {
  if (utf8Length(line) <= 75) return line;
  const result: string[] = [];
  let current = '';
  for (const character of line) {
    const prefixLength = result.length === 0 ? 0 : 1;
    if (current && utf8Length(current + character) + prefixLength > 75) {
      result.push(current);
      current = character;
    } else current += character;
  }
  if (current) result.push(current);
  return result.map((part, index) => (index === 0 ? part : ` ${part}`)).join(CRLF);
}

function dtstamp(value: Date): string {
  return value
    .toISOString()
    .replace(/[-:]/g, '')
    .replace(/\.\d{3}Z$/, 'Z');
}

const TIME_ZONE_BLOCK = [
  'BEGIN:VTIMEZONE',
  'TZID:Europe/Budapest',
  'X-LIC-LOCATION:Europe/Budapest',
  'BEGIN:DAYLIGHT',
  'TZOFFSETFROM:+0100',
  'TZOFFSETTO:+0200',
  'TZNAME:CEST',
  'DTSTART:19700329T020000',
  'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU',
  'END:DAYLIGHT',
  'BEGIN:STANDARD',
  'TZOFFSETFROM:+0200',
  'TZOFFSETTO:+0100',
  'TZNAME:CET',
  'DTSTART:19701025T030000',
  'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU',
  'END:STANDARD',
  'END:VTIMEZONE',
];

export function buildIcs(events: CalendarEvent[], generatedAt = new Date()): string {
  const lines = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Beosztáskezelő//HU',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH',
    ...TIME_ZONE_BLOCK,
  ];
  for (const item of events) {
    lines.push(
      'BEGIN:VEVENT',
      `UID:${stableUid(item)}`,
      `DTSTAMP:${dtstamp(generatedAt)}`,
      `DTSTART;TZID=Europe/Budapest:${formatIcsLocal(item.calendarTime.start)}`,
      `DTEND;TZID=Europe/Budapest:${formatIcsLocal(item.calendarTime.end)}`,
      `SUMMARY:${escapeIcsText(item.summary)}`,
      `DESCRIPTION:${escapeIcsText(calendarEventDescription(item))}`,
      'END:VEVENT',
    );
  }
  lines.push('END:VCALENDAR');
  return `${lines.map(foldIcsLine).join(CRLF)}${CRLF}`;
}

export function icsFileName(employeeName: string, year: number, month: number): string {
  return `${safeFileStem(employeeName)}-${year}-${HUNGARIAN_MONTHS[month - 1]}.ics`;
}

export function downloadIcs(content: string, fileName: string): void {
  const url = URL.createObjectURL(new Blob([content], { type: 'text/calendar;charset=utf-8' }));
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}
