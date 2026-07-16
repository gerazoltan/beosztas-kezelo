import { HUNGARIAN_MONTHS } from '../domain/types';

export function normalizeWhitespace(value: string): string {
  return value.trim().replace(/\s+/g, ' ');
}

export function normalizeLookup(value: string): string {
  return normalizeWhitespace(value)
    .toLocaleLowerCase('hu-HU')
    .replace(/[.:]+$/u, '');
}

export function normalizeMarker(value: string): string {
  return normalizeLookup(value).replace(/\s+/g, '');
}

export function monthFromText(value: string): number | undefined {
  const normalized = normalizeLookup(value);
  const index = HUNGARIAN_MONTHS.findIndex((month) => normalized.includes(month));
  return index >= 0 ? index + 1 : undefined;
}

export function safeFileStem(value: string): string {
  return normalizeWhitespace(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLocaleLowerCase('hu-HU')
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '');
}
