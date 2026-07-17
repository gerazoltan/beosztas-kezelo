import { describe, expect, it } from 'vitest';
import type { MonthSheet } from '../src/domain/types';
import { monthOptionLabel, monthOptionValue } from '../src/utils/monthOptions';

function month(sheetName: string): MonthSheet {
  return {
    sheetName,
    year: 2026,
    month: 6,
    headerRow: 4,
    nameColumn: 2,
    dayGroups: [],
    employees: [],
    warnings: [],
    legendStyles: { blue12: [], green12: [] },
  };
}

describe('hónapválasztó opciók', () => {
  it('az évvel együtt, pontosan egyszer jeleníti meg a hónap nevét', () => {
    const label = monthOptionLabel(month('Június'));

    expect(label).toBe('2026. június');
    expect(label.match(/június/gu)).toHaveLength(1);
    expect(label).toContain('2026');
  });

  it('a munkalap nevét csak a belső értékben őrzi meg a pontos kiválasztáshoz', () => {
    const primary = month('Június');
    const copy = month('Június másolat');
    const options = [primary, copy];
    const selectedValue = monthOptionValue(copy);

    expect(monthOptionLabel(copy)).toBe('2026. június');
    expect(monthOptionValue(primary)).not.toBe(selectedValue);
    expect(options.find((item) => monthOptionValue(item) === selectedValue)?.sheetName).toBe(
      'Június másolat',
    );
  });
});
