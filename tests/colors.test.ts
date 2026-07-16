import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { resolveCellStyle, resolveExcelColor, tintColor } from '../src/excel/colors';
import { classifyTwelve } from '../src/services/shifts';

describe('szín- és stílusfeloldás', () => {
  it('feloldja a közvetlen ARGB színt', () => {
    expect(resolveExcelColor({ argb: 'FF123456' }, [])).toBe('#123456');
  });

  it('feloldja a theme + tint színt', () => {
    expect(resolveExcelColor({ theme: 4, tint: 0.4 }, ['', '', '', '', '#0000FF'])).toBe('#6666FF');
    expect(tintColor('#80C080', -0.25)).toBe('#609060');
  });

  it('solid kitöltésnél az fgColor értékét használja', () => {
    const sheet = new ExcelJS.Workbook().addWorksheet('Teszt');
    const cell = sheet.getCell('A1');
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF00AAFF' },
      bgColor: { argb: 'FFFF0000' },
    };
    expect(resolveCellStyle(cell, []).fillColor).toBe('#00AAFF');
  });

  it('elkülöníti a kék, zöld és ismeretlen 12-est', () => {
    const legend = { blue12: [], green12: [] };
    expect(classifyTwelve({ fillColor: '#C5D9F1', italic: false, bold: false }, legend)).toBe(
      'blue',
    );
    expect(classifyTwelve({ fontColor: '#008000', italic: true, bold: false }, legend)).toBe(
      'green',
    );
    expect(classifyTwelve({ fontColor: '#000000', italic: false, bold: false }, legend)).toBe(
      'unknown',
    );
  });
});
