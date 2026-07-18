import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { fillIsVisible, resolveCellStyle, resolveExcelColor, tintColor } from '../src/excel/colors';
import { classifyTwelve } from '../src/services/shifts';

describe('szín- és stílusfeloldás', () => {
  it('feloldja a közvetlen ARGB színt', () => {
    expect(resolveExcelColor({ argb: 'FF123456' }, [])).toBe('#123456');
    expect(resolveExcelColor({ argb: 'FFFFFFFF' }, [])).toBe('#FFFFFF');
    expect(resolveExcelColor({ argb: 'FFFFFF' }, [])).toBe('#FFFFFF');
  });

  it('feloldja a theme + tint színt', () => {
    expect(resolveExcelColor({ theme: 4, tint: 0.4 }, ['', '', '', '', '#0000FF'])).toBe('#6666FF');
    expect(tintColor('#80C080', -0.25)).toBe('#609060');
  });

  it('feloldja a theme- és indexed-fehéret, de az ismeretlen indexed színt nem találja ki', () => {
    expect(resolveExcelColor({ theme: 0 }, ['#FFFFFF'])).toBe('#FFFFFF');
    expect(resolveExcelColor({ indexed: 1 }, [])).toBe('#FFFFFF');
    expect(resolveExcelColor({ indexed: 9 }, [])).toBe('#FFFFFF');
    expect(resolveExcelColor({ indexed: 42 }, [])).toBeUndefined();
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
    expect(resolveCellStyle(cell, [])).toMatchObject({
      fillColor: '#00AAFF',
      hasVisibleFill: true,
    });
  });

  it('megkülönbözteti a kitöltés nélküli, pattern none és feloldhatatlan kitöltést', () => {
    const sheet = new ExcelJS.Workbook().addWorksheet('Teszt');
    const emptyFill = sheet.getCell('A1');
    const noneFill = sheet.getCell('A2');
    noneFill.fill = { type: 'pattern', pattern: 'none' };
    const unresolvedFill = sheet.getCell('A3');
    unresolvedFill.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { theme: 99 },
    };

    expect(resolveCellStyle(emptyFill, []).hasVisibleFill).toBe(false);
    expect(resolveCellStyle(noneFill, []).hasVisibleFill).toBe(false);
    expect(resolveCellStyle(unresolvedFill, [])).toMatchObject({
      fillColor: undefined,
      hasVisibleFill: true,
    });
  });

  it('a default 000000 színmezőt solid minta nélkül noFill állapotként kezeli', () => {
    const sheet = new ExcelJS.Workbook().addWorksheet('Teszt');
    const cell = sheet.getCell('A1');
    cell.value = 12;
    cell.fill = {
      type: 'pattern',
      pattern: 'none',
      fgColor: { argb: '00000000' },
      bgColor: { argb: '000000' },
    };

    expect(fillIsVisible(cell.fill)).toBe(false);
    expect(resolveCellStyle(cell, [])).toMatchObject({
      fillType: 'pattern',
      fillPatternType: 'none',
      fillForegroundRaw: 'argb=00000000',
      fillBackgroundRaw: 'argb=000000',
      hasVisibleFill: false,
      fillColor: undefined,
    });
  });

  it('a valódi solid fekete kitöltést látható, nem támogatott 12-ként kezeli', () => {
    const sheet = new ExcelJS.Workbook().addWorksheet('Teszt');
    const cell = sheet.getCell('A1');
    cell.value = 12;
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF000000' },
    };
    const style = resolveCellStyle(cell, []);

    expect(style).toMatchObject({
      fillType: 'pattern',
      fillPatternType: 'solid',
      fillForegroundRaw: 'argb=FF000000',
      hasVisibleFill: true,
      fillColor: '#000000',
    });
    expect(classifyTwelve(style, { blue12: [], green12: [] })).toBe('unknown');
  });

  it('elkülöníti a kék, zöld, fehér és ismeretlen 12-est', () => {
    const legend = { blue12: [], green12: [] };
    expect(
      classifyTwelve(
        { fillColor: '#C5D9F1', hasVisibleFill: true, italic: false, bold: false },
        legend,
      ),
    ).toBe('blue');
    expect(classifyTwelve({ fontColor: '#008000', italic: true, bold: false }, legend)).toBe(
      'green',
    );
    expect(classifyTwelve({ hasVisibleFill: false, italic: false, bold: false }, legend)).toBe(
      'white',
    );
    expect(
      classifyTwelve(
        { fillColor: '#FFFFFF', hasVisibleFill: true, italic: false, bold: false },
        legend,
      ),
    ).toBe('white');
    expect(
      classifyTwelve(
        { fillColor: '#F4CCCC', hasVisibleFill: true, italic: false, bold: false },
        legend,
      ),
    ).toBe('unknown');
    expect(classifyTwelve({ hasVisibleFill: true, italic: false, bold: false }, legend)).toBe(
      'unknown',
    );
    expect(classifyTwelve({ fontColor: '#000000', italic: false, bold: false }, legend)).toBe(
      'white',
    );
  });
});
