import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { fillIsVisible, resolveCellStyle, resolveExcelColor, tintColor } from '../src/excel/colors';
import { classifyTwelve } from '../src/services/shifts';

describe('szín- és stílusfeloldás', () => {
  it('feloldja a közvetlen ARGB színt', () => {
    expect(resolveExcelColor({ argb: 'FF123456' }, [])).toBe('#123456');
    expect(resolveExcelColor({ argb: 'FFFFFFFF' }, [])).toBe('#FFFFFF');
    expect(resolveExcelColor({ argb: 'FFFFFF' }, [])).toBe('#FFFFFF');
    expect(resolveExcelColor({ rgb: '008000' }, [])).toBe('#008000');
  });

  it('feloldja a theme + tint színt', () => {
    expect(resolveExcelColor({ theme: 4, tint: 0.4 }, ['', '', '', '', '#0000FF'])).toBe('#6666FF');
    expect(tintColor('#80C080', -0.25)).toBe('#609060');
  });

  it('feloldja a theme- és indexed-színeket, de a rendszer indexed színét nem találja ki', () => {
    expect(resolveExcelColor({ theme: 0 }, ['#FFFFFF'])).toBe('#FFFFFF');
    expect(resolveExcelColor({ indexed: 1 }, [])).toBe('#FFFFFF');
    expect(resolveExcelColor({ indexed: 9 }, [])).toBe('#FFFFFF');
    expect(resolveExcelColor({ indexed: 2 }, [])).toBe('#FF0000');
    expect(resolveExcelColor({ indexed: 17 }, [])).toBe('#008000');
    expect(resolveExcelColor({ indexed: 64 }, [])).toBeUndefined();
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

  it('a valódi solid fekete hátteret technikailag megőrzi, de a 12-et a fekete betű alapján osztályozza', () => {
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
    expect(classifyTwelve(style)).toBe('party');
  });

  it('központilag normalizálja a font nyers színét, alapértelmezését és aláhúzását', () => {
    const sheet = new ExcelJS.Workbook().addWorksheet('Teszt');
    const defaultBlack = sheet.getCell('A1');
    const themed = sheet.getCell('A2');
    themed.font = { color: { theme: 1 }, underline: true };
    const indexed = sheet.getCell('A3');
    indexed.font = { color: { indexed: 2 } as unknown as ExcelJS.Color };
    const theme = ['#FFFFFF', '#000000'];

    expect(resolveCellStyle(defaultBlack, theme)).toMatchObject({
      fontColorRaw: undefined,
      fontColor: '#000000',
      underline: false,
    });
    expect(resolveCellStyle(themed, theme)).toMatchObject({
      fontColorRaw: 'theme=1',
      fontColor: '#000000',
      underline: true,
    });
    expect(resolveCellStyle(indexed, theme)).toMatchObject({
      fontColorRaw: 'indexed=2',
      fontColor: '#FF0000',
    });
  });

  it('csak a betűszín és az aláhúzás alapján különíti el a 12-eseket', () => {
    expect(
      classifyTwelve({
        fontColor: '#0000FF',
        fillColor: '#00FF00',
        underline: false,
        italic: false,
        bold: false,
      }),
    ).toBe('blue');
    expect(
      classifyTwelve({
        fontColor: '#008000',
        underline: true,
        italic: true,
        bold: false,
      }),
    ).toBe('tenCar');
    expect(
      classifyTwelve({
        fontColor: '#008000',
        underline: false,
        italic: true,
        bold: false,
      }),
    ).toBe('unknown');
    expect(
      classifyTwelve({
        fontColor: '#FF0000',
        fillColor: '#FFF2CC',
        underline: false,
        italic: false,
        bold: false,
      }),
    ).toBe('emergency');
    expect(
      classifyTwelve({
        fontColor: '#000000',
        fillColor: '#C6EFCE',
        underline: false,
        italic: false,
        bold: false,
      }),
    ).toBe('party');
  });
});
