import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { readMonthEntries } from '../src/excel/dayEntries';
import {
  chooseDefaultMonth,
  parseWorkbook,
  validateExcelFileName,
} from '../src/excel/workbookParser';
import { AppError } from '../src/domain/errors';
import { workbookBuffer } from './fixtures/syntheticWorkbook';

describe('Excel parser', () => {
  it('felismeri a havi lapokat, és kizárja a Munka* segédlapot', async () => {
    const session = await parseWorkbook(await workbookBuffer(), 'minta.xlsx');
    expect(session.months.map((month) => month.sheetName)).toEqual(['Augusztus']);
    expect(session.months[0]).toMatchObject({ year: 2026, month: 8, headerRow: 4, nameColumn: 2 });
  });

  it.each([
    ['Február', 2025, 'február', 28],
    ['Február', 2024, 'február', 29],
    ['Április', 2026, 'április', 30],
    ['Január', 2026, 'január', 31],
  ])('%s %i hónapban %i napot fogad el', async (name, year, monthName, days) => {
    const session = await parseWorkbook(
      await workbookBuffer([{ name, year, monthName, days }]),
      'minta.xlsx',
    );
    expect(session.months[0]?.dayGroups.filter((group) => group.valid)).toHaveLength(days);
  });

  it('jelzi a 31. napot egy 30 napos hónapban', async () => {
    const session = await parseWorkbook(
      await workbookBuffer([
        {
          name: 'Szeptember',
          year: 2026,
          monthName: 'szeptember',
          days: 30,
          includeInvalid31: true,
        },
      ]),
      'minta.xlsx',
    );
    expect(session.months[0]?.dayGroups.at(-1)).toMatchObject({ day: 31, valid: false });
    expect(session.months[0]?.warnings.join(' ')).toContain('31. nap');
  });

  it('kiolvassa a neveket az üres sor megszakítása nélkül, és az Összesen: sornál megáll', async () => {
    const session = await parseWorkbook(await workbookBuffer(), 'minta.xlsx');
    expect(session.months[0]?.employees).toHaveLength(2);
    expect(session.months[0]?.employees.map((employee) => employee.rows)).toEqual([[7], [5]]);
  });

  it('egy naphoz két fizikai oszlopot térképez', async () => {
    const session = await parseWorkbook(await workbookBuffer(), 'minta.xlsx');
    expect(session.months[0]?.dayGroups[0]).toMatchObject({ startColumn: 3, endColumn: 4 });
  });

  it('az összevont napfejléc slave cellájának ismételt napszámát deduplikálja', async () => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await workbookBuffer());
    const sheet = workbook.getWorksheet('Augusztus');
    if (!sheet) throw new Error('Hiányzó tesztlap.');
    sheet.mergeCells('C4:D4');
    const bytes = new Uint8Array(await workbook.xlsx.writeBuffer());
    const session = await parseWorkbook(bytes.buffer, 'minta.xlsx');
    expect(session.months[0]?.dayGroups).toHaveLength(31);
    expect(session.months[0]?.dayGroups[1]).toMatchObject({ day: 2, startColumn: 5 });
  });

  it('kezeli a merge-et, a jobb oldali értéket és a kettős bejegyzést', async () => {
    const session = await parseWorkbook(await workbookBuffer(), 'minta.xlsx');
    const month = session.months[0];
    if (!month) throw new Error('Hiányzó teszthónap.');
    const entries = readMonthEntries(session, month, 'teszt elek');
    expect(entries[0]).toMatchObject({ kind: 'single', marker: '12' });
    expect(entries[0]?.selectedDiagnostic?.isMerged).toBe(true);
    expect(entries[8]).toMatchObject({ kind: 'single', marker: 'sz' });
    expect(entries[8]?.selectedDiagnostic?.address).toBe('T5');
    expect(entries[9]).toMatchObject({ kind: 'double' });
  });

  it('a napi mátrixot a napfejlécek határolják, az oldalsó KMR nem kerül bele', async () => {
    const session = await parseWorkbook(await workbookBuffer(), 'minta.xlsx');
    const month = session.months[0];
    if (!month) throw new Error('Hiányzó teszthónap.');
    const entries = readMonthEntries(session, month, 'teszt elek');
    expect(entries).toHaveLength(31);
    expect(
      entries.flatMap((entry) => entry.diagnostics).some((item) => item.address === 'BR15'),
    ).toBe(false);
  });

  it('a következő hónapot választja, különben dokumentált fallbacket használ', async () => {
    const session = await parseWorkbook(await workbookBuffer(), 'minta.xlsx');
    expect(chooseDefaultMonth(session.months, new Date(2026, 6, 2)).usedFallback).toBe(false);
    expect(chooseDefaultMonth(session.months, new Date(2026, 0, 2)).usedFallback).toBe(true);
  });

  it('elutasítja a nem Excel-fájlt', () => {
    expect(() => validateExcelFileName('minta.csv')).toThrowError(AppError);
  });

  it('érthetően elutasítja a régi bináris .xls formátumot', () => {
    expect(() => validateExcelFileName('minta.xls')).toThrowError(/régi \.xls formátum/);
  });

  it('a Név fejléc hiányát célzott hibakóddal jelzi', async () => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await workbookBuffer());
    const sheet = workbook.getWorksheet('Augusztus');
    if (!sheet) throw new Error('Hiányzó tesztlap.');
    sheet.getCell('B4').value = 'Dolgozó';
    const bytes = new Uint8Array(await workbook.xlsx.writeBuffer());
    await expect(parseWorkbook(bytes.buffer, 'minta.xlsx')).rejects.toMatchObject({
      code: 'NAME_HEADER_MISSING',
    });
  });

  it('összevonja a csak kisbetűben és szóközben eltérő duplikált neveket, de megőrzi a sorokat', async () => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await workbookBuffer());
    const sheet = workbook.getWorksheet('Augusztus');
    if (!sheet) throw new Error('Hiányzó tesztlap.');
    sheet.getCell('B7').value = '  TESZT   ELEK  ';
    const bytes = new Uint8Array(await workbook.xlsx.writeBuffer());
    const session = await parseWorkbook(bytes.buffer, 'minta.xlsx');
    expect(session.months[0]?.employees).toHaveLength(1);
    expect(session.months[0]?.employees[0]?.rows).toEqual([5, 7]);
  });

  it('nem bízik a worksheet deklarált dimension értékében', async () => {
    const zip = await JSZip.loadAsync(await workbookBuffer());
    const sheet = zip.file('xl/worksheets/sheet1.xml');
    if (!sheet) throw new Error('Hiányzó worksheet XML.');
    const source = await sheet.async('string');
    zip.file(
      'xl/worksheets/sheet1.xml',
      source.replace(/<dimension ref="[^"]+"\/>/u, '<dimension ref="A1:XFD1048576"/>'),
    );
    const buffer = await zip.generateAsync({ type: 'arraybuffer' });
    const session = await parseWorkbook(buffer, 'minta.xlsx');
    expect(session.months[0]?.employees).toHaveLength(2);
    expect(session.months[0]?.dayGroups).toHaveLength(31);
  });
});
