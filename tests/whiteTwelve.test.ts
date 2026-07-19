import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { readMonthEntries } from '../src/excel/dayEntries';
import { parseWorkbook } from '../src/excel/workbookParser';
import { interpretSchedule } from '../src/services/shifts';

async function fontBasedTwelveWorkbook(): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Augusztus');
  sheet.getCell('B2').value = '2026. augusztus';
  sheet.getCell('B4').value = 'Név';
  for (let day = 1; day <= 12; day += 1) {
    sheet.getCell(4, 3 + (day - 1) * 2).value = day;
  }
  sheet.getCell('B5').value = 'Teszt Elek';
  sheet.getCell('B6').value = 'Összesen';

  sheet.getCell(5, 3).value = 12;
  sheet.getCell(5, 5).value = '12';
  sheet.getCell(5, 7).value = 12;
  sheet.getCell(5, 7).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFFFFF' },
  };
  sheet.getCell(5, 9).value = 12;
  sheet.getCell(5, 9).fill = {
    type: 'pattern',
    pattern: 'none',
    fgColor: { argb: '00000000' },
    bgColor: { argb: '000000' },
  };
  sheet.getCell(5, 11).value = 12;
  sheet.getCell(5, 11).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFC6EFCE' },
  };
  sheet.getCell(5, 13).value = 12;
  sheet.getCell(5, 13).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFF2CC' },
  };
  sheet.getCell(5, 15).value = 12;
  sheet.getCell(5, 15).font = { color: { argb: 'FFFF0000' } };
  sheet.getCell(5, 15).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFC6EFCE' },
  };
  sheet.getCell(5, 17).value = 12;
  sheet.getCell(5, 17).font = { color: { argb: 'FF0000FF' } };
  sheet.getCell(5, 17).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFF2CC' },
  };
  sheet.getCell(5, 19).value = 12;
  sheet.getCell(5, 19).font = {
    color: { argb: 'FF008000' },
    underline: true,
    italic: true,
  };
  sheet.getCell(5, 21).value = 12;
  sheet.getCell(5, 21).font = { color: { argb: 'FF008000' } };
  sheet.getCell(5, 23).value = 'x';
  sheet.getCell(5, 25).value = 12;
  sheet.getCell(5, 25).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF000000' },
  };

  const bytes = new Uint8Array(await workbook.xlsx.writeBuffer());
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

describe('betűformázás-alapú 12 integrációja', () => {
  it('a hátteret figyelmen kívül hagyva ismeri fel a támogatott 12-eseket', async () => {
    const session = await parseWorkbook(await fontBasedTwelveWorkbook(), 'font-12.xlsx');
    const month = session.months[0];
    if (!month) throw new Error('Hiányzó teszthónap.');
    const entries = readMonthEntries(session, month, 'teszt elek');
    const result = interpretSchedule(entries, { legend: month.legendStyles });

    for (const day of [1, 2, 3, 4, 5, 6, 12]) {
      expect(result.rows.find((row) => row.date.day === day)).toMatchObject({
        status: 'Exportálható',
        serviceCategory: 'Parti szolgálat',
        event: {
          shiftTime: {
            start: `2026-08-${String(day).padStart(2, '0')}T07:00:00`,
            end: `2026-08-${String(day).padStart(2, '0')}T19:00:00`,
          },
        },
      });
    }

    expect(result.rows.find((row) => row.date.day === 7)).toMatchObject({
      status: 'Exportálható',
      serviceCategory: 'Esetszolgálat',
    });
    expect(result.rows.find((row) => row.date.day === 8)).toMatchObject({
      status: 'Exportálható',
      serviceCategory: '6-os kocsi',
      event: {
        shiftTime: { start: '2026-08-08T06:00:00', end: '2026-08-08T18:00:00' },
      },
    });
    expect(result.rows.find((row) => row.date.day === 9)).toMatchObject({
      status: 'Exportálható',
      serviceCategory: '10-es kocsi',
      event: {
        shiftTime: { start: '2026-08-09T10:00:00', end: '2026-08-09T22:00:00' },
      },
    });
    expect(result.rows.find((row) => row.date.day === 10)).toMatchObject({
      status: 'Exportálható',
      serviceCategory: '10-es kocsi',
      event: {
        shiftTime: { start: '2026-08-10T10:00:00', end: '2026-08-10T22:00:00' },
      },
    });
    expect(result.rows.find((row) => row.date.day === 11)?.status).toBe('Kizárva');

    const noFillDiagnostic = result.rows
      .find((row) => row.date.day === 4)
      ?.diagnostics.find((diagnostic) => diagnostic.displayedText === '12');
    expect(noFillDiagnostic).toMatchObject({
      fillPatternType: 'none',
      fillForegroundRaw: 'argb=00000000',
      fillBackgroundRaw: 'argb=000000',
      hasVisibleFill: false,
      fontColor: '#000000',
      positionInDayGroup: 1,
    });
  });
});
