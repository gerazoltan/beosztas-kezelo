import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { readMonthEntries } from '../src/excel/dayEntries';
import { parseWorkbook } from '../src/excel/workbookParser';
import { interpretSchedule } from '../src/services/shifts';

async function whiteTwelveWorkbook(): Promise<ArrayBuffer> {
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
    fgColor: { argb: 'FFC5D9F1' },
  };
  sheet.getCell(5, 13).value = 12;
  sheet.getCell(5, 13).font = { color: { argb: 'FF008000' }, italic: true };
  sheet.getCell(5, 15).value = 12;
  sheet.getCell(5, 15).fill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFF4CCCC' },
  };
  sheet.getCell(5, 17).value = 7;
  sheet.getCell(5, 19).value = 5;
  sheet.getCell(5, 21).value = 17;
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

describe('fehér vagy kitöltés nélküli 12 integrációja', () => {
  it('a tényleges Excel-cellákból csak a támogatott 12-eseket teszi exportálhatóvá', async () => {
    const session = await parseWorkbook(await whiteTwelveWorkbook(), 'feher-12.xlsx');
    const month = session.months[0];
    if (!month) throw new Error('Hiányzó teszthónap.');
    const entries = readMonthEntries(session, month, 'teszt elek');
    const result = interpretSchedule(entries, { legend: month.legendStyles });

    for (const day of [1, 2, 3, 4]) {
      const event = result.events.find((item) =>
        item.shiftTime.start.startsWith(`2026-08-0${day}`),
      );
      expect(event).toMatchObject({
        summary: 'OMSZ',
        shiftType: 'Nappalos 07–19',
        shiftTime: {
          start: `2026-08-0${day}T07:00:00`,
          end: `2026-08-0${day}T19:00:00`,
        },
        calendarTime: {
          start: `2026-08-0${day}T07:00:00`,
          end: `2026-08-0${day}T19:00:00`,
        },
      });
      expect(result.rows.find((row) => row.date.day === day)).toMatchObject({
        status: 'Exportálható',
        note: 'Fehér vagy kitöltés nélküli 12 felismerve.',
      });
    }
    const noFillRow = result.rows.find((row) => row.date.day === 4);
    expect(
      noFillRow?.diagnostics.some(
        (diagnostic) =>
          diagnostic.fillPatternType === 'none' &&
          diagnostic.fillForegroundRaw === 'argb=00000000' &&
          diagnostic.fillBackgroundRaw === 'argb=000000' &&
          diagnostic.hasVisibleFill === false &&
          diagnostic.fillCategory === 'noFill',
      ),
    ).toBe(true);

    expect(
      result.events.find((item) => item.shiftTime.start.startsWith('2026-08-05')),
    ).toMatchObject({
      shiftType: 'Nappalos 06–18',
      shiftTime: { start: '2026-08-05T06:00:00', end: '2026-08-05T18:00:00' },
    });
    expect(
      result.events.find((item) => item.shiftTime.start.startsWith('2026-08-06')),
    ).toMatchObject({
      shiftType: 'Nappalos 10–22',
      shiftTime: { start: '2026-08-06T10:00:00', end: '2026-08-06T22:00:00' },
    });
    expect(result.rows.find((row) => row.date.day === 7)).toMatchObject({
      status: 'Bizonytalan',
    });
    const blackFillRow = result.rows.find((row) => row.date.day === 12);
    expect(blackFillRow?.status).toBe('Bizonytalan');
    expect(
      blackFillRow?.diagnostics.some(
        (diagnostic) =>
          diagnostic.fillPatternType === 'solid' &&
          diagnostic.fillColor === '#000000' &&
          diagnostic.hasVisibleFill === true &&
          diagnostic.fillCategory === 'unsupported',
      ),
    ).toBe(true);
    expect(
      result.events.some((item) =>
        ['2026-08-08', '2026-08-09', '2026-08-10', '2026-08-11', '2026-08-12'].some((date) =>
          item.shiftTime.start.startsWith(date),
        ),
      ),
    ).toBe(false);
  });
});
