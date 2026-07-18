import ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { readMonthEntries } from '../src/excel/dayEntries';
import { parseWorkbook } from '../src/excel/workbookParser';
import { buildIcs } from '../src/services/ics';
import { interpretSchedule } from '../src/services/shifts';

async function sevenFiveWorkbook(): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Július');
  sheet.getCell('B2').value = '2026. július';
  sheet.getCell('B4').value = 'Név';
  sheet.getCell('C4').value = 1;
  sheet.getCell('E4').value = 2;
  sheet.getCell('G4').value = 3;
  sheet.getCell('B5').value = 'Teszt Elek';
  sheet.getCell('B6').value = 'Összesen';

  sheet.getCell('C5').value = 5;
  sheet.getCell('C5').font = { color: { argb: 'FF000000' } };

  sheet.getCell('E5').value = 7;
  sheet.getCell('E5').font = { color: { argb: 'FF0000FF' } };
  sheet.getCell('F5').value = 5;
  sheet.getCell('F5').font = { color: { argb: 'FFFF0000' } };

  sheet.getCell('G5').value = 7;
  sheet.getCell('G5').font = { color: { argb: 'FF008000' } };

  const bytes = new Uint8Array(await workbook.xlsx.writeBuffer());
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

describe('7 / 5 dupla napi cellacsoport integrációja', () => {
  it('két folytonos éjszakai eseményt készít a cellák fizikai sorrendje alapján', async () => {
    const session = await parseWorkbook(await sevenFiveWorkbook(), '7-5-regresszio.xlsx');
    const month = session.months[0];
    if (!month) throw new Error('Hiányzó teszthónap.');

    const entries = readMonthEntries(session, month, 'teszt elek');
    const doubleDay = entries.find((entry) => entry.date.day === 2);
    expect(doubleDay).toMatchObject({
      kind: 'double',
      marker: '7 / 5',
      diagnostics: [
        {
          address: 'E5',
          displayedText: '7',
          positionInDayGroup: 1,
          fontColorRaw: 'argb=FF0000FF',
          fontColor: '#0000FF',
        },
        {
          address: 'F5',
          displayedText: '5',
          positionInDayGroup: 2,
          fontColorRaw: 'argb=FFFF0000',
          fontColor: '#FF0000',
        },
      ],
    });

    const result = interpretSchedule(entries, {});
    expect(result.events).toEqual([
      expect.objectContaining({
        serviceCategory: 'Parti szolgálat',
        calendarTime: {
          start: '2026-07-01T19:00:00',
          end: '2026-07-02T07:00:00',
        },
      }),
      expect.objectContaining({
        serviceCategory: 'Esetszolgálat',
        calendarTime: {
          start: '2026-07-02T19:00:00',
          end: '2026-07-03T07:00:00',
        },
      }),
    ]);

    const rows = result.rows.filter((row) => row.date.day === 2);
    expect(rows.map((row) => [row.marker, row.status, row.serviceCategory])).toEqual([
      ['7', 'Felismerve', 'Parti szolgálat'],
      ['5', 'Exportálható', 'Esetszolgálat'],
    ]);
    expect(rows[0]?.note).toContain('00:00–07:00');
    expect(rows[1]?.note).toContain('19:00–24:00');
    expect(rows.some((row) => row.status === 'Hibás párosítás')).toBe(false);
    expect(result.summary).toMatchObject({ recognized: 2, invalid: 0, exportable: 2 });

    const ics = buildIcs(result.events);
    expect(ics.match(/BEGIN:VEVENT/g)).toHaveLength(2);
    expect(ics).toContain('DTSTART;TZID=Europe/Budapest:20260701T190000');
    expect(ics).toContain('DTEND;TZID=Europe/Budapest:20260702T070000');
    expect(ics).toContain('DTSTART;TZID=Europe/Budapest:20260702T190000');
    expect(ics).toContain('DTEND;TZID=Europe/Budapest:20260703T070000');
  });
});
