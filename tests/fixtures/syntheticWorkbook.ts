import ExcelJS from 'exceljs';

export interface SyntheticMonth {
  name: string;
  year: number;
  monthName: string;
  days: number;
  includeInvalid31?: boolean;
  duplicateEmployee?: boolean;
}

export async function workbookBuffer(
  months: SyntheticMonth[] = [{ name: 'Augusztus', year: 2026, monthName: 'augusztus', days: 31 }],
): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  for (const month of months) {
    const sheet = workbook.addWorksheet(month.name);
    sheet.getCell('B2').value = `${month.year}. ${month.monthName}`;
    sheet.getCell('B4').value = 'Név';
    const dayCount = month.includeInvalid31 ? 31 : month.days;
    for (let day = 1; day <= dayCount; day += 1) {
      sheet.getCell(4, 3 + (day - 1) * 2).value = day;
    }
    sheet.getCell('B5').value = 'Teszt Elek';
    sheet.getCell('B7').value = month.duplicateEmployee ? 'Teszt Elek' : 'Minta Anna';
    sheet.getCell('B9').value = 'Összesen:';

    if (month.name === 'Augusztus') {
      const first = sheet.getCell(5, 3);
      first.value = 12;
      first.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC5D9F1' } };
      sheet.mergeCells(5, 3, 5, 4);
      const second = sheet.getCell(5, 5);
      second.value = 12;
      second.font = { color: { argb: 'FF008000' }, italic: true };
      sheet.getCell(5, 7).value = 17;
      sheet.getCell(5, 9).value = 7;
      sheet.getCell(5, 11).value = 5;
      sheet.getCell(5, 13).value = 7;
      sheet.getCell(5, 15).value = 'KMR';
      sheet.getCell(5, 17).value = 'x';
      sheet.getCell(5, 20).value = 'sz';
      sheet.getCell(5, 21).value = 12;
      sheet.getCell(5, 22).value = 'KMR';

      const blueLegend = sheet.getCell(12, 3);
      blueLegend.value = 12;
      blueLegend.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC5D9F1' } };
      sheet.getCell(12, 4).value = 'Nappalos 06-18';
      const greenLegend = sheet.getCell(13, 3);
      greenLegend.value = 12;
      greenLegend.font = { color: { argb: 'FF008000' }, italic: true };
      sheet.getCell(13, 4).value = 'Nappalos 10-22';
      sheet.getCell(15, 70).value = 'KMR';
      sheet.getCell(15, 71).value = 'Ápoló/GKV';
    }
  }
  workbook.addWorksheet('Munka1');
  const result = await workbook.xlsx.writeBuffer();
  const bytes = new Uint8Array(result);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

export function asFile(buffer: ArrayBuffer, name = 'anonim-minta.xlsx'): File {
  const file = new File([buffer], name, {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  Object.defineProperty(file, 'arrayBuffer', { value: () => Promise.resolve(buffer) });
  return file;
}
