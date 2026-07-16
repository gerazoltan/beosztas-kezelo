import { expect, test } from '@playwright/test';
import ExcelJS from 'exceljs';

async function syntheticWorkbook(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Augusztus');
  sheet.getCell('B2').value = '2026. augusztus';
  sheet.getCell('B4').value = 'Név';
  for (let day = 1; day <= 31; day += 1) sheet.getCell(4, 3 + (day - 1) * 2).value = day;
  sheet.getCell('B5').value = 'Teszt Elek';
  sheet.getCell('B6').value = 'Összesen';
  sheet.getCell('C5').value = 12;
  sheet.getCell('C5').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC5D9F1' } };
  sheet.mergeCells('C5:D5');
  const content = await workbook.xlsx.writeBuffer();
  return Buffer.from(content);
}

test('teljes helyi ICS-folyamat', async ({ page }) => {
  await page.goto('.');
  await expect(page.getByRole('heading', { name: 'Beosztáskezelő' })).toBeVisible();
  await expect(page.getByText(/A fájl feldolgozása helyben/)).toBeVisible();
  await page.getByTestId('file-input').setInputFiles({
    name: 'anonim-minta.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await syntheticWorkbook(),
  });
  await page.getByLabel('Dolgozó').selectOption('teszt elek');
  await page.getByRole('button', { name: 'Beosztás feldolgozása' }).click();
  await expect(page.getByText('Nappalos 06–18')).toBeVisible();
  await expect(page.getByRole('button', { name: 'ICS letöltése' })).toBeEnabled();
});

test('mobilnézetben a folyamat és a kártyák az oldalon belül maradnak', async ({ page }) => {
  await page.goto('.');
  const body = page.locator('body');
  await expect(body).toBeVisible();
  const overflow = await page.evaluate(
    () => document.documentElement.scrollWidth <= window.innerWidth + 1,
  );
  expect(overflow).toBe(true);
});
