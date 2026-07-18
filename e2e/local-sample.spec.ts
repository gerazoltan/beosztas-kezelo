import { expect, test } from '@playwright/test';
import { existsSync, readdirSync } from 'node:fs';
import { resolve } from 'node:path';

const directory = resolve(process.cwd(), 'local-samples');
const sampleName = existsSync(directory)
  ? readdirSync(directory).find((name) => name.toLocaleLowerCase('hu-HU').endsWith('.xlsx'))
  : undefined;

test.skip(!sampleName, 'A helyi, ignorált mintafájl nem érhető el.');

test('helyi valós mintafájlos böngészős smoke', async ({ page }) => {
  if (!sampleName) return;
  await page.goto('.');
  await page.getByTestId('file-input').setInputFiles(resolve(directory, sampleName));
  await expect(page.getByLabel('Hónap')).toBeVisible();
  await page.getByLabel('Dolgozó').selectOption({ index: 1 });
  await page.getByRole('button', { name: 'Beosztás feldolgozása' }).click();
  await expect(page.getByRole('heading', { name: 'Ellenőrzés' })).toBeVisible();
  await expect(page.getByRole('heading', { name: 'Export' })).toBeVisible();
  await expect(page.getByText('Felismert szolgálat', { exact: true })).toBeVisible();
});
