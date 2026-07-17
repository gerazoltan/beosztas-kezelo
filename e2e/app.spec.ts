import { expect, test, type Page } from '@playwright/test';
import ExcelJS from 'exceljs';

interface GoogleIdentityTestWindow extends Window {
  scrollIntoViewCalls?: ScrollIntoViewOptions[];
  google?: {
    accounts: {
      oauth2: {
        initTokenClient(config: { callback: (response: { access_token: string }) => void }): {
          requestAccessToken(): void;
        };
        revoke(): void;
      };
    };
  };
}

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
  sheet.getCell('E5').value = 'KMR';
  const content = await workbook.xlsx.writeBuffer();
  return Buffer.from(content);
}

async function twentyFourHourWorkbook(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Augusztus');
  sheet.getCell('B2').value = '2026. augusztus';
  sheet.getCell('B4').value = 'Név';
  for (let day = 1; day <= 31; day += 1) sheet.getCell(4, 3 + (day - 1) * 2).value = day;
  sheet.getCell('B5').value = 'Teszt Elek';
  sheet.getCell('B6').value = 'Összesen';
  sheet.getCell('G5').value = 17;
  sheet.getCell('I5').value = 7;
  const content = await workbook.xlsx.writeBuffer();
  return Buffer.from(content);
}

function processScheduleButton(page: Page) {
  return page.getByRole('button', { name: 'Beosztás feldolgozása' });
}

function icsButton(page: Page) {
  return page.getByRole('button', { name: 'ICS letöltése' });
}

function retryAction(page: Page, label: string) {
  return page.getByRole('button', { name: label });
}

function googleUploadButton(page: Page, eventCount: number) {
  return page.getByRole('button', {
    name: `${eventCount} kijelölt esemény hozzáadása a Google Naptárhoz`,
  });
}

function visibleShiftType(page: Page, shiftType: string) {
  return page.locator('.table-scroll').getByText(shiftType).first();
}

function visibleEventRow(page: Page, eventName: string) {
  return page.locator('tbody tr').filter({ hasText: eventName });
}

async function installGoogleIdentity(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const browserWindow = window as GoogleIdentityTestWindow;
    browserWindow.google = {
      accounts: {
        oauth2: {
          initTokenClient: ({ callback }) => ({
            requestAccessToken: () => callback({ access_token: 'e2e-access-token' }),
          }),
          revoke: () => undefined,
        },
      },
    };
  });
}

async function installScrollTracker(page: Page): Promise<void> {
  await page.addInitScript(() => {
    const browserWindow = window as GoogleIdentityTestWindow;
    browserWindow.scrollIntoViewCalls = [];
    HTMLElement.prototype.scrollIntoView = (options?: boolean | ScrollIntoViewOptions) => {
      if (typeof options === 'object') browserWindow.scrollIntoViewCalls?.push(options);
    };
  });
}

async function routeWritableCalendar(page: Page): Promise<void> {
  await page.route(
    'https://www.googleapis.com/calendar/v3/users/me/calendarList',
    async (route) => {
      await route.fulfill({
        contentType: 'application/json',
        body: JSON.stringify({
          items: [
            {
              id: 'primary',
              summary: 'Elsődleges tesztnaptár',
              primary: true,
              accessRole: 'owner',
            },
          ],
        }),
      });
    },
  );
}

async function openGoogleUpload(page: Page): Promise<void> {
  await page.goto('.');
  await page.getByTestId('file-input').setInputFiles({
    name: 'anonim-minta.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await syntheticWorkbook(),
  });
  await page.getByLabel('Dolgozó').selectOption('teszt elek');
  await processScheduleButton(page).click();
  await page.getByRole('button', { name: 'Google-bejelentkezés' }).click();
  await expect(page.getByLabel('Írható naptár')).toHaveValue('primary');
}

async function uploadGoogleEvents(page: Page, eventCount: number): Promise<void> {
  await googleUploadButton(page, eventCount).click();
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
  const selectedMonth = page.getByLabel('Hónap').locator('option:checked');
  await expect(selectedMonth).toHaveText('2026. augusztus');
  expect((await selectedMonth.textContent())?.match(/augusztus/gu)).toHaveLength(1);
  await page.getByLabel('Dolgozó').selectOption('teszt elek');
  await processScheduleButton(page).click();
  await expect(visibleShiftType(page, 'Nappalos 06–18')).toBeVisible();
  const [download] = await Promise.all([page.waitForEvent('download'), icsButton(page).click()]);
  expect(download.suggestedFilename()).toBe('teszt-elek-2026-augusztus.ics');
});

test('a 17–7 szolgálat listában 24 órás, ICS-ben és Google-ben 06:59-ig tart', async ({ page }) => {
  await installGoogleIdentity(page);
  await routeWritableCalendar(page);
  let eventRequestBody: unknown;
  await page.route(
    'https://www.googleapis.com/calendar/v3/calendars/primary/events**',
    async (route) => {
      if (route.request().method() === 'GET') {
        await route.fulfill({
          contentType: 'application/json',
          body: JSON.stringify({ items: [] }),
        });
        return;
      }
      eventRequestBody = JSON.parse(route.request().postData() ?? 'null') as unknown;
      await route.fulfill({
        contentType: 'application/json',
        body: JSON.stringify({ id: 'created-24-hour', colorId: '10' }),
      });
    },
  );
  await page.goto('.');
  await page.getByTestId('file-input').setInputFiles({
    name: '17-7-minta.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await twentyFourHourWorkbook(),
  });
  await page.getByLabel('Dolgozó').selectOption('teszt elek');
  await processScheduleButton(page).click();

  const serviceRow = page.locator('tbody tr').filter({ hasText: '24 órás szolgálat' });
  await expect(serviceRow.locator('td[data-label="Kezdés"]')).toHaveText('07:00');
  await expect(serviceRow.locator('td[data-label="Befejezés"]')).toHaveText('07:00');
  await serviceRow.getByText('Technikai részletek').click();
  await expect(
    serviceRow.getByText('A naptáresemény befejezése 06:59 a jobb naptári elkülönítés érdekében.'),
  ).toBeVisible();

  const [download] = await Promise.all([page.waitForEvent('download'), icsButton(page).click()]);
  const downloadStream = await download.createReadStream();
  downloadStream.setEncoding('utf8');
  const chunks: string[] = [];
  for await (const chunk of downloadStream as AsyncIterable<unknown>) {
    if (typeof chunk !== 'string') throw new Error('Az ICS-letöltés nem UTF-8 szöveget adott.');
    chunks.push(chunk);
  }
  const icsContent = chunks.join('');
  expect(icsContent).toContain('DTSTART;TZID=Europe/Budapest:20260803T070000');
  expect(icsContent).toContain('DTEND;TZID=Europe/Budapest:20260804T065900');

  await page.getByRole('button', { name: 'Google-bejelentkezés' }).click();
  await uploadGoogleEvents(page, 1);
  await expect(page.getByRole('heading', { name: 'Sikeres naptárfeltöltés' })).toBeVisible();
  expect(eventRequestBody).toMatchObject({
    summary: 'OMSZ',
    start: { dateTime: '2026-08-03T07:00:00', timeZone: 'Europe/Budapest' },
    end: { dateTime: '2026-08-04T06:59:00', timeZone: 'Europe/Budapest' },
  });
});

test('mobilnézetben az oldal vízszintes túlcsordulás nélkül megjelenik', async ({ page }) => {
  await page.goto('.');
  await expect(page.locator('body')).toBeVisible();
  const overflow = await page.evaluate(
    () => document.documentElement.scrollWidth <= window.innerWidth + 1,
  );
  expect(overflow).toBe(true);
});

test('OAuth után natív fetch-csel, Bearer tokennel lekéri az írható naptárakat', async ({
  page,
}) => {
  await installGoogleIdentity(page);

  let authorizationHeader: string | undefined;
  const eventRequestBodies: unknown[] = [];
  await page.route(
    'https://www.googleapis.com/calendar/v3/users/me/calendarList',
    async (route) => {
      authorizationHeader = route.request().headers().authorization;
      await route.fulfill({
        contentType: 'application/json',
        body: JSON.stringify({
          items: [
            {
              id: 'primary',
              summary: 'Elsődleges tesztnaptár',
              primary: true,
              accessRole: 'owner',
            },
          ],
        }),
      });
    },
  );
  await page.route(
    'https://www.googleapis.com/calendar/v3/calendars/primary/events**',
    async (route) => {
      if (route.request().method() === 'GET') {
        await route.fulfill({
          contentType: 'application/json',
          body: JSON.stringify({ items: [] }),
        });
        return;
      }
      const requestBody = JSON.parse(route.request().postData() ?? 'null') as unknown;
      eventRequestBodies.push(requestBody);
      await route.fulfill({
        contentType: 'application/json',
        body: JSON.stringify({
          ...(typeof requestBody === 'object' && requestBody !== null ? requestBody : {}),
          id: `created-${eventRequestBodies.length}`,
          colorId: '10',
        }),
      });
    },
  );

  await openGoogleUpload(page);

  const calendarSelect = page.getByLabel('Írható naptár');
  await expect(calendarSelect).toHaveValue('primary');
  await expect(calendarSelect.locator('option:checked')).toHaveText(
    'Elsődleges tesztnaptár (elsődleges)',
  );
  expect(authorizationHeader).toBe('Bearer e2e-access-token');

  await uploadGoogleEvents(page, 2);
  await expect
    .poll(() => eventRequestBodies)
    .toEqual(
      expect.arrayContaining([
        expect.objectContaining({ summary: 'OMSZ', colorId: '10' }),
        expect.objectContaining({ summary: 'KMR', colorId: '10' }),
      ]),
    );
  const visibleReview = page.locator('.table-scroll');
  await expect(visibleReview.getByText(/zöld Basil \(10\) színnel létrehozta/u)).toHaveCount(2);
  await expect(page.getByRole('heading', { name: 'Sikeres naptárfeltöltés' })).toBeVisible();
  await expect(
    page.getByRole('button', { name: /kijelölt esemény hozzáadása/u }),
  ).not.toBeVisible();
  await expect(visibleEventRow(page, 'OMSZ').getByRole('checkbox')).toBeDisabled();
  await expect(visibleEventRow(page, 'KMR').getByRole('checkbox')).toBeDisabled();

  const calendarLink = page.getByRole('link', { name: 'Google Naptár megnyitása' });
  await expect(calendarLink).toBeVisible();
  await expect(calendarLink).toHaveClass(/tertiary/u);
  await expect(calendarLink).toHaveAttribute('href', 'https://calendar.google.com/calendar/u/0/r');
  await expect(calendarLink).toHaveAttribute('target', '_blank');
  await expect(calendarLink).toHaveAttribute('rel', 'noopener noreferrer');
  await page.context().route('https://calendar.google.com/**', async (route) => {
    await route.fulfill({ contentType: 'text/html', body: '<title>Google Naptár teszt</title>' });
  });
  const [calendarPage] = await Promise.all([page.waitForEvent('popup'), calendarLink.click()]);
  await calendarPage.waitForLoadState();
  expect(calendarPage.url()).toBe('https://calendar.google.com/calendar/u/0/r');
  await calendarPage.close();

  await page.getByRole('button', { name: 'Kijelentkezés' }).click();
  await expect(
    page.getByText(
      'Kijelentkeztél a Google Naptárból. A korábban létrehozott események a naptárban maradnak.',
    ),
  ).toBeVisible();
  await expect(page.getByRole('heading', { name: 'Sikeres naptárfeltöltés' })).toBeVisible();
  await expect(page.getByText('2 esemény létrehozva.')).toBeVisible();
  await expect(page.getByRole('link', { name: 'Google Naptár megnyitása' })).not.toBeVisible();
  await expect(page.getByLabel('Írható naptár')).not.toBeVisible();
  await expect(page.getByRole('button', { name: 'Másik fiók csatlakoztatása' })).toBeVisible();
  await expect(visibleReview.getByText('Létrehozva', { exact: true })).toHaveCount(2);

  await page.getByRole('button', { name: 'Másik fiók csatlakoztatása' }).click();
  await expect(page.getByLabel('Írható naptár')).toHaveValue('primary');
  await expect(page.getByLabel('Írható naptár')).toBeEnabled();
});

test('az új beosztás feldolgozása nullázza az Excel-folyamatot, felgörget és megőrzi a Google-munkamenetet', async ({
  page,
}) => {
  await installGoogleIdentity(page);
  await installScrollTracker(page);
  await routeWritableCalendar(page);
  const eventMethods: string[] = [];
  await page.route(
    /https:\/\/www\.googleapis\.com\/calendar\/v3\/calendars\/[^/]+\/events/u,
    async (route) => {
      eventMethods.push(route.request().method());
      if (route.request().method() === 'GET') {
        await route.fulfill({
          contentType: 'application/json',
          body: JSON.stringify({ items: [] }),
        });
        return;
      }
      await route.fulfill({
        contentType: 'application/json',
        body: JSON.stringify({ id: 'created', colorId: '10' }),
      });
    },
  );
  await openGoogleUpload(page);
  await uploadGoogleEvents(page, 2);
  await expect(page.getByRole('heading', { name: 'Sikeres naptárfeltöltés' })).toBeVisible();

  const newScheduleButton = page.getByRole('button', { name: 'Új beosztás feldolgozása' });
  await expect(newScheduleButton).toHaveClass(/primary/u);
  await expect(page.getByRole('button', { name: 'Feltöltés másik naptárba' })).toHaveClass(
    /secondary/u,
  );
  await newScheduleButton.click();

  await expect(page.getByRole('heading', { name: 'Excel-fájl kiválasztása' })).toBeVisible();
  await expect(page.getByText(/Kiválasztva:/u)).not.toBeVisible();
  await expect(page.getByLabel('Hónap')).not.toBeVisible();
  await expect(page.getByLabel('Dolgozó')).not.toBeVisible();
  await expect(page.getByRole('heading', { name: 'Ellenőrzés' })).not.toBeVisible();
  await expect(page.getByRole('heading', { name: 'Sikeres naptárfeltöltés' })).not.toBeVisible();
  expect(
    await page.evaluate(() => (window as GoogleIdentityTestWindow).scrollIntoViewCalls ?? []),
  ).toContainEqual({ behavior: 'smooth', block: 'start' });

  await page.getByTestId('file-input').setInputFiles({
    name: 'új-anonim-minta.xlsx',
    mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    buffer: await syntheticWorkbook(),
  });
  await page.getByLabel('Dolgozó').selectOption('teszt elek');
  await processScheduleButton(page).click();

  await expect(page.getByLabel('Írható naptár')).toHaveValue('primary');
  await expect(page.getByRole('button', { name: 'Google-bejelentkezés' })).not.toBeVisible();
  expect(eventMethods).not.toContain('DELETE');
});

test('a másik naptár művelet megtartja az eseményeket, és üres naptárválasztással új kört nyit', async ({
  page,
}) => {
  await installGoogleIdentity(page);
  await page.route(
    'https://www.googleapis.com/calendar/v3/users/me/calendarList',
    async (route) => {
      await route.fulfill({
        contentType: 'application/json',
        body: JSON.stringify({
          items: [
            {
              id: 'primary',
              summary: 'Elsődleges tesztnaptár',
              primary: true,
              accessRole: 'owner',
            },
            {
              id: 'secondary',
              summary: 'Második tesztnaptár',
              accessRole: 'writer',
            },
          ],
        }),
      });
    },
  );
  const eventMethods: string[] = [];
  await page.route(
    /https:\/\/www\.googleapis\.com\/calendar\/v3\/calendars\/[^/]+\/events/u,
    async (route) => {
      eventMethods.push(route.request().method());
      if (route.request().method() === 'GET') {
        await route.fulfill({
          contentType: 'application/json',
          body: JSON.stringify({ items: [] }),
        });
        return;
      }
      await route.fulfill({
        contentType: 'application/json',
        body: JSON.stringify({ id: 'created', colorId: '10' }),
      });
    },
  );
  await openGoogleUpload(page);
  await uploadGoogleEvents(page, 2);
  await expect(page.getByRole('heading', { name: 'Sikeres naptárfeltöltés' })).toBeVisible();

  await page.getByRole('button', { name: 'Feltöltés másik naptárba' }).click();

  await expect(
    page.getByText('Ugyanezek az események egy másik naptárban is létrejönnek.'),
  ).toBeVisible();
  await expect(page.getByLabel('Írható naptár')).toHaveValue('');
  await expect(page.getByRole('heading', { name: 'Ellenőrzés' })).toBeVisible();
  await expect(visibleShiftType(page, 'Nappalos 06–18')).toBeVisible();
  await expect(page.getByText('Létrehozva', { exact: true })).toHaveCount(0);
  await expect(visibleEventRow(page, 'OMSZ').getByRole('checkbox')).toBeEnabled();
  await expect(visibleEventRow(page, 'KMR').getByRole('checkbox')).toBeEnabled();

  await page.getByLabel('Írható naptár').selectOption('secondary');
  await expect(googleUploadButton(page, 2)).toBeEnabled();
  expect(eventMethods).not.toContain('DELETE');
});

test('részleges siker után csak a sikertelen eseményt próbálja újra', async ({ page }) => {
  await installGoogleIdentity(page);
  await routeWritableCalendar(page);
  const postCounts = new Map<string, number>();
  await page.route(
    'https://www.googleapis.com/calendar/v3/calendars/primary/events**',
    async (route) => {
      if (route.request().method() === 'GET') {
        await route.fulfill({
          contentType: 'application/json',
          body: JSON.stringify({ items: [] }),
        });
        return;
      }
      const body = JSON.parse(route.request().postData() ?? 'null') as unknown;
      const summary =
        typeof body === 'object' &&
        body !== null &&
        'summary' in body &&
        typeof body.summary === 'string'
          ? body.summary
          : '';
      const count = (postCounts.get(summary) ?? 0) + 1;
      postCounts.set(summary, count);
      if (summary === 'KMR' && count === 1) {
        await route.fulfill({
          status: 500,
          contentType: 'application/json',
          body: JSON.stringify({ error: { message: 'teszt API-hiba' } }),
        });
        return;
      }
      await route.fulfill({
        contentType: 'application/json',
        body: JSON.stringify({ id: `created-${summary}`, colorId: '10' }),
      });
    },
  );

  await openGoogleUpload(page);
  await uploadGoogleEvents(page, 2);
  await expect(
    page.getByRole('heading', { name: 'Részben sikeres naptárfeltöltés' }),
  ).toBeVisible();
  await expect(page.getByText('1 esemény létrehozva.')).toBeVisible();
  await expect(page.getByText('1 sikertelen művelet.')).toBeVisible();

  await retryAction(page, 'Csak a sikertelenek újrapróbálása').click();
  await expect(page.getByRole('heading', { name: 'Sikeres naptárfeltöltés' })).toBeVisible();
  expect(postCounts.get('OMSZ')).toBe(1);
  expect(postCounts.get('KMR')).toBe(2);
});

test('teljes hiba esetén piros eredménykártyát és újrapróbálást mutat', async ({ page }) => {
  await installGoogleIdentity(page);
  await routeWritableCalendar(page);
  await page.route(
    'https://www.googleapis.com/calendar/v3/calendars/primary/events**',
    async (route) => {
      if (route.request().method() === 'GET') {
        await route.fulfill({
          contentType: 'application/json',
          body: JSON.stringify({ items: [] }),
        });
        return;
      }
      await route.fulfill({
        status: 500,
        contentType: 'application/json',
        body: JSON.stringify({ error: { message: 'teljes teszthiba' } }),
      });
    },
  );

  await openGoogleUpload(page);
  await uploadGoogleEvents(page, 2);

  const errorCard = page.locator('.upload-result-error');
  await expect(
    errorCard.getByRole('heading', { name: 'Sikertelen naptárfeltöltés' }),
  ).toBeVisible();
  await expect(errorCard.getByText('A Google Naptár API hibát jelzett.')).toBeVisible();
  await expect(retryAction(page, 'Újrapróbálás')).toBeEnabled();
  await expect(errorCard.getByText('Technikai részletek')).toBeVisible();
});
