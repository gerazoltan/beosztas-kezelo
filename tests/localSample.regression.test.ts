import { readdir, readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import type { WorkbookSession } from '../src/domain/types';
import { parseWorkbook } from '../src/excel/workbookParser';
import { readMonthEntries } from '../src/excel/dayEntries';
import { classifyTwelve, interpretSchedule } from '../src/services/shifts';
import { AppError } from '../src/domain/errors';
import { buildIcs } from '../src/services/ics';
import { GoogleCalendarClient } from '../src/services/googleCalendar';

const sampleDirectory = resolve(process.cwd(), 'local-samples');
const hasSamples = existsSync(sampleDirectory);

async function loadFirstSample(): Promise<WorkbookSession> {
  const files = (await readdir(sampleDirectory)).filter((name) =>
    name.toLowerCase().endsWith('.xlsx'),
  );
  expect(files.length).toBeGreaterThan(0);
  const fileName = files[0];
  if (!fileName) throw new Error('Hiányzó helyi mintafájl.');
  const bytes = await readFile(resolve(sampleDirectory, fileName));
  const buffer = Uint8Array.from(bytes).buffer;
  try {
    return await parseWorkbook(buffer, fileName);
  } catch (error) {
    if (error instanceof AppError) throw new Error(error.technicalDetails ?? error.message);
    throw error;
  }
}

describe('helyi, nem követett Excel-minta regresszió', () => {
  it.skipIf(!hasSamples)(
    'felismeri a valós havi szerkezetet és mindkét bizonyított 12-stílust',
    async () => {
      const session = await loadFirstSample();
      expect(session.months).toHaveLength(12);
      expect(session.months.every((month) => month.dayGroups.length >= 28)).toBe(true);

      const kinds = new Set<string>();
      const shiftTypes = new Set<string>();
      for (const month of session.months) {
        for (const employee of month.employees) {
          if (employee.rows.length !== 1) continue;
          const entries = readMonthEntries(session, month, employee.normalizedName);
          for (const entry of entries) {
            if (entry.normalizedMarker === '12') {
              kinds.add(classifyTwelve(entry.selectedDiagnostic, month.legendStyles));
            }
          }
          for (const calendarEvent of interpretSchedule(entries, {
            legend: month.legendStyles,
          }).events) {
            shiftTypes.add(calendarEvent.shiftType);
          }
        }
      }
      expect(kinds.has('blue')).toBe(true);
      expect(kinds.has('green')).toBe(true);
      expect(kinds.has('white')).toBe(true);
      expect([...shiftTypes]).toEqual(
        expect.arrayContaining([
          'Nappalos 06–18',
          'Nappalos 07–19',
          'Nappalos 10–22',
          '24 órás szolgálat',
          'Éjszakai szolgálat',
          'KMR',
        ]),
      );
    },
  );

  it.skipIf(!hasSamples)(
    'az augusztusi BA7 merge master 709-es fehér stílusát 07:00–19:00 eseménnyé alakítja',
    async () => {
      const session = await loadFirstSample();
      const month = session.months.find((item) => item.year === 2026 && item.month === 8);
      if (!month) throw new Error('Hiányzó augusztusi munkalap.');
      const employee = month.employees.find((item) => item.rows.includes(7));
      if (!employee) throw new Error('A BA7 sorához nem található dolgozó.');

      expect(session.ooxml.themeColors.slice(0, 4)).toEqual([
        '#FFFFFF',
        '#000000',
        '#E7E6E6',
        '#44546A',
      ]);

      const entries = readMonthEntries(session, month, employee.normalizedName, 7);
      const ba7Entry = entries.find((entry) => entry.selectedDiagnostic?.address === 'BA7');
      expect(ba7Entry?.selectedDiagnostic).toMatchObject({
        address: 'BA7',
        rawValue: '12',
        displayedText: '12',
        isMerged: true,
        mergeMaster: 'BA7',
        styleId: 709,
        fillType: 'pattern',
        fillPatternType: 'solid',
        fillForegroundRaw: 'theme=0',
        fillBackgroundRaw: 'argb=FFFFFFFF',
        hasVisibleFill: true,
        fillColor: '#FFFFFF',
      });

      const result = interpretSchedule(entries, { legend: month.legendStyles });
      const row = result.rows.find((item) =>
        item.diagnostics.some((diagnostic) => diagnostic.address === 'BA7'),
      );
      expect(row).toMatchObject({
        shiftType: 'Nappalos 07–19',
        summary: 'OMSZ',
        status: 'Exportálható',
        note: 'Fehér vagy kitöltés nélküli 12 felismerve.',
        diagnostics: [expect.objectContaining({ fillCategory: 'white' })],
        event: {
          shiftTime: {
            start: '2026-08-26T07:00:00',
            end: '2026-08-26T19:00:00',
          },
          calendarTime: {
            start: '2026-08-26T07:00:00',
            end: '2026-08-26T19:00:00',
          },
        },
      });
      if (!row?.event) throw new Error('A BA7-ből nem készült exportálható esemény.');

      const ics = buildIcs([row.event]);
      expect(ics).toContain('DTSTART;TZID=Europe/Budapest:20260826T070000');
      expect(ics).toContain('DTEND;TZID=Europe/Budapest:20260826T190000');

      let requestBody: unknown;
      const fetcher: typeof fetch = (_input, init) => {
        if (typeof init?.body !== 'string') throw new Error('Hiányzó Google request body.');
        requestBody = JSON.parse(init.body) as unknown;
        return Promise.resolve(
          new Response(JSON.stringify({ id: 'created-ba7', colorId: '10' }), {
            status: 200,
            headers: { 'Content-Type': 'application/json' },
          }),
        );
      };
      await new GoogleCalendarClient('token', fetcher).insertEvent('primary', row.event);
      expect(requestBody).toMatchObject({
        summary: 'OMSZ',
        start: { dateTime: '2026-08-26T07:00:00', timeZone: 'Europe/Budapest' },
        end: { dateTime: '2026-08-26T19:00:00', timeZone: 'Europe/Budapest' },
        colorId: '10',
      });
    },
  );
});
