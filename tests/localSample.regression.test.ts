import { readdir, readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import type { WorkbookSession } from '../src/domain/types';
import { parseWorkbook } from '../src/excel/workbookParser';
import {
  readMonthEntries,
  readWorksheetScheduleEntries,
} from '../src/excel/dayEntries';
import { classifyTwelve, interpretSchedule } from '../src/services/shifts';
import { AppError } from '../src/domain/errors';
import { buildDailyServicePatterns } from '../src/services/dailyServiceInference';
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
    'felismeri a valós havi szerkezetet és a fontalapú szolgálati kategóriákat',
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
              kinds.add(classifyTwelve(entry.selectedDiagnostic));
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
      expect(kinds.has('tenCar')).toBe(true);
      expect(kinds.has('party')).toBe(true);
      expect(kinds.has('emergency')).toBe(true);
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
    'az augusztusi BA7 merge master fekete 12-esét Parti szolgálattá alakítja',
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
        fontColorRaw: undefined,
        fontColor: '#000000',
        underline: false,
        positionInDayGroup: 1,
      });

      const result = interpretSchedule(entries, { legend: month.legendStyles });
      const row = result.rows.find((item) =>
        item.diagnostics.some((diagnostic) => diagnostic.address === 'BA7'),
      );
      expect(row).toMatchObject({
        shiftType: 'Nappalos 07–19',
        serviceCategory: 'Parti szolgálat',
        summary: 'OMSZ',
        status: 'Exportálható',
        note: 'Fekete 12 felismerve: Parti szolgálat.',
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
      expect(ics).toContain('DESCRIPTION:Szolgálati jelleg: Parti szolgálat');

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
        description: 'Szolgálati jelleg: Parti szolgálat',
        start: { dateTime: '2026-08-26T07:00:00', timeZone: 'Europe/Budapest' },
        end: { dateTime: '2026-08-26T19:00:00', timeZone: 'Europe/Budapest' },
        colorId: '10',
      });
    },
  );

  it.skipIf(!hasSamples)(
    'a valós fájl fekete 12-eseit eltérő naptári háttereken is Parti szolgálatként kezeli',
    async () => {
      const session = await loadFirstSample();
      const month = session.months.find((item) => item.month === 1);
      if (!month) throw new Error('Hiányzó januári munkalap.');

      const examples = [
        { row: 14, address: 'W14' },
        { row: 15, address: 'G15' },
      ];
      const fillColors = new Set<string>();
      for (const example of examples) {
        const employee = month.employees.find((item) => item.rows.includes(example.row));
        if (!employee) throw new Error(`Hiányzó dolgozói sor: ${example.row}.`);
        const entries = readMonthEntries(
          session,
          month,
          employee.normalizedName,
          example.row,
        );
        const result = interpretSchedule(entries, { legend: month.legendStyles });
        const row = result.rows.find((item) =>
          item.diagnostics.some((diagnostic) => diagnostic.address === example.address),
        );
        expect(row?.marker).toBe('12');
        expect(row?.status).toBe('Exportálható');
        expect(row?.serviceCategory).toBe('Parti szolgálat');
        expect(row?.event?.shiftTime.start).toContain('T07:00');
        const diagnostic = row?.diagnostics.find((item) => item.address === example.address);
        expect(diagnostic?.fontColor).toBe('#000000');
        if (diagnostic?.fillColor) fillColors.add(diagnostic.fillColor);
      }
      expect(fillColors.size).toBe(2);
    },
  );

  it.skipIf(!hasSamples)(
    'a valós fájl napi következtetései kizárólag az egyértelmű teljes munkalapos mintákra épülnek',
    async () => {
      const session = await loadFirstSample();
      const corrections = session.months.flatMap((month) =>
        [...buildDailyServicePatterns(readWorksheetScheduleEntries(session, month)).values()]
          .filter((pattern) => pattern.correction)
          .map((pattern) => ({ month: month.month, pattern })),
      );

      expect(corrections.length).toBeGreaterThan(0);
      for (const { pattern } of corrections) {
        expect(pattern.partyTwentyFourHourCount).toBe(1);
        expect(pattern.blackTwelveCandidateCount).toBe(1);
        expect(pattern.conflictingServiceMarkerCount).toBe(0);
        expect(
          (pattern.blueTwelveCount === 1 && pattern.tenCarTwelveCount === 0) ||
            (pattern.blueTwelveCount === 0 && pattern.tenCarTwelveCount === 1),
        ).toBe(true);
      }
    },
  );
});
