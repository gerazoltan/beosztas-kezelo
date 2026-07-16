import { readdir, readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { parseWorkbook } from '../src/excel/workbookParser';
import { readMonthEntries } from '../src/excel/dayEntries';
import { classifyTwelve, interpretSchedule } from '../src/services/shifts';
import { AppError } from '../src/domain/errors';

const sampleDirectory = resolve(process.cwd(), 'local-samples');
const hasSamples = existsSync(sampleDirectory);

describe('helyi, nem követett Excel-minta regresszió', () => {
  it.skipIf(!hasSamples)(
    'felismeri a valós havi szerkezetet és mindkét bizonyított 12-stílust',
    async () => {
      const files = (await readdir(sampleDirectory)).filter((name) =>
        name.toLowerCase().endsWith('.xlsx'),
      );
      expect(files.length).toBeGreaterThan(0);
      const fileName = files[0];
      if (!fileName) throw new Error('Hiányzó helyi mintafájl.');
      const bytes = await readFile(resolve(sampleDirectory, fileName));
      const buffer = Uint8Array.from(bytes).buffer;
      let session;
      try {
        session = await parseWorkbook(buffer, fileName);
      } catch (error) {
        if (error instanceof AppError) throw new Error(error.technicalDetails ?? error.message);
        throw error;
      }
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
      expect([...shiftTypes]).toEqual(
        expect.arrayContaining([
          'Nappalos 06–18',
          'Nappalos 10–22',
          '24 órás szolgálat',
          'Éjszakai szolgálat',
          'KMR',
        ]),
      );
    },
  );
});
