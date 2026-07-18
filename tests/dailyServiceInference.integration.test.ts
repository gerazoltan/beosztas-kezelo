import { describe, expect, it } from 'vitest';
import {
  readEmployeeScheduleEntries,
  readWorksheetScheduleEntries,
} from '../src/excel/dayEntries';
import { parseWorkbook } from '../src/excel/workbookParser';
import { buildDailyServicePatterns } from '../src/services/dailyServiceInference';
import { GoogleCalendarClient } from '../src/services/googleCalendar';
import { buildIcs } from '../src/services/ics';
import { interpretSchedule } from '../src/services/shifts';
import { dailyInferenceWorkbookBuffer } from './fixtures/syntheticWorkbook';

function response(body: unknown): Response {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
  });
}

function requestDescription(body: unknown): string {
  if (!body || typeof body !== 'object' || !('description' in body)) {
    throw new Error('Hiányzó Google eseményleírás.');
  }
  const description = body.description;
  if (typeof description !== 'string') throw new Error('Érvénytelen Google eseményleírás.');
  return description;
}

describe('teljes munkalapos következtetés Excel-integrációja', () => {
  it.each([
    {
      presentVehicle: 'blue' as const,
      target: 'tenCar',
      category: '10-es kocsi',
      start: '2026-08-01T10:00:00',
      end: '2026-08-01T22:00:00',
      icsStart: '20260801T100000',
      icsEnd: '20260801T220000',
    },
    {
      presentVehicle: 'tenCar' as const,
      target: 'blue',
      category: '6-os kocsi',
      start: '2026-08-01T06:00:00',
      end: '2026-08-01T18:00:00',
      icsStart: '20260801T060000',
      icsEnd: '20260801T180000',
    },
  ])(
    'a nem első, kijelölt dolgozó fekete 12-esét $category szolgálattá korrigálja',
    async ({ presentVehicle, target, category, start, end, icsStart, icsEnd }) => {
      const session = await parseWorkbook(
        await dailyInferenceWorkbookBuffer(presentVehicle),
        'napi-osszefugges.xlsx',
      );
      const month = session.months[0];
      if (!month) throw new Error('Hiányzó teszthónap.');

      const worksheetSchedules = readWorksheetScheduleEntries(session, month);
      expect(worksheetSchedules.map((schedule) => schedule.row)).toContain(7);
      expect(worksheetSchedules).toHaveLength(3);

      const dailyServicePatterns = buildDailyServicePatterns(worksheetSchedules);
      expect(dailyServicePatterns.get('2026-08-01')?.correction).toMatchObject({
        candidateAddress: 'C7',
        target,
      });

      const selected = readEmployeeScheduleEntries(
        session,
        month,
        'jelölt dolgozó',
        7,
      );
      const result = interpretSchedule(selected.current, {
        previous: selected.previous,
        next: selected.next,
        dailyServicePatterns,
      });
      const event = result.events[0];
      if (!event) throw new Error('Nem készült következtetett esemény.');

      expect(event).toMatchObject({
        summary: 'OMSZ',
        serviceCategory: category,
        calendarTime: { start, end },
        inference: { source: 'daily-service-pattern', target },
      });
      expect(result.rows[0]?.note).toContain('napi 24 órás Parti szolgálat');
      expect(result.rows[0]?.diagnostics[0]).toMatchObject({
        address: 'C7',
        fontColor: '#000000',
        underline: false,
        fillColor: '#FFF2CC',
      });

      const ics = buildIcs([event]);
      expect(ics).toContain(`DTSTART;TZID=Europe/Budapest:${icsStart}`);
      expect(ics).toContain(`DTEND;TZID=Europe/Budapest:${icsEnd}`);
      expect(ics).toContain(`DESCRIPTION:Szolgálati jelleg: ${category}`);
      expect(ics).toContain('napi szolgálati összeállításból lett következtetve');

      let requestBody: unknown;
      const fetcher: typeof fetch = (_input, init) => {
        if (typeof init?.body !== 'string') throw new Error('Hiányzó Google request body.');
        requestBody = JSON.parse(init.body) as unknown;
        return Promise.resolve(response({ id: 'created-inferred', colorId: '10' }));
      };
      await new GoogleCalendarClient('token', fetcher).insertEvent('primary', event);
      expect(requestBody).toMatchObject({
        summary: 'OMSZ',
        start: { dateTime: start, timeZone: 'Europe/Budapest' },
        end: { dateTime: end, timeZone: 'Europe/Budapest' },
        colorId: '10',
      });
      expect(requestDescription(requestBody)).toContain(`Szolgálati jelleg: ${category}`);
      expect(requestDescription(requestBody)).toContain(
        'napi szolgálati összeállításból lett következtetve',
      );
    },
  );
});
