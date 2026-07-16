import { describe, expect, it, vi } from 'vitest';
import type { CalendarEvent } from '../src/domain/types';
import type { AppError } from '../src/domain/errors';
import { GoogleCalendarClient, GoogleTokenSession } from '../src/services/googleCalendar';
import { requestGoogleAccessToken } from '../src/services/googleOAuth';

const item: CalendarEvent = {
  id: 'event-1',
  summary: 'OMSZ',
  shiftType: 'Nappalos 06–18',
  start: '2026-08-10T06:00:00',
  end: '2026-08-10T18:00:00',
  timeZone: 'Europe/Budapest',
};

function response(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

describe('Google Naptár szolgáltatás', () => {
  it('csak pontos summary/start/end egyezést tekint duplikációnak', async () => {
    const fetcher = vi.fn<typeof fetch>().mockResolvedValue(
      response({
        items: [
          {
            summary: 'OMSZ',
            start: { dateTime: '2026-08-10T04:00:00Z' },
            end: { dateTime: '2026-08-10T16:00:00Z' },
          },
        ],
      }),
    );
    await expect(
      new GoogleCalendarClient('token', fetcher).isDuplicate('primary', item),
    ).resolves.toBe(true);
  });

  it('eltérő név vagy időpont nem duplikáció', async () => {
    const fetcher = vi
      .fn<typeof fetch>()
      .mockResolvedValueOnce(
        response({
          items: [
            {
              summary: 'KMR',
              start: { dateTime: '2026-08-10T04:00:00Z' },
              end: { dateTime: '2026-08-10T16:00:00Z' },
            },
          ],
        }),
      )
      .mockResolvedValueOnce(
        response({
          items: [
            {
              summary: 'OMSZ',
              start: { dateTime: '2026-08-10T05:00:00Z' },
              end: { dateTime: '2026-08-10T16:00:00Z' },
            },
          ],
        }),
      );
    const client = new GoogleCalendarClient('token', fetcher);
    await expect(client.isDuplicate('primary', item)).resolves.toBe(false);
    await expect(client.isDuplicate('primary', item)).resolves.toBe(false);
  });

  it('részleges API-hibát eseményenként jelez', async () => {
    const second = {
      ...item,
      id: 'event-2',
      start: '2026-08-11T06:00:00',
      end: '2026-08-11T18:00:00',
    };
    const fetcher = vi
      .fn<typeof fetch>()
      .mockResolvedValueOnce(response({ event: {} }))
      .mockResolvedValueOnce(response({ items: [] }))
      .mockResolvedValueOnce(response({ id: 'created-1' }))
      .mockResolvedValueOnce(response({ items: [] }))
      .mockResolvedValueOnce(response({ error: { message: 'quota' } }, 429));
    const results = await new GoogleCalendarClient('token', fetcher).addEvents('primary', [
      item,
      second,
    ]);
    expect(results.map((result) => result.status)).toEqual(['Létrehozva', 'Sikertelen']);
  });

  it('a legközelebbi sötétzöld eseményszínt választja', async () => {
    const fetcher = vi
      .fn<typeof fetch>()
      .mockResolvedValue(
        response({ event: { '1': { background: '#ff0000' }, '2': { background: '#0a5f3c' } } }),
      );
    await expect(
      new GoogleCalendarClient('token', fetcher).nearestDarkGreenColorId(),
    ).resolves.toBe('2');
  });

  it('konfiguráció nélkül érthető hibát ad', async () => {
    await expect(requestGoogleAccessToken('')).rejects.toMatchObject({
      code: 'GOOGLE_NOT_CONFIGURED',
    } satisfies Partial<AppError>);
  });

  it('kijelentkezéskor memóriából törli és visszavonja a tokent', () => {
    const session = new GoogleTokenSession();
    const revoke = vi.fn();
    session.set('secret-token');
    session.signOut(revoke);
    expect(session.get()).toBeUndefined();
    expect(revoke).toHaveBeenCalledWith('secret-token');
  });
});
