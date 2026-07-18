import { afterEach, describe, expect, it, vi } from 'vitest';
import type { CalendarEvent } from '../src/domain/types';
import type { AppError } from '../src/domain/errors';
import { GoogleCalendarClient, GoogleTokenSession } from '../src/services/googleCalendar';
import { requestGoogleAccessToken } from '../src/services/googleOAuth';

const item: CalendarEvent = {
  id: 'event-1',
  summary: 'OMSZ',
  shiftType: 'Nappalos 06–18',
  shiftTime: { start: '2026-08-10T06:00:00', end: '2026-08-10T18:00:00' },
  calendarTime: { start: '2026-08-10T06:00:00', end: '2026-08-10T18:00:00' },
  timeZone: 'Europe/Budapest',
};

const kmrItem: CalendarEvent = {
  ...item,
  id: 'event-kmr',
  summary: 'KMR',
  shiftType: 'KMR',
  shiftTime: { start: '2026-08-11T05:00:00', end: '2026-08-12T01:00:00' },
  calendarTime: { start: '2026-08-11T05:00:00', end: '2026-08-12T01:00:00' },
};

const twentyFourHourItem: CalendarEvent = {
  ...item,
  id: 'event-24-hour',
  shiftType: '24 órás szolgálat',
  shiftTime: { start: '2026-08-31T07:00:00', end: '2026-09-01T07:00:00' },
  calendarTime: { start: '2026-08-31T07:00:00', end: '2026-09-01T06:59:00' },
};

const whiteTwelveItem: CalendarEvent = {
  ...item,
  id: 'event-white-12',
  shiftType: 'Nappalos 07–19',
  shiftTime: { start: '2026-08-12T07:00:00', end: '2026-08-12T19:00:00' },
  calendarTime: { start: '2026-08-12T07:00:00', end: '2026-08-12T19:00:00' },
};

function response(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

describe('Google Naptár szolgáltatás', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('a böngészős fetch-et window kontextussal és Bearer tokennel hívja', async () => {
    let receivedHeaders: Headers | undefined;
    const strictWindowFetch = vi.fn(function (
      this: typeof globalThis,
      _input: RequestInfo | URL,
      init?: RequestInit,
    ): Promise<Response> {
      if (this !== window) throw new TypeError('Illegal invocation');
      receivedHeaders = new Headers(init?.headers);
      return Promise.resolve(
        response({
          items: [{ id: 'primary', summary: 'Elsődleges', accessRole: 'owner' }],
        }),
      );
    });
    vi.stubGlobal('fetch', strictWindowFetch);

    const calendars = await new GoogleCalendarClient(
      'regression-access-token',
    ).listWritableCalendars();

    expect(calendars).toHaveLength(1);
    expect(strictWindowFetch).toHaveBeenCalledOnce();
    expect(receivedHeaders?.get('Authorization')).toBe('Bearer regression-access-token');
  });

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

  it('17–7 esetén kizárólag a 06:59-es naptári befejezést tekinti duplikációnak', async () => {
    const fetcher = vi
      .fn<typeof fetch>()
      .mockResolvedValueOnce(
        response({
          items: [
            {
              summary: 'OMSZ',
              start: { dateTime: '2026-08-31T07:00:00+02:00' },
              end: { dateTime: '2026-09-01T06:59:00+02:00' },
            },
          ],
        }),
      )
      .mockResolvedValueOnce(
        response({
          items: [
            {
              summary: 'OMSZ',
              start: { dateTime: '2026-08-31T07:00:00+02:00' },
              end: { dateTime: '2026-09-01T06:55:00+02:00' },
            },
          ],
        }),
      )
      .mockResolvedValueOnce(
        response({
          items: [
            {
              summary: 'OMSZ',
              start: { dateTime: '2026-08-31T07:00:00+02:00' },
              end: { dateTime: '2026-09-01T07:00:00+02:00' },
            },
          ],
        }),
      );
    const client = new GoogleCalendarClient('token', fetcher);

    await expect(client.isDuplicate('primary', twentyFourHourItem)).resolves.toBe(true);
    await expect(client.isDuplicate('primary', twentyFourHourItem)).resolves.toBe(false);
    await expect(client.isDuplicate('primary', twentyFourHourItem)).resolves.toBe(false);
  });

  it('17–7 feltöltésnél a request body a 07:00–másnap 06:59 naptári időt használja', async () => {
    let requestBody: unknown;
    const fetcher = vi.fn<typeof fetch>().mockImplementation((_input, init) => {
      if (typeof init?.body !== 'string') throw new Error('Hiányzó JSON request body.');
      requestBody = JSON.parse(init.body) as unknown;
      return Promise.resolve(response({ id: 'created-24-hour', colorId: '10' }));
    });

    await new GoogleCalendarClient('token', fetcher).insertEvent('primary', twentyFourHourItem);

    expect(requestBody).toMatchObject({
      summary: 'OMSZ',
      start: {
        dateTime: '2026-08-31T07:00:00',
        timeZone: 'Europe/Budapest',
      },
      end: {
        dateTime: '2026-09-01T06:59:00',
        timeZone: 'Europe/Budapest',
      },
    });
  });

  it('fehér 12 feltöltésnél a request body a 07:00–19:00 naptári időt használja', async () => {
    let requestBody: unknown;
    const fetcher = vi.fn<typeof fetch>().mockImplementation((_input, init) => {
      if (typeof init?.body !== 'string') throw new Error('Hiányzó JSON request body.');
      requestBody = JSON.parse(init.body) as unknown;
      return Promise.resolve(response({ id: 'created-white-12', colorId: '10' }));
    });

    await new GoogleCalendarClient('token', fetcher).insertEvent('primary', whiteTwelveItem);

    expect(requestBody).toMatchObject({
      summary: 'OMSZ',
      start: {
        dateTime: '2026-08-12T07:00:00',
        timeZone: 'Europe/Budapest',
      },
      end: {
        dateTime: '2026-08-12T19:00:00',
        timeZone: 'Europe/Budapest',
      },
      colorId: '10',
    });
  });

  it('részleges API-hibát eseményenként jelez', async () => {
    const second = {
      ...item,
      id: 'event-2',
      shiftTime: { start: '2026-08-11T06:00:00', end: '2026-08-11T18:00:00' },
      calendarTime: { start: '2026-08-11T06:00:00', end: '2026-08-11T18:00:00' },
    };
    const fetcher = vi
      .fn<typeof fetch>()
      .mockResolvedValueOnce(response({ items: [] }))
      .mockResolvedValueOnce(response({ id: 'created-1', colorId: '10' }))
      .mockResolvedValueOnce(response({ items: [] }))
      .mockResolvedValueOnce(response({ error: { message: 'quota' } }, 429));
    const results = await new GoogleCalendarClient('token', fetcher).addEvents('primary', [
      item,
      second,
    ]);
    expect(results.map((result) => result.status)).toEqual(['Létrehozva', 'Sikertelen']);
  });

  it('az eseményenkénti kezdő- és eredmény callbacket folyamatosan meghívja', async () => {
    const onStart = vi.fn();
    const onResult = vi.fn();
    const fetcher = vi
      .fn<typeof fetch>()
      .mockResolvedValueOnce(response({ items: [] }))
      .mockResolvedValueOnce(response({ id: 'created-1', colorId: '10' }))
      .mockResolvedValueOnce(response({ items: [] }))
      .mockResolvedValueOnce(response({ id: 'created-2', colorId: '10' }));

    await new GoogleCalendarClient('token', fetcher).addEvents('primary', [item, kmrItem], {
      onStart,
      onResult,
    });

    expect(onStart).toHaveBeenNthCalledWith(1, item);
    expect(onStart).toHaveBeenNthCalledWith(2, kmrItem);
    expect(onResult).toHaveBeenNthCalledWith(1, expect.objectContaining({ eventId: 'event-1' }));
    expect(onResult).toHaveBeenNthCalledWith(2, expect.objectContaining({ eventId: 'event-kmr' }));
  });

  it('lejárt tokennél megáll, és a hátralévő eseményt nem jelöli sikertelennek', async () => {
    const onResult = vi.fn();
    const fetcher = vi.fn<typeof fetch>().mockResolvedValue(response({ error: 'expired' }, 401));

    const results = await new GoogleCalendarClient('token', fetcher).addEvents(
      'primary',
      [item, kmrItem],
      { onResult },
    );

    expect(results).toHaveLength(1);
    expect(results[0]).toMatchObject({
      eventId: 'event-1',
      status: 'Sikertelen',
      errorCode: 'GOOGLE_TOKEN_EXPIRED',
    });
    expect(onResult).toHaveBeenCalledOnce();
    expect(fetcher).toHaveBeenCalledOnce();
  });

  it.each([
    {
      name: 'hálózati hibát',
      fetcher: () => vi.fn<typeof fetch>().mockRejectedValue(new TypeError('offline')),
      errorCode: 'GOOGLE_NETWORK_ERROR',
    },
    {
      name: 'nem írható naptárt',
      fetcher: () => vi.fn<typeof fetch>().mockResolvedValue(response({ error: 'forbidden' }, 403)),
      errorCode: 'GOOGLE_CALENDAR_NOT_WRITABLE',
    },
    {
      name: 'API-hibát',
      fetcher: () => vi.fn<typeof fetch>().mockResolvedValue(response({ error: 'backend' }, 500)),
      errorCode: 'GOOGLE_API_ERROR',
    },
  ] as const)('külön eredménykóddal jelzi: $name', async ({ fetcher, errorCode }) => {
    const [result] = await new GoogleCalendarClient('token', fetcher()).addEvents('primary', [
      item,
    ]);
    expect(result).toMatchObject({ status: 'Sikertelen', errorCode });
  });

  it.each([
    ['OMSZ', item],
    ['KMR', kmrItem],
  ] as const)(
    'a(z) %s esemény request body-jában explicit colorId 10-et küld, és kiolvassa a válaszból',
    async (_eventType, calendarEvent) => {
      let requestBody: unknown;
      const fetcher = vi.fn<typeof fetch>().mockImplementation((_input, init) => {
        if (typeof init?.body !== 'string') throw new Error('Hiányzó JSON request body.');
        requestBody = JSON.parse(init.body) as unknown;
        return Promise.resolve(response({ id: `created-${calendarEvent.id}`, colorId: '10' }));
      });

      const created = await new GoogleCalendarClient('token', fetcher).insertEvent(
        'primary',
        calendarEvent,
      );

      expect(requestBody).toMatchObject({
        summary: calendarEvent.summary,
        colorId: '10',
      });
      expect(created.colorId).toBe('10');
    },
  );

  it('jelzi, ha a Google válasza nem a kért colorId 10-et igazolja vissza', async () => {
    const fetcher = vi
      .fn<typeof fetch>()
      .mockResolvedValueOnce(response({ items: [] }))
      .mockResolvedValueOnce(response({ id: 'created-1', colorId: '5' }));

    const [result] = await new GoogleCalendarClient('token', fetcher).addEvents('primary', [item]);

    expect(result).toMatchObject({ status: 'Létrehozva' });
    expect(result?.message).toContain('nem a kért zöld Basil (10) színt igazolta vissza');
  });

  it('a már létező eseményt nem írja és nem színezi át', async () => {
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

    const [result] = await new GoogleCalendarClient('token', fetcher).addEvents('primary', [item]);

    expect(result?.status).toBe('Már szerepel a naptárban');
    expect(fetcher).toHaveBeenCalledOnce();
    expect(fetcher.mock.calls[0]?.[1]?.method).toBeUndefined();
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
