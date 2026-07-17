import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type { CalendarEvent } from '../src/domain/types';
import { GooglePanel } from '../src/components/GooglePanel';

const omszEvent: CalendarEvent = {
  id: 'event-omsz',
  summary: 'OMSZ',
  shiftType: 'Nappalos 06–18',
  shiftTime: { start: '2026-08-10T06:00:00', end: '2026-08-10T18:00:00' },
  calendarTime: { start: '2026-08-10T06:00:00', end: '2026-08-10T18:00:00' },
  timeZone: 'Europe/Budapest',
};

const kmrEvent: CalendarEvent = {
  id: 'event-kmr',
  summary: 'KMR',
  shiftType: 'KMR',
  shiftTime: { start: '2026-08-11T05:00:00', end: '2026-08-12T01:00:00' },
  calendarTime: { start: '2026-08-11T05:00:00', end: '2026-08-12T01:00:00' },
  timeZone: 'Europe/Budapest',
};

const revokeToken = vi.fn();

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

function requestUrl(input: RequestInfo | URL): string {
  if (typeof input === 'string') return input;
  if (input instanceof URL) return input.href;
  return input.url;
}

function requestSummary(init?: RequestInit): string | undefined {
  if (typeof init?.body !== 'string') return undefined;
  const parsed = JSON.parse(init.body) as unknown;
  if (typeof parsed !== 'object' || parsed === null || !('summary' in parsed)) return undefined;
  return typeof parsed.summary === 'string' ? parsed.summary : undefined;
}

function installGoogleIdentity(): void {
  window.google = {
    accounts: {
      oauth2: {
        initTokenClient: ({ callback }) => ({
          requestAccessToken: () => callback({ access_token: 'panel-access-token' }),
        }),
        revoke: revokeToken,
      },
    },
  };
}

function calendars(items = [{ id: 'primary', summary: 'Teszt naptár', accessRole: 'owner' }]) {
  return jsonResponse({ items });
}

function renderGooglePanel(events: CalendarEvent[] = [omszEvent, kmrEvent], resetKey = 0) {
  const callbacks = {
    onEventStart: vi.fn(),
    onResult: vi.fn(),
    onCalendarChange: vi.fn(),
    onNewSchedule: vi.fn(),
  };
  const view = render(<GooglePanel events={events} resetKey={resetKey} {...callbacks} />);
  return { ...view, callbacks };
}

async function signIn(user: ReturnType<typeof userEvent.setup>): Promise<void> {
  await user.click(screen.getByRole('button', { name: 'Google-bejelentkezés' }));
  await screen.findByLabelText('Írható naptár');
}

async function startUpload(
  user: ReturnType<typeof userEvent.setup>,
  eventCount: number,
): Promise<void> {
  await user.click(
    screen.getByRole('button', {
      name: `${eventCount} kijelölt esemény hozzáadása a Google Naptárhoz`,
    }),
  );
}

describe('Google feltöltési panel', () => {
  beforeEach(() => {
    vi.stubEnv('VITE_GOOGLE_CLIENT_ID', 'test-client-id');
    revokeToken.mockReset();
    installGoogleIdentity();
  });

  afterEach(() => {
    vi.unstubAllEnvs();
    vi.unstubAllGlobals();
    window.google = undefined;
    vi.restoreAllMocks();
  });

  it('a kijelölt eseményszámot mutatja, feltöltéskor tiltott, és dupla kattintásra is csak egyszer indul', async () => {
    const user = userEvent.setup();
    let resolveDuplicateCheck: (response: Response) => void = () => undefined;
    const duplicateCheck = new Promise<Response>((resolve) => {
      resolveDuplicateCheck = resolve;
    });
    const fetcher = vi.fn<typeof fetch>().mockImplementation((input, init) => {
      const url = requestUrl(input);
      if (url.endsWith('/users/me/calendarList')) return Promise.resolve(calendars());
      if (init?.method === 'POST') {
        return Promise.resolve(jsonResponse({ id: 'created', colorId: '10' }));
      }
      return duplicateCheck;
    });
    vi.stubGlobal('fetch', fetcher);
    renderGooglePanel([omszEvent]);
    await signIn(user);

    const uploadButton = screen.getByRole('button', {
      name: '1 kijelölt esemény hozzáadása a Google Naptárhoz',
    });
    fireEvent.click(uploadButton);
    fireEvent.click(uploadButton);

    const progressButton = await screen.findByRole('button', {
      name: /^Feltöltés folyamatban… 0 \/ 1$/u,
    });
    expect(progressButton).toBeDisabled();
    expect(
      fetcher.mock.calls.filter(([input]) => requestUrl(input).includes('/events?')),
    ).toHaveLength(1);

    resolveDuplicateCheck(jsonResponse({ items: [] }));
    expect(await screen.findByRole('heading', { name: 'Sikeres naptárfeltöltés' })).toBeVisible();
  });

  it('teljes siker után eltünteti a feltöltési gombot, összesít és helyes címen nyitja a naptárt', async () => {
    const user = userEvent.setup();
    const fetcher = vi.fn<typeof fetch>().mockImplementation((input, init) => {
      if (requestUrl(input).endsWith('/users/me/calendarList')) return Promise.resolve(calendars());
      if (init?.method === 'POST') {
        return Promise.resolve(jsonResponse({ id: 'created', colorId: '10' }));
      }
      return Promise.resolve(jsonResponse({ items: [] }));
    });
    vi.stubGlobal('fetch', fetcher);
    renderGooglePanel();
    await signIn(user);

    await startUpload(user, 2);

    expect(await screen.findByRole('heading', { name: 'Sikeres naptárfeltöltés' })).toBeVisible();
    expect(screen.getByText('2 esemény létrehozva.')).toBeVisible();
    expect(screen.getByText('0 esemény már szerepelt a naptárban.')).toBeVisible();
    expect(screen.getByText('0 sikertelen művelet.')).toBeVisible();
    expect(screen.getByText(/Kiválasztott naptár:/u)).toHaveTextContent('Teszt naptár');
    expect(
      screen.queryByRole('button', { name: /kijelölt esemény hozzáadása/u }),
    ).not.toBeInTheDocument();

    const calendarLink = screen.getByRole('link', { name: 'Google Naptár megnyitása' });
    const clickListener = vi.fn((event: Event) => event.preventDefault());
    calendarLink.addEventListener('click', clickListener);
    expect(calendarLink).toHaveClass('button', 'tertiary');
    expect(calendarLink).toHaveAttribute('href', 'https://calendar.google.com/calendar/u/0/r');
    expect(calendarLink).toHaveAttribute('target', '_blank');
    expect(calendarLink).toHaveAttribute('rel', 'noopener noreferrer');

    await user.click(calendarLink);
    expect(clickListener).toHaveBeenCalledOnce();
  });

  it('részleges siker után kizárólag a sikertelen eseményt próbálja újra', async () => {
    const user = userEvent.setup();
    const postCounts = new Map<string, number>();
    const fetcher = vi.fn<typeof fetch>().mockImplementation((input, init) => {
      if (requestUrl(input).endsWith('/users/me/calendarList')) return Promise.resolve(calendars());
      if (init?.method !== 'POST') return Promise.resolve(jsonResponse({ items: [] }));
      const summary = requestSummary(init) ?? '';
      const count = (postCounts.get(summary) ?? 0) + 1;
      postCounts.set(summary, count);
      if (summary === 'KMR' && count === 1) {
        return Promise.resolve(jsonResponse({ error: { message: 'quota' } }, 500));
      }
      return Promise.resolve(jsonResponse({ id: `created-${summary}`, colorId: '10' }));
    });
    vi.stubGlobal('fetch', fetcher);
    renderGooglePanel();
    await signIn(user);

    await startUpload(user, 2);
    expect(
      await screen.findByRole('heading', { name: 'Részben sikeres naptárfeltöltés' }),
    ).toBeVisible();
    expect(screen.getByText('1 esemény létrehozva.')).toBeVisible();
    expect(screen.getByText('1 sikertelen művelet.')).toBeVisible();
    expect(
      screen.queryByRole('button', { name: /kijelölt esemény hozzáadása/u }),
    ).not.toBeInTheDocument();

    await user.click(screen.getByRole('button', { name: 'Csak a sikertelenek újrapróbálása' }));
    expect(await screen.findByRole('heading', { name: 'Sikeres naptárfeltöltés' })).toBeVisible();
    expect(postCounts.get('OMSZ')).toBe(1);
    expect(postCounts.get('KMR')).toBe(2);
  });

  it('a duplikált eseményt teljes sikernek, nem hibának számolja', async () => {
    const user = userEvent.setup();
    const fetcher = vi.fn<typeof fetch>().mockImplementation((input, init) => {
      if (requestUrl(input).endsWith('/users/me/calendarList')) return Promise.resolve(calendars());
      if (init?.method === 'POST') throw new Error('Duplikátumhoz nem indulhat POST.');
      return Promise.resolve(
        jsonResponse({
          items: [
            {
              summary: 'OMSZ',
              start: { dateTime: '2026-08-10T04:00:00Z' },
              end: { dateTime: '2026-08-10T16:00:00Z' },
            },
          ],
        }),
      );
    });
    vi.stubGlobal('fetch', fetcher);
    renderGooglePanel([omszEvent]);
    await signIn(user);

    await startUpload(user, 1);

    expect(await screen.findByRole('heading', { name: 'Sikeres naptárfeltöltés' })).toBeVisible();
    expect(screen.getByText('0 esemény létrehozva.')).toBeVisible();
    expect(screen.getByText('1 esemény már szerepelt a naptárban.')).toBeVisible();
    expect(screen.getByText('0 sikertelen művelet.')).toBeVisible();
  });

  it('teljes hiba után érthető hibát, technikai részleteket és újrapróbálást mutat', async () => {
    const user = userEvent.setup();
    const fetcher = vi.fn<typeof fetch>().mockImplementation((input, init) => {
      if (requestUrl(input).endsWith('/users/me/calendarList')) return Promise.resolve(calendars());
      if (init?.method === 'POST') {
        return Promise.resolve(jsonResponse({ error: { message: 'backend error' } }, 500));
      }
      return Promise.resolve(jsonResponse({ items: [] }));
    });
    vi.stubGlobal('fetch', fetcher);
    renderGooglePanel([omszEvent]);
    await signIn(user);

    await startUpload(user, 1);

    expect(
      await screen.findByRole('heading', { name: 'Sikertelen naptárfeltöltés' }),
    ).toBeVisible();
    expect(screen.getByText('A Google Naptár API hibát jelzett.')).toBeVisible();
    expect(screen.getByRole('button', { name: 'Újrapróbálás' })).toBeEnabled();
    expect(screen.getByText('Technikai részletek')).toBeVisible();
  });

  it('lejárt tokennél új bejelentkezést kér, és a hátralévő eseményt függőben hagyja', async () => {
    const user = userEvent.setup();
    const fetcher = vi.fn<typeof fetch>().mockImplementation((input) => {
      if (requestUrl(input).endsWith('/users/me/calendarList')) return Promise.resolve(calendars());
      return Promise.resolve(jsonResponse({ error: 'expired' }, 401));
    });
    vi.stubGlobal('fetch', fetcher);
    const { callbacks } = renderGooglePanel();
    await signIn(user);

    await startUpload(user, 2);

    expect(
      await screen.findByRole('button', { name: 'Google-bejelentkezés megújítása' }),
    ).toBeVisible();
    expect(screen.getByText('1 sikertelen művelet.')).toBeVisible();
    expect(screen.getByText('1 esemény még nem került feldolgozásra.')).toBeVisible();
    expect(callbacks.onResult).toHaveBeenCalledOnce();
  });

  it('resetKey vagy másik naptár választása törli az eredménykártyát', async () => {
    const user = userEvent.setup();
    const availableCalendars = [
      { id: 'primary', summary: 'Első naptár', accessRole: 'owner' },
      { id: 'secondary', summary: 'Második naptár', accessRole: 'writer' },
    ];
    const fetcher = vi.fn<typeof fetch>().mockImplementation((input, init) => {
      if (requestUrl(input).endsWith('/users/me/calendarList')) {
        return Promise.resolve(calendars(availableCalendars));
      }
      if (init?.method === 'POST') {
        return Promise.resolve(jsonResponse({ id: 'created', colorId: '10' }));
      }
      return Promise.resolve(jsonResponse({ items: [] }));
    });
    vi.stubGlobal('fetch', fetcher);
    const { rerender, callbacks } = renderGooglePanel([omszEvent]);
    await signIn(user);
    await startUpload(user, 1);
    await screen.findByRole('heading', { name: 'Sikeres naptárfeltöltés' });

    await user.selectOptions(screen.getByLabelText('Írható naptár'), 'secondary');
    expect(callbacks.onCalendarChange).toHaveBeenCalledOnce();
    expect(screen.queryByText('Sikeres naptárfeltöltés')).not.toBeInTheDocument();

    rerender(
      <GooglePanel
        events={[omszEvent]}
        resetKey={1}
        onEventStart={callbacks.onEventStart}
        onResult={callbacks.onResult}
        onCalendarChange={callbacks.onCalendarChange}
        onNewSchedule={callbacks.onNewSchedule}
      />,
    );
    await waitFor(() =>
      expect(screen.getByRole('button', { name: /1 kijelölt esemény hozzáadása/u })).toBeVisible(),
    );
  });

  it('teljes siker után egy újonnan kijelölt eseményhez új feltöltési kört enged', async () => {
    const user = userEvent.setup();
    const fetcher = vi.fn<typeof fetch>().mockImplementation((input, init) => {
      if (requestUrl(input).endsWith('/users/me/calendarList')) return Promise.resolve(calendars());
      if (init?.method === 'POST') {
        return Promise.resolve(jsonResponse({ id: 'created', colorId: '10' }));
      }
      return Promise.resolve(jsonResponse({ items: [] }));
    });
    vi.stubGlobal('fetch', fetcher);
    const callbacks = {
      onEventStart: vi.fn(),
      onResult: vi.fn(),
      onCalendarChange: vi.fn(),
      onNewSchedule: vi.fn(),
    };
    const { rerender } = render(<GooglePanel events={[omszEvent]} resetKey={0} {...callbacks} />);
    await signIn(user);
    await startUpload(user, 1);
    await screen.findByRole('heading', { name: 'Sikeres naptárfeltöltés' });

    rerender(<GooglePanel events={[kmrEvent]} resetKey={0} {...callbacks} />);

    expect(
      await screen.findByRole('button', {
        name: '1 kijelölt esemény hozzáadása a Google Naptárhoz',
      }),
    ).toBeVisible();
    expect(screen.queryByText('Sikeres naptárfeltöltés')).not.toBeInTheDocument();
  });

  it('kijelentkezéskor törli a hitelesítési adatokat, de megőrzi az összesítést', async () => {
    const user = userEvent.setup();
    let calendarListRequests = 0;
    const fetcher = vi.fn<typeof fetch>().mockImplementation((input, init) => {
      if (requestUrl(input).endsWith('/users/me/calendarList')) {
        calendarListRequests += 1;
        return Promise.resolve(calendars());
      }
      if (init?.method === 'POST') {
        return Promise.resolve(jsonResponse({ id: 'created', colorId: '10' }));
      }
      return Promise.resolve(jsonResponse({ items: [] }));
    });
    vi.stubGlobal('fetch', fetcher);
    renderGooglePanel([omszEvent]);
    await signIn(user);
    await startUpload(user, 1);
    await screen.findByRole('heading', { name: 'Sikeres naptárfeltöltés' });

    await user.click(screen.getByRole('button', { name: 'Kijelentkezés' }));

    expect(revokeToken).toHaveBeenCalledWith('panel-access-token');
    expect(screen.queryByLabelText('Írható naptár')).not.toBeInTheDocument();
    expect(screen.getByRole('heading', { name: 'Sikeres naptárfeltöltés' })).toBeVisible();
    expect(screen.getByText('1 esemény létrehozva.')).toBeVisible();
    expect(
      screen.queryByRole('link', { name: 'Google Naptár megnyitása' }),
    ).not.toBeInTheDocument();
    expect(
      screen.getByText(
        'Kijelentkeztél a Google Naptárból. A korábban létrehozott események a naptárban maradnak.',
      ),
    ).toBeVisible();

    await user.click(screen.getByRole('button', { name: 'Másik fiók csatlakoztatása' }));
    expect(await screen.findByLabelText('Írható naptár')).toBeVisible();
    expect(calendarListRequests).toBe(2);
  });

  it('kijelentkezés után nem engedi a sikertelen esemény API-újrapróbálását', async () => {
    const user = userEvent.setup();
    const postCounts = new Map<string, number>();
    const fetcher = vi.fn<typeof fetch>().mockImplementation((input, init) => {
      if (requestUrl(input).endsWith('/users/me/calendarList')) return Promise.resolve(calendars());
      if (init?.method !== 'POST') return Promise.resolve(jsonResponse({ items: [] }));
      const summary = requestSummary(init) ?? '';
      postCounts.set(summary, (postCounts.get(summary) ?? 0) + 1);
      if (summary === 'KMR') {
        return Promise.resolve(jsonResponse({ error: { message: 'quota' } }, 500));
      }
      return Promise.resolve(jsonResponse({ id: 'created', colorId: '10' }));
    });
    vi.stubGlobal('fetch', fetcher);
    renderGooglePanel();
    await signIn(user);
    await startUpload(user, 2);
    await screen.findByRole('heading', { name: 'Részben sikeres naptárfeltöltés' });
    await user.click(screen.getByRole('button', { name: 'Kijelentkezés' }));
    const retryButton = screen.getByRole('button', {
      name: 'Csak a sikertelenek újrapróbálása',
    });
    const callsBeforeRetry = fetcher.mock.calls.length;

    expect(retryButton).toBeDisabled();
    fireEvent.click(retryButton);
    expect(fetcher).toHaveBeenCalledTimes(callsBeforeRetry);
    expect(postCounts.get('OMSZ')).toBe(1);
    expect(postCounts.get('KMR')).toBe(1);
  });

  it('az Új beosztás feldolgozása műveletet átadja az alkalmazásnak', async () => {
    const user = userEvent.setup();
    const fetcher = vi.fn<typeof fetch>().mockImplementation((input, init) => {
      if (requestUrl(input).endsWith('/users/me/calendarList')) return Promise.resolve(calendars());
      if (init?.method === 'POST') {
        return Promise.resolve(jsonResponse({ id: 'created', colorId: '10' }));
      }
      return Promise.resolve(jsonResponse({ items: [] }));
    });
    vi.stubGlobal('fetch', fetcher);
    const { callbacks } = renderGooglePanel([omszEvent]);
    await signIn(user);
    await startUpload(user, 1);
    await screen.findByRole('heading', { name: 'Sikeres naptárfeltöltés' });

    await user.click(screen.getByRole('button', { name: 'Új beosztás feldolgozása' }));

    expect(callbacks.onNewSchedule).toHaveBeenCalledOnce();
    expect(callbacks.onResult).toHaveBeenCalledOnce();
  });

  it('másik naptárhoz megtartja az eseményeket, törli a választást és nem küld DELETE kérést', async () => {
    const user = userEvent.setup();
    const availableCalendars = [
      { id: 'primary', summary: 'Első naptár', accessRole: 'owner' },
      { id: 'secondary', summary: 'Második naptár', accessRole: 'writer' },
    ];
    const fetcher = vi.fn<typeof fetch>().mockImplementation((input, init) => {
      if (requestUrl(input).endsWith('/users/me/calendarList')) {
        return Promise.resolve(calendars(availableCalendars));
      }
      if (init?.method === 'POST') {
        return Promise.resolve(jsonResponse({ id: 'created', colorId: '10' }));
      }
      return Promise.resolve(jsonResponse({ items: [] }));
    });
    vi.stubGlobal('fetch', fetcher);
    const { callbacks } = renderGooglePanel([omszEvent]);
    await signIn(user);
    await startUpload(user, 1);
    await screen.findByRole('heading', { name: 'Sikeres naptárfeltöltés' });

    await user.click(screen.getByRole('button', { name: 'Feltöltés másik naptárba' }));

    expect(screen.getByLabelText<HTMLSelectElement>('Írható naptár')).toHaveValue('');
    expect(
      screen.getByText('Ugyanezek az események egy másik naptárban is létrejönnek.'),
    ).toBeVisible();
    expect(screen.queryByText('Sikeres naptárfeltöltés')).not.toBeInTheDocument();
    expect(callbacks.onCalendarChange).toHaveBeenCalledOnce();
    expect(fetcher.mock.calls.some(([, init]) => init?.method?.toUpperCase() === 'DELETE')).toBe(
      false,
    );

    await user.selectOptions(screen.getByLabelText('Írható naptár'), 'secondary');
    expect(
      screen.getByRole('button', {
        name: '1 kijelölt esemény hozzáadása a Google Naptárhoz',
      }),
    ).toBeEnabled();
  });
});
