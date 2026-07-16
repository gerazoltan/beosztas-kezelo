import type { CalendarEvent, ReviewStatus } from '../domain/types';
import { AppError } from '../domain/errors';
import { colorDistance } from '../excel/colors';
import { instantToLocal, zonedLocalToInstant } from './dates';

const API_ROOT = 'https://www.googleapis.com/calendar/v3';
const DARK_GREEN = '#0B5D3B';

export interface GoogleCalendarListItem {
  id: string;
  summary: string;
  primary?: boolean;
  accessRole: string;
}

interface GoogleEventItem {
  id?: string;
  summary?: string;
  colorId?: string;
  start?: { dateTime?: string; timeZone?: string };
  end?: { dateTime?: string; timeZone?: string };
}

interface CalendarListResponse {
  items?: GoogleCalendarListItem[];
}

interface EventsResponse {
  items?: GoogleEventItem[];
}

interface ColorsResponse {
  event?: Record<string, { background?: string; foreground?: string }>;
}

export interface GoogleWriteResult {
  eventId: string;
  status: Extract<ReviewStatus, 'Már szerepel a naptárban' | 'Létrehozva' | 'Sikertelen'>;
  message: string;
}

async function parseGoogleError(response: Response): Promise<AppError> {
  const details = await response.text();
  if (response.status === 401 || response.status === 403) {
    return new AppError('GOOGLE_ACCESS_DENIED', details);
  }
  return new AppError('GOOGLE_API_ERROR', `${response.status}: ${details}`);
}

export class GoogleCalendarClient {
  constructor(
    private readonly accessToken: string,
    private readonly fetcher: typeof fetch = fetch,
  ) {}

  private async request<T>(path: string, init?: RequestInit): Promise<T> {
    let response: Response;
    try {
      response = await this.fetcher(`${API_ROOT}${path}`, {
        ...init,
        headers: {
          Authorization: `Bearer ${this.accessToken}`,
          'Content-Type': 'application/json',
          ...init?.headers,
        },
      });
    } catch (error) {
      throw new AppError(
        'GOOGLE_API_ERROR',
        error instanceof Error ? error.message : String(error),
      );
    }
    if (!response.ok) throw await parseGoogleError(response);
    return (await response.json()) as T;
  }

  async listWritableCalendars(): Promise<GoogleCalendarListItem[]> {
    const response = await this.request<CalendarListResponse>('/users/me/calendarList');
    return (response.items ?? []).filter((item) => ['owner', 'writer'].includes(item.accessRole));
  }

  async nearestDarkGreenColorId(): Promise<string | undefined> {
    const response = await this.request<ColorsResponse>('/colors');
    return Object.entries(response.event ?? {})
      .filter((entry): entry is [string, { background: string }] => Boolean(entry[1].background))
      .map(([id, value]) => ({ id, distance: colorDistance(value.background, DARK_GREEN) }))
      .sort((a, b) => a.distance - b.distance)[0]?.id;
  }

  async isDuplicate(calendarId: string, item: CalendarEvent): Promise<boolean> {
    const start = zonedLocalToInstant(item.start);
    const end = zonedLocalToInstant(item.end);
    const query = new URLSearchParams({
      timeMin: new Date(start.getTime() - 1000).toISOString(),
      timeMax: new Date(end.getTime() + 1000).toISOString(),
      singleEvents: 'true',
      showDeleted: 'false',
    });
    const response = await this.request<EventsResponse>(
      `/calendars/${encodeURIComponent(calendarId)}/events?${query.toString()}`,
    );
    return (response.items ?? []).some((candidate) => {
      const startValue = candidate.start?.dateTime;
      const endValue = candidate.end?.dateTime;
      if (!startValue || !endValue || candidate.summary !== item.summary) return false;
      return instantToLocal(startValue) === item.start && instantToLocal(endValue) === item.end;
    });
  }

  async insertEvent(
    calendarId: string,
    item: CalendarEvent,
    colorId?: string,
  ): Promise<GoogleEventItem> {
    return this.request<GoogleEventItem>(`/calendars/${encodeURIComponent(calendarId)}/events`, {
      method: 'POST',
      body: JSON.stringify({
        summary: item.summary,
        start: { dateTime: item.start, timeZone: item.timeZone },
        end: { dateTime: item.end, timeZone: item.timeZone },
        ...(colorId ? { colorId } : {}),
      }),
    });
  }

  async addEvents(calendarId: string, events: CalendarEvent[]): Promise<GoogleWriteResult[]> {
    let colorId: string | undefined;
    try {
      colorId = await this.nearestDarkGreenColorId();
    } catch {
      colorId = undefined;
    }
    const results: GoogleWriteResult[] = [];
    for (const item of events) {
      try {
        if (await this.isDuplicate(calendarId, item)) {
          results.push({
            eventId: item.id,
            status: 'Már szerepel a naptárban',
            message: 'Azonos nevű és időpontú esemény már létezik.',
          });
          continue;
        }
        const created = await this.insertEvent(calendarId, item, colorId);
        const colorConfirmed = colorId !== undefined && created.colorId === colorId;
        results.push({
          eventId: item.id,
          status: 'Létrehozva',
          message: !colorId
            ? 'Az esemény létrejött, de nem volt elérhető beállítható sötétzöld eseményszín.'
            : colorConfirmed
              ? 'Az eseményt a Google Naptár a kért színnel létrehozta.'
              : 'Az esemény létrejött, de a Google nem igazolta vissza a kért színt.',
        });
      } catch (error) {
        results.push({
          eventId: item.id,
          status: 'Sikertelen',
          message: error instanceof AppError ? error.message : 'Ismeretlen Google API-hiba.',
        });
      }
    }
    return results;
  }
}

export class GoogleTokenSession {
  private token?: string;

  set(token: string): void {
    this.token = token;
  }

  get(): string | undefined {
    return this.token;
  }

  clear(): void {
    this.token = undefined;
  }

  signOut(revoke?: (token: string) => void): void {
    const current = this.token;
    this.clear();
    if (current && revoke) revoke(current);
  }
}
