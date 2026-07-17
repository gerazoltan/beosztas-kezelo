import type { CalendarEvent, ReviewStatus } from '../domain/types';
import { AppError, type AppErrorCode } from '../domain/errors';
import { instantToLocal, zonedLocalToInstant } from './dates';

const API_ROOT = 'https://www.googleapis.com/calendar/v3';
const GOOGLE_EVENT_COLOR_ID = '10';

const browserFetch: typeof fetch = (input, init) => window.fetch(input, init);

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

export interface GoogleWriteResult {
  eventId: string;
  status: Extract<ReviewStatus, 'Már szerepel a naptárban' | 'Létrehozva' | 'Sikertelen'>;
  message: string;
  errorCode?: Extract<
    AppErrorCode,
    | 'GOOGLE_TOKEN_EXPIRED'
    | 'GOOGLE_NETWORK_ERROR'
    | 'GOOGLE_CALENDAR_NOT_WRITABLE'
    | 'GOOGLE_API_ERROR'
  >;
  technicalDetails?: string;
}

export interface GoogleUploadOptions {
  signal?: AbortSignal;
  onStart?: (event: CalendarEvent) => void;
  onResult?: (result: GoogleWriteResult) => void;
}

async function parseGoogleError(response: Response): Promise<AppError> {
  const details = await response.text();
  if (response.status === 401) return new AppError('GOOGLE_TOKEN_EXPIRED', details);
  if (response.status === 403) return new AppError('GOOGLE_CALENDAR_NOT_WRITABLE', details);
  return new AppError('GOOGLE_API_ERROR', `${response.status}: ${details}`);
}

export class GoogleCalendarClient {
  constructor(
    private readonly accessToken: string,
    private readonly fetcher: typeof fetch = browserFetch,
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
        'GOOGLE_NETWORK_ERROR',
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

  async isDuplicate(
    calendarId: string,
    item: CalendarEvent,
    signal?: AbortSignal,
  ): Promise<boolean> {
    const start = zonedLocalToInstant(item.calendarTime.start);
    const end = zonedLocalToInstant(item.calendarTime.end);
    const query = new URLSearchParams({
      timeMin: new Date(start.getTime() - 1000).toISOString(),
      timeMax: new Date(end.getTime() + 1000).toISOString(),
      singleEvents: 'true',
      showDeleted: 'false',
    });
    const response = await this.request<EventsResponse>(
      `/calendars/${encodeURIComponent(calendarId)}/events?${query.toString()}`,
      { signal },
    );
    return (response.items ?? []).some((candidate) => {
      const startValue = candidate.start?.dateTime;
      const endValue = candidate.end?.dateTime;
      if (!startValue || !endValue || candidate.summary !== item.summary) return false;
      return (
        instantToLocal(startValue) === item.calendarTime.start &&
        instantToLocal(endValue) === item.calendarTime.end
      );
    });
  }

  async insertEvent(
    calendarId: string,
    item: CalendarEvent,
    signal?: AbortSignal,
  ): Promise<GoogleEventItem> {
    return this.request<GoogleEventItem>(`/calendars/${encodeURIComponent(calendarId)}/events`, {
      method: 'POST',
      signal,
      body: JSON.stringify({
        summary: item.summary,
        start: { dateTime: item.calendarTime.start, timeZone: item.timeZone },
        end: { dateTime: item.calendarTime.end, timeZone: item.timeZone },
        colorId: GOOGLE_EVENT_COLOR_ID,
      }),
    });
  }

  async addEvents(
    calendarId: string,
    events: CalendarEvent[],
    options: GoogleUploadOptions = {},
  ): Promise<GoogleWriteResult[]> {
    const results: GoogleWriteResult[] = [];
    for (const item of events) {
      if (options.signal?.aborted) break;
      options.onStart?.(item);
      let result: GoogleWriteResult;
      try {
        if (await this.isDuplicate(calendarId, item, options.signal)) {
          result = {
            eventId: item.id,
            status: 'Már szerepel a naptárban',
            message: 'Azonos nevű és időpontú esemény már létezik.',
          };
        } else {
          const created = await this.insertEvent(calendarId, item, options.signal);
          const colorConfirmed = created.colorId === GOOGLE_EVENT_COLOR_ID;
          result = {
            eventId: item.id,
            status: 'Létrehozva',
            message: colorConfirmed
              ? 'Az eseményt a Google Naptár a zöld Basil (10) színnel létrehozta.'
              : 'Az esemény létrejött, de a Google nem a kért zöld Basil (10) színt igazolta vissza.',
          };
        }
      } catch (error) {
        if (options.signal?.aborted) break;
        const appError =
          error instanceof AppError
            ? error
            : new AppError(
                'GOOGLE_API_ERROR',
                error instanceof Error ? error.message : String(error),
              );
        result = {
          eventId: item.id,
          status: 'Sikertelen',
          message: appError.message,
          errorCode:
            appError.code === 'GOOGLE_TOKEN_EXPIRED' ||
            appError.code === 'GOOGLE_NETWORK_ERROR' ||
            appError.code === 'GOOGLE_CALENDAR_NOT_WRITABLE' ||
            appError.code === 'GOOGLE_API_ERROR'
              ? appError.code
              : 'GOOGLE_API_ERROR',
          technicalDetails: appError.technicalDetails,
        };
      }
      results.push(result);
      options.onResult?.(result);
      if (result.errorCode === 'GOOGLE_TOKEN_EXPIRED') break;
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
