import { useRef, useState } from 'react';
import type { CalendarEvent } from '../domain/types';
import { AppError, toAppError } from '../domain/errors';
import {
  GoogleCalendarClient,
  GoogleTokenSession,
  type GoogleCalendarListItem,
  type GoogleWriteResult,
} from '../services/googleCalendar';
import { requestGoogleAccessToken, revokeGoogleToken } from '../services/googleOAuth';
import { ErrorNotice } from './ErrorNotice';

interface GooglePanelProps {
  events: CalendarEvent[];
  onResults: (results: GoogleWriteResult[]) => void;
}

export function GooglePanel({ events, onResults }: GooglePanelProps) {
  const clientId = import.meta.env.VITE_GOOGLE_CLIENT_ID?.trim() ?? '';
  const tokenSession = useRef(new GoogleTokenSession());
  const [calendars, setCalendars] = useState<GoogleCalendarListItem[]>([]);
  const [calendarId, setCalendarId] = useState('');
  const [signedIn, setSignedIn] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<AppError>();

  if (!clientId) {
    return (
      <section className="panel google-panel" aria-labelledby="google-heading">
        <span className="eyebrow">Google Naptár</span>
        <h2 id="google-heading">Közvetlen hozzáadás</h2>
        <p className="notice neutral">
          A Google-integráció nincs konfigurálva. Az ICS-export ettől függetlenül teljesen
          használható.
        </p>
      </section>
    );
  }

  const signIn = async () => {
    setBusy(true);
    setError(undefined);
    try {
      const token = await requestGoogleAccessToken(clientId);
      tokenSession.current.set(token);
      const client = new GoogleCalendarClient(token);
      const available = await client.listWritableCalendars();
      setCalendars(available);
      setCalendarId(available.find((item) => item.primary)?.id ?? available[0]?.id ?? '');
      setSignedIn(true);
    } catch (caught) {
      tokenSession.current.clear();
      setError(toAppError(caught));
    } finally {
      setBusy(false);
    }
  };

  const signOut = () => {
    tokenSession.current.signOut(revokeGoogleToken);
    setSignedIn(false);
    setCalendars([]);
    setCalendarId('');
    setError(undefined);
  };

  const addToCalendar = async () => {
    const token = tokenSession.current.get();
    if (!token || !calendarId) {
      setError(new AppError('GOOGLE_ACCESS_DENIED'));
      return;
    }
    setBusy(true);
    setError(undefined);
    try {
      const results = await new GoogleCalendarClient(token).addEvents(calendarId, events);
      onResults(results);
    } catch (caught) {
      setError(toAppError(caught));
    } finally {
      setBusy(false);
    }
  };

  return (
    <section className="panel google-panel" aria-labelledby="google-heading">
      <div className="section-heading">
        <span className="eyebrow">Google Naptár</span>
        <h2 id="google-heading">Közvetlen hozzáadás</h2>
      </div>
      <ErrorNotice error={error} />
      {!signedIn ? (
        <button
          type="button"
          className="button secondary"
          onClick={() => void signIn()}
          disabled={busy}
        >
          {busy ? 'Kapcsolódás…' : 'Google-bejelentkezés'}
        </button>
      ) : (
        <div className="google-actions">
          {calendars.length === 0 && (
            <p className="notice warning">
              Nem található írható naptár. Ellenőrizd a fiók jogosultságait.
            </p>
          )}
          <label>
            Írható naptár
            <select value={calendarId} onChange={(event) => setCalendarId(event.target.value)}>
              {calendars.map((calendar) => (
                <option key={calendar.id} value={calendar.id}>
                  {calendar.summary}
                  {calendar.primary ? ' (elsődleges)' : ''}
                </option>
              ))}
            </select>
          </label>
          <div className="button-row">
            <button
              type="button"
              className="button primary"
              onClick={() => void addToCalendar()}
              disabled={busy || events.length === 0 || !calendarId}
            >
              {busy ? 'Feldolgozás…' : `${events.length} esemény hozzáadása`}
            </button>
            <button type="button" className="button text" onClick={signOut}>
              Kijelentkezés
            </button>
          </div>
        </div>
      )}
    </section>
  );
}
