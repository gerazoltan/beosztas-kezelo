import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { CalendarEvent } from '../domain/types';
import { AppError } from '../domain/errors';
import {
  GoogleCalendarClient,
  GoogleTokenSession,
  type GoogleCalendarListItem,
  type GoogleWriteResult,
} from '../services/googleCalendar';
import { requestGoogleAccessToken, revokeGoogleToken } from '../services/googleOAuth';
import { ErrorNotice } from './ErrorNotice';

const GOOGLE_CALENDAR_URL = 'https://calendar.google.com/calendar/u/0/r';

type UploadPhase = 'idle' | 'uploading' | 'success' | 'partial' | 'failure';

interface UploadCounts {
  created: number;
  duplicate: number;
  failed: number;
  pending: number;
}

interface GooglePanelProps {
  visible?: boolean;
  events: CalendarEvent[];
  resetKey: number;
  onEventStart: (eventId: string) => void;
  onResult: (result: GoogleWriteResult) => void;
  onCalendarChange: () => void;
  onNewSchedule: () => void;
}

function asGoogleError(error: unknown): AppError {
  if (error instanceof AppError) return error;
  return new AppError('GOOGLE_API_ERROR', error instanceof Error ? error.message : String(error));
}

function uploadCounts(
  scope: CalendarEvent[],
  outcomes: ReadonlyMap<string, GoogleWriteResult>,
): UploadCounts {
  let created = 0;
  let duplicate = 0;
  let failed = 0;
  let pending = 0;
  for (const event of scope) {
    const outcome = outcomes.get(event.id);
    if (!outcome) pending += 1;
    else if (outcome.status === 'Létrehozva') created += 1;
    else if (outcome.status === 'Már szerepel a naptárban') duplicate += 1;
    else failed += 1;
  }
  return { created, duplicate, failed, pending };
}

export function GooglePanel({
  visible = true,
  events,
  resetKey,
  onEventStart,
  onResult,
  onCalendarChange,
  onNewSchedule,
}: GooglePanelProps) {
  const clientId = import.meta.env.VITE_GOOGLE_CLIENT_ID?.trim() ?? '';
  const tokenSession = useRef(new GoogleTokenSession());
  const uploadLock = useRef(false);
  const runId = useRef(0);
  const abortController = useRef<AbortController | undefined>(undefined);
  const [calendars, setCalendars] = useState<GoogleCalendarListItem[]>([]);
  const [calendarId, setCalendarId] = useState('');
  const [signedIn, setSignedIn] = useState(false);
  const [authBusy, setAuthBusy] = useState(false);
  const [phase, setPhase] = useState<UploadPhase>('idle');
  const [progress, setProgress] = useState({ completed: 0, total: 0 });
  const [scope, setScope] = useState<CalendarEvent[]>([]);
  const [outcomes, setOutcomes] = useState<Map<string, GoogleWriteResult>>(new Map());
  const [calendarName, setCalendarName] = useState('');
  const [authRequired, setAuthRequired] = useState(false);
  const [signedOutNotice, setSignedOutNotice] = useState(false);
  const [anotherCalendarNotice, setAnotherCalendarNotice] = useState(false);
  const [error, setError] = useState<AppError>();

  const resetUploadResult = useCallback(() => {
    abortController.current?.abort();
    abortController.current = undefined;
    runId.current += 1;
    uploadLock.current = false;
    setPhase('idle');
    setProgress({ completed: 0, total: 0 });
    setScope([]);
    setOutcomes(new Map());
    setCalendarName('');
    setAuthRequired(false);
    setError(undefined);
  }, []);

  useEffect(() => {
    resetUploadResult();
    setAnotherCalendarNotice(false);
  }, [resetKey, resetUploadResult]);

  useEffect(
    () => () => {
      runId.current += 1;
      abortController.current?.abort();
    },
    [],
  );

  useEffect(() => {
    if (phase !== 'success' || events.length === 0) return;
    const completedScopeIds = new Set(scope.map((event) => event.id));
    if (events.some((event) => !completedScopeIds.has(event.id))) resetUploadResult();
  }, [events, phase, resetUploadResult, scope]);

  const counts = useMemo(() => uploadCounts(scope, outcomes), [outcomes, scope]);
  const failedResults = useMemo(
    () => [...outcomes.values()].filter((result) => result.status === 'Sikertelen'),
    [outcomes],
  );
  const retryEvents = useMemo(() => {
    const selectedIds = new Set(events.map((event) => event.id));
    return scope.filter((event) => {
      const outcome = outcomes.get(event.id);
      return selectedIds.has(event.id) && (!outcome || outcome.status === 'Sikertelen');
    });
  }, [events, outcomes, scope]);
  if (!visible) return null;

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
    if (authBusy) return;
    setAuthBusy(true);
    setError(undefined);
    try {
      const token = await requestGoogleAccessToken(clientId);
      tokenSession.current.set(token);
      const available = await new GoogleCalendarClient(token).listWritableCalendars();
      const nextCalendarId =
        available.find((item) => item.id === calendarId)?.id ??
        available.find((item) => item.primary)?.id ??
        available[0]?.id ??
        '';
      if (calendarId && nextCalendarId !== calendarId) {
        resetUploadResult();
        onCalendarChange();
      }
      setCalendars(available);
      setCalendarId(nextCalendarId);
      setSignedIn(true);
      setSignedOutNotice(false);
    } catch (caught) {
      tokenSession.current.clear();
      setSignedIn(false);
      setError(asGoogleError(caught));
    } finally {
      setAuthBusy(false);
    }
  };

  const signOut = () => {
    if (phase === 'uploading') return;
    tokenSession.current.signOut(revokeGoogleToken);
    setSignedIn(false);
    setCalendars([]);
    setCalendarId('');
    setAuthBusy(false);
    setAuthRequired(false);
    setSignedOutNotice(true);
    setError(undefined);
  };

  const selectCalendar = (nextCalendarId: string) => {
    if (nextCalendarId === calendarId || phase === 'uploading') return;
    setCalendarId(nextCalendarId);
    resetUploadResult();
    onCalendarChange();
  };

  const uploadToAnotherCalendar = () => {
    if (phase === 'uploading') return;
    setCalendarId('');
    resetUploadResult();
    setAnotherCalendarNotice(true);
    onCalendarChange();
  };

  const upload = async (targets: CalendarEvent[], startsNewScope: boolean) => {
    if (
      uploadLock.current ||
      targets.length === 0 ||
      !signedIn ||
      !calendarId ||
      phase === 'uploading'
    ) {
      return;
    }
    const token = tokenSession.current.get();
    if (!token) {
      setSignedIn(false);
      setError(new AppError('GOOGLE_TOKEN_EXPIRED'));
      return;
    }

    uploadLock.current = true;
    const currentRunId = runId.current + 1;
    runId.current = currentRunId;
    const controller = new AbortController();
    abortController.current = controller;
    const uploadScope = startsNewScope ? targets : scope;
    const currentOutcomes = startsNewScope
      ? new Map<string, GoogleWriteResult>()
      : new Map(outcomes);
    const selectedCalendar = calendars.find((calendar) => calendar.id === calendarId);

    if (startsNewScope) {
      setScope(targets);
      setOutcomes(new Map());
      setCalendarName(selectedCalendar?.summary ?? calendarId);
    }
    setProgress({ completed: 0, total: targets.length });
    setPhase('uploading');
    setAuthRequired(false);
    setError(undefined);

    let completed = 0;
    try {
      const results = await new GoogleCalendarClient(token).addEvents(calendarId, targets, {
        signal: controller.signal,
        onStart: (event) => {
          if (runId.current === currentRunId) onEventStart(event.id);
        },
        onResult: (result) => {
          if (runId.current !== currentRunId) return;
          currentOutcomes.set(result.eventId, result);
          setOutcomes(new Map(currentOutcomes));
          completed += 1;
          setProgress({ completed, total: targets.length });
          onResult(result);
        },
      });
      if (runId.current !== currentRunId || controller.signal.aborted) return;

      const currentCounts = uploadCounts(uploadScope, currentOutcomes);
      const tokenExpired = results.some((result) => result.errorCode === 'GOOGLE_TOKEN_EXPIRED');
      setAuthRequired(tokenExpired);
      if (tokenExpired) {
        tokenSession.current.clear();
        setSignedIn(false);
        setError(new AppError('GOOGLE_TOKEN_EXPIRED'));
      }

      const successful = currentCounts.created + currentCounts.duplicate;
      if (currentCounts.failed === 0 && currentCounts.pending === 0) setPhase('success');
      else if (successful > 0) setPhase('partial');
      else setPhase('failure');
    } catch (caught) {
      if (runId.current !== currentRunId || controller.signal.aborted) return;
      setError(asGoogleError(caught));
      setPhase('failure');
    } finally {
      if (runId.current === currentRunId) {
        uploadLock.current = false;
        abortController.current = undefined;
      }
    }
  };

  const technicalDetails = [
    ...new Set(
      failedResults.flatMap((result) => (result.technicalDetails ? [result.technicalDetails] : [])),
    ),
  ];
  const failureMessage =
    failedResults[0]?.message ?? error?.message ?? 'A naptárfeltöltés nem sikerült.';
  const retryDisabled = !signedIn || !calendarId || retryEvents.length === 0;

  return (
    <section className="panel google-panel" aria-labelledby="google-heading">
      <div className="section-heading">
        <span className="eyebrow">Google Naptár</span>
        <h2 id="google-heading">Közvetlen hozzáadás</h2>
      </div>
      <ErrorNotice error={error} />
      {signedOutNotice && (
        <p className="notice neutral" role="status">
          Kijelentkeztél a Google Naptárból. A korábban létrehozott események a naptárban maradnak.
        </p>
      )}
      {anotherCalendarNotice && (
        <p className="notice warning" role="status">
          Ugyanezek az események egy másik naptárban is létrejönnek.
        </p>
      )}

      {!signedIn ? (
        <button
          type="button"
          className="button secondary"
          onClick={() => void signIn()}
          disabled={authBusy || phase === 'uploading'}
        >
          {authBusy
            ? 'Kapcsolódás…'
            : authRequired
              ? 'Google-bejelentkezés megújítása'
              : signedOutNotice
                ? 'Másik fiók csatlakoztatása'
                : 'Google-bejelentkezés'}
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
            <select
              value={calendarId}
              onChange={(event) => selectCalendar(event.target.value)}
              disabled={phase === 'uploading'}
            >
              <option value="">Válassz naptárt…</option>
              {calendars.map((calendar) => (
                <option key={calendar.id} value={calendar.id}>
                  {calendar.summary}
                  {calendar.primary ? ' (elsődleges)' : ''}
                </option>
              ))}
            </select>
          </label>
        </div>
      )}

      {phase === 'idle' && signedIn && (
        <button
          type="button"
          className="button primary google-upload-button"
          onClick={() => void upload(events, true)}
          disabled={events.length === 0 || !calendarId || uploadLock.current}
        >
          {events.length} kijelölt esemény hozzáadása a Google Naptárhoz
        </button>
      )}

      {phase === 'uploading' && (
        <button
          type="button"
          className="button primary google-upload-button"
          disabled
          aria-live="polite"
        >
          <span className="spinner" aria-hidden="true" />
          Feltöltés folyamatban… {progress.completed} / {progress.total}
        </button>
      )}

      {phase === 'success' && (
        <div className="upload-result upload-result-success" role="status">
          <h3>Sikeres naptárfeltöltés</h3>
          <UploadSummary counts={counts} calendarName={calendarName} />
          <div className="button-row">
            <button type="button" className="button primary" onClick={onNewSchedule}>
              Új beosztás feldolgozása
            </button>
            <button
              type="button"
              className="button secondary"
              onClick={uploadToAnotherCalendar}
              disabled={!signedIn}
            >
              Feltöltés másik naptárba
            </button>
            {signedIn && (
              <a
                className="button tertiary"
                href={GOOGLE_CALENDAR_URL}
                target="_blank"
                rel="noopener noreferrer"
              >
                Google Naptár megnyitása
                <span className="external-link-icon" aria-hidden="true">
                  ↗
                </span>
              </a>
            )}
          </div>
        </div>
      )}

      {phase === 'partial' && (
        <div className="upload-result upload-result-warning" role="status">
          <h3>Részben sikeres naptárfeltöltés</h3>
          <UploadSummary counts={counts} calendarName={calendarName} />
          {counts.pending > 0 && <p>{counts.pending} esemény még nem került feldolgozásra.</p>}
          <button
            type="button"
            className="button secondary"
            onClick={() => void upload(retryEvents, false)}
            disabled={retryDisabled}
          >
            {authRequired ? 'Feltöltés folytatása' : 'Csak a sikertelenek újrapróbálása'}
          </button>
        </div>
      )}

      {phase === 'failure' && (
        <div className="upload-result upload-result-error" role="alert">
          <h3>Sikertelen naptárfeltöltés</h3>
          <UploadSummary counts={counts} calendarName={calendarName} />
          {counts.pending > 0 && <p>{counts.pending} esemény még nem került feldolgozásra.</p>}
          <p>{failureMessage}</p>
          {(technicalDetails.length > 0 || error?.technicalDetails) && (
            <details>
              <summary>Technikai részletek</summary>
              <code>{technicalDetails.join('\n') || error?.technicalDetails}</code>
            </details>
          )}
          <button
            type="button"
            className="button secondary"
            onClick={() => void upload(retryEvents, false)}
            disabled={retryDisabled}
          >
            Újrapróbálás
          </button>
        </div>
      )}

      {signedIn && (
        <button
          type="button"
          className="button text"
          onClick={signOut}
          disabled={phase === 'uploading'}
        >
          Kijelentkezés
        </button>
      )}
    </section>
  );
}

function UploadSummary({ counts, calendarName }: { counts: UploadCounts; calendarName: string }) {
  return (
    <div className="upload-summary">
      <p>{counts.created} esemény létrehozva.</p>
      <p>{counts.duplicate} esemény már szerepelt a naptárban.</p>
      <p>{counts.failed} sikertelen művelet.</p>
      <p>
        Kiválasztott naptár: <strong>{calendarName || '—'}</strong>
      </p>
    </div>
  );
}
