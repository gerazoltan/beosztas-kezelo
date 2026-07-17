import { useMemo, useRef, useState } from 'react';
import type { GoogleEventState, ScheduleResult, WorkbookSession } from './domain/types';
import { AppError, toAppError } from './domain/errors';
import { readEmployeeScheduleEntries } from './excel/dayEntries';
import { buildIcs, downloadIcs, icsFileName } from './services/ics';
import type { GoogleWriteResult } from './services/googleCalendar';
import { interpretSchedule } from './services/shifts';
import { monthOptionLabel, monthOptionValue } from './utils/monthOptions';
import { isGoogleSelectionLocked, isGoogleUploadComplete } from './utils/googleUpload';
import { deriveWorkflowProgress, type WorkflowStepId } from './utils/workflowProgress';
import { BackToTopButton } from './components/BackToTopButton';
import { ErrorNotice } from './components/ErrorNotice';
import { FileUpload } from './components/FileUpload';
import { GooglePanel } from './components/GooglePanel';
import { ReviewTable } from './components/ReviewTable';
import { Stepper } from './components/Stepper';
import { SummaryCards } from './components/SummaryCards';
import './styles.css';

export default function App() {
  const uploadSectionRef = useRef<HTMLElement>(null);
  const selectionSectionRef = useRef<HTMLElement>(null);
  const reviewSectionRef = useRef<HTMLElement>(null);
  const exportSectionRef = useRef<HTMLElement>(null);
  const [session, setSession] = useState<WorkbookSession>();
  const [selectedMonthKey, setSelectedMonthKey] = useState('');
  const [employeeName, setEmployeeName] = useState('');
  const [employeeRow, setEmployeeRow] = useState<number>();
  const [result, setResult] = useState<ScheduleResult>();
  const [selectedEvents, setSelectedEvents] = useState<Set<string>>(new Set());
  const [error, setError] = useState<AppError>();
  const [notice, setNotice] = useState('');
  const [busy, setBusy] = useState(false);
  const [googleEventStates, setGoogleEventStates] = useState<Map<string, GoogleEventState>>(
    new Map(),
  );
  const [googleUploadResetKey, setGoogleUploadResetKey] = useState(0);
  const [icsExported, setIcsExported] = useState(false);

  const selectedMonth = session?.months.find(
    (month) => monthOptionValue(month) === selectedMonthKey,
  );
  const employee = selectedMonth?.employees.find((item) => item.normalizedName === employeeName);
  const selectedCalendarEvents = useMemo(
    () =>
      result?.events.filter(
        (event) =>
          selectedEvents.has(event.id) && !isGoogleUploadComplete(googleEventStates.get(event.id)),
      ) ?? [],
    [googleEventStates, result, selectedEvents],
  );
  const employeeSelectionComplete = Boolean(
    employeeName && employee && (employee.rows.length === 1 || employeeRow !== undefined),
  );
  const hasSelectedExportableEvent =
    result?.events.some((event) => selectedEvents.has(event.id)) ?? false;
  const hasCompletedGoogleEvent =
    result?.events.some((event) => isGoogleUploadComplete(googleEventStates.get(event.id))) ??
    false;
  const googleUploadInProgress = [...googleEventStates.values()].some(
    (state) => state.status === 'Létrehozás folyamatban',
  );
  const googleUploadFailed = [...googleEventStates.values()].some(
    (state) => state.status === 'Sikertelen',
  );
  const workflowSteps = deriveWorkflowProgress({
    fileLoaded: Boolean(session),
    monthSelected: Boolean(selectedMonth),
    employeeSelected: employeeSelectionComplete,
    resultReady: Boolean(result),
    hasSelectedExportableEvent,
    hasCompletedGoogleEvent,
    googleUploadInProgress,
    googleUploadFailed,
    icsExported,
    errorCode: error?.code,
  });
  const resetGoogleUpload = () => {
    setGoogleEventStates(new Map());
    setGoogleUploadResetKey((current) => current + 1);
  };

  const resetAfterFile = () => {
    setEmployeeName('');
    setEmployeeRow(undefined);
    setResult(undefined);
    setSelectedEvents(new Set());
    setNotice('');
    setError(undefined);
    setIcsExported(false);
    resetGoogleUpload();
  };

  const handleFile = async (file: File) => {
    resetAfterFile();
    setSession(undefined);
    setSelectedMonthKey('');
    setBusy(true);
    try {
      const { chooseDefaultMonth, parseWorkbook } = await import('./excel/workbookParser');
      const parsed = await parseWorkbook(await file.arrayBuffer(), file.name);
      const defaultSelection = chooseDefaultMonth(parsed.months);
      setSession(parsed);
      setSelectedMonthKey(monthOptionValue(defaultSelection.month));
      if (defaultSelection.usedFallback) {
        setNotice(
          'A következő naptári hónap nem található; az első kitöltött havi lapot választottuk ki.',
        );
      }
    } catch (caught) {
      setError(toAppError(caught));
    } finally {
      setBusy(false);
    }
  };

  const selectMonth = (value: string) => {
    setSelectedMonthKey(value);
    setEmployeeName('');
    setEmployeeRow(undefined);
    setResult(undefined);
    setSelectedEvents(new Set());
    setError(undefined);
    setIcsExported(false);
    resetGoogleUpload();
  };

  const selectEmployee = (value: string) => {
    setEmployeeName(value);
    const nextEmployee = selectedMonth?.employees.find((item) => item.normalizedName === value);
    setEmployeeRow(nextEmployee?.rows.length === 1 ? nextEmployee.rows[0] : undefined);
    setResult(undefined);
    setSelectedEvents(new Set());
    setError(undefined);
    setIcsExported(false);
    resetGoogleUpload();
  };

  const selectEmployeeRow = (row: number | undefined) => {
    setEmployeeRow(row);
    setResult(undefined);
    setSelectedEvents(new Set());
    setError(undefined);
    setIcsExported(false);
    resetGoogleUpload();
  };

  const processSchedule = () => {
    setError(undefined);
    setIcsExported(false);
    resetGoogleUpload();
    if (!session || !selectedMonth || !employeeName) {
      setError(new AppError('EMPLOYEE_NOT_FOUND'));
      return;
    }
    if (employee && employee.rows.length > 1 && employeeRow === undefined) {
      setError(
        new AppError('EMPLOYEE_DUPLICATE', `Választható sorok: ${employee.rows.join(', ')}.`),
      );
      return;
    }
    try {
      const entries = readEmployeeScheduleEntries(
        session,
        selectedMonth,
        employeeName,
        employeeRow,
      );
      const interpreted = interpretSchedule(entries.current, {
        legend: selectedMonth.legendStyles,
        previous: entries.previous,
        next: entries.next,
      });
      setResult(interpreted);
      setSelectedEvents(new Set(interpreted.events.map((event) => event.id)));
    } catch (caught) {
      setError(toAppError(caught));
    }
  };

  const toggleEvent = (id: string) => {
    if (isGoogleSelectionLocked(googleEventStates.get(id))) return;
    setIcsExported(false);
    setSelectedEvents((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const selectAll = (checked: boolean) => {
    setIcsExported(false);
    setSelectedEvents(
      checked && result
        ? new Set(
            result.events
              .filter((event) => !isGoogleSelectionLocked(googleEventStates.get(event.id)))
              .map((event) => event.id),
          )
        : new Set(),
    );
  };

  const exportIcs = () => {
    if (!selectedMonth || !employee || selectedCalendarEvents.length === 0) return;
    downloadIcs(
      buildIcs(selectedCalendarEvents),
      icsFileName(employee.name, selectedMonth.year, selectedMonth.month),
    );
    setIcsExported(true);
  };

  const markGoogleEventStarted = (eventId: string) => {
    setGoogleEventStates((current) => {
      const next = new Map(current);
      next.set(eventId, {
        status: 'Létrehozás folyamatban',
        message: 'A Google Naptár ellenőrzése és az esemény létrehozása folyamatban van.',
      });
      return next;
    });
  };

  const applyGoogleResult = (googleResult: GoogleWriteResult) => {
    setGoogleEventStates((current) => {
      const next = new Map(current);
      next.set(googleResult.eventId, googleResult);
      return next;
    });
    if (
      googleResult.status === 'Létrehozva' ||
      googleResult.status === 'Már szerepel a naptárban'
    ) {
      setSelectedEvents((current) => {
        const next = new Set(current);
        next.delete(googleResult.eventId);
        return next;
      });
    }
  };

  const resetAfterCalendarChange = () => {
    setGoogleEventStates(new Map());
    setSelectedEvents(new Set(result?.events.map((event) => event.id) ?? []));
    setIcsExported(false);
  };

  const startNewSchedule = () => {
    setSession(undefined);
    setSelectedMonthKey('');
    setEmployeeName('');
    setEmployeeRow(undefined);
    setResult(undefined);
    setSelectedEvents(new Set());
    setError(undefined);
    setNotice('');
    setBusy(false);
    setIcsExported(false);
    resetGoogleUpload();
    uploadSectionRef.current?.scrollIntoView?.({
      behavior: 'smooth',
      block: 'start',
    });
  };

  const navigateWorkflow = (stepId: WorkflowStepId) => {
    const selectionTarget = selectionSectionRef.current ?? uploadSectionRef.current;
    const target =
      stepId === 'file'
        ? uploadSectionRef.current
        : stepId === 'month' || stepId === 'employee' || stepId === 'processing'
          ? selectionTarget
          : stepId === 'review'
            ? reviewSectionRef.current
            : exportSectionRef.current;
    target?.scrollIntoView({
      behavior: window.matchMedia?.('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth',
      block: 'start',
    });
  };

  return (
    <div className="app-shell">
      <header className="hero">
        <div className="brand-mark" aria-hidden="true">
          B
        </div>
        <div>
          <p className="eyebrow">Biztonságos, helyi feldolgozás</p>
          <h1>Beosztáskezelő</h1>
          <p>Alakítsd át az Excel-beosztást ellenőrzött naptáreseményekké.</p>
        </div>
      </header>

      <main>
        <Stepper steps={workflowSteps} onNavigate={navigateWorkflow} />
        <FileUpload
          sectionRef={uploadSectionRef}
          fileName={session?.fileName}
          disabled={busy}
          onFile={(file) => void handleFile(file)}
        />
        {busy && (
          <div className="notice neutral" role="status">
            A munkafüzet feldolgozása…
          </div>
        )}
        <ErrorNotice error={error} />
        {notice && (
          <div className="notice warning" role="status">
            {notice}
          </div>
        )}
        {session?.warnings.map((warning) => (
          <div className="notice warning" key={warning}>
            {warning}
          </div>
        ))}

        {session && (
          <section
            ref={selectionSectionRef}
            className="panel workflow-section"
            aria-labelledby="selection-heading"
          >
            <div className="section-heading">
              <span className="eyebrow">2–4. lépés</span>
              <h2 id="selection-heading">Beosztás kiválasztása</h2>
            </div>
            <div className="form-grid">
              <label>
                Hónap
                <select
                  value={selectedMonthKey}
                  onChange={(event) => selectMonth(event.target.value)}
                >
                  {session.months.map((month) => (
                    <option key={monthOptionValue(month)} value={monthOptionValue(month)}>
                      {monthOptionLabel(month)}
                    </option>
                  ))}
                </select>
              </label>
              <label>
                Dolgozó
                <select
                  value={employeeName}
                  onChange={(event) => selectEmployee(event.target.value)}
                >
                  <option value="">Válassz dolgozót…</option>
                  {selectedMonth?.employees.map((item) => (
                    <option key={item.normalizedName} value={item.normalizedName}>
                      {item.name}
                    </option>
                  ))}
                </select>
              </label>
              {employee && employee.rows.length > 1 && (
                <label>
                  Sor kézi kiválasztása
                  <select
                    value={employeeRow ?? ''}
                    onChange={(event) => selectEmployeeRow(Number(event.target.value) || undefined)}
                  >
                    <option value="">Válassz sort…</option>
                    {employee.rows.map((row) => (
                      <option key={row} value={row}>
                        {row}. sor
                      </option>
                    ))}
                  </select>
                </label>
              )}
            </div>
            {selectedMonth && selectedMonth.warnings.length > 0 && (
              <div className="notice warning">
                <strong>Forrásadat-ellenőrzés</strong>
                <ul>
                  {selectedMonth.warnings.map((warning) => (
                    <li key={warning}>{warning}</li>
                  ))}
                </ul>
              </div>
            )}
            <button
              type="button"
              className="button primary"
              disabled={
                !employeeName ||
                (Boolean(employee && employee.rows.length > 1) && employeeRow === undefined)
              }
              onClick={processSchedule}
            >
              Beosztás feldolgozása
            </button>
          </section>
        )}

        {result && (
          <>
            <SummaryCards summary={result.summary} />
            <ReviewTable
              sectionRef={reviewSectionRef}
              rows={result.rows}
              selected={selectedEvents}
              googleStates={googleEventStates}
              onToggle={toggleEvent}
              onSelectAll={selectAll}
            />
            <section
              ref={exportSectionRef}
              className="panel export-panel workflow-section"
              aria-labelledby="export-heading"
            >
              <div>
                <span className="eyebrow">6. lépés</span>
                <h2 id="export-heading">Export</h2>
                <p>
                  <strong>{selectedCalendarEvents.length}</strong> kijelölt, biztos esemény kerül az
                  ICS-fájlba.
                </p>
              </div>
              <button
                type="button"
                className="button primary"
                onClick={exportIcs}
                disabled={selectedCalendarEvents.length === 0}
              >
                ICS letöltése
              </button>
            </section>
          </>
        )}
        <GooglePanel
          visible={Boolean(result)}
          events={selectedCalendarEvents}
          resetKey={googleUploadResetKey}
          onEventStart={markGoogleEventStarted}
          onResult={applyGoogleResult}
          onCalendarChange={resetAfterCalendarChange}
          onNewSchedule={startNewSchedule}
        />
      </main>
      <BackToTopButton />
      <footer>Az alkalmazás nem küldi el és nem tárolja a feltöltött beosztást.</footer>
    </div>
  );
}
