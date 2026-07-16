import { useMemo, useState } from 'react';
import type { MonthSheet, ScheduleResult, WorkbookSession } from './domain/types';
import { HUNGARIAN_MONTHS } from './domain/types';
import { AppError, toAppError } from './domain/errors';
import { readEmployeeScheduleEntries } from './excel/dayEntries';
import { buildIcs, downloadIcs, icsFileName } from './services/ics';
import type { GoogleWriteResult } from './services/googleCalendar';
import { interpretSchedule } from './services/shifts';
import { ErrorNotice } from './components/ErrorNotice';
import { FileUpload } from './components/FileUpload';
import { GooglePanel } from './components/GooglePanel';
import { ReviewTable } from './components/ReviewTable';
import { Stepper } from './components/Stepper';
import { SummaryCards } from './components/SummaryCards';
import './styles.css';

function monthKey(month: MonthSheet): string {
  return `${month.year}-${month.month}-${month.sheetName}`;
}

export default function App() {
  const [session, setSession] = useState<WorkbookSession>();
  const [selectedMonthKey, setSelectedMonthKey] = useState('');
  const [employeeName, setEmployeeName] = useState('');
  const [employeeRow, setEmployeeRow] = useState<number>();
  const [result, setResult] = useState<ScheduleResult>();
  const [selectedEvents, setSelectedEvents] = useState<Set<string>>(new Set());
  const [error, setError] = useState<AppError>();
  const [notice, setNotice] = useState('');
  const [busy, setBusy] = useState(false);

  const selectedMonth = session?.months.find((month) => monthKey(month) === selectedMonthKey);
  const employee = selectedMonth?.employees.find((item) => item.normalizedName === employeeName);
  const selectedCalendarEvents = useMemo(
    () => result?.events.filter((event) => selectedEvents.has(event.id)) ?? [],
    [result, selectedEvents],
  );
  const currentStep = result ? 6 : employeeName ? 4 : selectedMonth ? 2 : session ? 1 : 1;

  const resetAfterFile = () => {
    setEmployeeName('');
    setEmployeeRow(undefined);
    setResult(undefined);
    setSelectedEvents(new Set());
    setNotice('');
    setError(undefined);
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
      setSelectedMonthKey(monthKey(defaultSelection.month));
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
  };

  const selectEmployee = (value: string) => {
    setEmployeeName(value);
    const nextEmployee = selectedMonth?.employees.find((item) => item.normalizedName === value);
    setEmployeeRow(nextEmployee?.rows.length === 1 ? nextEmployee.rows[0] : undefined);
    setResult(undefined);
    setSelectedEvents(new Set());
    setError(undefined);
  };

  const processSchedule = () => {
    setError(undefined);
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
    setSelectedEvents((current) => {
      const next = new Set(current);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  };

  const selectAll = (checked: boolean) => {
    setSelectedEvents(
      checked && result ? new Set(result.events.map((event) => event.id)) : new Set(),
    );
  };

  const exportIcs = () => {
    if (!selectedMonth || !employee || selectedCalendarEvents.length === 0) return;
    downloadIcs(
      buildIcs(selectedCalendarEvents),
      icsFileName(employee.name, selectedMonth.year, selectedMonth.month),
    );
  };

  const applyGoogleResults = (googleResults: GoogleWriteResult[]) => {
    const statuses = new Map(googleResults.map((item) => [item.eventId, item]));
    setResult((current) =>
      current
        ? {
            ...current,
            rows: current.rows.map((row) => {
              const status = row.event ? statuses.get(row.event.id) : undefined;
              return status ? { ...row, status: status.status, note: status.message } : row;
            }),
          }
        : current,
    );
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
        <Stepper current={currentStep} />
        <FileUpload
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
          <section className="panel" aria-labelledby="selection-heading">
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
                    <option key={monthKey(month)} value={monthKey(month)}>
                      {month.year}. {HUNGARIAN_MONTHS[month.month - 1]} — {month.sheetName}
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
                    onChange={(event) => setEmployeeRow(Number(event.target.value) || undefined)}
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
              rows={result.rows}
              selected={selectedEvents}
              onToggle={toggleEvent}
              onSelectAll={selectAll}
            />
            <section className="panel export-panel" aria-labelledby="export-heading">
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
            <GooglePanel events={selectedCalendarEvents} onResults={applyGoogleResults} />
          </>
        )}
      </main>
      <footer>Az alkalmazás nem küldi el és nem tárolja a feltöltött beosztást.</footer>
    </div>
  );
}
