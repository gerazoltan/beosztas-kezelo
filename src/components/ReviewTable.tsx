import type { Ref } from 'react';
import type { GoogleEventState, ReviewRow } from '../domain/types';
import { formatHungarianDate, weekdayHungarian } from '../services/dates';
import { isGoogleSelectionLocked } from '../utils/googleUpload';

interface ReviewTableProps {
  sectionRef?: Ref<HTMLElement>;
  rows: ReviewRow[];
  selected: Set<string>;
  googleStates: ReadonlyMap<string, GoogleEventState>;
  onToggle: (eventId: string) => void;
  onSelectAll: (selected: boolean) => void;
}

function time(value?: string): string {
  return value?.slice(11, 16) ?? '—';
}

export function ReviewTable({
  sectionRef,
  rows,
  selected,
  googleStates,
  onToggle,
  onSelectAll,
}: ReviewTableProps) {
  const exportable = rows.filter(
    (row) => row.event && !isGoogleSelectionLocked(googleStates.get(row.event.id)),
  );
  const allSelected =
    exportable.length > 0 && exportable.every((row) => selected.has(row.event?.id ?? ''));

  return (
    <section
      ref={sectionRef}
      className="panel review-panel workflow-section"
      aria-labelledby="review-heading"
    >
      <div className="section-heading review-title-row">
        <div>
          <span className="eyebrow">5. lépés</span>
          <h2 id="review-heading">Ellenőrzés</h2>
        </div>
        <label className="select-all">
          <input
            type="checkbox"
            checked={allSelected}
            onChange={(event) => onSelectAll(event.target.checked)}
            disabled={exportable.length === 0}
          />
          Összes biztos kijelölése
        </label>
      </div>
      <div className="table-scroll">
        <table>
          <thead>
            <tr>
              <th scope="col">Kijelölés</th>
              <th scope="col">Dátum</th>
              <th scope="col">Nap</th>
              <th scope="col">Felismert Excel-jelölés</th>
              <th scope="col">Szolgálat típusa</th>
              <th scope="col">Kezdés</th>
              <th scope="col">Befejezés</th>
              <th scope="col">Naptáresemény neve</th>
              <th scope="col">Állapot</th>
              <th scope="col">Ellenőrzési megjegyzés</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => {
              const eventId = row.event?.id;
              const googleState = eventId ? googleStates.get(eventId) : undefined;
              const displayedStatus = googleState?.status ?? row.status;
              const issue =
                displayedStatus === 'Bizonytalan' ||
                displayedStatus === 'Hibás párosítás' ||
                displayedStatus === 'Sikertelen';
              return (
                <tr
                  key={row.id}
                  className={issue ? 'row-issue' : displayedStatus === 'Kizárva' ? 'row-muted' : ''}
                >
                  <td data-label="Kijelölés">
                    <input
                      type="checkbox"
                      aria-label={`${formatHungarianDate(row.date)} exportálása`}
                      checked={eventId ? selected.has(eventId) : false}
                      disabled={!eventId || isGoogleSelectionLocked(googleState)}
                      onChange={() => eventId && onToggle(eventId)}
                    />
                  </td>
                  <td data-label="Dátum">{formatHungarianDate(row.date)}</td>
                  <td data-label="Nap">{weekdayHungarian(row.date)}</td>
                  <td data-label="Excel-jelölés">
                    <strong>{row.marker || '—'}</strong>
                  </td>
                  <td data-label="Szolgálat">{row.shiftType ?? '—'}</td>
                  <td data-label="Kezdés">{time(row.event?.shiftTime.start)}</td>
                  <td data-label="Befejezés">{time(row.event?.shiftTime.end)}</td>
                  <td data-label="Esemény">{row.summary ?? '—'}</td>
                  <td data-label="Állapot">
                    <span
                      className={`status status-${
                        displayedStatus === 'Létrehozás folyamatban'
                          ? 'pending'
                          : issue
                            ? 'issue'
                            : 'ok'
                      }`}
                    >
                      {displayedStatus}
                    </span>
                  </td>
                  <td data-label="Megjegyzés">
                    {googleState ? <strong>{googleState.message}</strong> : row.note}
                    {googleState?.technicalDetails && (
                      <details className="diagnostics">
                        <summary>Google API technikai részletek</summary>
                        <code>{googleState.technicalDetails}</code>
                      </details>
                    )}
                    <details className="diagnostics">
                      <summary>Technikai részletek</summary>
                      {row.technicalNote && <p>{row.technicalNote}</p>}
                      {row.diagnostics.map((item) => (
                        <dl key={item.address}>
                          <dt>Cella</dt>
                          <dd>
                            {item.address}
                            {item.isMerged ? ` (merge master: ${item.mergeMaster})` : ''}
                          </dd>
                          <dt>Nyers / megjelenített</dt>
                          <dd>
                            {item.rawValue || '∅'} / {item.displayedText || '∅'}
                          </dd>
                          <dt>Stílus</dt>
                          <dd>
                            #{item.styleId ?? '—'}, háttér {item.fillColor ?? '—'}, betű{' '}
                            {item.fontColor ?? '—'}, dőlt {item.italic ? 'igen' : 'nem'}, félkövér{' '}
                            {item.bold ? 'igen' : 'nem'}
                          </dd>
                        </dl>
                      ))}
                    </details>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </section>
  );
}
