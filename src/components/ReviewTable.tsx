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
              <th scope="col">Szolgálati jelleg</th>
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
              const inferredDiagnostic = row.dailyInference
                ? row.diagnostics.find((diagnostic) => diagnostic.displayedText.trim() === '12')
                : undefined;
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
                  <td data-label="Szolgálati jelleg">
                    {row.serviceCategory
                      ? `${row.serviceCategory}${
                          row.dailyInference?.correctionApplied ||
                          row.serviceResolution?.dailyInferenceApplied
                            ? ' – következtetett'
                            : ''
                        }`
                      : '—'}
                  </td>
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
                      {row.timeRule && (
                        <p>
                          <strong>Felismert időszabály:</strong> {row.timeRule}
                        </p>
                      )}
                      {row.pairingReferences?.map((reference) => (
                        <p key={`${reference.direction}-${reference.address}`}>
                          <strong>Párosításhoz használt cella:</strong>{' '}
                          {reference.direction === 'previous' ? 'előző' : 'következő'} –{' '}
                          {reference.address}
                        </p>
                      ))}
                      {row.dailyInference && (
                        <dl>
                          <dt>Eredeti érték</dt>
                          <dd>{inferredDiagnostic?.displayedText ?? row.marker}</dd>
                          <dt>Eredeti betűszín</dt>
                          <dd>{inferredDiagnostic?.fontColor ?? '—'}</dd>
                          <dt>Eredeti aláhúzás</dt>
                          <dd>{inferredDiagnostic?.underline ? 'igen' : 'nem'}</dd>
                          <dt>24 órás Parti szolgálat jelen van</dt>
                          <dd>
                            {row.dailyInference.partyTwentyFourHourPresent ? 'igen' : 'nem'}
                          </dd>
                          <dt>Kék 12 jelen van</dt>
                          <dd>{row.dailyInference.blueTwelvePresent ? 'igen' : 'nem'}</dd>
                          <dt>Zöld-aláhúzott 12 jelen van</dt>
                          <dd>{row.dailyInference.tenCarTwelvePresent ? 'igen' : 'nem'}</dd>
                          <dt>Fekete 12 jelöltek száma</dt>
                          <dd>{row.dailyInference.blackTwelveCandidateCount}</dd>
                          <dt>Következtetett korrekció történt</dt>
                          <dd>{row.dailyInference.correctionApplied ? 'igen' : 'nem'}</dd>
                          <dt>Eredeti szolgálattípus</dt>
                          <dd>
                            {row.dailyInference.originalServiceCategory},{' '}
                            {row.dailyInference.originalShiftType}
                          </dd>
                          <dt>Végső szolgálattípus</dt>
                          <dd>
                            {row.dailyInference.finalServiceCategory},{' '}
                            {row.dailyInference.finalShiftType}
                          </dd>
                          <dt>Végső időintervallum</dt>
                          <dd>
                            {time(row.dailyInference.finalTime.start)}–
                            {time(row.dailyInference.finalTime.end)}
                          </dd>
                        </dl>
                      )}
                      {row.serviceResolution && (
                        <dl>
                          <dt>Eredeti szolgálati kategória</dt>
                          <dd>{row.serviceResolution.originalServiceCategory}</dd>
                          <dt>Végső szolgálati kategória</dt>
                          <dd>{row.serviceResolution.finalServiceCategory ?? '—'}</dd>
                          <dt>Formázási korrekció történt</dt>
                          <dd>
                            {row.serviceResolution.formattingCorrectionApplied ? 'igen' : 'nem'}
                          </dd>
                          <dt>Napi összeállításból következtetve</dt>
                          <dd>{row.serviceResolution.dailyInferenceApplied ? 'igen' : 'nem'}</dd>
                          <dt>Feltételezett hónaphatár-párosítás</dt>
                          <dd>{row.serviceResolution.assumedBoundaryPairing ? 'igen' : 'nem'}</dd>
                          <dt>Párosítás forrása</dt>
                          <dd>{row.serviceResolution.pairingSource ?? 'nem szükséges'}</dd>
                          <dt>Tényleges vagy feltételezett párosító cella</dt>
                          <dd>{row.serviceResolution.pairingCell ?? 'nem szükséges'}</dd>
                          <dt>Végső listaidő</dt>
                          <dd>
                            {row.serviceResolution.finalShiftTime
                              ? `${row.serviceResolution.finalShiftTime.start} – ${row.serviceResolution.finalShiftTime.end}`
                              : '—'}
                          </dd>
                          <dt>Végső naptáridő</dt>
                          <dd>
                            {row.serviceResolution.finalCalendarTime
                              ? `${row.serviceResolution.finalCalendarTime.start} – ${row.serviceResolution.finalCalendarTime.end}`
                              : '—'}
                          </dd>
                        </dl>
                      )}
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
                          <dt>Napi cellacsoporton belüli pozíció</dt>
                          <dd>{item.positionInDayGroup}.</dd>
                          <dt>Stílus</dt>
                          <dd>
                            #{item.styleId ?? '—'}, dőlt {item.italic ? 'igen' : 'nem'}, félkövér{' '}
                            {item.bold ? 'igen' : 'nem'}
                          </dd>
                          <dt>Betűszín nyers értéke</dt>
                          <dd>{item.fontColorRaw ?? 'alapértelmezett'}</dd>
                          <dt>Betűszín normalizált értéke</dt>
                          <dd>{item.fontColor ?? '—'}</dd>
                          <dt>Aláhúzott</dt>
                          <dd>{item.underline ? 'igen' : 'nem'}</dd>
                          <dt>Fill típusa</dt>
                          <dd>{item.fillType ?? '—'}</dd>
                          <dt>patternType</dt>
                          <dd>{item.fillPatternType ?? '—'}</dd>
                          <dt>fgColor nyers értéke</dt>
                          <dd>{item.fillForegroundRaw ?? '—'}</dd>
                          <dt>bgColor nyers értéke</dt>
                          <dd>{item.fillBackgroundRaw ?? '—'}</dd>
                          <dt>Van látható kitöltés</dt>
                          <dd>{item.hasVisibleFill ? 'igen' : 'nem'}</dd>
                          <dt>Normalizált szín</dt>
                          <dd>{item.fillColor ?? '—'}</dd>
                          <dt>Végső fill kategória</dt>
                          <dd>{item.fillCategory ?? '—'}</dd>
                          <dt>Felismert szolgálati kategória</dt>
                          <dd>{row.serviceCategory ?? '—'}</dd>
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
