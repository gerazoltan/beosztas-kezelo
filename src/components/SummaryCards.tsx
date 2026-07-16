import type { ScheduleResult } from '../domain/types';

export function SummaryCards({ summary }: { summary: ScheduleResult['summary'] }) {
  const cards = [
    ['Felismert szolgálat', summary.recognized],
    ['OMSZ', summary.omsz],
    ['KMR', summary.kmr],
    ['Bizonytalan', summary.uncertain],
    ['Hibás', summary.invalid],
    ['Exportálható', summary.exportable],
  ];
  return (
    <section aria-label="Feldolgozás összesítése" className="summary-grid">
      {cards.map(([label, value]) => (
        <article key={String(label)} className="summary-card">
          <strong>{value}</strong>
          <span>{label}</span>
        </article>
      ))}
    </section>
  );
}
