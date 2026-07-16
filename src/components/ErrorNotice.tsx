import type { AppError } from '../domain/errors';

export function ErrorNotice({ error }: { error?: AppError }) {
  if (!error) return null;
  return (
    <div className="notice error" role="alert">
      <strong>Nem sikerült folytatni</strong>
      <p>{error.message}</p>
      {error.technicalDetails && (
        <details>
          <summary>Technikai részletek</summary>
          <code>{error.technicalDetails}</code>
        </details>
      )}
    </div>
  );
}
