import { useRef, useState, type ChangeEvent, type DragEvent } from 'react';

interface FileUploadProps {
  fileName?: string;
  disabled?: boolean;
  onFile: (file: File) => void;
}

export function FileUpload({ fileName, disabled, onFile }: FileUploadProps) {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);

  const choose = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file) onFile(file);
    event.target.value = '';
  };

  const drop = (event: DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    setDragging(false);
    const file = event.dataTransfer.files[0];
    if (file) onFile(file);
  };

  return (
    <section className="panel" aria-labelledby="upload-heading">
      <div className="section-heading">
        <span className="eyebrow">1. lépés</span>
        <h2 id="upload-heading">Excel-fájl kiválasztása</h2>
      </div>
      <div
        className={`drop-zone ${dragging ? 'is-dragging' : ''}`}
        onDragEnter={(event) => {
          event.preventDefault();
          setDragging(true);
        }}
        onDragOver={(event) => event.preventDefault()}
        onDragLeave={() => setDragging(false)}
        onDrop={drop}
      >
        <div className="upload-icon" aria-hidden="true">
          ↥
        </div>
        <p>
          <strong>Húzd ide a beosztást</strong>
        </p>
        <p className="muted">vagy válaszd ki a számítógépedről (.xlsx)</p>
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.xls"
          onChange={choose}
          hidden
          data-testid="file-input"
        />
        <button
          type="button"
          className="button secondary"
          onClick={() => inputRef.current?.click()}
          disabled={disabled}
        >
          Fájl kiválasztása
        </button>
        {fileName && (
          <p className="selected-file" aria-live="polite">
            Kiválasztva: <strong>{fileName}</strong>
          </p>
        )}
      </div>
      <p className="privacy-note">
        <span aria-hidden="true">◆</span>A fájl feldolgozása helyben, a böngészőben történik. A
        beosztás nem kerül feltöltésre vagy eltárolásra.
      </p>
    </section>
  );
}
