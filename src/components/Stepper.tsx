const STEPS = ['Fájl', 'Hónap', 'Dolgozó', 'Feldolgozás', 'Ellenőrzés', 'Export'];

export function Stepper({ current }: { current: number }) {
  return (
    <nav aria-label="Feldolgozás lépései" className="stepper">
      <ol>
        {STEPS.map((step, index) => (
          <li
            key={step}
            className={index + 1 <= current ? 'step-active' : ''}
            aria-current={index + 1 === current ? 'step' : undefined}
          >
            <span>{index + 1}</span>
            {step}
          </li>
        ))}
      </ol>
    </nav>
  );
}
