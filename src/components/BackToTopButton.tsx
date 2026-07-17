import { useEffect, useState } from 'react';

export const BACK_TO_TOP_THRESHOLD = 450;

function prefersReducedMotion(): boolean {
  return window.matchMedia?.('(prefers-reduced-motion: reduce)').matches ?? false;
}

export function BackToTopButton() {
  const [visible, setVisible] = useState(() => window.scrollY >= BACK_TO_TOP_THRESHOLD);

  useEffect(() => {
    const updateVisibility = () => setVisible(window.scrollY >= BACK_TO_TOP_THRESHOLD);
    updateVisibility();
    window.addEventListener('scroll', updateVisibility, { passive: true });
    return () => window.removeEventListener('scroll', updateVisibility);
  }, []);

  if (!visible) return null;

  return (
    <button
      type="button"
      className="back-to-top"
      aria-label="Vissza az oldal tetejére"
      onClick={() =>
        window.scrollTo({
          top: 0,
          behavior: prefersReducedMotion() ? 'auto' : 'smooth',
        })
      }
    >
      <svg viewBox="0 0 24 24" aria-hidden="true" focusable="false">
        <path d="m6 14 6-6 6 6" />
      </svg>
    </button>
  );
}
