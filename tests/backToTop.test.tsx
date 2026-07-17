import { fireEvent, render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { BACK_TO_TOP_THRESHOLD, BackToTopButton } from '../src/components/BackToTopButton';

function setScrollY(value: number): void {
  Object.defineProperty(window, 'scrollY', {
    configurable: true,
    value,
  });
}

function setReducedMotion(matches: boolean): void {
  vi.stubGlobal(
    'matchMedia',
    vi.fn().mockReturnValue({
      matches,
      media: '(prefers-reduced-motion: reduce)',
      onchange: null,
      addListener: vi.fn(),
      removeListener: vi.fn(),
      addEventListener: vi.fn(),
      removeEventListener: vi.fn(),
      dispatchEvent: vi.fn(),
    } satisfies MediaQueryList),
  );
}

describe('BackToTopButton', () => {
  beforeEach(() => {
    setScrollY(0);
    setReducedMotion(false);
  });

  afterEach(() => {
    vi.unstubAllGlobals();
    vi.restoreAllMocks();
    setScrollY(0);
  });

  it('a tetején rejtett, a 450 pixeles küszöbnél látható, majd fent ismét eltűnik', () => {
    render(<BackToTopButton />);
    expect(
      screen.queryByRole('button', { name: 'Vissza az oldal tetejére' }),
    ).not.toBeInTheDocument();

    setScrollY(BACK_TO_TOP_THRESHOLD - 1);
    fireEvent.scroll(window);
    expect(
      screen.queryByRole('button', { name: 'Vissza az oldal tetejére' }),
    ).not.toBeInTheDocument();

    setScrollY(BACK_TO_TOP_THRESHOLD);
    fireEvent.scroll(window);
    expect(screen.getByRole('button', { name: 'Vissza az oldal tetejére' })).toBeVisible();

    setScrollY(0);
    fireEvent.scroll(window);
    expect(
      screen.queryByRole('button', { name: 'Vissza az oldal tetejére' }),
    ).not.toBeInTheDocument();
  });

  it('kattintáskor simán az oldal tetejére görget', async () => {
    const user = userEvent.setup();
    const scrollTo = vi.spyOn(window, 'scrollTo').mockImplementation(() => undefined);
    setScrollY(BACK_TO_TOP_THRESHOLD);
    render(<BackToTopButton />);

    await user.click(screen.getByRole('button', { name: 'Vissza az oldal tetejére' }));

    expect(scrollTo).toHaveBeenCalledWith({ top: 0, behavior: 'smooth' });
  });

  it('csökkentett mozgásnál animáció nélkül görget', async () => {
    const user = userEvent.setup();
    const scrollTo = vi.spyOn(window, 'scrollTo').mockImplementation(() => undefined);
    setReducedMotion(true);
    setScrollY(BACK_TO_TOP_THRESHOLD);
    render(<BackToTopButton />);

    await user.click(screen.getByRole('button', { name: 'Vissza az oldal tetejére' }));

    expect(scrollTo).toHaveBeenCalledWith({ top: 0, behavior: 'auto' });
  });
});
