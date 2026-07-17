import { render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { describe, expect, it, vi } from 'vitest';
import App from '../src/App';
import { asFile, workbookBuffer } from './fixtures/syntheticWorkbook';

vi.mock('../src/components/GooglePanel', async () => {
  const { useEffect, useState } = await import('react');
  return {
    GooglePanel: ({
      visible = true,
      resetKey,
      onNewSchedule,
    }: {
      visible?: boolean;
      resetKey: number;
      onNewSchedule: () => void;
    }) => {
      const [signedIn, setSignedIn] = useState(false);
      const [completed, setCompleted] = useState(false);
      useEffect(() => setCompleted(false), [resetKey]);
      if (!visible) return null;
      if (!signedIn) {
        return (
          <button type="button" onClick={() => setSignedIn(true)}>
            Teszt Google-bejelentkezés
          </button>
        );
      }
      return completed ? (
        <section>
          <h3>Sikeres naptárfeltöltés</h3>
          <button type="button" onClick={onNewSchedule}>
            Új beosztás feldolgozása
          </button>
        </section>
      ) : (
        <button type="button" onClick={() => setCompleted(true)}>
          Tesztfeltöltés befejezése
        </button>
      );
    },
  };
});

async function prepareProcessedSchedule(user: ReturnType<typeof userEvent.setup>): Promise<void> {
  render(<App />);
  await user.upload(screen.getByTestId('file-input'), asFile(await workbookBuffer()));
  await user.selectOptions(await screen.findByLabelText('Dolgozó'), 'teszt elek');
  await user.click(screen.getByRole('button', { name: 'Beosztás feldolgozása' }));
  await user.click(await screen.findByRole('button', { name: 'Teszt Google-bejelentkezés' }));
  await user.click(await screen.findByRole('button', { name: 'Tesztfeltöltés befejezése' }));
  expect(await screen.findByText('Sikeres naptárfeltöltés')).toBeVisible();
}

describe('Google feltöltési eredmény visszaállítása', () => {
  it('új Excel-fájl választásakor törli az eredményt', async () => {
    const user = userEvent.setup();
    await prepareProcessedSchedule(user);

    await user.upload(
      screen.getByTestId('file-input'),
      asFile(await workbookBuffer(), 'másik-minta.xlsx'),
    );

    await waitFor(() =>
      expect(screen.queryByText('Sikeres naptárfeltöltés')).not.toBeInTheDocument(),
    );
  });

  it('másik hónap választásakor törli az eredményt', async () => {
    const user = userEvent.setup();
    render(<App />);
    await user.upload(
      screen.getByTestId('file-input'),
      asFile(
        await workbookBuffer([
          { name: 'Július', year: 2026, monthName: 'július', days: 31 },
          { name: 'Augusztus', year: 2026, monthName: 'augusztus', days: 31 },
        ]),
      ),
    );
    await user.selectOptions(await screen.findByLabelText('Dolgozó'), 'teszt elek');
    await user.click(screen.getByRole('button', { name: 'Beosztás feldolgozása' }));
    await user.click(await screen.findByRole('button', { name: 'Teszt Google-bejelentkezés' }));
    await user.click(await screen.findByRole('button', { name: 'Tesztfeltöltés befejezése' }));
    expect(await screen.findByText('Sikeres naptárfeltöltés')).toBeVisible();

    const monthSelect = screen.getByLabelText<HTMLSelectElement>('Hónap');
    const otherMonth = [...monthSelect.options].find(
      (option) => option.value !== monthSelect.value,
    );
    if (!otherMonth) throw new Error('Hiányzó második teszthónap.');
    await user.selectOptions(monthSelect, otherMonth.value);

    expect(screen.queryByText('Sikeres naptárfeltöltés')).not.toBeInTheDocument();
  });

  it('másik dolgozó választásakor törli az eredményt', async () => {
    const user = userEvent.setup();
    await prepareProcessedSchedule(user);

    await user.selectOptions(screen.getByLabelText('Dolgozó'), 'minta anna');

    expect(screen.queryByText('Sikeres naptárfeltöltés')).not.toBeInTheDocument();
  });

  it('újrafeldolgozáskor törli az eredményt', async () => {
    const user = userEvent.setup();
    await prepareProcessedSchedule(user);

    await user.click(screen.getByRole('button', { name: 'Beosztás feldolgozása' }));

    expect(screen.queryByText('Sikeres naptárfeltöltés')).not.toBeInTheDocument();
    expect(screen.getByRole('button', { name: 'Tesztfeltöltés befejezése' })).toBeVisible();
  });

  it('egyszerű kijelölésmódosításkor megőrzi az eredményt', async () => {
    const user = userEvent.setup();
    await prepareProcessedSchedule(user);

    await user.click(screen.getAllByRole('checkbox')[1] as HTMLElement);

    expect(screen.getByText('Sikeres naptárfeltöltés')).toBeVisible();
  });

  it('az új beosztás gomb törli a feldolgozási állapotot, simán felgörget és megőrzi a Google-bejelentkezést', async () => {
    const user = userEvent.setup();
    const scrollIntoView = vi.fn();
    Object.defineProperty(HTMLElement.prototype, 'scrollIntoView', {
      configurable: true,
      value: scrollIntoView,
    });
    await prepareProcessedSchedule(user);

    await user.click(screen.getByRole('button', { name: 'Új beosztás feldolgozása' }));

    expect(scrollIntoView).toHaveBeenCalledWith({ behavior: 'smooth', block: 'start' });
    expect(screen.queryByText(/Kiválasztva:/u)).not.toBeInTheDocument();
    expect(screen.queryByLabelText('Hónap')).not.toBeInTheDocument();
    expect(screen.queryByLabelText('Dolgozó')).not.toBeInTheDocument();
    expect(screen.queryByRole('heading', { name: 'Ellenőrzés' })).not.toBeInTheDocument();
    expect(screen.queryByText('Sikeres naptárfeltöltés')).not.toBeInTheDocument();

    await user.upload(
      screen.getByTestId('file-input'),
      asFile(await workbookBuffer(), 'új-minta.xlsx'),
    );
    await user.selectOptions(await screen.findByLabelText('Dolgozó'), 'teszt elek');
    await user.click(screen.getByRole('button', { name: 'Beosztás feldolgozása' }));

    expect(await screen.findByRole('button', { name: 'Tesztfeltöltés befejezése' })).toBeVisible();
    expect(
      screen.queryByRole('button', { name: 'Teszt Google-bejelentkezés' }),
    ).not.toBeInTheDocument();

    Reflect.deleteProperty(HTMLElement.prototype, 'scrollIntoView');
  });
});
