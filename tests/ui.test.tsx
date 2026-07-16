import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { describe, expect, it, vi } from 'vitest';
import App from '../src/App';
import { asFile, workbookBuffer } from './fixtures/syntheticWorkbook';

describe('felhasználói felület', () => {
  it('megjeleníti az adatvédelmi tájékoztatást és a kikapcsolt Google-integrációt', () => {
    render(<App />);
    expect(screen.getByText(/A fájl feldolgozása helyben/)).toBeInTheDocument();
    expect(screen.queryByText(/Google-integráció nincs konfigurálva/)).not.toBeInTheDocument();
  });

  it('elutasítja a hibás fájltípust', async () => {
    render(<App />);
    const file = asFile(new TextEncoder().encode('adat').buffer, 'minta.csv');
    fireEvent.change(screen.getByTestId('file-input'), { target: { files: [file] } });
    expect(await screen.findByText(/Nem támogatott fájltípus/)).toBeInTheDocument();
  });

  it('érthető hibát ad sérült Excelre', async () => {
    const user = userEvent.setup();
    render(<App />);
    const file = asFile(new TextEncoder().encode('nem zip').buffer, 'minta.xlsx');
    await user.upload(screen.getByTestId('file-input'), file);
    expect(await screen.findByText(/nem olvasható vagy sérült/)).toBeInTheDocument();
  });

  it('feltöltés, hónap- és dolgozóválasztás után feldolgoz és csak biztos eseményt exportál', async () => {
    const user = userEvent.setup();
    const createObjectURL = vi.fn(() => 'blob:test');
    Object.defineProperty(URL, 'createObjectURL', { configurable: true, value: createObjectURL });
    Object.defineProperty(URL, 'revokeObjectURL', { configurable: true, value: vi.fn() });
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(() => undefined);
    render(<App />);
    await user.upload(screen.getByTestId('file-input'), asFile(await workbookBuffer()));
    expect(await screen.findByRole('option', { name: /2026\. augusztus/ })).toBeInTheDocument();
    await user.selectOptions(screen.getByLabelText('Dolgozó'), 'teszt elek');
    await user.click(screen.getByRole('button', { name: 'Beosztás feldolgozása' }));
    expect(await screen.findByText('Nappalos 06–18')).toBeInTheDocument();
    expect(screen.getByText('Nappalos 10–22')).toBeInTheDocument();
    expect(screen.getAllByText('Hibás párosítás').length).toBeGreaterThan(0);
    const exportButton = screen.getByRole('button', { name: 'ICS letöltése' });
    expect(exportButton).toBeEnabled();
    await user.click(exportButton);
    expect(createObjectURL).toHaveBeenCalledOnce();
    expect(screen.getByText(/Google-integráció nincs konfigurálva/)).toBeInTheDocument();
  });

  it('tiltja az exportot, ha minden biztos eseményt kizártak', async () => {
    const user = userEvent.setup();
    render(<App />);
    await user.upload(screen.getByTestId('file-input'), asFile(await workbookBuffer()));
    await user.selectOptions(await screen.findByLabelText('Dolgozó'), 'teszt elek');
    await user.click(screen.getByRole('button', { name: 'Beosztás feldolgozása' }));
    const selectAll = await screen.findByLabelText('Összes biztos kijelölése');
    await user.click(selectAll);
    await waitFor(() =>
      expect(screen.getByRole('button', { name: 'ICS letöltése' })).toBeDisabled(),
    );
  });
});
