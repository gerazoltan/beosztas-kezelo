import { fireEvent, render, screen, waitFor, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import App from '../src/App';
import { asFile, workbookBuffer } from './fixtures/syntheticWorkbook';

describe('felhasználói felület', () => {
  beforeEach(() => {
    vi.stubEnv('VITE_GOOGLE_CLIENT_ID', '');
  });

  afterEach(() => {
    vi.unstubAllEnvs();
  });

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
    expect(screen.getByRole('button', { name: /Fájl.*Hiba/u })).toBeVisible();
    expect(screen.getByRole('button', { name: /Hónap.*Nem elérhető/u })).toBeDisabled();
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
    expect(screen.getByRole('button', { name: /Fájl.*Aktuális/u })).toHaveAttribute(
      'aria-current',
      'step',
    );
    expect(screen.getByRole('button', { name: /Export.*Nem elérhető/u })).toBeDisabled();
    await user.upload(screen.getByTestId('file-input'), asFile(await workbookBuffer()));
    const monthOption = await screen.findByRole('option', { name: '2026. augusztus' });
    expect(monthOption).toBeInTheDocument();
    expect(monthOption.textContent?.match(/augusztus/gu)).toHaveLength(1);
    expect(screen.getByLabelText<HTMLSelectElement>('Hónap')).toHaveValue('2026-8-Augusztus');
    expect(screen.getByRole('button', { name: /Fájl.*Teljesítve/u })).toBeEnabled();
    expect(screen.getByRole('button', { name: /Hónap.*Teljesítve/u })).toBeEnabled();
    expect(screen.getByRole('button', { name: /Dolgozó.*Aktuális/u })).toBeEnabled();
    await user.selectOptions(screen.getByLabelText('Dolgozó'), 'teszt elek');
    expect(screen.getByRole('button', { name: /Feldolgozás.*Aktuális/u })).toBeEnabled();
    await user.click(screen.getByRole('button', { name: 'Beosztás feldolgozása' }));
    const reviewTable = await screen.findByRole('table');
    expect(screen.getByRole('button', { name: /Ellenőrzés.*Teljesítve/u })).toBeEnabled();
    expect(screen.getByRole('button', { name: /Export.*Aktuális/u })).toBeEnabled();
    expect(within(reviewTable).getByText('Nappalos 06–18')).toBeInTheDocument();
    expect(within(reviewTable).getByText('Nappalos 10–22')).toBeInTheDocument();
    const twentyFourHourRow = within(reviewTable).getByText('24 órás szolgálat').closest('tr');
    if (!twentyFourHourRow) throw new Error('Hiányzó 24 órás szolgálati sor.');
    expect(within(twentyFourHourRow).getAllByText('07:00')).toHaveLength(2);
    expect(
      within(twentyFourHourRow).getByText(
        'A naptáresemény befejezése 06:59 a jobb naptári elkülönítés érdekében.',
      ),
    ).toBeInTheDocument();
    expect(screen.getAllByText('Hibás párosítás').length).toBeGreaterThan(0);
    const exportButton = screen.getByRole('button', { name: 'ICS letöltése' });
    expect(exportButton).toBeEnabled();
    await user.click(exportButton);
    expect(createObjectURL).toHaveBeenCalledOnce();
    expect(screen.getByRole('button', { name: /Export.*Teljesítve/u })).toBeEnabled();
    expect(screen.getByText(/Google-integráció nincs konfigurálva/)).toBeInTheDocument();

    await user.selectOptions(screen.getByLabelText('Dolgozó'), 'minta anna');
    expect(screen.getByRole('button', { name: /Feldolgozás.*Aktuális/u })).toBeEnabled();
    expect(screen.getByRole('button', { name: /Ellenőrzés.*Nem elérhető/u })).toBeDisabled();
    expect(screen.getByRole('button', { name: /Export.*Nem elérhető/u })).toBeDisabled();
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
    expect(screen.getByRole('button', { name: /Ellenőrzés.*Aktuális/u })).toBeEnabled();
    expect(screen.getByRole('button', { name: /Export.*Nem elérhető/u })).toBeDisabled();
  });

  it('teljesített lépésre kattintva a megfelelő valódi szakaszhoz görget', async () => {
    const user = userEvent.setup();
    const scrollIntoView = vi.fn();
    Object.defineProperty(HTMLElement.prototype, 'scrollIntoView', {
      configurable: true,
      value: scrollIntoView,
    });
    render(<App />);
    await user.upload(screen.getByTestId('file-input'), asFile(await workbookBuffer()));
    await screen.findByLabelText('Hónap');

    await user.click(screen.getByRole('button', { name: /Fájl.*Teljesítve/u }));
    expect(scrollIntoView).toHaveBeenLastCalledWith({ behavior: 'smooth', block: 'start' });

    scrollIntoView.mockClear();
    await user.click(screen.getByRole('button', { name: /Hónap.*Teljesítve/u }));
    expect(scrollIntoView).toHaveBeenCalledWith({ behavior: 'smooth', block: 'start' });

    Reflect.deleteProperty(HTMLElement.prototype, 'scrollIntoView');
  });

  it('többször szereplő névnél csak konkrét sorral teljesít, sorváltáskor pedig visszaáll', async () => {
    const user = userEvent.setup();
    render(<App />);
    await user.upload(
      screen.getByTestId('file-input'),
      asFile(
        await workbookBuffer([
          {
            name: 'Augusztus',
            year: 2026,
            monthName: 'augusztus',
            days: 31,
            duplicateEmployee: true,
          },
        ]),
      ),
    );
    await user.selectOptions(await screen.findByLabelText('Dolgozó'), 'teszt elek');

    expect(screen.getByRole('button', { name: /Dolgozó.*Aktuális/u })).toBeEnabled();
    expect(screen.getByRole('button', { name: /Feldolgozás.*Nem elérhető/u })).toBeDisabled();

    const rowSelect = screen.getByLabelText('Sor kézi kiválasztása');
    await user.selectOptions(rowSelect, '5');
    expect(screen.getByRole('button', { name: /Dolgozó.*Teljesítve/u })).toBeEnabled();
    expect(screen.getByRole('button', { name: /Feldolgozás.*Aktuális/u })).toBeEnabled();

    await user.click(screen.getByRole('button', { name: 'Beosztás feldolgozása' }));
    expect(await screen.findByRole('heading', { name: 'Ellenőrzés' })).toBeVisible();

    await user.selectOptions(rowSelect, '7');
    expect(screen.queryByRole('heading', { name: 'Ellenőrzés' })).not.toBeInTheDocument();
    expect(screen.getByRole('button', { name: /Feldolgozás.*Aktuális/u })).toBeEnabled();
    expect(screen.getByRole('button', { name: /Ellenőrzés.*Nem elérhető/u })).toBeDisabled();
    expect(screen.getByRole('button', { name: /Export.*Nem elérhető/u })).toBeDisabled();
  });
});
