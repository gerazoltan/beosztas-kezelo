import { render, screen, within } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { describe, expect, it, vi } from 'vitest';
import { ReviewTable } from '../src/components/ReviewTable';
import type { GoogleEventState, ReviewRow } from '../src/domain/types';

describe('érintett számos jelölések technikai részletei', () => {
  it('megjeleníti a stílust, a korrekciót, a párosítást, az időket és a Google-átfedést', async () => {
    const user = userEvent.setup();
    const row: ReviewRow = {
      id: 'row-1',
      date: { year: 2026, month: 8, day: 31 },
      marker: '17',
      shiftType: '24 órás szolgálat',
      serviceCategory: 'Parti szolgálat',
      summary: 'OMSZ',
      status: 'Exportálható',
      note: 'Helyreállított zöld 17.',
      timeRule: 'Hónapvégi 17 → 07:00–másnap 06:59',
      technicalNote: 'Napi összeállításból helyreállítva.',
      diagnostics: [
        {
          address: 'C5',
          rawValue: '17',
          displayedText: '17',
          isMerged: true,
          mergeMaster: 'C5',
          positionInDayGroup: 1,
          fillColor: '#FFF2CC',
          fontColorRaw: 'argb=FF008000',
          fontColor: '#008000',
          underline: false,
          italic: false,
          bold: false,
        },
      ],
      serviceResolution: {
        originalServiceCategory: 'Nem meghatározható',
        finalServiceCategory: 'Parti szolgálat',
        formattingCorrectionApplied: true,
        dailyInferenceApplied: true,
        assumedBoundaryPairing: true,
        pairingSource: 'assumed',
        pairingCell: 'feltételezett következő havi 7',
        finalShiftTime: {
          start: '2026-08-31T07:00:00',
          end: '2026-09-01T07:00:00',
        },
        finalCalendarTime: {
          start: '2026-08-31T07:00:00',
          end: '2026-09-01T06:59:00',
        },
      },
      event: {
        id: 'event-1',
        summary: 'OMSZ',
        shiftType: '24 órás szolgálat',
        serviceCategory: 'Parti szolgálat',
        shiftTime: {
          start: '2026-08-31T07:00:00',
          end: '2026-09-01T07:00:00',
        },
        calendarTime: {
          start: '2026-08-31T07:00:00',
          end: '2026-09-01T06:59:00',
        },
        timeZone: 'Europe/Budapest',
      },
    };
    const googleState: GoogleEventState = {
      status: 'Már szerepel a naptárban',
      message: 'Már szerepel.',
      technicalDetails: 'Átfedő előző havi teljes esemény található: igen.',
    };

    render(
      <ReviewTable
        rows={[row]}
        selected={new Set(['event-1'])}
        googleStates={new Map([['event-1', googleState]])}
        onToggle={vi.fn()}
        onSelectAll={vi.fn()}
      />,
    );

    const tableRow = screen.getByText('Parti szolgálat – következtetett').closest('tr');
    if (!tableRow) throw new Error('Hiányzó technikai tesztsor.');
    expect(
      within(tableRow).getByText('Átfedő előző havi teljes esemény található: igen.'),
    ).toBeInTheDocument();
    await user.click(within(tableRow).getByText('Technikai részletek'));
    expect(within(tableRow).getByText('Eredeti szolgálati kategória')).toBeVisible();
    expect(within(tableRow).getByText('Nem meghatározható')).toBeVisible();
    expect(within(tableRow).getByText('Formázási korrekció történt')).toBeVisible();
    expect(within(tableRow).getByText('Napi összeállításból következtetve')).toBeVisible();
    expect(within(tableRow).getByText('Feltételezett hónaphatár-párosítás')).toBeVisible();
    expect(within(tableRow).getByText('feltételezett következő havi 7')).toBeVisible();
    expect(within(tableRow).getByText('C5 (merge master: C5)')).toBeVisible();
    expect(within(tableRow).getByText('argb=FF008000')).toBeVisible();
    expect(within(tableRow).getByText('#008000')).toBeVisible();
    expect(
      within(tableRow).getByText(
        '2026-08-31T07:00:00 – 2026-09-01T06:59:00',
      ),
    ).toBeVisible();
  });
});
