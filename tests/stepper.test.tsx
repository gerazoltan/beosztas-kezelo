import { fireEvent, render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { describe, expect, it, vi } from 'vitest';
import { Stepper } from '../src/components/Stepper';
import { deriveWorkflowProgress } from '../src/utils/workflowProgress';

const initialSteps = deriveWorkflowProgress({
  fileLoaded: false,
  monthSelected: false,
  employeeSelected: false,
  resultReady: false,
  hasSelectedExportableEvent: false,
  hasCompletedGoogleEvent: false,
  googleUploadInProgress: false,
  googleUploadFailed: false,
  icsExported: false,
});

describe('Stepper', () => {
  it('szövegesen is jelzi az egyetlen aktuális és a tiltott lépéseket', () => {
    render(<Stepper steps={initialSteps} onNavigate={vi.fn()} />);

    expect(screen.getByRole('button', { name: /Fájl.*Aktuális/u })).toHaveAttribute(
      'aria-current',
      'step',
    );
    expect(screen.getAllByText('Nem elérhető')).toHaveLength(5);
    expect(screen.getByRole('button', { name: /Export.*Nem elérhető/u })).toBeDisabled();
  });

  it('a tiltott lépés nem navigál, az aktuális és teljesített lépés igen', async () => {
    const user = userEvent.setup();
    const onNavigate = vi.fn();
    const completedSteps = deriveWorkflowProgress({
      fileLoaded: true,
      monthSelected: true,
      employeeSelected: false,
      resultReady: false,
      hasSelectedExportableEvent: false,
      hasCompletedGoogleEvent: false,
      googleUploadInProgress: false,
      googleUploadFailed: false,
      icsExported: false,
    });
    render(<Stepper steps={completedSteps} onNavigate={onNavigate} />);

    await user.click(screen.getByRole('button', { name: /Export.*Nem elérhető/u }));
    expect(onNavigate).not.toHaveBeenCalled();

    await user.click(screen.getByRole('button', { name: /Fájl.*Teljesítve/u }));
    await user.click(screen.getByRole('button', { name: /Dolgozó.*Aktuális/u }));
    expect(onNavigate).toHaveBeenNthCalledWith(1, 'file');
    expect(onNavigate).toHaveBeenNthCalledWith(2, 'employee');
  });

  it('a görgetés nem változtatja meg a kapott folyamatállapotot', () => {
    render(<Stepper steps={initialSteps} onNavigate={vi.fn()} />);
    const before = [...document.querySelectorAll<HTMLElement>('[data-step-id]')].map(
      (step) => step.dataset.state,
    );

    fireEvent.scroll(window);

    expect(
      [...document.querySelectorAll<HTMLElement>('[data-step-id]')].map(
        (step) => step.dataset.state,
      ),
    ).toEqual(before);
  });
});
