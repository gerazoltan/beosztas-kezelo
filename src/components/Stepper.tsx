import type { WorkflowStep, WorkflowStepId, WorkflowStepState } from '../utils/workflowProgress';

const STATE_META: Record<WorkflowStepState, { icon: string; label: string }> = {
  current: { icon: '●', label: 'Aktuális' },
  complete: { icon: '✓', label: 'Teljesítve' },
  unavailable: { icon: '–', label: 'Nem elérhető' },
  error: { icon: '!', label: 'Hiba' },
};

interface StepperProps {
  steps: WorkflowStep[];
  onNavigate: (stepId: WorkflowStepId) => void;
}

export function Stepper({ steps, onNavigate }: StepperProps) {
  return (
    <nav aria-label="Feldolgozás lépései" className="stepper">
      <ol>
        {steps.map((step, index) => {
          const state = STATE_META[step.state];
          const disabled = step.state === 'unavailable';
          return (
            <li key={step.id} className={`step-${step.state}`}>
              <button
                type="button"
                className="step-button"
                data-step-id={step.id}
                data-state={step.state}
                aria-current={step.state === 'current' ? 'step' : undefined}
                disabled={disabled}
                onClick={() => onNavigate(step.id)}
              >
                <span className="step-number">{index + 1}</span>
                <span className="step-copy">
                  <span className="step-name">{step.label}</span>
                  <span className="step-state">
                    <span aria-hidden="true">{state.icon}</span> {state.label}
                  </span>
                </span>
              </button>
            </li>
          );
        })}
      </ol>
    </nav>
  );
}
