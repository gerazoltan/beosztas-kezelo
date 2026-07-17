import { describe, expect, it } from 'vitest';
import {
  deriveWorkflowProgress,
  type WorkflowProgressInput,
  type WorkflowStepId,
  type WorkflowStepState,
} from '../src/utils/workflowProgress';

const initialInput: WorkflowProgressInput = {
  fileLoaded: false,
  monthSelected: false,
  employeeSelected: false,
  resultReady: false,
  hasSelectedExportableEvent: false,
  hasCompletedGoogleEvent: false,
  googleUploadInProgress: false,
  googleUploadFailed: false,
  icsExported: false,
};

function states(
  input: Partial<WorkflowProgressInput> = {},
): Record<WorkflowStepId, WorkflowStepState> {
  return Object.fromEntries(
    deriveWorkflowProgress({ ...initialInput, ...input }).map((step) => [step.id, step.state]),
  ) as Record<WorkflowStepId, WorkflowStepState>;
}

describe('a folyamatjelző levezetett állapota', () => {
  it('induláskor csak a fájl lépés aktuális', () => {
    expect(states()).toEqual({
      file: 'current',
      month: 'unavailable',
      employee: 'unavailable',
      processing: 'unavailable',
      review: 'unavailable',
      export: 'unavailable',
    });
  });

  it('a tényleges fájl-, hónap- és dolgozóválasztást követi', () => {
    expect(states({ fileLoaded: true })).toMatchObject({
      file: 'complete',
      month: 'current',
      employee: 'unavailable',
    });
    expect(states({ fileLoaded: true, monthSelected: true })).toMatchObject({
      file: 'complete',
      month: 'complete',
      employee: 'current',
      processing: 'unavailable',
    });
    expect(states({ fileLoaded: true, monthSelected: true, employeeSelected: true })).toMatchObject(
      {
        employee: 'complete',
        processing: 'current',
        review: 'unavailable',
      },
    );
  });

  it('feldolgozás után az ellenőrzés, kijelölés után az export lesz aktuális', () => {
    const processed = {
      fileLoaded: true,
      monthSelected: true,
      employeeSelected: true,
      resultReady: true,
    };
    expect(states(processed)).toMatchObject({
      processing: 'complete',
      review: 'current',
      export: 'unavailable',
    });
    expect(states({ ...processed, hasSelectedExportableEvent: true })).toMatchObject({
      review: 'complete',
      export: 'current',
    });
  });

  it('sikeres ICS-export után teljesítettnek jelöli az exportot', () => {
    expect(
      states({
        fileLoaded: true,
        monthSelected: true,
        employeeSelected: true,
        resultReady: true,
        hasSelectedExportableEvent: true,
        icsExported: true,
      }).export,
    ).toBe('complete');
  });

  it('teljes Google-siker után teljesített, folyamatban vagy hiba esetén nem teljesített az export', () => {
    const processed = {
      fileLoaded: true,
      monthSelected: true,
      employeeSelected: true,
      resultReady: true,
      hasCompletedGoogleEvent: true,
    };
    expect(states(processed).export).toBe('complete');
    expect(
      states({
        ...processed,
        hasSelectedExportableEvent: true,
        googleUploadInProgress: true,
      }).export,
    ).toBe('current');
    expect(
      states({
        ...processed,
        hasSelectedExportableEvent: true,
        googleUploadFailed: true,
      }).export,
    ).toBe('error');
    expect(
      states({
        ...processed,
        hasCompletedGoogleEvent: false,
        hasSelectedExportableEvent: true,
        googleUploadFailed: true,
      }).export,
    ).toBe('error');
  });

  it('a hibás fájlt nem, a felismerhető Excel hónaphibáját viszont a megfelelő lépésen jelzi', () => {
    expect(states({ errorCode: 'INVALID_FILE_TYPE' })).toMatchObject({
      file: 'error',
      month: 'unavailable',
    });
    expect(states({ errorCode: 'AMBIGUOUS_MONTH' })).toMatchObject({
      file: 'complete',
      month: 'error',
      employee: 'unavailable',
    });
  });

  it('új fájl, hónap- és dolgozóváltás után a későbbi lépések visszaállnak', () => {
    expect(
      states({
        fileLoaded: true,
        monthSelected: true,
        employeeSelected: true,
        resultReady: true,
        hasSelectedExportableEvent: true,
        icsExported: true,
      }).export,
    ).toBe('complete');

    expect(states()).toEqual({
      file: 'current',
      month: 'unavailable',
      employee: 'unavailable',
      processing: 'unavailable',
      review: 'unavailable',
      export: 'unavailable',
    });
    expect(states({ fileLoaded: true, monthSelected: true })).toMatchObject({
      employee: 'current',
      processing: 'unavailable',
      review: 'unavailable',
      export: 'unavailable',
    });
    expect(states({ fileLoaded: true, monthSelected: true, employeeSelected: true })).toMatchObject(
      {
        processing: 'current',
        review: 'unavailable',
        export: 'unavailable',
      },
    );
  });

  it('másik naptár választásakor megtartja az előzményeket, de az export ismét aktuális', () => {
    expect(
      states({
        fileLoaded: true,
        monthSelected: true,
        employeeSelected: true,
        resultReady: true,
        hasSelectedExportableEvent: true,
      }),
    ).toEqual({
      file: 'complete',
      month: 'complete',
      employee: 'complete',
      processing: 'complete',
      review: 'complete',
      export: 'current',
    });
  });
});
