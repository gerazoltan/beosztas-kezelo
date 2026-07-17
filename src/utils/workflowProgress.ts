import type { AppErrorCode } from '../domain/errors';

export type WorkflowStepId = 'file' | 'month' | 'employee' | 'processing' | 'review' | 'export';

export type WorkflowStepState = 'current' | 'complete' | 'unavailable' | 'error';

export interface WorkflowStep {
  id: WorkflowStepId;
  label: string;
  state: WorkflowStepState;
}

export interface WorkflowProgressInput {
  fileLoaded: boolean;
  monthSelected: boolean;
  employeeSelected: boolean;
  resultReady: boolean;
  hasSelectedExportableEvent: boolean;
  hasCompletedGoogleEvent: boolean;
  googleUploadInProgress: boolean;
  googleUploadFailed: boolean;
  icsExported: boolean;
  errorCode?: AppErrorCode;
}

const FILE_ERROR_CODES = new Set<AppErrorCode>([
  'INVALID_FILE_TYPE',
  'LEGACY_XLS',
  'CORRUPT_WORKBOOK',
  'PROTECTED_WORKBOOK',
]);

const MONTH_ERROR_CODES = new Set<AppErrorCode>([
  'NO_MONTH_SHEET',
  'AMBIGUOUS_MONTH',
  'AMBIGUOUS_YEAR',
  'NAME_HEADER_MISSING',
  'NAME_HEADER_MULTIPLE',
  'NO_EMPLOYEE',
  'INVALID_DAY_COUNT',
]);

const EMPLOYEE_ERROR_CODES = new Set<AppErrorCode>(['EMPLOYEE_NOT_FOUND', 'EMPLOYEE_DUPLICATE']);

export function deriveWorkflowProgress(input: WorkflowProgressInput): WorkflowStep[] {
  const fileError =
    !input.fileLoaded && Boolean(input.errorCode && FILE_ERROR_CODES.has(input.errorCode));
  const monthError =
    !input.fileLoaded && Boolean(input.errorCode && MONTH_ERROR_CODES.has(input.errorCode));
  const fileComplete = input.fileLoaded || monthError;
  const employeeError =
    input.monthSelected && Boolean(input.errorCode && EMPLOYEE_ERROR_CODES.has(input.errorCode));
  const processingError =
    input.employeeSelected && !input.resultReady && Boolean(input.errorCode) && !employeeError;
  const reviewComplete =
    input.resultReady && (input.hasSelectedExportableEvent || input.hasCompletedGoogleEvent);
  const googleExportComplete =
    input.hasCompletedGoogleEvent &&
    !input.hasSelectedExportableEvent &&
    !input.googleUploadInProgress &&
    !input.googleUploadFailed;
  const exportComplete =
    !input.googleUploadInProgress &&
    !input.googleUploadFailed &&
    (input.icsExported || googleExportComplete);

  return [
    {
      id: 'file',
      label: 'Fájl',
      state: fileError ? 'error' : fileComplete ? 'complete' : 'current',
    },
    {
      id: 'month',
      label: 'Hónap',
      state: monthError
        ? 'error'
        : !fileComplete
          ? 'unavailable'
          : input.monthSelected
            ? 'complete'
            : 'current',
    },
    {
      id: 'employee',
      label: 'Dolgozó',
      state: !input.monthSelected
        ? 'unavailable'
        : employeeError
          ? 'error'
          : input.employeeSelected
            ? 'complete'
            : 'current',
    },
    {
      id: 'processing',
      label: 'Feldolgozás',
      state: !input.employeeSelected
        ? 'unavailable'
        : input.resultReady
          ? 'complete'
          : processingError
            ? 'error'
            : 'current',
    },
    {
      id: 'review',
      label: 'Ellenőrzés',
      state: !input.resultReady ? 'unavailable' : reviewComplete ? 'complete' : 'current',
    },
    {
      id: 'export',
      label: 'Export',
      state: !reviewComplete
        ? 'unavailable'
        : input.googleUploadFailed
          ? 'error'
          : exportComplete
            ? 'complete'
            : 'current',
    },
  ];
}
