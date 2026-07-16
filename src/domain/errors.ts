export type AppErrorCode =
  | 'INVALID_FILE_TYPE'
  | 'LEGACY_XLS'
  | 'CORRUPT_WORKBOOK'
  | 'PROTECTED_WORKBOOK'
  | 'NO_MONTH_SHEET'
  | 'AMBIGUOUS_MONTH'
  | 'AMBIGUOUS_YEAR'
  | 'NAME_HEADER_MISSING'
  | 'NAME_HEADER_MULTIPLE'
  | 'NO_EMPLOYEE'
  | 'EMPLOYEE_NOT_FOUND'
  | 'EMPLOYEE_DUPLICATE'
  | 'INVALID_DAY_COUNT'
  | 'GOOGLE_NOT_CONFIGURED'
  | 'GOOGLE_ACCESS_DENIED'
  | 'GOOGLE_API_ERROR';

const MESSAGES: Record<AppErrorCode, string> = {
  INVALID_FILE_TYPE: 'Nem támogatott fájltípus. Válassz .xlsx Excel-fájlt.',
  LEGACY_XLS: 'A régi .xls formátum nem támogatott. Mentsd a munkafüzetet .xlsx formátumban.',
  CORRUPT_WORKBOOK: 'Az Excel-fájl nem olvasható vagy sérült.',
  PROTECTED_WORKBOOK: 'A jelszóval védett vagy titkosított Excel-fájl nem dolgozható fel.',
  NO_MONTH_SHEET: 'A munkafüzetben nem található felismerhető havi munkalap.',
  AMBIGUOUS_MONTH: 'A munkalap hónapja nem állapítható meg egyértelműen.',
  AMBIGUOUS_YEAR: 'A munkalap éve nem állapítható meg egyértelműen.',
  NAME_HEADER_MISSING: 'A havi munkalapon nem található egyértelmű „Név” fejléc.',
  NAME_HEADER_MULTIPLE: 'A havi munkalapon több „Név” fejléc található.',
  NO_EMPLOYEE: 'A kiválasztott havi munkalapon nem található dolgozó.',
  EMPLOYEE_NOT_FOUND: 'A kiválasztott dolgozó nem található a havi munkalapon.',
  EMPLOYEE_DUPLICATE: 'A kiválasztott név több sorban is szerepel; kézi döntés szükséges.',
  INVALID_DAY_COUNT: 'A munkalap a hónaphoz nem tartozó naptári napot is tartalmaz.',
  GOOGLE_NOT_CONFIGURED: 'A Google Naptár-integráció nincs konfigurálva.',
  GOOGLE_ACCESS_DENIED: 'A Google-hozzáférés nem engedélyezett.',
  GOOGLE_API_ERROR: 'A Google Naptár elérése hálózati vagy API-hiba miatt sikertelen.',
};

export class AppError extends Error {
  readonly code: AppErrorCode;
  readonly technicalDetails?: string;

  constructor(code: AppErrorCode, technicalDetails?: string) {
    super(MESSAGES[code]);
    this.name = 'AppError';
    this.code = code;
    this.technicalDetails = technicalDetails;
  }
}

export function toAppError(error: unknown): AppError {
  if (error instanceof AppError) return error;
  const detail = error instanceof Error ? error.message : String(error);
  const lowered = detail.toLocaleLowerCase('hu-HU');
  if (lowered.includes('password') || lowered.includes('encrypted')) {
    return new AppError('PROTECTED_WORKBOOK', detail);
  }
  return new AppError('CORRUPT_WORKBOOK', detail);
}
