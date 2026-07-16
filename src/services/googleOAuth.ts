import { AppError } from '../domain/errors';

const SCRIPT_ID = 'google-identity-services';
const SCRIPT_URL = 'https://accounts.google.com/gsi/client';
const SCOPE =
  'https://www.googleapis.com/auth/calendar.events https://www.googleapis.com/auth/calendar.calendarlist.readonly';

export interface GoogleTokenResponse {
  access_token?: string;
  error?: string;
  error_description?: string;
}

interface GoogleTokenClient {
  requestAccessToken(options?: { prompt?: string }): void;
}

interface GoogleAccountsOAuth2 {
  initTokenClient(config: {
    client_id: string;
    scope: string;
    callback: (response: GoogleTokenResponse) => void;
    error_callback?: (error: unknown) => void;
  }): GoogleTokenClient;
  revoke(token: string, callback?: () => void): void;
}

declare global {
  interface Window {
    google?: { accounts?: { oauth2?: GoogleAccountsOAuth2 } };
  }
}

export async function loadGoogleIdentityServices(): Promise<void> {
  if (window.google?.accounts?.oauth2) return;
  await new Promise<void>((resolve, reject) => {
    const existing = document.getElementById(SCRIPT_ID) as HTMLScriptElement | null;
    if (existing) {
      existing.addEventListener('load', () => resolve(), { once: true });
      existing.addEventListener('error', () => reject(new AppError('GOOGLE_API_ERROR')), {
        once: true,
      });
      return;
    }
    const script = document.createElement('script');
    script.id = SCRIPT_ID;
    script.src = SCRIPT_URL;
    script.async = true;
    script.defer = true;
    script.addEventListener('load', () => resolve(), { once: true });
    script.addEventListener('error', () => reject(new AppError('GOOGLE_API_ERROR')), {
      once: true,
    });
    document.head.append(script);
  });
}

export async function requestGoogleAccessToken(clientId: string): Promise<string> {
  if (!clientId) throw new AppError('GOOGLE_NOT_CONFIGURED');
  await loadGoogleIdentityServices();
  const oauth2 = window.google?.accounts?.oauth2;
  if (!oauth2)
    throw new AppError('GOOGLE_API_ERROR', 'A Google Identity Services nem töltődött be.');
  return new Promise<string>((resolve, reject) => {
    const client = oauth2.initTokenClient({
      client_id: clientId,
      scope: SCOPE,
      callback: (response) => {
        if (response.access_token) resolve(response.access_token);
        else
          reject(
            new AppError('GOOGLE_ACCESS_DENIED', response.error_description ?? response.error),
          );
      },
      error_callback: (error) =>
        reject(
          new AppError(
            'GOOGLE_ACCESS_DENIED',
            error instanceof Error ? error.message : String(error),
          ),
        ),
    });
    client.requestAccessToken({ prompt: 'consent' });
  });
}

export function revokeGoogleToken(token: string): void {
  window.google?.accounts?.oauth2?.revoke(token);
}
