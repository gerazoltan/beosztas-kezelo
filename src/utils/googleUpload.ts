import type { GoogleEventState } from '../domain/types';

export function isGoogleUploadComplete(state: GoogleEventState | undefined): boolean {
  return state?.status === 'Létrehozva' || state?.status === 'Már szerepel a naptárban';
}

export function isGoogleSelectionLocked(state: GoogleEventState | undefined): boolean {
  return state?.status === 'Létrehozás folyamatban' || isGoogleUploadComplete(state);
}
