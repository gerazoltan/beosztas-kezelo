import { describe, expect, it } from 'vitest';
import type { GoogleEventState } from '../src/domain/types';
import { isGoogleSelectionLocked, isGoogleUploadComplete } from '../src/utils/googleUpload';

function state(status: GoogleEventState['status']): GoogleEventState {
  return { status, message: 'teszt' };
}

describe('Google esemény kijelölhetősége', () => {
  it.each(['Létrehozva', 'Már szerepel a naptárban'] as const)(
    '%s állapotban lezárja és kizárja az eseményt az újabb feltöltésből',
    (status) => {
      expect(isGoogleUploadComplete(state(status))).toBe(true);
      expect(isGoogleSelectionLocked(state(status))).toBe(true);
    },
  );

  it('folyamatban tiltja a kijelölést, de még nem tekinti befejezettnek', () => {
    expect(isGoogleUploadComplete(state('Létrehozás folyamatban'))).toBe(false);
    expect(isGoogleSelectionLocked(state('Létrehozás folyamatban'))).toBe(true);
  });

  it('sikertelen állapotban engedi az újrapróbálási kijelölést', () => {
    expect(isGoogleUploadComplete(state('Sikertelen'))).toBe(false);
    expect(isGoogleSelectionLocked(state('Sikertelen'))).toBe(false);
  });
});
