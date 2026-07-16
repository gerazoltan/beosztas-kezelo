import { describe, expect, it } from 'vitest';
import { instantToLocal, zonedLocalToInstant } from '../src/services/dates';

describe('Europe/Budapest dátumkezelés', () => {
  it('téli időpontot CET szerint alakít', () => {
    expect(zonedLocalToInstant('2026-01-15T06:00:00').toISOString()).toBe(
      '2026-01-15T05:00:00.000Z',
    );
  });

  it('nyári időpontot CEST szerint alakít', () => {
    expect(zonedLocalToInstant('2026-08-15T06:00:00').toISOString()).toBe(
      '2026-08-15T04:00:00.000Z',
    );
    expect(instantToLocal('2026-08-15T04:00:00Z')).toBe('2026-08-15T06:00:00');
  });
});
