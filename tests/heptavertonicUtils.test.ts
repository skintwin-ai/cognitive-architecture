import { describe, expect, it } from 'vitest';
import { FINANCIAL_DATA } from '../src/components/heptavertonicData';
import { deriveBalanceState, formatCurrency, getBalanceColor } from '../src/components/heptavertonicUtils';

describe('heptavertonicUtils', () => {
  it('formats currency values for USD display', () => {
    expect(formatCurrency(1234567)).toBe('$1,234,567');
  });

  it('returns semantic balance colors', () => {
    expect(getBalanceColor(0.9)).toBe('#10B981');
    expect(getBalanceColor(0.75)).toBe('#F59E0B');
    expect(getBalanceColor(0.4)).toBe('#EF4444');
  });

  it('derives normalized balance metrics from financial data', () => {
    const balanceState = deriveBalanceState(FINANCIAL_DATA);

    expect(balanceState.potential).toBeGreaterThan(0);
    expect(balanceState.potential).toBeLessThanOrEqual(1);
    expect(balanceState.commitment).toBeGreaterThan(0);
    expect(balanceState.performance).toBeGreaterThan(0);
    expect(balanceState.coherence).toBeGreaterThan(0);
  });
});
