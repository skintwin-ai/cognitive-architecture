import { describe, expect, it } from 'vitest';
import { FLOWS, TRIADS, VALID_NODE_IDS } from '../src/components/heptavertonicData';

describe('heptavertonic data integrity', () => {
  it('defines all seven triads with exactly three vertices each', () => {
    expect(TRIADS).toHaveLength(7);
    TRIADS.forEach((triad) => {
      expect(triad.vertices).toHaveLength(3);
    });
  });

  it('ensures every flow endpoint references a known node', () => {
    FLOWS.forEach((flow) => {
      expect(VALID_NODE_IDS.has(flow.from)).toBe(true);
      expect(VALID_NODE_IDS.has(flow.to)).toBe(true);
    });
  });
});
