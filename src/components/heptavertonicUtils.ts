import { FINANCIAL_DATA, type FinancialData } from './heptavertonicData';

export interface BalanceState {
  potential: number;
  commitment: number;
  performance: number;
  coherence: number;
}

export interface VisualizationStateHash {
  showFlows: boolean;
  showBalance: boolean;
  showFinancial: boolean;
  search: string;
}

export function clamp(value: number, min = 0, max = 1) {
  return Math.min(max, Math.max(min, value));
}

export function formatCurrency(value: number) {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    maximumFractionDigits: 0,
  }).format(value);
}

export function getBalanceColor(value: number) {
  if (value >= 0.85) {
    return '#10B981';
  }
  if (value >= 0.7) {
    return '#F59E0B';
  }
  return '#EF4444';
}

export function deriveBalanceState(financialData: FinancialData = FINANCIAL_DATA): BalanceState {
  const centerTotal = financialData.centers.reduce((sum, center) => sum + center.total, 0);
  const idea = financialData.centers[0]?.total ?? 0;
  const cerebral = financialData.centers[1]?.total ?? 0;
  const routine = financialData.centers[3]?.total ?? 0;
  const form = financialData.centers[4]?.total ?? 0;
  const adjustments = financialData.adjustments.reduce((sum, item) => sum + item.value, 0);
  const transformations = financialData.transformations.reduce((sum, item) => sum + item.value, 0);
  const eliminations = Math.abs(financialData.eliminations.reduce((sum, item) => sum + item.value, 0));

  return {
    potential: clamp((idea + form * 0.35) / centerTotal),
    commitment: clamp((routine + cerebral * 0.2) / centerTotal),
    performance: clamp((cerebral - eliminations * 0.25) / centerTotal + 0.5),
    coherence: clamp((adjustments + transformations) / (centerTotal * 0.25)),
  };
}

export function serializeStateToHash(state: VisualizationStateHash) {
  const params = new URLSearchParams();
  params.set('flows', state.showFlows ? '1' : '0');
  params.set('balance', state.showBalance ? '1' : '0');
  params.set('financial', state.showFinancial ? '1' : '0');
  if (state.search.trim()) {
    params.set('search', state.search.trim());
  }
  return params.toString();
}

export function parseStateFromHash(hash: string): VisualizationStateHash {
  const params = new URLSearchParams(hash.replace(/^#/, ''));
  return {
    showFlows: params.get('flows') !== '0',
    showBalance: params.get('balance') !== '0',
    showFinancial: params.get('financial') === '1',
    search: params.get('search') ?? '',
  };
}
