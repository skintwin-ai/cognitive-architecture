import { render, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import HeptavertonicVisualization from '../src/components/HeptavertonicVisualization';

describe('HeptavertonicVisualization', () => {
  it('renders core controls and simulation content', () => {
    render(<HeptavertonicVisualization />);

    expect(screen.getByRole('heading', { name: /Heptavertonic Organizational Model/i })).toBeInTheDocument();
    expect(screen.getByLabelText(/Search architecture/i)).toBeInTheDocument();
    expect(screen.getByRole('button', { name: /Simulate/i })).toBeInTheDocument();
    expect(screen.getByText(/Step 1 of 7/i)).toBeInTheDocument();
  });
});
