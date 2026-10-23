// xlsx chart save parity (`test-support/chart-save-parity.ts`) over the Change Colors evidence parts.
import { describe } from 'vitest';
import { describeChartSaveParity } from './test-support/chart-save-parity';

describe('xlsx chart save parity', () => {
	describeChartSaveParity((id) => id.startsWith('excel-chart-palette-edits.json#'));
});
