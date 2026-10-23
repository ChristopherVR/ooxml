// xlsx chart save parity (`test-support/chart-save-parity.ts`) over the gradient and transparency evidence parts.
import { describe } from 'vitest';
import { describeChartSaveParity } from './test-support/chart-save-parity';

describe('xlsx chart save parity', () => {
	describeChartSaveParity((id) => /gradient|transparency/.test(id));
});
