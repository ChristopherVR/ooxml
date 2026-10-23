// xlsx chart save parity (`test-support/chart-save-parity.ts`) over the packaged workbooks' charts and the type, spacing and fill evidence parts.
import { describe } from 'vitest';
import { describeChartSaveParity } from './test-support/chart-save-parity';

describe('xlsx chart save parity', () => {
	describeChartSaveParity(
		(id) => !id.includes('.json#') || /excel-chart-(types|spacing|series-fill)/.test(id),
	);
});
