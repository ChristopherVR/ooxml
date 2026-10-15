import { describe, expect, it } from 'vitest';
import { loadXlsx } from '../read/index';
import { createWorkbook } from '../workbook';
import { saveXlsx } from './index';

describe('sheet view on save', () => {
	it('saves a sheet built without a view as the default view', async () => {
		const workbook = createWorkbook();
		const sheet = workbook.sheets[0]!;
		delete (sheet as { view?: unknown }).view;
		const reopened = await loadXlsx(await saveXlsx(workbook));
		expect(reopened.sheets[0]?.view.showGridLines).toBe(true);
		expect(reopened.sheets[0]?.view.freeze).toBeUndefined();
	});
});
