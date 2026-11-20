import { expect, test } from '@playwright/test';
import { heapReport, openAndCloseEditor } from '../../scripts/heap-snapshot.mjs';
import { excelFixture } from './fixture-files';
import { newWorkbook } from './helpers';

// Shapes, form controls and charts paint with `url(#id)` clip paths and fills. Chromium's style
// cache kept a closed editor alive through them until some later stylesheet change.
const FIXTURES = ['excel-shapes.xlsx', 'excel-chart-families.xlsx'];
const CYCLES = 2;
const PROBE = /^<xlsx-editor[^>]*data-leak-probe/;

test('editors opened and closed repeatedly are collected', async ({ page }) => {
	test.setTimeout(240_000);
	await newWorkbook(page);
	const session = await page.context().newCDPSession(page);
	const start = await heapReport(session, PROBE);
	let references = 0;
	for (const fileName of FIXTURES) {
		const bytes = [...excelFixture(fileName)];
		for (let cycle = 1; cycle <= CYCLES; cycle++) {
			references += await page.evaluate(openAndCloseEditor, {
				tag: 'xlsx-editor',
				bytes,
				fileName,
			});
			const report = await heapReport(session, PROBE);
			test.info().annotations.push({
				type: 'heap',
				description: `${fileName} close ${cycle}: ${report.count} closed editors alive, ${report.detachedBytes} detached bytes, heap ${report.heapBytes - start.heapBytes} bytes over the start`,
			});
			expect.soft(report.count, `${fileName}, close ${cycle}`).toBe(0);
		}
	}
	expect(references).toBeGreaterThan(0);
});
