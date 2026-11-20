import { expect, test, type Page } from '@playwright/test';
import { createWorkbook, putCell, saveXlsx, type Workbook } from 'ooxml-core/xlsx';
import { editor, grid, landingFileInput, openLanding, pageErrors } from './helpers';

/**
 * Opening a workbook runs the calculation it asks for on open (`calcPr fullCalcOnLoad`): briefly on
 * the page, then in a worker when it is slow, so the page keeps painting while it runs, and a
 * calculation past its time limit leaves saved values with a notice.
 */

function workbookWith(fill: (workbook: Workbook) => void): Promise<Uint8Array> {
	const workbook = createWorkbook({ sheets: ['Sheet1'] });
	fill(workbook);
	workbook.fullCalcOnLoad = true;
	return saveXlsx(workbook);
}

/** Starts a timer on the page that records the longest gap between its ticks. */
async function watchMainThread(page: Page): Promise<() => Promise<number>> {
	await page.evaluate(() => {
		const state = { last: performance.now(), longest: 0 };
		(window as unknown as { __mainThread: typeof state }).__mainThread = state;
		setInterval(() => {
			const now = performance.now();
			state.longest = Math.max(state.longest, now - state.last);
			state.last = now;
		}, 20);
	});
	return () =>
		page.evaluate(
			() => (window as unknown as { __mainThread: { longest: number } }).__mainThread.longest,
		);
}

async function open(page: Page, name: string, bytes: Uint8Array) {
	await (
		await landingFileInput(page)
	).setInputFiles({
		name,
		mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
		buffer: Buffer.from(bytes),
	});
}

test('a stale workbook is calculated on open', async ({ page }) => {
	const errors = pageErrors(page);
	const bytes = await workbookWith((wb) => {
		putCell(wb.sheets[0]!, 0, 0, { value: 2 });
		putCell(wb.sheets[0]!, 0, 1, { value: 99, formula: 'A1*3' });
	});
	await openLanding(page);
	await open(page, 'stale.xlsx', bytes);
	await expect(grid(page)).toContainText('6');
	await expect(grid(page)).not.toContainText('99');
	expect(errors).toEqual([]);
});

test('a formula that would take tens of seconds stops at the budget without freezing the page', async ({
	page,
}) => {
	const bytes = await workbookWith((wb) => {
		putCell(wb.sheets[0]!, 0, 0, {
			value: 0,
			formula: 'SUM(MMULT(SEQUENCE(1500,1500),SEQUENCE(1500,1500)))',
		});
	});
	await openLanding(page);
	const longest = await watchMainThread(page);
	await open(page, 'mmult.xlsx', bytes);
	await expect(grid(page)).toContainText('#NUM!');
	expect(await longest()).toBeLessThan(2000);
});

test('a calculation past the time limit leaves saved values and says so', async ({ page }) => {
	test.setTimeout(120_000);
	const rows = 20_000;
	const bytes = await workbookWith((wb) => {
		const sheet = wb.sheets[0]!;
		for (let r = 0; r < rows; r++) {
			putCell(sheet, r, 0, { value: (r * 7919) % rows });
			// One scan of the column per formula: 400 million tests, far past the ten-second limit.
			putCell(sheet, r, 1, { value: -1, formula: `COUNTIF($A$1:$A$${rows},">"&A${r + 1})` });
		}
	});
	await openLanding(page);
	const longest = await watchMainThread(page);
	await open(page, 'rank.xlsx', bytes);
	await expect(editor(page).locator('.xve-toast')).toContainText(
		'Formulas were not all recalculated on open',
		{ timeout: 60_000 },
	);
	// The rows the page reached before it handed over are calculated; the rest keep the saved -1.
	const saved = await editor(page).evaluate((element) => {
		const rows = (element as unknown as { workbook: Workbook }).workbook.sheets[0]!.rows;
		return [...rows.values()].filter((row) => row.get(1)?.value === -1).length;
	});
	expect(saved).toBeGreaterThan(0);
	expect(await longest()).toBeLessThan(2000);
});
