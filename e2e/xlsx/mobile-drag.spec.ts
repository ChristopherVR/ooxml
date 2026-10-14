import { expect, test } from '@playwright/test';
import { createWorkbook, createEditSession, saveXlsx, type Workbook } from 'ooxml-core/xlsx';
import { editor, openLanding, FRAMEWORKS } from './helpers';

interface DrawingEditor extends HTMLElement {
	readonly workbook: Workbook;
	undo(): void;
}
test.use({ viewport: { width: 390, height: 844 }, hasTouch: true, isMobile: true });

for (const framework of FRAMEWORKS)
	test(`${framework}: native touch moves and resizes a chart with undo`, async ({ page }) => {
		await openLanding(page, framework);
		const workbook = createWorkbook();
		createEditSession(workbook).addChart(0, {
			chartType: 'column',
			showLegend: false,
			title: 'Touch chart',
			series: [{ name: 'Sales', categories: ['A', 'B'], values: [10, 20] }],
			anchor: {
				from: { row: 1, col: 0, rowOffset: 0, colOffset: 0 },
				ext: { cx: 200 * 9525, cy: 140 * 9525 },
			},
		});
		await page.locator('#landing-file').setInputFiles({
			name: 'touch-chart.xlsx',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			buffer: Buffer.from(await saveXlsx(workbook)),
		});
		const host = editor(page);
		const chart = host.locator('.xg-obj').first();
		await expect(chart).toBeVisible();
		const anchor = () =>
			host.evaluate((el) => (el as DrawingEditor).workbook.sheets[0]!.drawings[0]!.anchor);
		const before = await anchor();
		const box = (await chart.boundingBox())!;
		const x = Math.round(Math.min(box.x + box.width / 2, 300)),
			y = Math.round(Math.min(box.y + box.height / 2, 650));
		const cdp = await page.context().newCDPSession(page);
		await cdp.send('Input.dispatchTouchEvent', { type: 'touchStart', touchPoints: [{ x, y }] });
		for (let i = 1; i <= 10; i++) {
			await cdp.send('Input.dispatchTouchEvent', {
				type: 'touchMove',
				touchPoints: [{ x: x + i * 5, y: y + i * 3 }],
			});
			await page.waitForTimeout(25);
		}
		await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
		await expect.poll(anchor).not.toEqual(before);
		await host.evaluate((el) => (el as DrawingEditor).undo());
		await expect.poll(anchor).toEqual(before);
		// Native cancellation must remove the preview without writing an anchor.
		const cancelBox = (await chart.boundingBox())!;
		const cx = Math.round(cancelBox.x + cancelBox.width / 2),
			cy = Math.round(cancelBox.y + cancelBox.height / 2);
		await cdp.send('Input.dispatchTouchEvent', {
			type: 'touchStart',
			touchPoints: [{ x: cx, y: cy }],
		});
		for (let i = 1; i <= 5; i++) {
			await cdp.send('Input.dispatchTouchEvent', {
				type: 'touchMove',
				touchPoints: [{ x: cx + i * 4, y: cy + i * 3 }],
			});
			await page.waitForTimeout(25);
		}
		await expect.poll(async () => (await chart.boundingBox())!.x).toBeGreaterThan(cancelBox.x + 10);
		await cdp.send('Input.dispatchTouchEvent', { type: 'touchCancel', touchPoints: [] });
		await expect.poll(anchor).toEqual(before);
		await expect.poll(async () => (await chart.boundingBox())!.x).toBeCloseTo(cancelBox.x, 1);
		await chart.tap();
		const grip = chart.locator('[data-grip="se"]');
		const handle = (await grip.boundingBox())!;
		const hx = Math.round(handle.x + handle.width / 2),
			hy = Math.round(handle.y + handle.height / 2);
		await cdp.send('Input.dispatchTouchEvent', {
			type: 'touchStart',
			touchPoints: [{ x: hx, y: hy }],
		});
		for (let i = 1; i <= 10; i++) {
			await cdp.send('Input.dispatchTouchEvent', {
				type: 'touchMove',
				touchPoints: [{ x: hx + i * 3, y: hy + i * 2 }],
			});
			await page.waitForTimeout(25);
		}
		await cdp.send('Input.dispatchTouchEvent', { type: 'touchEnd', touchPoints: [] });
		await expect
			.poll(async () => (await chart.boundingBox())!.width)
			.toBeGreaterThan(box.width + 20);
		await host.evaluate((el) => (el as DrawingEditor).undo());
		await expect.poll(anchor).toEqual(before);
		for (const selector of ['.xfb-name-toggle', '.xfb-fx', '.xfb-expand', '.xst-btn', '.xst-tab']) {
			const bounds = (await host.locator(selector).first().boundingBox())!;
			expect(bounds.width).toBeGreaterThanOrEqual(44);
			expect(bounds.height).toBeGreaterThanOrEqual(44);
		}
	});
