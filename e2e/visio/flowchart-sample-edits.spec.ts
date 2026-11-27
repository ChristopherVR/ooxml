import { readFile } from 'node:fs/promises';
import { expect, test, type Locator, type Page } from '@playwright/test';
import { openDemo } from './demo-page';

const SAMPLE = new URL('../../samples/visio/order-to-delivery-flowchart.vsdx', import.meta.url);

async function openSample(page: Page): Promise<Locator> {
	await page.setViewportSize({ width: 1500, height: 1000 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await page.locator('#file').setInputFiles({
		name: 'order-to-delivery-flowchart.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await readFile(SAMPLE),
	});
	await expect(page.locator('#file-name')).toHaveText('order-to-delivery-flowchart.vsdx');
	await expect(viewer.locator('svg.paper [data-page-id="0"][data-shape-id="21"]')).toBeVisible();
	return viewer;
}

type Edit = { busy: boolean; dirty: boolean; error: string | null; selected: string[] };
const edit = (viewer: Locator) =>
	viewer.evaluate((element): Edit => {
		const { state } = (
			element as unknown as {
				controller: {
					state: {
						edit: { busy: boolean; dirty: boolean; error: Error | null };
						selectedShapes: readonly { id: string }[];
					};
				};
			}
		).controller;
		return {
			busy: state.edit.busy,
			dirty: state.edit.dirty,
			error: state.edit.error?.message ?? null,
			selected: state.selectedShapes.map((shape) => shape.id),
		};
	});
const centre = async (shape: Locator) => {
	const box = (await shape.boundingBox())!;
	return { x: box.x + box.width / 2, y: box.y + box.height / 2 };
};

test('clicking one flowchart step after another selects without editing', async ({ page }) => {
	const viewer = await openSample(page);
	await viewer.evaluate((element) => {
		const changes: unknown[] = [];
		(window as unknown as { changes: unknown[] }).changes = changes;
		element.addEventListener('document-change', (event) =>
			changes.push((event as CustomEvent).detail?.kind),
		);
	});
	for (const id of ['17', '21']) {
		const { x, y } = await centre(
			viewer.locator(`svg.paper [data-page-id="0"][data-shape-id="${id}"]`).first(),
		);
		await page.mouse.click(x, y);
		await expect.poll(async () => (await edit(viewer)).selected).toEqual([id]);
	}
	await page.waitForTimeout(300);
	expect(await edit(viewer)).toEqual({ busy: false, dirty: false, error: null, selected: ['21'] });
	expect(await page.evaluate(() => (window as unknown as { changes: unknown[] }).changes)).toEqual(
		[],
	);
	await expect(viewer).not.toContainText('Edit rejected');
});

// Every master instance on the page is proven independent of a moved shape. Each master was once
// rebuilt for every instance, so moving any shape of this sample exceeded the construction budget.
test('a shape of the flowchart sample moves within the master proof budget', async ({ page }) => {
	const viewer = await openSample(page);
	// The legend's "Reading the chart" heading, a local shape among the legend's master swatches.
	const heading = viewer.locator('svg.paper [data-page-id="0"][data-shape-id="65"]').first();
	const { x, y } = await centre(heading);
	await page.mouse.click(x, y);
	await expect.poll(async () => (await edit(viewer)).selected).toEqual(['65']);
	await page.mouse.move(x, y);
	await page.mouse.down();
	await page.mouse.move(x - 12, y + 6, { steps: 4 });
	await page.mouse.up();
	await expect.poll(async () => (await edit(viewer)).dirty, { timeout: 15_000 }).toBe(true);
	expect(await edit(viewer)).toEqual({ busy: false, dirty: true, error: null, selected: ['65'] });
});
