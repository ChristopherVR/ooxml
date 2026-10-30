import { expect, test, type Locator } from '@playwright/test';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { openDemo } from './demo-page';

const TAB = '[data-ribbon-tab="org-chart"]';

/** Ids of the shapes that carry the org chart Name row, and the glued connector count. */
async function chart(viewer: Locator): Promise<{ positions: string[]; connectors: number }> {
	return viewer.evaluate((node) => {
		const page = (node as VisioViewerElement).controller.state.document!.pages[0]!;
		return {
			positions: page.shapes
				.filter((shape) => shape.shapeData?.some((item) => item.name === 'Name'))
				.map((shape) => shape.id),
			connectors: page.connectors.length / 2,
		};
	});
}

test('the New tiles show thumbnails and the org chart gets its own tab', async ({ page }) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.addInitScript(() => localStorage.clear());
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page, '/demo/?sample=1');
	const viewer = page.locator('visio-viewer');
	await viewer.locator('office-ui-ribbon .file').click();
	await viewer.locator('[data-backstage-item="new"]').click();

	// Every template tile shows an image rendered from its template.
	const tiles = viewer.locator('.backstage-templates .backstage-action');
	await expect(tiles).toHaveCount(4);
	await expect(viewer.locator('.backstage-templates img.backstage-thumb')).toHaveCount(4);
	for (const image of await viewer.locator('.backstage-templates img.backstage-thumb').all())
		await expect
			.poll(() => image.evaluate((node) => (node as HTMLImageElement).naturalWidth))
			.toBeGreaterThan(0);

	await viewer.locator('[data-backstage-action="new-template:organization-chart"]').click();
	await expect(viewer.locator('svg.paper')).toContainText('Chief Executive');
	await expect(viewer.locator(`${TAB}[data-tab-hidden]`)).toHaveCount(1);

	// Selecting a position shows the Org Chart tab.
	await viewer
		.locator('svg.paper [data-shape-id="3"]')
		.click({ force: true, position: { x: 10, y: 10 } });
	await expect(viewer.locator(`${TAB}[data-tab-hidden]`)).toHaveCount(0);
	await viewer.locator('office-ui-ribbon #tab-org-chart').click();
	const panel = viewer.locator(TAB);
	expect((await chart(viewer)).positions).toHaveLength(6);

	await panel.locator('[command="org-subordinate"]').click();
	await expect.poll(async () => (await chart(viewer)).positions.length).toBe(7);
	expect((await chart(viewer)).connectors).toBe(6);
	// One undo step per command.
	await viewer.evaluate((node) => (node as VisioViewerElement).controller.undo());
	await expect.poll(async () => (await chart(viewer)).positions.length).toBe(6);
	await viewer.evaluate((node) => (node as VisioViewerElement).controller.redo());
	await expect.poll(async () => (await chart(viewer)).positions.length).toBe(7);

	// Delete the new position: it and its connector go.
	const added = (await chart(viewer)).positions.at(-1)!;
	await viewer
		.locator(`svg.paper [data-shape-id="${added}"]`)
		.click({ force: true, position: { x: 10, y: 10 } });
	await panel.locator('[command="org-delete"]').click();
	await expect.poll(async () => (await chart(viewer)).positions.length).toBe(6);
	expect((await chart(viewer)).connectors).toBe(5);
	expect(errors).toEqual([]);
});
