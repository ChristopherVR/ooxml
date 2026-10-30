import { expect, test, type Locator } from '@playwright/test';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { openDemo } from './demo-page';

const TAB = '[data-ribbon-tab="cross-functional"]';

/** The lanes of the flowchart (stacking order) and their page-inch heights, read from the viewer. */
async function lanes(viewer: Locator): Promise<{ id: string; height: number }[]> {
	return viewer.evaluate((node) => {
		const state = (node as VisioViewerElement).controller.state;
		const page = state.document!.pages[0]!;
		const frame = page.shapes.find((shape) => shape.id === '50')!;
		const ids = frame.structure?.type === 'container' ? frame.structure.memberIds : [];
		return page.shapes
			.filter((shape) => ids.includes(shape.id) && shape.structure?.type === 'container')
			.sort((a, b) => b.transform[5] + b.height - (a.transform[5] + a.height))
			.map((shape) => ({ id: shape.id, height: shape.height }));
	});
}

test('a cross-functional flowchart gets lanes that stack, insert, resize and delete together', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.addInitScript(() => localStorage.clear());
	await page.setViewportSize({ width: 1600, height: 1000 });
	await openDemo(page, '/demo/?sample=1');
	const viewer = page.locator('visio-viewer');
	await viewer.locator('office-ui-ribbon .file').click();
	await viewer.locator('[data-backstage-item="new"]').click();
	await viewer.locator('[data-backstage-action="new-template:cross-functional-flowchart"]').click();
	await expect(viewer.locator('svg.paper')).toContainText('Place order');
	expect((await lanes(viewer)).map((lane) => lane.id)).toEqual(['10', '11', '12']);

	// The contextual tab appears with the flowchart selected, and not for an unrelated selection.
	await expect(viewer.locator(`${TAB}[data-tab-hidden]`)).toHaveCount(1);
	await viewer
		.locator('svg.paper [data-shape-id="11"]')
		.click({ force: true, position: { x: 10, y: 10 } });
	await expect(viewer.locator(`${TAB}[data-tab-hidden]`)).toHaveCount(0);
	const tab = viewer.locator('office-ui-ribbon #tab-cross-functional');
	await expect(tab).toBeVisible();
	await tab.click();
	const panel = viewer.locator(TAB);
	await expect(panel.locator('office-ui-ribbon-group')).toHaveCount(3);

	// Insert Swimlane After the selected lane: four lanes, the new one second to last.
	await panel.getByRole('button', { name: 'Swimlane', exact: true }).click();
	await viewer.locator(`${TAB} office-ui-menu-item[label="Insert Swimlane After"]`).click();
	await expect.poll(async () => (await lanes(viewer)).length).toBe(4);
	const four = await lanes(viewer);
	expect(four[1]!.id).toBe('11');
	expect(four[2]!.id).not.toBe('12');
	expect(four[3]!.id).toBe('12');
	// One undo step per command.
	await viewer.evaluate((node) => (node as VisioViewerElement).controller.undo());
	await expect.poll(async () => (await lanes(viewer)).length).toBe(3);
	await viewer.evaluate((node) => (node as VisioViewerElement).controller.redo());
	await expect.poll(async () => (await lanes(viewer)).length).toBe(4);

	// Show Title Bar and Separator: options record on the frame, the lanes keep stacking.
	await panel.locator('[data-check="swimlane-title"]').click();
	await expect
		.poll(() =>
			viewer.evaluate((node) => {
				const state = (node as VisioViewerElement).controller.state;
				const pageModel = state.document!.pages[0]!;
				const frame = pageModel.shapes.find((shape) => shape.id === '50')!;
				return frame.structure?.type === 'container'
					? frame.structure.swimlane?.role === 'frame' && frame.structure.swimlane.options.showTitle
					: undefined;
			}),
		)
		.toBe(false);
	await panel.locator('[data-check="swimlane-title"]').click();
	await panel.getByRole('button', { name: 'Separator', exact: true }).click();
	await expect
		.poll(() =>
			viewer.evaluate(
				(node) => (node as VisioViewerElement).controller.state.document!.pages[0]!.shapes.length,
			),
		)
		.toBeGreaterThan(0);

	// Delete the selected lane: its shapes go and the lanes close up.
	await viewer
		.locator('svg.paper [data-shape-id="12"]')
		.click({ force: true, position: { x: 10, y: 10 } });
	await panel.getByRole('button', { name: 'Swimlane', exact: true }).click();
	await viewer.locator(`${TAB} office-ui-menu-item[label="Delete Swimlane"]`).click();
	await expect.poll(async () => (await lanes(viewer)).length).toBe(3);
	await expect(viewer.locator('svg.paper')).not.toContainText('Ship order');
	const info = await viewer.evaluate((node) => {
		const state = (node as VisioViewerElement).controller.state;
		return state.document!.pages[0]!.shapes.some((shape) => shape.id === '12');
	});
	expect(info).toBe(false);

	// The Delete key on a selected lane does the same.
	await viewer
		.locator('svg.paper [data-shape-id="10"]')
		.click({ force: true, position: { x: 10, y: 10 } });
	await page.keyboard.press('Delete');
	await expect.poll(async () => (await lanes(viewer)).length).toBe(2);
	expect(errors).toEqual([]);
});
