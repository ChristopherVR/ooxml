/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec, `test`/`expect` come from @playwright/test */
/**
 * A 3D SmartArt node is found by where the pointer is over it, not by what the
 * browser happens to hit.
 *
 * `<pptx-three-view>` keeps the 2D SVG only as its fallback, so each binding lays
 * an invisible copy of the diagram over the scene. A node group of that copy is
 * hit only where it paints (often its label alone), and the scene underneath is
 * drawn in perspective, so the editor and the hover fill swatches used to appear
 * only sometimes. Every binding now resolves the node by geometry
 * (`smartArtNodeAtPoint` in shared).
 *
 * Also covers the layout switch on a 3D diagram: the scene follows the new layout
 * and keeps the quick style's bevel instead of drawing the old arrangement.
 */
import { expect, test } from '@playwright/test';
import type { Locator, Page } from '@playwright/test';

import { fixture, loadDeckAt, slideStage, thumbnail } from './support/deck';

test.use({ viewport: { width: 1440, height: 900 } });

let webglAvailable = true;

test.beforeAll(async ({ browser }) => {
	const page = await browser.newPage();
	webglAvailable = await page.evaluate(() => {
		const canvas = document.createElement('canvas');
		return Boolean(canvas.getContext('webgl2') ?? canvas.getContext('webgl'));
	});
	await page.close();
});

const DECK = fixture('three-d-parity/three-d-smartart.pptx');

/** Slide 6 of the deck: a diagram with the "Polished" bevel quick style. */
async function openBevelSlide(page: Page): Promise<Locator> {
	await loadDeckAt(page, '/?smartArt3D=1', DECK);
	await thumbnail(page, 6).click();
	const view = slideStage(page).locator('pptx-three-view').first();
	await view.waitFor({ state: 'attached', timeout: 30_000 });
	await expect.poll(async () => view.getAttribute('data-state'), { timeout: 30_000 }).toBe('ready');
	return view;
}

interface NodeBox {
	left: number;
	top: number;
	width: number;
	height: number;
}

/** The laid-out boxes of the nodes of the diagram on the main canvas. */
async function nodeBoxes(page: Page): Promise<NodeBox[]> {
	return slideStage(page).evaluate((stage) => {
		const seen = new Set<string>();
		const boxes: NodeBox[] = [];
		for (const node of stage.querySelectorAll('[data-smartart-node-id]')) {
			const box = node.getBoundingClientRect();
			const key = [box.left, box.top, box.width, box.height].map(Math.round).join(',');
			if (box.width > 40 && box.height > 40 && !seen.has(key)) {
				seen.add(key);
				boxes.push({ left: box.left, top: box.top, width: box.width, height: box.height });
			}
		}
		return boxes;
	});
}

/** Points on a node's fill: near its corners and edges, away from the centred label. */
function fillPoints(box: NodeBox): Array<[number, number]> {
	return [
		[0.12, 0.2],
		[0.88, 0.2],
		[0.12, 0.85],
		[0.88, 0.85],
	].map(([fx, fy]) => [box.left + box.width * (fx ?? 0), box.top + box.height * (fy ?? 0)]);
}

const editor = (page: Page): Locator => page.locator('[data-pptx-viewport] textarea:visible');
/** The hover fill swatches: "Set fill to #..." (React, Vue, Angular) or "Fill Color #..." (Svelte, Vanilla). */
const swatches = (page: Page): Locator =>
	page.locator('[data-pptx-viewport] button[aria-label*="fill" i]');

test('double-clicking the fill of a 3D SmartArt node, away from its text, opens the editor', async ({
	page,
}) => {
	test.setTimeout(180_000);
	test.skip(!webglAvailable, 'headless Chromium has no WebGL context in this environment');
	await openBevelSlide(page);
	const boxes = await nodeBoxes(page);
	expect(boxes.length).toBeGreaterThanOrEqual(4);

	// Select the diagram once, then edit every node from several points on its fill.
	const first = boxes[0]!;
	await page.mouse.click(first.left + first.width / 2, first.top + first.height / 2);
	for (const box of boxes) {
		for (const [x, y] of fillPoints(box)) {
			await page.mouse.move(x, y);
			await page.mouse.dblclick(x, y);
			await expect(editor(page), `editor at ${Math.round(x)},${Math.round(y)}`).toHaveCount(1);
			await editor(page).press('Escape');
			await expect(editor(page)).toHaveCount(0);
		}
	}
});

test('hovering the fill of a 3D SmartArt node shows its fill swatches', async ({ page }) => {
	test.setTimeout(180_000);
	test.skip(!webglAvailable, 'headless Chromium has no WebGL context in this environment');
	await openBevelSlide(page);
	const boxes = await nodeBoxes(page);
	const first = boxes[0]!;
	await page.mouse.click(first.left + first.width / 2, first.top + first.height / 2);

	for (const box of boxes) {
		const [x, y] = fillPoints(box)[2]!;
		await page.mouse.move(x - 8, y - 8);
		await page.mouse.move(x, y);
		await expect(
			swatches(page).first(),
			`swatches at ${Math.round(x)},${Math.round(y)}`,
		).toBeVisible();
	}

	// Off every node (the gap between the columns) they go again.
	const a = boxes[0]!;
	const b = boxes.find((box) => box.left > a.left + a.width) ?? boxes[1]!;
	await page.mouse.move((a.left + a.width + b.left) / 2, a.top + a.height / 2);
	await expect(swatches(page)).toHaveCount(0, { timeout: 5_000 });
});

/** Where the scene draws each node, from the view's own spec. */
async function scenePositions(view: Locator): Promise<string[]> {
	return view.evaluate((el) => {
		const spec = (
			el as unknown as {
				spec?: { spec?: { meshes?: Array<{ position: { x: number; y: number } }> } };
			}
		).spec?.spec;
		return (spec?.meshes ?? []).map(
			(mesh) => `${Math.round(mesh.position.x)},${Math.round(mesh.position.y)}`,
		);
	});
}

test('switching the layout of a 3D SmartArt redraws the scene and keeps the quick style', async ({
	page,
}) => {
	test.setTimeout(180_000);
	test.skip(!webglAvailable, 'headless Chromium has no WebGL context in this environment');
	const view = await openBevelSlide(page);
	const before = await scenePositions(view);
	expect(before.length).toBeGreaterThanOrEqual(4);

	const boxes = await nodeBoxes(page);
	const first = boxes[0]!;
	await page.mouse.click(first.left + first.width / 2, first.top + first.height / 2);
	await page
		.getByRole('button', { name: /^cycle$/iu })
		.first()
		.click();

	await expect.poll(async () => scenePositions(view), { timeout: 15_000 }).not.toEqual(before);
	const style = await view.evaluate(
		(el) =>
			(el as unknown as { spec?: { spec?: { styleCategory?: string } } }).spec?.spec?.styleCategory,
	);
	expect(style).toBe('bevel');
});
