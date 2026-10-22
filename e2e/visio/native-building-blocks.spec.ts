import { expect, test, type Page } from '@playwright/test';
import { createVsdxFixture } from './fixture.mjs';
import type { ViewerController } from 'ooxml-ui/visio';
declare global {
	interface Window {
		nativeBlocks: { controller: ViewerController };
	}
}
const frameworks = ['react', 'vue', 'angular', 'svelte', 'solid', 'vanilla'];
for (const framework of frameworks)
	test(`${framework}: native composed canvas edits, reloads and follows permissions`, async ({
		page,
	}) => {
		await page.goto(`/native-blocks.html?framework=${framework}`);
		// Granted permission is not enough until a drawing that can be saved back is open.
		await expect(page.locator('output')).toHaveText('view:1');
		await page.locator('#content').setInputFiles({
			name: 'native.vsdx',
			mimeType: 'application/octet-stream',
			buffer: await createVsdxFixture('Native composition'),
		});
		await expect(page.locator('svg.paper text')).toContainText('Native composition');
		await expect(page.locator('output')).toHaveText('edit:1');
		await expect(page.locator('office-ui-ribbon, office-ui-backstage, visio-viewer')).toHaveCount(
			0,
		);
		await page.locator('svg.paper [data-shape-id="1"]').click();
		const changed = await page.evaluate(async () => {
			const controller = window.nativeBlocks.controller;
			await controller.replacePlainText('1', '1', 'Native edited');
			return Array.from(controller.exportVsdx().bytes);
		});
		await expect(page.locator('svg.paper text')).toContainText('Native edited');
		await page.getByRole('button', { name: 'Undo', exact: true }).click();
		await expect(page.locator('svg.paper text')).toContainText('Native composition');
		await page.getByRole('button', { name: 'Redo', exact: true }).click();
		await expect(page.locator('svg.paper text')).toContainText('Native edited');
		const undo = page.getByRole('button', { name: 'Undo', exact: true });
		await undo.focus();
		await page.evaluate(() => window.nativeBlocks.controller.setZoom(1.25));
		await expect(undo).toBeFocused();
		await page.locator('#readonly').click();
		await page.locator('#host-zoom').click();
		await expect(page.locator('output')).toHaveText('view:2');
		await expect(page.getByRole('button', { name: 'Undo', exact: true })).toBeDisabled();
		expect(
			await page.evaluate(async () => {
				try {
					await window.nativeBlocks.controller.replacePlainText('1', '1', 'blocked');
					return false;
				} catch {
					return true;
				}
			}),
		).toBe(true);
		await page.locator('#content').setInputFiles({
			name: 'roundtrip.vsdx',
			mimeType: 'application/octet-stream',
			buffer: Buffer.from(changed),
		});
		await expect(page.locator('svg.paper text')).toContainText('Native edited');
		await page.locator('#content').setInputFiles({
			name: 'bad.vsdx',
			mimeType: 'application/octet-stream',
			buffer: Buffer.from('bad'),
		});
		await expect
			.poll(() => page.evaluate(() => window.nativeBlocks.controller.state.error?.message))
			.toBeTruthy();
		await page.locator('#content').setInputFiles({
			name: 'recover.vsdx',
			mimeType: 'application/octet-stream',
			buffer: await createVsdxFixture('Recovered'),
		});
		await expect(page.locator('svg.paper text')).toContainText('Recovered');
		await page.locator('#unmount').click();
		await expect(page.locator('svg.paper')).toHaveCount(0);
		expect(
			await page.evaluate(() => {
				try {
					window.nativeBlocks.controller.setZoom(4);
					return false;
				} catch {
					return true;
				}
			}),
		).toBe(true);
	});

for (const framework of frameworks)
	test(`${framework}: library toolbar, drawing tools and save acknowledgement follow permission`, async ({
		page,
	}) => {
		await page.goto(`/native-blocks.html?framework=${framework}`);
		await page.locator('#content').setInputFiles({
			name: 'tools.vsdx',
			mimeType: 'application/octet-stream',
			buffer: await createVsdxFixture('Toolbar'),
		});
		await expect(page.locator('output')).toHaveText('edit:1');
		const control = (id: string) => page.locator(`[role="toolbar"] [data-control="${id}"]`);
		for (const id of [
			'zoom-in',
			'zoom-out',
			'fit',
			'duplicate',
			'delete',
			'copy',
			'tool-rectangle',
		])
			await expect(control(id)).toBeVisible();
		// No host dialog callback was given, so no dialog buttons are offered.
		await expect(control('settings')).toHaveCount(0);
		await control('zoom-in').click();
		await expect(page.locator('output')).toHaveText('edit:1.25');
		// A shape-creation tool draws on the standalone canvas, like the full viewer's.
		const count = () => page.locator('svg.paper [data-shape-id]').count();
		const before = await count();
		await control('tool-rectangle').click();
		await expect(control('tool-rectangle')).toHaveAttribute('aria-pressed', 'true');
		const paper = await visiblePaper(page);
		await page.mouse.move(paper.x + 20, paper.y + 20);
		await page.mouse.down();
		await page.mouse.move(paper.x + 80, paper.y + 60, { steps: 4 });
		await page.mouse.up();
		await expect.poll(count).toBeGreaterThan(before);
		await expect(
			page.evaluate(() => window.nativeBlocks.controller.state.edit.dirty),
		).resolves.toBe(true);
		// The host persisted the bytes and acknowledges the save.
		await page.evaluate(() => window.nativeBlocks.controller.markClean());
		await expect(
			page.evaluate(() => window.nativeBlocks.controller.state.edit.dirty),
		).resolves.toBe(false);
		await page.locator('#readonly').click();
		await expect(page.locator('output')).toHaveText('view:1.25');
		for (const id of ['delete', 'duplicate', 'cut', 'paste', 'tool-rectangle'])
			await expect(control(id)).toBeDisabled();
		// Viewing controls stay available.
		await expect(control('zoom-out')).toBeEnabled();
		await expect(control('fit')).toBeEnabled();
	});

/** The paper's box with the page and canvas scrolled to the top, so its corner is on screen. */
async function visiblePaper(page: Page) {
	await page.evaluate(() => {
		window.scrollTo(0, 0);
		document.querySelectorAll('*').forEach((element) => {
			element.shadowRoot?.querySelector('[part="canvas"]')?.scrollTo(0, 0);
		});
	});
	return (await page.locator('svg.paper').boundingBox())!;
}
const shapeCount = (page: Page) => page.locator('svg.paper [data-shape-id]').count();
const shapeAngle = (page: Page) =>
	page.evaluate(() => {
		const state = window.nativeBlocks.controller.state;
		return state.document?.pages[state.pageIndex]?.shapes.find((shape) => shape.id === '1')
			?.rotation?.angle;
	});
async function openEditable(page: Page, framework: string, name: string) {
	await page.goto(`/native-blocks.html?framework=${framework}`);
	await page.locator('#content').setInputFiles({
		name: `${name}.vsdx`,
		mimeType: 'application/octet-stream',
		buffer: await createVsdxFixture(name),
	});
	await expect(page.locator('output')).toHaveText('edit:1');
}
/** A drag on empty paper, in the lower right where the fixture's shape is not. */
async function dragOnPaper(page: Page) {
	const paper = await visiblePaper(page);
	const x = paper.x + paper.width * 0.6,
		y = paper.y + paper.height * 0.3;
	await page.mouse.move(x, y);
	await page.mouse.down();
	await page.mouse.move(x + 90, y + 50, { steps: 5 });
	await page.mouse.up();
}

for (const framework of frameworks) {
	test(`${framework}: the rotation handle turns the selected shape and is absent when read-only`, async ({
		page,
	}) => {
		await openEditable(page, framework, 'Rotation');
		const handle = page.locator('[data-rotation-handle]');
		await expect(handle).toHaveCount(0);
		await page.locator('svg.paper [data-shape-id="1"]').click();
		await expect(handle).toHaveCount(1);
		const before = await shapeAngle(page);
		const box = (await handle.boundingBox())!;
		const shape = (await page.locator('svg.paper [data-shape-id="1"]').boundingBox())!;
		const centre = { x: shape.x + shape.width / 2, y: shape.y + shape.height / 2 };
		await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
		await page.mouse.down();
		// A quarter turn clockwise: from above the shape to its right.
		await page.mouse.move(centre.x + 60, centre.y, { steps: 6 });
		await page.mouse.up();
		await expect.poll(() => shapeAngle(page)).not.toBe(before);
		await expect(
			page.evaluate(() => window.nativeBlocks.controller.state.edit.dirty),
		).resolves.toBe(true);
		await page.locator('#readonly').click();
		await expect(page.locator('output')).toHaveText(/^view:/);
		// The selection survives, but the handle is gone and nothing can start a rotation.
		await expect(page.locator('[data-rotation-handle]')).toHaveCount(0);
		const angle = await shapeAngle(page);
		await page.locator('svg.paper [data-shape-id="1"]').click();
		await expect(page.locator('[data-rotation-handle]')).toHaveCount(0);
		expect(await shapeAngle(page)).toBe(angle);
	});

	test(`${framework}: the text tool adds a text box and the connector tool a connector, only when editable`, async ({
		page,
	}) => {
		await openEditable(page, framework, 'Object tools');
		const control = (id: string) => page.locator(`[role="toolbar"] [data-control="${id}"]`);
		await expect(control('tool-text')).toBeEnabled();
		await expect(control('tool-connector')).toBeEnabled();
		const before = await shapeCount(page);
		await control('tool-text').click();
		await expect(control('tool-text')).toHaveAttribute('aria-pressed', 'true');
		await dragOnPaper(page);
		const input = page.locator('[data-text-draft] textarea');
		await expect(input).toBeVisible();
		await input.fill('Typed on canvas');
		await input.press('Control+Enter');
		await expect(page.locator('svg.paper text', { hasText: 'Typed on canvas' })).toHaveCount(1);
		expect(await shapeCount(page)).toBe(before + 1);
		await control('tool-connector').click();
		await expect(control('tool-connector')).toHaveAttribute('aria-pressed', 'true');
		// Drag from the first shape into empty paper: the connector is glued at its start.
		const shape = (await page.locator('svg.paper [data-shape-id="1"]').boundingBox())!;
		const x = shape.x + shape.width / 2,
			y = shape.y + shape.height / 2;
		await page.mouse.move(x, y);
		await page.mouse.down();
		await page.mouse.move(x + 150, y + 120, { steps: 6 });
		await page.mouse.up();
		await expect.poll(() => shapeCount(page)).toBe(before + 2);
		await page.locator('#readonly').click();
		await expect(control('tool-text')).toBeDisabled();
		await expect(control('tool-connector')).toBeDisabled();
		await expect(page.evaluate(() => window.nativeBlocks.controller.state.tool)).resolves.toBe(
			'pointer',
		);
		const kept = await shapeCount(page);
		await page.evaluate(() => {
			window.nativeBlocks.controller.setTool('text');
			window.nativeBlocks.controller.setTool('connector');
		});
		await dragOnPaper(page);
		await expect(page.locator('[data-text-draft]')).toHaveCount(0);
		expect(await shapeCount(page)).toBe(kept);
	});
}
