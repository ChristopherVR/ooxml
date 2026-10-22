import { expect, test } from '@playwright/test';
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
		const paper = (await page.locator('svg.paper').boundingBox())!;
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
