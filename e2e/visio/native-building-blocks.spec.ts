import { expect, test } from '@playwright/test';
import { createVsdxFixture } from './fixture.mjs';
import type { ViewerController } from 'ooxml-ui/visio';
declare global {
	interface Window {
		nativeBlocks: { controller: ViewerController };
	}
}
for (const framework of ['react', 'vue', 'angular', 'svelte', 'solid'])
	test(`${framework}: native composed canvas edits, reloads and follows permissions`, async ({
		page,
	}) => {
		await page.goto(`/native-blocks.html?framework=${framework}`);
		await expect(page.locator('output')).toHaveText('edit:1');
		await page.locator('#content').setInputFiles({
			name: 'native.vsdx',
			mimeType: 'application/octet-stream',
			buffer: await createVsdxFixture('Native composition'),
		});
		await expect(page.locator('svg.paper text')).toContainText('Native composition');
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
