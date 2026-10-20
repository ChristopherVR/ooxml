import { expect, test } from '@playwright/test';
import { openDemo } from './demo-page';
import { createVsdxFixture } from './fixture.mjs';
import type { VisioViewerElement } from '../../src/ui/src/visio/viewer-element';

for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid']) {
	test(`${framework}: host permission blocks diagram edits and preserves viewing interactions`, async ({
		page,
	}) => {
		await openDemo(
			page,
			framework === 'vanilla' ? '/demo/?sample=1' : `/demo-${framework}/?sample=1`,
		);
		await page.locator('#file').setInputFiles({
			name: 'composition.vsdx',
			mimeType: 'application/vnd.ms-visio.drawing',
			buffer: await createVsdxFixture('Host controlled diagram'),
		});
		await expect(page.locator('visio-viewer svg.paper text')).toContainText(
			'Host controlled diagram',
		);
		await page.evaluate(() => {
			(document.querySelector('visio-viewer') as VisioViewerElement).canEdit = false;
		});
		await page.locator('visio-viewer svg.paper [data-shape-id="1"]').click();
		const result = await page.evaluate(async () => {
			const element = document.querySelector('visio-viewer') as VisioViewerElement;
			const controller = element.controller;
			const selected = controller.state.selectedShapes.length;
			controller.setZoom(2);
			controller.setSearchQuery('controlled');
			let blocked = false;
			try {
				await controller.replacePlainText('1', '1', 'Disallowed');
			} catch {
				blocked = true;
			}
			return {
				blocked,
				selected,
				zoom: controller.state.zoom,
				bytes: controller.exportVsdx().bytes.length,
			};
		});
		expect(result.blocked).toBe(true);
		expect(result.selected).toBe(1);
		expect(result.zoom).toBe(2);
		expect(result.bytes).toBeGreaterThan(0);
		await page.locator('visio-viewer svg.paper [data-shape-id="1"]').dblclick();
		await expect(page.locator('visio-viewer #edit-text')).toHaveCount(0);
		await expect(page.locator('visio-viewer svg.paper text')).toContainText(
			'Host controlled diagram',
		);
	});
}
