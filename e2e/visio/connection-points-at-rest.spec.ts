import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { openDemo } from './demo-page';

const SAMPLE = new URL('../../samples/visio/order-to-delivery-flowchart.vsdx', import.meta.url);

// Visio shows connection points only while connecting, never as a blue x on every shape at rest.
test('the flowchart sample shows connection points only with the Connector tool', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await page.setViewportSize({ width: 1500, height: 1000 });
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await page.locator('#file').setInputFiles({
		name: 'order-to-delivery-flowchart.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await readFile(SAMPLE),
	});
	await expect(page.locator('#file-name')).toHaveText('order-to-delivery-flowchart.vsdx');
	await expect(viewer.locator('svg.paper [data-shape-id]').first()).toBeVisible();
	await expect(viewer.locator('[data-connection-point]')).toHaveCount(0);
	await viewer.locator('svg.paper').click({ position: { x: 5, y: 5 } });
	await page.keyboard.press('Control+3');
	await expect.poll(() => viewer.locator('[data-connection-point]').count()).toBeGreaterThan(0);
	await page.keyboard.press('Control+1');
	await expect(viewer.locator('[data-connection-point]')).toHaveCount(0);
	expect(errors).toEqual([]);
});
