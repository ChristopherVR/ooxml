import { readFile } from 'node:fs/promises';
import { expect, test } from '@playwright/test';
import { parseVsdx } from 'ooxml-core/visio';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { openDemo } from './demo-page';

test('renders embedded Word math and saves the Visio foreign shape after moving it', async ({
	page,
}) => {
	const errors: string[] = [];
	page.on('pageerror', (error) => errors.push(error.message));
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await page.locator('#file').setInputFiles({
		name: 'Equation.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: await readFile(
			new URL('../../src/core/visio/__fixtures__/word-equation-native.vsdx', import.meta.url),
		),
	});
	await expect(page.locator('#file-name')).toHaveText('Equation.vsdx');
	const shape = viewer.locator('svg.paper [data-shape-id="1"]');
	await expect(shape.locator('image')).toBeVisible();
	const before = await shape.getAttribute('transform');
	await viewer.evaluate(async (node: VisioViewerElement) => {
		const doc = node.controller.state.document;
		const target = doc?.pages[0];
		if (!target) throw new Error('Drawing did not open');
		await node.applyEdits([{ type: 'move-shape', pageId: target.id, shapeId: '1', x: 4, y: 5 }]);
	});
	await expect(shape).not.toHaveAttribute('transform', before!);
	const saved = await viewer.evaluate((node: VisioViewerElement) =>
		Array.from(node.exportVsdx().bytes),
	);
	const reopened = await parseVsdx(Uint8Array.from(saved));
	expect(reopened.pages[0]?.shapes[0]?.image?.mimeType).toBe('image/svg+xml');
	expect(errors).toEqual([]);
});
