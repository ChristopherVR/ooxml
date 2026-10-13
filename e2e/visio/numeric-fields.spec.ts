import { test, expect } from '@playwright/test';
import { readFileSync } from 'node:fs';
import type { VisioViewerElement } from 'ooxml-ui/visio';
import { openDemo } from './demo-page';

const native = new URL(
	'../../src/core/visio/__fixtures__/numeric-field-native.vsdx',
	import.meta.url,
);
const evidence = JSON.parse(
	readFileSync(
		new URL('../../src/core/visio/__fixtures__/numeric-field-native.json', import.meta.url),
		'utf8',
	).replace(/^\uFEFF/, ''),
) as { cases: { id: number; text: string }[] };

test('inherited decimal fields render native values and survive browser save/reopen', async ({
	page,
}) => {
	await openDemo(page);
	const viewer = page.locator('visio-viewer');
	await page.locator('#file').setInputFiles({
		name: 'Numbers.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: readFileSync(native),
	});
	await expect(page.locator('#file-name')).toHaveText('Numbers.vsdx');
	const check = async () => {
		for (const item of evidence.cases)
			await expect(viewer.locator(`svg.paper [data-shape-id="${item.id}"] text`)).toHaveText(
				item.text,
			);
	};
	await check();
	const bytes = await viewer.evaluate((node) =>
		Array.from((node as VisioViewerElement).exportVsdx().bytes),
	);
	await page.locator('#file').setInputFiles({
		name: 'Saved numbers.vsdx',
		mimeType: 'application/vnd.ms-visio.drawing',
		buffer: Buffer.from(bytes),
	});
	await expect(page.locator('#file-name')).toHaveText('Saved numbers.vsdx');
	await check();
});
