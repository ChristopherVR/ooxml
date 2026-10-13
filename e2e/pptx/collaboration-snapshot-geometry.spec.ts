/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright spec */
import { readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';

import { expect, test } from '@playwright/test';

import { savedShapeGeometry } from './support/collaboration-snapshot-geometry';
import { fixture } from './support/deck';

const coreRequire = createRequire(createRequire(import.meta.url).resolve('pptx-viewer-core'));
const { PptxHandler } = coreRequire('pptx-viewer-core') as typeof import('pptx-viewer-core');

test('snapshot geometry checks full-precision model EMUs at the CSS half-pixel boundary', async ({
	page,
}, info) => {
	const handler = new PptxHandler();
	const data = await handler.load(new Uint8Array(readFileSync(fixture('sample-deck.pptx'))));
	const elementId = 'ppt/slides/slide1.xml-shape-6';
	const model = data.slides[0]!.elements.find((element) => element.id === elementId)!;
	model.x = 1031.49999999;
	model.y = 174.49999999;
	const cssX = await page.evaluate((x) => {
		const element = document.createElement('div');
		element.style.position = 'absolute';
		element.style.left = `${x}px`;
		document.body.appendChild(element);
		const value = Number.parseFloat(getComputedStyle(element).left);
		element.remove();
		return value;
	}, model.x);
	expect(cssX).toBe(1031.5);
	const path = info.outputPath('half-pixel-snapshot.pptx');
	writeFileSync(path, await handler.save(data.slides));
	const parsed = await savedShapeGeometry(path, elementId, model);
	expect(parsed.slice(0, 2)).toEqual([1031, 174]);
	expect(parsed[0]).not.toBe(Math.round(cssX));
	// A real one-pixel save error must still fail the exact EMU assertion.
	await expect(savedShapeGeometry(path, elementId, { ...model, x: model.x + 1 })).rejects.toThrow();
});
