import { writeFile } from 'node:fs/promises';

/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright provides test and expect */
import { test, expect } from '@playwright/test';
import { PptxHandler } from 'pptx-viewer-core';

import { loadDeck } from './support/deck';

test('unrelated incoming artwork keeps its orientation during a morph (#105)', async ({
	page,
}, info) => {
	const { handler, data, createSlide } = await PptxHandler.createBlank({ initialSlideCount: 0 });
	const svg = (color: string) =>
		`data:image/svg+xml;base64,${Buffer.from(`<svg xmlns="http://www.w3.org/2000/svg" width="200" height="200"><path fill="${color}" d="M10 10H190V190H10Z"/></svg>`).toString('base64')}`;
	const first = createSlide('Blank')
		.addImage(svg('blue'), {
			x: 302,
			y: 22,
			width: 675,
			height: 675,
			altText: 'Outgoing wheel',
		})
		.build();
	first.elements[0]!.rotation = 45;
	const second = createSlide('Blank')
		.addText('New scene', { x: 10, y: 10, width: 200, height: 30 })
		.addImage(svg('green'), { x: 349, y: 67, width: 706, height: 407, altText: 'Incoming radar' })
		.addImage(svg('red'), { x: 750, y: 100, width: 80, height: 40, altText: 'Incoming aircraft' })
		.setTransition({ type: 'morph', duration: 10000 })
		.build();
	data.slides.push(first, second);
	const path = info.outputPath('unrelated-artwork.pptx');
	await writeFile(path, await handler.save(data.slides));
	await loadDeck(page, path);
	await page
		.getByRole('button', { name: /^present$|slide show/iu })
		.first()
		.click();
	await expect(page.locator('pptx-ui-present-toolbar:visible')).toHaveCount(1);
	await page.keyboard.press('ArrowRight');
	await page.locator('[data-pptx-morph-outgoing]').first().waitFor({ state: 'attached' });
	for (const label of ['Incoming radar', 'Incoming aircraft']) {
		const image = page.getByRole('img', { name: label }).last();
		await image.waitFor();
		const angle = await image.evaluate((node) => {
			const element = node.closest('[data-element-id]')!;
			const transform = getComputedStyle(element).transform;
			const matrix = new DOMMatrix(transform === 'none' ? undefined : transform);
			return (Math.atan2(matrix.b, matrix.a) * 180) / Math.PI;
		});
		expect(Math.abs(angle)).toBeLessThan(0.01);
	}
});
