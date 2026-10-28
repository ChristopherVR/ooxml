/* oxlint-disable vitest/prefer-importing-vitest-globals -- Playwright owns test and expect */
import { readFileSync } from 'node:fs';

import { expect, test } from '@playwright/test';
import JSZip from 'jszip';

import { withNativePictureLayout } from '../../src/core/diagram/test-support/picture-layout';
import { savePptxViaBackstage } from './save-pptx';
import { resetTabSession, slideStage, loadDeck } from './support/deck';
import { unobstructedNodePoint } from './support/smartart-hit-target';

for (const layout of ['picture-accent-list--flat3', 'picture-grid--flat3']) {
	test(`${layout} keeps image fills after caption edit, reflow and save/reopen`, async ({
		page,
	}) => {
		const bytes = await withNativePictureLayout(
			readFileSync(
				new URL(
					`../../src/core/pptx/__tests__/fixtures/smartart-gallery/${layout}.pptx`,
					import.meta.url,
				),
			),
			'ppt',
			layout,
		);
		await resetTabSession(page);
		await page.goto('/');
		await page.locator('#file-input').setInputFiles({
			name: `${layout}.pptx`,
			mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
			buffer: Buffer.from(bytes),
		});
		const stage = slideStage(page);
		await expect(stage.locator('image')).toHaveCount(3);
		const caption = stage
			.locator('[data-smartart-node-id]')
			.filter({ has: page.locator('text').filter({ hasText: /^Alpha$/ }) })
			.first();
		const point = await unobstructedNodePoint(caption);
		await page.mouse.dblclick(point.x, point.y);
		const editor = page.locator('[data-pptx-viewport] textarea:visible');
		await expect(editor).toHaveCount(1);
		await editor.fill('Updated picture caption');
		await page
			.getByRole('toolbar', { name: 'Presentation toolbar' })
			.getByRole('tab', { name: 'Home', exact: true })
			.click();
		await expect(editor).toHaveCount(0);
		await expect(stage.locator('image')).toHaveCount(3);
		await expect(stage).toContainText('Updated picture caption');
		const download = await savePptxViaBackstage(page);
		const path = await download.path();
		if (!path) throw new Error('Saved deck download missing');
		await loadDeck(page, path);
		await expect(slideStage(page).locator('image')).toHaveCount(3);
		await expect(slideStage(page)).toContainText('Updated picture caption');
	});
}

test('nonlinear SmartArt picture paint refreshes the mounted image after ordered pixel processing', async ({
	page,
}) => {
	const bytes = await withNativePictureLayout(
		readFileSync(
			new URL(
				'../../src/core/pptx/__tests__/fixtures/smartart-gallery/picture-grid--flat3.pptx',
				import.meta.url,
			),
		),
		'ppt',
		'picture-grid--flat3',
	);
	const zip = await JSZip.loadAsync(bytes);
	for (const name of ['data', 'drawing']) {
		const path = `ppt/diagrams/${name}1.xml`;
		zip.file(
			path,
			(await zip.file(path)!.async('string')).replaceAll(
				'<a:grayscl/>',
				'<a:hsl hue="0" sat="100000" lum="0"/>',
			),
		);
	}
	await resetTabSession(page);
	await page.goto('/');
	await page.locator('#file-input').setInputFiles({
		name: 'nonlinear-pictures.pptx',
		mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		buffer: await zip.generateAsync({ type: 'nodebuffer' }),
	});
	const images = slideStage(page).locator('image[data-office-picture-pixels-ready="true"]');
	await expect(images).toHaveCount(3);
	const pixel = await images.first().evaluate(async (element) => {
		const svg = new DOMParser().parseFromString(
			decodeURIComponent(element.getAttribute('href')!.split(',').slice(1).join(',')),
			'image/svg+xml',
		);
		const url = svg.querySelector('pattern image')!.getAttribute('href')!;
		const image = new Image();
		image.src = url;
		await image.decode();
		const canvas = document.createElement('canvas');
		canvas.width = image.naturalWidth;
		canvas.height = image.naturalHeight;
		const context = canvas.getContext('2d')!;
		context.drawImage(image, 0, 0);
		return [...context.getImageData(0, 0, 1, 1).data];
	});
	expect(pixel[0]).toBeGreaterThan(120);
	expect(pixel[0]).toBeLessThan(135);
	expect(Math.abs(pixel[0]! - pixel[1]!)).toBeLessThanOrEqual(1);
	expect(Math.abs(pixel[1]! - pixel[2]!)).toBeLessThanOrEqual(1);
	expect(pixel[3]).toBeGreaterThan(185);
	expect(pixel[3]).toBeLessThan(195);
});
